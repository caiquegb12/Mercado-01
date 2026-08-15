import io
import re
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import quote
from fastapi import FastAPI, Form, Request, UploadFile, File
from fastapi.responses import HTMLResponse, JSONResponse, PlainTextResponse, RedirectResponse, StreamingResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pathlib import Path

from openpyxl import Workbook

from app.database import SessionLocal, Empresa, Contato, Contrato, NF, Anexo

app = FastAPI(title="Mercado Central - Gestão de Contratos")

LOGIN_USER = "mercado"
LOGIN_PASSWORD = "mercado123"
REMEMBER_ME_DAYS = 30

BASE_DIR = Path(__file__).resolve().parent.parent
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


def is_authenticated(request: Request) -> bool:
    return request.cookies.get("mc_session") == "authenticated"


def set_auth_cookie(response: RedirectResponse, remember_me: bool = False):
    max_age = REMEMBER_ME_DAYS * 24 * 60 * 60 if remember_me else None
    response.set_cookie(
        key="mc_session",
        value="authenticated",
        httponly=True,
        samesite="lax",
        max_age=max_age,
    )


def set_remembered_user(response: RedirectResponse, username: str, remember_me: bool):
    max_age = REMEMBER_ME_DAYS * 24 * 60 * 60 if remember_me else None
    if remember_me:
        response.set_cookie(
            key="mc_username",
            value=username,
            httponly=False,
            samesite="lax",
            max_age=max_age,
        )
        response.set_cookie(
            key="mc_remember_me",
            value="1",
            httponly=False,
            samesite="lax",
            max_age=max_age,
        )
    else:
        response.delete_cookie("mc_username")
        response.delete_cookie("mc_remember_me")


def parse_date(value):
    if not value:
        return None
    value = str(value).strip()
    if not value:
        return None

    try:
        return date.fromisoformat(value)
    except ValueError:
        pass

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    return None


def parse_money_to_float(value):
    if value is None:
        return 0.0

    text = str(value).strip()
    if not text:
        return 0.0

    normalized = text.replace("R$", "").replace(" ", "")
    normalized = normalized.replace(".", "").replace(",", ".")

    try:
        return float(Decimal(normalized))
    except (InvalidOperation, ValueError):
        return 0.0


def extract_pdf_text(pdf_bytes: bytes) -> str:
    if not pdf_bytes:
        return ""

    if pdf_bytes.startswith(b"%PDF"):
        try:
            from pypdf import PdfReader
        except Exception:
            return ""

        try:
            reader = PdfReader(io.BytesIO(pdf_bytes))
            pages = []
            for page in reader.pages:
                text = page.extract_text() or ""
                pages.append(text)
            joined = "\n".join(pages)
            if joined.strip():
                return joined
        except Exception:
            pass

    for encoding in ("utf-8", "latin-1"):
        try:
            text = pdf_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
        if text.strip():
            return text

    return ""


def parse_invoice_from_text(text: str, filename: str = "") -> dict:
    raw_text = (text or "")
    raw_text = raw_text.replace("�\x89", "é").replace("\x89", "é").replace("\x80", "")
    raw_text = raw_text.replace("\r", "\n")
    full_text = raw_text + "\n" + (filename or "")
    result = {
        "empresa_nome": "",
        "numero_nf": "",
        "valor_nf": 0.0,
        "data_nf": "",
        "razao_nf": "",
    }

    cleaned_lines = [line.strip() for line in re.split(r"[\n\r]+", full_text) if line.strip()]
    candidate_lines = []
    for line in cleaned_lines:
        repair_line = line.replace("�\x89", "é").replace("\x89", "é").replace("\x80", "")
        if len(repair_line) < 6:
            continue
        if re.search(r"(?:NF|NFE|NOTA|FATURA|DANFE|CNPJ|IE|CHAVE|SÉRIE|SERIE|TOTAL|VALOR|PAGAMENTO|EMISS|DATA|R\$)", repair_line, flags=re.I):
            continue
        if re.search(r"[A-Za-zÀ-ÿ]", repair_line) and not re.search(r"\d", repair_line):
            candidate_lines.append(repair_line)

    if candidate_lines:
        result["empresa_nome"] = max(candidate_lines, key=len).strip(" -_")

    file_name = Path(filename or "").stem or ""
    if file_name:
        match = re.search(r"(?i)([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ0-9 .&-]{4,80})(?:\s+(?:NF|NFE|NOTA|FATURA)[-\s]?[A-Z0-9/.-]+)", file_name)
        if match and not result["empresa_nome"]:
            result["empresa_nome"] = match.group(1).strip(" -_")

    numero_patterns = [
        r"(?:NF(?:[- ]?E)?|NFE|NOTA FISCAL|N[º°])\s*[:\-]*\s*[Nº°]?\s*([0-9]{1,12})",
        r"(?:N[º°]|NÚMERO|NUMERO)[\s:]*([0-9]{1,12})",
        r"(?:N[º°]|N°)\s*([0-9]{1,12})",
        r"(?:NF|NOTA FISCAL|N[º°])[^\d]{0,15}([0-9]{1,12})",
    ]
    for pattern in numero_patterns:
        numero_match = re.search(pattern, full_text, flags=re.I)
        if numero_match:
            result["numero_nf"] = numero_match.group(1).strip(" -_")
            break

    if not result["numero_nf"]:
        for pattern in numero_patterns:
            filename_match = re.search(pattern, file_name, flags=re.I)
            if filename_match:
                result["numero_nf"] = filename_match.group(1).strip(" -_")
                break

    total_match = re.search(
        r"(?:valor\s*(?:total|liquido)|total\s*(?:da\s*)?(?:nota|nf)|total\s*geral)[^\n]{0,40}R\$\s*([0-9]{1,3}(?:\.[0-9]{3})*(?:[.,][0-9]{2})|[0-9]+(?:[.,][0-9]{2}))",
        full_text,
        flags=re.I,
    )
    if total_match:
        result["valor_nf"] = parse_money_to_float(total_match.group(1))
    else:
        money_values = re.findall(r"R\$\s*([0-9]{1,3}(?:\.[0-9]{3})*(?:[.,][0-9]{2})|[0-9]+(?:[.,][0-9]{2}))", full_text, flags=re.I)
        if money_values:
            result["valor_nf"] = max(parse_money_to_float(value) for value in money_values)

    if result["valor_nf"] == 0:
        filename_value_match = re.search(r"(\d+(?:[.,]\d{2}))", file_name)
        if filename_value_match:
            result["valor_nf"] = parse_money_to_float(filename_value_match.group(1))

    date_context_match = re.search(r"(?:data\s*(?:de\s*)?(?:emiss[ãa]o|recebimento|emissao)|emissao)[:\s-]*([0-9]{2}/[0-9]{2}/[0-9]{4}|[0-9]{2}-[0-9]{2}-[0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})", full_text, flags=re.I)
    if date_context_match:
        result["data_nf"] = date_context_match.group(1)
    else:
        date_match = re.search(r"(\d{2}/\d{2}/\d{4}|\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})", full_text)
        if date_match:
            result["data_nf"] = date_match.group(1)

    if not result["data_nf"]:
        date_match = re.search(r"(\d{2}/\d{2}/\d{4}|\d{2}-\d{2}-\d{4}|\d{4}-\d{2}-\d{2})", file_name)
        if date_match:
            result["data_nf"] = date_match.group(1)

    result["razao_nf"] = result["empresa_nome"]
    return result


def round_money(value):
    if value is None:
        return 0.0
    try:
        return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError, TypeError):
        return 0.0


def format_currency_br(value):
    try:
        numeric = Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        formatted = f"{numeric:,.2f}"
        return formatted.replace(",", "_").replace(".", ",").replace("_", ".")
    except (InvalidOperation, ValueError, TypeError):
        return "0,00"


def normalize_cnpj(value: str) -> str:
    digits = re.sub(r"\D", "", str(value or ""))
    return digits[:14]


def is_valid_cnpj(value: str) -> bool:
    cnpj = normalize_cnpj(value)
    if len(cnpj) != 14 or cnpj == cnpj[0] * 14:
        return False

    def calc_digit(digits: str, weight: list[int]) -> str:
        total = sum(int(digit) * factor for digit, factor in zip(digits, weight))
        remainder = total % 11
        return "0" if remainder < 2 else str(11 - remainder)

    digits_base = cnpj[:12]
    first = calc_digit(digits_base, list(range(5, 1, -1)) + list(range(9, 1, -1)))
    second = calc_digit(digits_base + first, list(range(6, 1, -1)) + list(range(9, 1, -1)))
    return cnpj == digits_base + first + second


def format_cnpj(value: str) -> str:
    digits = normalize_cnpj(value)
    if len(digits) != 14:
        return digits
    return f"{digits[:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"


@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    if is_authenticated(request):
        return RedirectResponse(url="/", status_code=307)
    saved_username = request.cookies.get("mc_username", "").strip()
    remember_me = request.cookies.get("mc_remember_me") == "1"
    return templates.TemplateResponse(
        "login.html",
        {"request": request, "saved_username": saved_username, "remember_me": remember_me},
    )


@app.post("/login")
def login_submit(
    request: Request,
    username: str = Form(...),
    password: str = Form(...),
    remember_me: str = Form("0"),
):
    remember_selected = remember_me == "1"
    if username == LOGIN_USER and password == LOGIN_PASSWORD:
        response = RedirectResponse(url="/", status_code=303)
        set_auth_cookie(response, remember_me=remember_selected)
        set_remembered_user(response, username, remember_selected)
        return response

    return templates.TemplateResponse(
        "login.html",
        {
            "request": request,
            "error": "Credenciais inválidas. Tente novamente.",
            "saved_username": username,
            "remember_me": remember_selected,
        },
    )


@app.get("/logout")
def logout():
    response = RedirectResponse(url="/login", status_code=307)
    response.delete_cookie("mc_session")
    response.delete_cookie("mc_username")
    response.delete_cookie("mc_remember_me")
    return response


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    db = SessionLocal()
    empresas = db.query(Empresa).all()
    contatos = db.query(Contato).all()
    contratos = db.query(Contrato).all()
    nfs = db.query(NF).all()

    hoje = datetime.utcnow().date()
    inicio_dia = datetime.combine(hoje, datetime.min.time())
    fim_dia = inicio_dia + timedelta(days=1)

    resumo_dia = {
        "empresas": db.query(Empresa).filter(Empresa.data_cadastro >= inicio_dia, Empresa.data_cadastro < fim_dia).count(),
        "contatos": db.query(Contato).filter(Contato.data_cadastro >= inicio_dia, Contato.data_cadastro < fim_dia).count(),
        "contratos": db.query(Contrato).filter(Contrato.data_cadastro >= inicio_dia, Contrato.data_cadastro < fim_dia).count(),
        "nfs": db.query(NF).filter(NF.data_cadastro >= inicio_dia, NF.data_cadastro < fim_dia).count(),
    }
    db.close()

    success_message = request.query_params.get("msg")
    error_message = request.query_params.get("error")

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "empresas": empresas,
            "contatos": contatos,
            "contratos": contratos,
            "nfs": nfs,
            "success_message": success_message,
            "error_message": error_message,
            "resumo_dia": resumo_dia,
            "format_currency_br": format_currency_br,
        },
    )


@app.post("/empresas")
def add_empresa(
    nome: str = Form(...),
    cnpj: str = Form(""),
    tipo: str = Form(""),
    status: str = Form("ativo"),
    observacoes: str = Form(""),
):
    cnpj_normalizado = normalize_cnpj(cnpj)
    if not cnpj_normalizado or not is_valid_cnpj(cnpj_normalizado):
        mensagem = "CNPJ obrigatório e inválido. Use o formato 00.000.000/0000-00."
        return RedirectResponse(url=f"/?error={quote(mensagem)}", status_code=303)

    db = SessionLocal()
    empresa = Empresa(
        nome=nome,
        cnpj=format_cnpj(cnpj_normalizado),
        tipo=tipo,
        status=status,
        observacoes=observacoes,
    )
    db.add(empresa)
    db.commit()
    db.close()

    mensagem = f"Empresa {nome} cadastrada"
    return RedirectResponse(url=f"/?msg={quote(mensagem)}", status_code=303)


@app.post("/contatos")
def add_contato(
    empresa_id: str = Form(""),
    empresa_nome: str = Form(""),
    nome: str = Form(...),
    cargo: str = Form(""),
    email: str = Form(""),
    telefone: str = Form(""),
    observacoes: str = Form(""),
):
    db = SessionLocal()
    empresa_id_value = None
    empresa_id_text = (empresa_id or "").strip()
    if empresa_id_text:
        try:
            empresa_id_value = int(empresa_id_text)
        except ValueError:
            empresa_id_value = None

    empresa_nome = (empresa_nome or "").strip()
    if not empresa_id_value and empresa_nome:
        empresa = db.query(Empresa).filter(Empresa.nome == empresa_nome).first()
        if not empresa:
            empresa = Empresa(
                nome=empresa_nome,
                cnpj="",
                tipo="Cliente",
                status="ativo",
                observacoes="Empresa criada automaticamente no cadastro de contato",
            )
            db.add(empresa)
            db.commit()
        empresa_id_value = empresa.id

    if not empresa_id_value:
        db.close()
        mensagem = "Selecione uma empresa ou informe o nome de uma nova empresa."
        return RedirectResponse(url=f"/?error={quote(mensagem)}", status_code=303)

    contato = Contato(
        empresa_id=empresa_id_value,
        nome=nome,
        cargo=cargo,
        email=email,
        telefone=telefone,
        observacoes=observacoes,
    )
    db.add(contato)
    db.commit()
    db.close()

    mensagem = f"Contato {nome} cadastrado"
    return RedirectResponse(url=f"/?msg={quote(mensagem)}", status_code=303)


@app.post("/contratos")
def add_contrato(
    empresa_id: str = Form(""),
    empresa_nome: str = Form(""),
    descricao: str = Form(""),
    categoria: str = Form("fixo"),
    tipo_servico: str = Form(""),
    data_assinatura: str = Form(""),
    vigencia: str = Form(""),
    renovacao: str = Form(""),
    denuncias: str = Form(""),
    valor_atual: float = Form(0.0),
    responsavel: str = Form(""),
    status: str = Form("ativo"),
    observacoes: str = Form(""),
):
    descricao_final = descricao.strip() or "Contrato"
    empresa_nome = (empresa_nome or "").strip()

    db = SessionLocal()
    empresa_id_value = None
    empresa_id_text = (empresa_id or "").strip()
    if empresa_id_text:
        try:
            empresa_id_value = int(empresa_id_text)
        except ValueError:
            empresa_id_value = None

    if not empresa_id_value and empresa_nome:
        empresa = db.query(Empresa).filter(Empresa.nome == empresa_nome).first()
        if not empresa:
            empresa = Empresa(
                nome=empresa_nome,
                cnpj="",
                tipo="Cliente",
                status="ativo",
                observacoes="Empresa criada automaticamente no cadastro de contrato",
            )
            db.add(empresa)
            db.commit()
        empresa_id_value = empresa.id

    if not empresa_id_value:
        db.close()
        mensagem = "Selecione uma empresa ou informe o nome de uma nova empresa para este contrato."
        return RedirectResponse(url=f"/?error={quote(mensagem)}", status_code=303)

    contrato = Contrato(
        empresa_id=empresa_id_value,
        descricao=descricao_final,
        categoria=categoria,
        tipo_servico=tipo_servico,
        data_assinatura=parse_date(data_assinatura),
        vigencia=vigencia,
        renovacao=renovacao,
        denuncias=denuncias.strip() if denuncias else "",
        valor_atual=round_money(valor_atual),
        responsavel=responsavel,
        status=status,
        observacoes=observacoes,
    )
    db.add(contrato)
    db.commit()
    db.close()

    nome_contrato = tipo_servico.strip() or categoria.strip() or "Serviço"
    mensagem = f"{nome_contrato.title()} cadastrado"
    return RedirectResponse(url=f"/?msg={quote(mensagem)}", status_code=303)


@app.get("/nfs/upload-page", response_class=HTMLResponse)
def upload_nf_page(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    db = SessionLocal()
    contratos = db.query(Contrato).all()
    db.close()

    return templates.TemplateResponse(
        "upload_nf.html",
        {
            "request": request,
            "contratos": contratos,
        },
    )


@app.post("/nfs/preview")
def preview_nf_upload(request: Request, file: UploadFile = File(...)):
    if not is_authenticated(request):
        return JSONResponse({"error": "sessao_invalidada"}, status_code=401)

    if not file or not getattr(file, "filename", None):
        return JSONResponse({"error": "arquivo_nao_enviado"}, status_code=400)

    uploaded_bytes = file.file.read()
    text = extract_pdf_text(uploaded_bytes)
    parsed = parse_invoice_from_text(text, file.filename)
    return {
        "empresa_nome": parsed.get("empresa_nome", ""),
        "numero_nf": parsed.get("numero_nf", ""),
        "valor_nf": float(parsed.get("valor_nf", 0.0) or 0.0),
        "data_nf": parsed.get("data_nf", ""),
        "razao_nf": parsed.get("razao_nf", ""),
    }


@app.post("/nfs")
def add_nf(
    contrato_id: str = Form(""),
    empresa_nome: str = Form(""),
    categoria: str = Form(""),
    tipo_servico: str = Form(""),
    descricao_contrato: str = Form(""),
    razao_nf: str = Form(""),
    data_nf: str = Form(""),
    numero_nf: str = Form(""),
    valor_nf: float = Form(0.0),
    data_pagamento: str = Form(""),
    forma_pagamento: str = Form(""),
    descricao_item: str = Form(""),
    observacoes: str = Form(""),
    status_conferencia: str = Form("pendente"),
    file: UploadFile = File(None),
):
    db = SessionLocal()
    categoria = (categoria or "").strip() or "fixo"

    parsed = {}
    if file and getattr(file, "filename", None):
        uploaded_bytes = file.file.read()
        parsed = parse_invoice_from_text(extract_pdf_text(uploaded_bytes), file.filename)

    empresa_nome = (empresa_nome or parsed.get("empresa_nome") or "").strip()
    razao_nf = (razao_nf or parsed.get("razao_nf") or empresa_nome or "").strip()
    numero_nf = (numero_nf or parsed.get("numero_nf") or "").strip()
    valor_nf = round_money(valor_nf or parsed.get("valor_nf") or 0)
    data_nf = parsed.get("data_nf") or data_nf
    data_pagamento = data_pagamento or ""

    contrato_id_value = None
    contrato_id_text = (contrato_id or "").strip()
    if contrato_id_text:
        try:
            contrato_id_value = int(contrato_id_text)
        except ValueError:
            contrato_id_value = None

    if not contrato_id_value:
        empresa = None
        if empresa_nome:
            empresa = db.query(Empresa).filter(Empresa.nome == empresa_nome).first()
            if not empresa:
                empresa = Empresa(
                    nome=empresa_nome,
                    cnpj="",
                    tipo="Cliente",
                    status="ativo",
                    observacoes="Empresa criada automaticamente no cadastro de NF",
                )
                db.add(empresa)
                db.commit()

        if empresa is None:
            db.close()
            mensagem = "Selecione um contrato ou informe a empresa para criar automaticamente."
            return RedirectResponse(url=f"/?error={quote(mensagem)}", status_code=303)

        contrato = db.query(Contrato).filter(
            Contrato.empresa_id == empresa.id,
            Contrato.categoria == categoria,
            Contrato.tipo_servico == (tipo_servico or "")
        ).order_by(Contrato.id.desc()).first()

        if not contrato:
            contrato = Contrato(
                empresa_id=empresa.id,
                descricao=(descricao_contrato or f"Contrato de {tipo_servico or categoria}").strip() or "Contrato automático",
                categoria=categoria,
                tipo_servico=tipo_servico or "Serviço",
                valor_atual=round_money(valor_nf),
                status="ativo",
                observacoes="Contrato gerado automaticamente no cadastro de NF",
            )
            db.add(contrato)
            db.commit()

        contrato_id_value = contrato.id

    nf = NF(
        contrato_id=contrato_id_value,
        razao_nf=razao_nf or empresa_nome or "",
        data_nf=parse_date(data_nf),
        numero_nf=numero_nf,
        valor_nf=round_money(valor_nf),
        data_pagamento=parse_date(data_pagamento),
        forma_pagamento=forma_pagamento,
        descricao_item=descricao_item,
        observacoes=observacoes,
        status_conferencia=status_conferencia,
    )
    db.add(nf)
    db.commit()
    db.close()

    mensagem = f"NF {numero_nf or 'registrada'} cadastrada"
    return RedirectResponse(url=f"/?msg={quote(mensagem)}", status_code=303)


@app.post("/nfs/upload")
def upload_nf_file(
    request: Request,
    empresa_nome: str = Form(""),
    categoria: str = Form("fixo"),
    tipo_servico: str = Form(""),
    descricao_contrato: str = Form(""),
    file: UploadFile = File(...),
):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    fake_form = {
        "empresa_nome": empresa_nome,
        "categoria": categoria,
        "tipo_servico": tipo_servico,
        "descricao_contrato": descricao_contrato,
    }
    form = {**fake_form}
    return add_nf(
        contrato_id=form.get("contrato_id", ""),
        empresa_nome=form.get("empresa_nome", ""),
        categoria=form.get("categoria", "fixo"),
        tipo_servico=form.get("tipo_servico", ""),
        descricao_contrato=form.get("descricao_contrato", ""),
        razao_nf="",
        data_nf="",
        numero_nf="",
        valor_nf=0.0,
        data_pagamento="",
        forma_pagamento="",
        descricao_item="",
        observacoes="",
        status_conferencia="pendente",
        file=file,
    )


def get_relatorio_data(categoria: str | None = None):
    db = SessionLocal()
    nfs = db.query(NF).all()
    filtradas = []
    for nf in nfs:
        contrato = db.query(Contrato).filter(Contrato.id == nf.contrato_id).first() if nf.contrato_id else None
        nf.contrato = contrato
        nf.anexos = db.query(Anexo).filter(Anexo.nf_id == nf.id).all()
        if categoria is None or (contrato and contrato.categoria == categoria):
            filtradas.append(nf)
    db.close()
    return filtradas


def get_nf_alerts(limit: int = 5):
    hoje = date.today()
    alertas = []
    for nf in get_relatorio_data():
        if nf.data_pagamento is None:
            continue
        dias_restantes = (nf.data_pagamento - hoje).days
        if dias_restantes < 0:
            continue
        if dias_restantes <= 30:
            alertas.append({
                "nf": nf,
                "dias_restantes": dias_restantes,
                "nivel": "alerta" if dias_restantes <= 7 else "aviso",
            })
    alertas.sort(key=lambda item: item["dias_restantes"])
    return alertas[:limit]


def get_nf_status_summary(nfs):
    hoje = date.today()
    summary = {"pagas": 0, "pendentes": 0, "vencidas": 0}
    for nf in nfs:
        status = (nf.status_conferencia or "pendente").lower()
        if status == "confirmada":
            summary["pagas"] += 1
            continue
        if nf.data_pagamento is not None:
            if nf.data_pagamento < hoje:
                summary["vencidas"] += 1
            else:
                summary["pendentes"] += 1
        else:
            summary["pendentes"] += 1
    return summary


@app.get("/relatorio")
def relatorio(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    nfs = get_relatorio_data()
    fixos = get_relatorio_data("fixo")
    pontuais = get_relatorio_data("pontual")
    alertas = get_nf_alerts()
    status_totals = get_nf_status_summary(nfs)
    status_fixos = get_nf_status_summary(fixos)
    status_pontuais = get_nf_status_summary(pontuais)
    return templates.TemplateResponse(
        "relatorio.html",
        {
            "request": request,
            "nfs": nfs,
            "alertas": alertas,
            "categoria": "todos",
            "titulo": "Relatório geral",
            "fixos": fixos,
            "pontuais": pontuais,
            "status_totals": status_totals,
            "status_fixos": status_fixos,
            "status_pontuais": status_pontuais,
            "format_currency_br": format_currency_br,
        },
    )


@app.get("/relatorio/fixo")
def relatorio_fixo(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    nfs = get_relatorio_data("fixo")
    fixos = get_relatorio_data("fixo")
    pontuais = get_relatorio_data("pontual")
    return templates.TemplateResponse(
        "relatorio.html",
        {
            "request": request,
            "nfs": nfs,
            "alertas": get_nf_alerts(),
            "categoria": "fixo",
            "titulo": "Contratos fixos",
            "fixos": fixos,
            "pontuais": pontuais,
            "status_totals": get_nf_status_summary(nfs),
            "status_fixos": get_nf_status_summary(fixos),
            "status_pontuais": get_nf_status_summary(pontuais),
            "format_currency_br": format_currency_br,
        },
    )


@app.get("/relatorio/pontual")
def relatorio_pontual(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    nfs = get_relatorio_data("pontual")
    fixos = get_relatorio_data("fixo")
    pontuais = get_relatorio_data("pontual")
    return templates.TemplateResponse(
        "relatorio.html",
        {
            "request": request,
            "nfs": nfs,
            "alertas": get_nf_alerts(),
            "categoria": "pontual",
            "titulo": "Serviços pontuais",
            "fixos": fixos,
            "pontuais": pontuais,
            "status_totals": get_nf_status_summary(nfs),
            "status_fixos": get_nf_status_summary(fixos),
            "status_pontuais": get_nf_status_summary(pontuais),
            "format_currency_br": format_currency_br,
        },
    )


@app.post("/anexos")
def upload_anexo(
    request: Request,
    empresa_id: int = Form(...),
    contrato_id: int = Form(None),
    nf_id: int = Form(None),
    file: UploadFile = File(...),
):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    conteudo = file.file.read()
    db = SessionLocal()
    anexo = Anexo(
        empresa_id=empresa_id,
        contrato_id=contrato_id or None,
        nf_id=nf_id or None,
        arquivo_nome=file.filename or "arquivo",
        mime_type=file.content_type or "application/octet-stream",
        tamanho=len(conteudo),
        conteudo=conteudo,
    )
    db.add(anexo)
    db.commit()
    db.close()

    mensagem = "Arquivo enviado"
    return RedirectResponse(url=f"/?msg={quote(mensagem)}", status_code=303)


@app.get("/anexos/{anexo_id}")
def download_anexo(anexo_id: int):
    db = SessionLocal()
    anexo = db.query(Anexo).filter(Anexo.id == anexo_id).first()
    db.close()
    if not anexo:
        return RedirectResponse(url="/", status_code=307)

    return Response(
        content=anexo.conteudo,
        media_type=anexo.mime_type or "application/octet-stream",
        headers={"Content-Disposition": f'attachment; filename="{anexo.arquivo_nome}"'},
    )


@app.get("/relatorio/excel")
def relatorio_excel(request: Request):
    if not is_authenticated(request):
        return RedirectResponse(url="/login", status_code=307)

    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    wb = Workbook()
    ws = wb.active
    ws.title = "Relatorio"

    header_fill = PatternFill("solid", fgColor="8B1D1D")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="E7D4CF")

    ws.append(["Contrato", "Categoria", "Razão NF", "Data NF", "Nº NF", "Valor", "Pagamento", "Status"])

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for nf in get_relatorio_data():
        ws.append([
            nf.contrato.descricao if nf.contrato else "Contrato não vinculado",
            nf.contrato.categoria if nf.contrato else "sem contrato",
            nf.razao_nf or "",
            nf.data_nf.isoformat() if nf.data_nf else "",
            nf.numero_nf or "",
            round_money(nf.valor_nf or 0),
            nf.forma_pagamento or "",
            nf.status_conferencia or "",
        ])

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=8):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
            cell.alignment = Alignment(vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=6, max_col=6):
        for cell in row:
            cell.number_format = "R$ #,##0.00"

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=8, max_col=8):
        for cell in row:
            status = (cell.value or "").lower()
            if status == "confirmada":
                cell.fill = PatternFill("solid", fgColor="D9F7E8")
            elif status == "pendente":
                cell.fill = PatternFill("solid", fgColor="FFF3D6")
            elif status == "rejeitada":
                cell.fill = PatternFill("solid", fgColor="FDE1E1")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{ws.max_row}"
    ws.sheet_view.showGridLines = False

    widths = [28, 18, 28, 14, 16, 16, 18, 18]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + idx)].width = width

    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    response = StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["Content-Disposition"] = "attachment; filename=relatorio_mercado_central.xlsx"
    return response
