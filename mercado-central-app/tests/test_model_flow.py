import os
import unittest
from datetime import date
from unittest.mock import patch
from fastapi.testclient import TestClient

os.environ["DATABASE_URL"] = "sqlite:///:memory:"

from app.database import SessionLocal, Empresa, Contato, Contrato, NF, Base, engine, seed_demo_data
from app.main import app, parse_invoice_from_text


class ModelFlowTest(unittest.TestCase):
    def setUp(self):
        Base.metadata.create_all(bind=engine)
        self.db = SessionLocal()
        self.db.query(NF).delete()
        self.db.query(Contrato).delete()
        self.db.query(Contato).delete()
        self.db.query(Empresa).delete()
        self.db.commit()

    def tearDown(self):
        self.db.query(NF).delete()
        self.db.query(Contrato).delete()
        self.db.query(Contato).delete()
        self.db.query(Empresa).delete()
        self.db.commit()
        self.db.close()

    def test_default_seed_data_creates_demo_company_and_nf_samples(self):
        self.db.query(NF).delete()
        self.db.query(Contrato).delete()
        self.db.query(Contato).delete()
        self.db.query(Empresa).delete()
        self.db.commit()

        seed_demo_data()

        empresa = self.db.query(Empresa).filter(Empresa.nome == "Empresa Demo").first()
        self.assertIsNotNone(empresa)
        self.assertGreaterEqual(self.db.query(Contrato).count(), 2)
        self.assertGreaterEqual(self.db.query(NF).count(), 3)

    def test_empresa_contact_and_nf_can_be_saved(self):
        empresa = Empresa(nome="Mercado Central", cnpj="00.000.000/0001-00", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contato = Contato(
            empresa_id=empresa.id,
            nome="Maria",
            cargo="Financeiro",
            email="maria@empresa.com",
            telefone="(31) 3333-4444",
        )
        self.db.add(contato)

        contrato = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato de manutenção elétrica",
            categoria="fixo",
            tipo_servico="Manutenção",
            valor_atual=2500.0,
            status="ativo",
        )
        self.db.add(contrato)
        self.db.commit()

        nf = NF(
            contrato_id=contrato.id,
            razao_nf="Mercado Central",
            numero_nf="001/2026",
            valor_nf=2500.0,
            forma_pagamento="PIX",
            status_conferencia="pendente",
        )
        self.db.add(nf)
        self.db.commit()

        self.assertEqual(self.db.query(Empresa).count(), 1)
        self.assertEqual(self.db.query(Contato).count(), 1)
        self.assertEqual(self.db.query(Contrato).count(), 1)
        self.assertEqual(self.db.query(NF).count(), 1)

    def test_parse_invoice_from_text_keeps_company_name_when_status_is_same_line(self):
        sample = (
            "DANFSe v2.0\n"
            "Documento Auxiliar da NFS-e\n"
            "NÚMERO DA NFS-E\n"
            "1969\n"
            "COMPETÊNCIA DA NFS-E\n"
            "03/08/2026\n"
            "SITUAÇÃO DA NFS-E: EMITIDA PHD AMBIENTAL LTDA\n"
            "VALOR LÍQUIDO NFS-e + IBS/CBS\n"
            "R$ 3.524,29\n"
        )

        parsed = parse_invoice_from_text(sample, "danfse_status_accented.pdf")

        self.assertEqual(parsed["numero_nf"], "1969")
        self.assertEqual(parsed["data_nf"], "03/08/2026")
        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_report_page_handles_orphan_nf_without_crashing(self):
        empresa = Empresa(nome="Empresa Teste", cnpj="00.000.000/0001-99", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato pontual de limpeza",
            categoria="pontual",
            tipo_servico="Limpeza",
            valor_atual=1500.0,
            status="ativo",
        )
        self.db.add(contrato)
        self.db.commit()

        nf_orfana = NF(
            contrato_id=9999,
            razao_nf="Empresa sem contrato",
            numero_nf="NF-ORFANA",
            valor_nf=456.0,
            forma_pagamento="DINHEIRO",
            status_conferencia="pendente",
        )
        self.db.add(nf_orfana)
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})
        response = client.get("/relatorio")

        self.assertEqual(response.status_code, 200)
        self.assertIn("Relatório de conferência", response.text)

    def test_report_page_has_polished_layout_markers(self):
        empresa = Empresa(nome="Empresa Layout", cnpj="00.000.000/0001-77", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato de layout",
            categoria="fixo",
            tipo_servico="Design",
            valor_atual=900.0,
            status="ativo",
        )
        self.db.add(contrato)
        self.db.commit()

        nf = NF(
            contrato_id=contrato.id,
            razao_nf="Empresa Layout",
            numero_nf="NF-LAYOUT",
            valor_nf=900.0,
            forma_pagamento="PIX",
            status_conferencia="confirmada",
        )
        self.db.add(nf)
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})
        response = client.get("/relatorio")

        self.assertEqual(response.status_code, 200)
        self.assertIn("report-summary", response.text)
        self.assertIn("status-badge", response.text)
        self.assertIn("Exportar Excel", response.text)

    def test_report_pages_keep_excel_export_for_filtered_categories(self):
        empresa = Empresa(nome="Empresa Filtrada", cnpj="00.111.000/0001-99", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato_fixo = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato fixo filtrado",
            categoria="fixo",
            tipo_servico="Manutenção",
            valor_atual=500.0,
            status="ativo",
        )
        self.db.add(contrato_fixo)

        contrato_pontual = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato pontual filtrado",
            categoria="pontual",
            tipo_servico="Instalação",
            valor_atual=200.0,
            status="ativo",
        )
        self.db.add(contrato_pontual)
        self.db.commit()

        self.db.add_all([
            NF(
                contrato_id=contrato_fixo.id,
                razao_nf="Empresa Filtrada",
                numero_nf="NF-FIXA",
                valor_nf=500.0,
                forma_pagamento="PIX",
                status_conferencia="confirmada",
            ),
            NF(
                contrato_id=contrato_pontual.id,
                razao_nf="Empresa Filtrada",
                numero_nf="NF-PONTUAL",
                valor_nf=200.0,
                forma_pagamento="PIX",
                status_conferencia="pendente",
            ),
        ])
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response_fixo = client.get("/relatorio/fixo")
        self.assertEqual(response_fixo.status_code, 200)
        self.assertIn("Exportar Excel", response_fixo.text)
        self.assertIn("/relatorio/excel?categoria=fixo", response_fixo.text)

        response_pontual = client.get("/relatorio/pontual")
        self.assertEqual(response_pontual.status_code, 200)
        self.assertIn("Exportar Excel", response_pontual.text)
        self.assertIn("/relatorio/excel?categoria=pontual", response_pontual.text)

    def test_contact_creation_can_auto_create_company(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.post(
            "/contatos",
            data={
                "empresa_id": "",
                "empresa_nome": "Nova Empresa Automatica",
                "nome": "Joao Silva",
                "cargo": "Financeiro",
                "email": "joao@novaempresa.com",
                "telefone": "(31) 99999-0000",
                "observacoes": "",
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertEqual(self.db.query(Empresa).filter(Empresa.nome == "Nova Empresa Automatica").count(), 1)
        self.assertEqual(self.db.query(Contato).filter(Contato.nome == "Joao Silva").count(), 1)

    def test_contract_form_uses_brazilian_date_mask(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.get("/")

        self.assertEqual(response.status_code, 200)
        self.assertIn('name="data_assinatura"', response.text)
        self.assertIn('placeholder="dd/mm/aaaa"', response.text)

    def test_nf_form_starts_without_preselected_contract_or_category(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.get("/")

        self.assertEqual(response.status_code, 200)
        self.assertIn('<option value="" selected>Selecione um contrato existente</option>', response.text)
        self.assertIn('<option value="" selected>Selecione a categoria</option>', response.text)

    def test_parse_invoice_from_real_nf_pdf_text(self):
        sample_text = """
        NF-e N° 1065444
        Série 1
        DATA DE RECEBIMENTO 24/12/2024
        PRINCIPAL VAREJO DE COSMÉTICOS
        CNPJ 00.000.000/0000-00
        TOTAL DA NOTA R$ 67,90
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["numero_nf"], "1065444")
        self.assertEqual(parsed["valor_nf"], 67.9)
        self.assertEqual(parsed["data_nf"], "24/12/2024")
        self.assertEqual(parsed["empresa_nome"], "PRINCIPAL VAREJO DE COSMÉTICOS")

    def test_parse_invoice_from_real_nfs_e_pdf_text(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        PRESTADOR / FORNECEDOR
        NOME / NOME EMPRESARIAL
        PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_31062002218053816000149000000000196926082644043565_01.pdf")

        self.assertEqual(parsed["numero_nf"], "1969")
        self.assertEqual(parsed["valor_nf"], 3524.29)
        self.assertEqual(parsed["data_nf"], "03/08/2026")
        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_keeps_company_name_and_legal_name_separate(self):
        sample_text = """
        DANFSe v2.0
        NOME FANTASIA: REDE DE VAREJO LTDA
        RAZÃO SOCIAL: PHD AMBIENTAL LTDA
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 1.250,00
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_31062002218053816000149000000000196926082644043565_01.pdf")

        self.assertEqual(parsed["empresa_nome"], "REDE DE VAREJO LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_ignores_nfs_status_label_as_company_name(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        SITUAÇÃO DA NFS-E
        EMITIDA
        PRESTADOR / FORNECEDOR
        NOME / NOME EMPRESARIAL
        PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_status_emitida.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_ignores_status_label_and_colon_variant(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        SITUACAO DA NFS-E: EMITIDA
        PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_status_colon.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_ignores_finalidade_label_and_keeps_company_name(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        SITUAÇÃO DA NFS-E
        NFS-e Gerada
        FINALIDADE
        NFS-e regular
        PRESTADOR / FORNECEDOR
        NOME / NOME EMPRESARIAL
        PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_finalidade.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_prefers_business_name_over_cnpj_in_real_pdf_text(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        Município: Belo Horizonte - MG
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        SITUAÇÃO DA NFS-E
        NFS-e Gerada
        FINALIDADE
        NFS-e regular
        PRESTADOR / FORNECEDOR
        CNPJ / CPF / NIF
        18.053.816/0001-49
        NOME / NOME EMPRESARIAL
        PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_31062002218053816000149000000000196926082644043565_01.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_ignores_accented_status_label_variant(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        SITUAÇÃO DA NFS-E: EMITIDA PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_status_accented.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_strips_status_label_from_same_line_company_name(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        SITUACAO DA NFS-E: EMITIDA PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_status_company_same_line.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_ignores_generic_danfe_filename_fallback(self):
        filename = "danfe_310620022180538160014900000000019629082644043565_01.pdf"

        parsed = parse_invoice_from_text("", filename)

        self.assertEqual(parsed["numero_nf"], "")
        self.assertEqual(parsed["valor_nf"], 0.0)

    def test_parse_invoice_ignores_mercado_central_as_company_name(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        PRESTADOR / FORNECEDOR
        NOME / NOME EMPRESARIAL
        PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "mercado-central-2026.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_prefers_provider_name_over_municipality_name(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        PRESTADOR / FORNECEDOR
        NOME / NOME EMPRESARIAL
        PHD AMBIENTAL LTDA
        MUNICÍPIO: JUNDIAÍ
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_31062002218053816000149000000000196926082644043565_01.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_ignores_indicator_municipal_label(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        INDICADOR MUNICIPAL (INSCRIÇÃO)
        PHD AMBIENTAL LTDA
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_indicator_municipal.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_ignores_uf_suffix_after_city_name(self):
        sample_text = """
        DANFSe v2.0
        Documento Auxiliar da NFS-e
        NÚMERO DA NFS-E
        1969
        COMPETÊNCIA DA NFS-E
        03/08/2026
        PRESTADOR / FORNECEDOR
        NOME / NOME EMPRESARIAL
        PHD AMBIENTAL LTDA
        MUNICÍPIO: JUNDIAÍ/SP
        UF: SP
        VALOR LÍQUIDO NFS-e + IBS/CBS
        R$ 3.524,29
        """

        parsed = parse_invoice_from_text(sample_text, "danfse_31062002218053816000149000000000196926082644043565_01.pdf")

        self.assertEqual(parsed["empresa_nome"], "PHD AMBIENTAL LTDA")
        self.assertEqual(parsed["razao_nf"], "PHD AMBIENTAL LTDA")

    def test_parse_invoice_uses_total_value_not_larger_item_value(self):
        sample_text = """
        NF-e N° 1065444
        ITEM 1 R$ 500,00
        ITEM 2 R$ 230,00
        TOTAL DA NOTA R$ 67,90
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["valor_nf"], 67.9)

    def test_parse_invoice_prefers_total_value_even_without_currency_symbol(self):
        sample_text = """
        NF-e N° 1065444
        ITEM 1 R$ 500,00
        ITEM 2 R$ 230,00
        VALOR TOTAL DA NOTA 220,00
        FRETE R$ 10,00
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["valor_nf"], 220.0)

    def test_parse_invoice_prefers_total_value_even_when_other_amounts_follow(self):
        sample_text = """
        NF-e N° 1065444
        DATA DE EMISSAO 24/12/2024
        PRINCIPAL VAREJO DE COSMÉTICOS
        VALOR TOTAL 67,90
        FRETE R$ 15,00
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["numero_nf"], "1065444")
        self.assertEqual(parsed["valor_nf"], 67.9)
        self.assertEqual(parsed["data_nf"], "24/12/2024")
        self.assertEqual(parsed["empresa_nome"], "PRINCIPAL VAREJO DE COSMÉTICOS")

    def test_parse_invoice_prefers_note_total_over_product_total_line(self):
        sample_text = """
        NF-e N° 1065444
        VALOR TOTAL DOS PRODUTOS: R$ 205,00
        VALOR TOTAL DA NOTA: R$ 220,00
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["valor_nf"], 220.0)

    def test_parse_invoice_does_not_treat_product_total_as_note_total(self):
        sample_text = """
        NF-e N° 1065444
        VALOR TOTAL DOS PRODUTOS: R$ 20.500,00
        VALOR TOTAL DA NOTA: R$ 220,00
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["valor_nf"], 220.0)

    def test_parse_invoice_number_prefers_nf_label_over_total_line(self):
        sample_text = """
        VALOR TOTAL DA NOTA: R$ 220,00
        NF-e N° 1065444
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["numero_nf"], "1065444")

    def test_parse_invoice_number_removes_zero_padding(self):
        sample_text = """
        NF-e N° 000123
        VALOR TOTAL DA NOTA: R$ 220,00
        """

        parsed = parse_invoice_from_text(sample_text, "nf-000123.pdf")

        self.assertEqual(parsed["numero_nf"], "123")

    def test_parse_invoice_handles_mojibake_from_pdf_extraction(self):
        sample_text = """
        NF-e
        NÂ° 1065444
        SÃ©rie 1
        DATA DE RECEBIMENTO: 24/12/2024
        PRINCIPAL VAREJO DE COSMÃ©TICOS
        TOTAL DA NOTA: R$ 67,90
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["numero_nf"], "1065444")
        self.assertEqual(parsed["valor_nf"], 67.9)
        self.assertEqual(parsed["data_nf"], "24/12/2024")
        self.assertEqual(parsed["empresa_nome"], "PRINCIPAL VAREJO DE COSMÉTICOS")
        self.assertEqual(parsed["razao_nf"], "PRINCIPAL VAREJO DE COSMÉTICOS")

    def test_parse_invoice_from_text_strips_company_prefix_labels(self):
        sample_text = """
        NF-e N° 1065444
        RAZÃO SOCIAL: ALFA TECNOLOGIA E COMÉRCIO LTDA.
        NATUREZA DA OPERAÇÃO: COMPRA DE MERCADORIA
        TOTAL DA NOTA R$ 67,90
        """

        parsed = parse_invoice_from_text(sample_text, "nf-1065444.pdf")

        self.assertEqual(parsed["empresa_nome"], "ALFA TECNOLOGIA E COMÉRCIO LTDA.")
        self.assertEqual(parsed["razao_nf"], "ALFA TECNOLOGIA E COMÉRCIO LTDA.")

    def test_nf_preview_route_parses_uploaded_file(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        sample = "NF-e N° 1065444\nDATA DE RECEBIMENTO 24/12/2024\nPRINCIPAL VAREJO DE COSMÉTICOS\nTOTAL DA NOTA R$ 67,90".encode("utf-8")
        response = client.post(
            "/nfs/preview",
            files={"file": ("nf-1065444.txt", sample, "text/plain")},
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(payload["numero_nf"], "1065444")
        self.assertEqual(payload["valor_nf"], 67.9)
        self.assertEqual(payload["data_nf"], "24/12/2024")
        self.assertEqual(payload["empresa_nome"], "PRINCIPAL VAREJO DE COSMÉTICOS")

    @patch("app.main.pytesseract.image_to_string", return_value="NF-e N° 1065444\nDATA DE RECEBIMENTO 24/12/2024\nPRINCIPAL VAREJO DE COSMÉTICOS\nTOTAL DA NOTA: R$ 67,90\n")
    def test_nf_preview_route_can_ocr_uploaded_image(self, mock_image_to_string):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        png_bytes = b"\x89PNG\r\n\x1a\n" + b"0" * 32
        response = client.post(
            "/nfs/preview",
            files={"file": ("nf-1065444.png", png_bytes, "image/png")},
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 200)
        payload = response.json()
        self.assertEqual(payload["numero_nf"], "1065444")
        self.assertEqual(payload["valor_nf"], 67.9)
        self.assertEqual(payload["data_nf"], "24/12/2024")
        self.assertEqual(payload["empresa_nome"], "PRINCIPAL VAREJO DE COSMÉTICOS")

    def test_nf_creation_can_auto_create_company_and_contract(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.post(
            "/nfs",
            data={
                "contrato_id": "",
                "empresa_nome": "Empresa NF Automatica",
                "categoria": "pontual",
                "tipo_servico": "Limpeza",
                "descricao_contrato": "Contrato de limpeza",
                "razao_nf": "Empresa NF Automatica",
                "data_nf": "2026-08-15",
                "numero_nf": "NF-001",
                "valor_nf": "250.50",
                "data_pagamento": "2026-08-20",
                "forma_pagamento": "PIX",
                "status_conferencia": "pendente",
                "descricao_item": "",
                "observacoes": "",
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertEqual(self.db.query(Empresa).filter(Empresa.nome == "Empresa NF Automatica").count(), 1)
        self.assertEqual(self.db.query(Contrato).filter(Contrato.empresa_id == self.db.query(Empresa).filter(Empresa.nome == "Empresa NF Automatica").first().id).count(), 1)
        self.assertEqual(self.db.query(NF).filter(NF.numero_nf == "NF-001").count(), 1)

    def test_report_page_formats_total_value_in_brazilian_currency(self):
        empresa = Empresa(nome="Empresa Valor", cnpj="00.000.000/0001-66", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato valor",
            categoria="fixo",
            tipo_servico="Serviço",
            valor_atual=1500.0,
            status="ativo",
        )
        self.db.add(contrato)
        self.db.commit()

        self.db.add_all([
            NF(
                contrato_id=contrato.id,
                razao_nf="Empresa Valor",
                numero_nf="VAL-1",
                valor_nf=1250.5,
                forma_pagamento="PIX",
                status_conferencia="confirmada",
            ),
            NF(
                contrato_id=contrato.id,
                razao_nf="Empresa Valor",
                numero_nf="VAL-2",
                valor_nf=500.0,
                forma_pagamento="PIX",
                status_conferencia="pendente",
            ),
        ])
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})
        response = client.get("/relatorio")

        self.assertEqual(response.status_code, 200)
        self.assertIn("R$ 1.750,50", response.text)

    def test_report_page_allows_printing_by_category(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.get("/relatorio")

        self.assertEqual(response.status_code, 200)
        self.assertIn('name="print-category"', response.text)
        self.assertIn('value="todos"', response.text)
        self.assertIn('value="fixo"', response.text)
        self.assertIn('value="pontual"', response.text)

    def test_report_page_shows_split_status_counts_by_category(self):
        empresa = Empresa(nome="Empresa Split", cnpj="00.000.000/0001-88", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato_fixo = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato fixo",
            categoria="fixo",
            tipo_servico="Manutenção",
            valor_atual=800.0,
            status="ativo",
        )
        contrato_pontual = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato pontual",
            categoria="pontual",
            tipo_servico="Limpeza",
            valor_atual=500.0,
            status="ativo",
        )
        self.db.add_all([contrato_fixo, contrato_pontual])
        self.db.commit()

        self.db.add_all([
            NF(
                contrato_id=contrato_fixo.id,
                razao_nf="Empresa Split",
                numero_nf="FIX-1",
                valor_nf=400.0,
                forma_pagamento="PIX",
                status_conferencia="confirmada",
            ),
            NF(
                contrato_id=contrato_fixo.id,
                razao_nf="Empresa Split",
                numero_nf="FIX-2",
                valor_nf=400.0,
                forma_pagamento="PIX",
                status_conferencia="pendente",
            ),
            NF(
                contrato_id=contrato_pontual.id,
                razao_nf="Empresa Split",
                numero_nf="PON-1",
                valor_nf=500.0,
                forma_pagamento="PIX",
                status_conferencia="pendente",
                data_pagamento=date(2025, 1, 1),
            ),
        ])
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})
        response = client.get("/relatorio")

        self.assertEqual(response.status_code, 200)
        self.assertIn("Fixos", response.text)
        self.assertIn("Pontuais", response.text)
        self.assertIn("Pagas fixas", response.text)
        self.assertIn("Pendentes fixas", response.text)
        self.assertIn("Vencidas fixas", response.text)
        self.assertIn("Pagas pontuais", response.text)
        self.assertIn("Pendentes pontuais", response.text)
        self.assertIn("Vencidas pontuais", response.text)

    def test_report_excel_export_works(self):
        empresa = Empresa(nome="Empresa Export", cnpj="00.000.000/0001-55", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato de manutenção",
            categoria="fixo",
            tipo_servico="Manutenção",
            valor_atual=1500.0,
            status="ativo",
        )
        self.db.add(contrato)
        self.db.commit()

        nf = NF(
            contrato_id=contrato.id,
            razao_nf="Empresa Export",
            numero_nf="NF-100",
            valor_nf=1500.0,
            forma_pagamento="PIX",
            status_conferencia="confirmada",
        )
        self.db.add(nf)
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})
        response = client.get("/relatorio/excel")

        self.assertEqual(response.status_code, 200)
        self.assertEqual(response.headers.get("content-type"), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.assertIn("relatorio_mercado_central.xlsx", response.headers.get("content-disposition", ""))

    def test_invalid_cnpj_is_rejected(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.post(
            "/empresas",
            data={
                "nome": "Empresa Invalida",
                "cnpj": "12345678900",
                "tipo": "Cliente",
                "status": "ativo",
                "observacoes": "",
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertIn("error=CNPJ%20obrigat", response.headers.get("location", ""))
        self.assertIn("inv", response.headers.get("location", ""))

    def test_blank_cnpj_is_rejected(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.post(
            "/empresas",
            data={
                "nome": "Empresa Sem CNPJ",
                "cnpj": "",
                "tipo": "Cliente",
                "status": "ativo",
                "observacoes": "",
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertIn("error=CNPJ%20obrigat", response.headers.get("location", ""))
        self.assertIn("inv", response.headers.get("location", ""))

    def test_success_message_redirect_after_company_registration(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.post(
            "/empresas",
            data={
                "nome": "Empresa Confirmacao",
                "cnpj": "12.345.678/0001-95",
                "tipo": "Cliente",
                "status": "ativo",
                "observacoes": "",
            },
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertIn("msg=Empresa%20Empresa%20Confirmacao%20cadastrada", response.headers.get("location", ""))

    def test_pages_require_login(self):
        client = TestClient(app)
        response = client.get("/", follow_redirects=False)

        self.assertEqual(response.status_code, 307)
        self.assertEqual(response.headers.get("location"), "/login")

    def test_remember_me_persists_username_cookie(self):
        client = TestClient(app)
        response = client.post(
            "/login",
            data={"username": "mercado", "password": "mercado123", "remember_me": "1"},
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertEqual(response.cookies.get("mc_session"), "authenticated")
        self.assertEqual(response.cookies.get("mc_username"), "mercado")

    def test_login_page_prefills_saved_username(self):
        client = TestClient(app)
        client.cookies.set("mc_username", "mercado")

        response = client.get("/login")

        self.assertEqual(response.status_code, 200)
        self.assertIn('value="mercado"', response.text)

    def test_report_has_category_sections_and_nf_alerts(self):
        empresa = Empresa(nome="Empresa alerta", cnpj="00.000.000/0001-33", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato_fixo = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato fixo alerta",
            categoria="fixo",
            tipo_servico="Manutenção",
            valor_atual=500.0,
            status="ativo",
        )
        self.db.add(contrato_fixo)
        self.db.commit()

        nf = NF(
            contrato_id=contrato_fixo.id,
            razao_nf="Empresa alerta",
            data_nf=date(2026, 8, 15),
            numero_nf="NF-ALERTA",
            valor_nf=500.0,
            data_pagamento=date(2026, 8, 20),
            forma_pagamento="PIX",
            status_conferencia="pendente",
        )
        self.db.add(nf)
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.get("/relatorio/fixo")
        self.assertEqual(response.status_code, 200)
        self.assertIn("Contratos fixos", response.text)
        self.assertIn("NF-ALERTA", response.text)

        alerts = client.get("/relatorio")
        self.assertEqual(alerts.status_code, 200)
        self.assertIn("próximo do vencimento", alerts.text.lower())

    def test_upload_nf_file_can_auto_create_company_contract_and_nf(self):
        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.post(
            "/nfs/upload",
            data={
                "empresa_nome": "Empresa Upload",
                "categoria": "fixo",
                "tipo_servico": "Manutenção",
                "descricao_contrato": "Contrato upload",
            },
            files={"file": ("Empresa Upload NF-005 850 15-08-2026.pdf", b"Empresa Upload\nNF-005\n850,00\n15/08/2026\n", "application/pdf")},
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertEqual(self.db.query(Empresa).filter(Empresa.nome == "Empresa Upload").count(), 1)
        self.assertEqual(self.db.query(Contrato).filter(Contrato.empresa_id == self.db.query(Empresa).filter(Empresa.nome == "Empresa Upload").first().id).count(), 1)
        self.assertEqual(self.db.query(NF).filter(NF.numero_nf == "NF-005").count(), 1)

    def test_can_upload_attachment_for_nf(self):
        empresa = Empresa(nome="Empresa doc", cnpj="00.000.000/0001-88", tipo="Cliente")
        self.db.add(empresa)
        self.db.commit()

        contrato = Contrato(
            empresa_id=empresa.id,
            descricao="Contrato doc",
            categoria="fixo",
            tipo_servico="Manutenção",
            valor_atual=1500.0,
            status="ativo",
        )
        self.db.add(contrato)
        self.db.commit()

        nf = NF(
            contrato_id=contrato.id,
            razao_nf="Empresa doc",
            data_nf=date(2026, 8, 15),
            numero_nf="NF-DOC",
            valor_nf=1500.0,
            data_pagamento=date(2026, 8, 20),
            forma_pagamento="PIX",
            status_conferencia="pendente",
        )
        self.db.add(nf)
        self.db.commit()

        client = TestClient(app)
        client.post("/login", data={"username": "mercado", "password": "mercado123"})

        response = client.post(
            "/anexos",
            data={"empresa_id": str(empresa.id), "contrato_id": str(contrato.id), "nf_id": str(nf.id)},
            files={"file": ("nota.pdf", b"conteudo pdf", "application/pdf")},
            follow_redirects=False,
        )

        self.assertEqual(response.status_code, 303)
        self.assertIn("msg=Arquivo%20enviado", response.headers.get("location", ""))


if __name__ == "__main__":
    unittest.main()
