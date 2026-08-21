from fastapi.testclient import TestClient

from app.database import SessionLocal, Empresa, Contrato, NF
from app.main import app


def ensure_demo_data() -> None:
    db = SessionLocal()
    db.query(NF).delete()
    db.query(Contrato).delete()
    db.query(Empresa).delete()

    empresa = Empresa(
        nome="Empresa Demo",
        cnpj="12.345.678/0001-95",
        tipo="Cliente",
        status="ativo",
    )
    db.add(empresa)
    db.commit()

    contrato = Contrato(
        empresa_id=empresa.id,
        descricao="Contrato de manutenção",
        categoria="fixo",
        tipo_servico="Manutenção",
        valor_atual=2500.0,
        status="ativo",
    )
    db.add(contrato)
    db.commit()

    dados = [
        ("NF-001", 1200.00, "confirmada", "PIX"),
        ("NF-002", 800.50, "pendente", "BOLETO"),
        ("NF-003", 1500.25, "confirmada", "TRANSFERÊNCIA"),
    ]

    for numero_nf, valor, status, pagamento in dados:
        db.add(
            NF(
                contrato_id=contrato.id,
                razao_nf="Empresa Demo",
                numero_nf=numero_nf,
                valor_nf=valor,
                forma_pagamento=pagamento,
                status_conferencia=status,
            )
        )

    db.commit()
    db.close()


def main() -> None:
    ensure_demo_data()

    client = TestClient(app)
    client.post("/login", data={"username": "mercado", "password": "mercado123"})
    response = client.get("/relatorio/excel")

    if response.status_code != 200:
        raise SystemExit(f"Falha ao gerar Excel: status={response.status_code}")

    with open("relatorio_mercado_central.xlsx", "wb") as file:
        file.write(response.content)

    print("Arquivo gerado em: relatorio_mercado_central.xlsx")
    print(f"Bytes: {len(response.content)}")


if __name__ == "__main__":
    main()
