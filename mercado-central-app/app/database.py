import os
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, DateTime, ForeignKey, Text, LargeBinary, inspect, text
from sqlalchemy.orm import declarative_base, relationship, sessionmaker
from sqlalchemy.pool import StaticPool
from datetime import date, datetime

DATABASE_URL = os.getenv("DATABASE_URL", "sqlite:///./data/mercado_central.db")

Base = declarative_base()

engine_kwargs = {"connect_args": {"check_same_thread": False}}
if DATABASE_URL.startswith("sqlite:///:memory:") or DATABASE_URL == "sqlite://":
    engine_kwargs["poolclass"] = StaticPool

engine = create_engine(DATABASE_URL, **engine_kwargs)
SessionLocal = sessionmaker(bind=engine, autoflush=False, autocommit=False)


def initialize_database():
    Base.metadata.create_all(bind=engine)

    inspector = inspect(engine)

    if inspector.has_table("contratos"):
        columns = [col["name"] for col in inspector.get_columns("contratos")]
        if "categoria" not in columns:
            with engine.begin() as conn:
                conn.execute(text("ALTER TABLE contratos ADD COLUMN categoria VARCHAR(50) DEFAULT 'fixo'"))

    if inspector.has_table("empresas"):
        columns = [col["name"] for col in inspector.get_columns("empresas")]
        if "tipo" not in columns:
            with engine.begin() as conn:
                conn.execute(text("ALTER TABLE empresas ADD COLUMN tipo VARCHAR(100)"))

    seed_demo_data()


def seed_demo_data():
    db = SessionLocal()
    if db.query(Empresa).count() > 0:
        db.close()
        return

    empresa = Empresa(
        nome="Empresa Demo",
        cnpj="12.345.678/0001-95",
        tipo="Cliente",
        status="ativo",
        observacoes="Empresa padrão para uso imediato no sistema.",
    )
    db.add(empresa)
    db.commit()

    contato = Contato(
        empresa_id=empresa.id,
        nome="Maria Financeiro",
        cargo="Financeiro",
        email="financeiro@empresademo.com",
        telefone="(31) 3333-4444",
        observacoes="Contato principal da empresa demo.",
    )

    contrato_fixo = Contrato(
        empresa_id=empresa.id,
        descricao="Contrato fixo de manutenção",
        categoria="fixo",
        tipo_servico="Manutenção",
        data_assinatura=date(2026, 1, 15),
        vigencia="12 meses",
        valor_atual=1800.0,
        responsavel="Maria Financeiro",
        status="ativo",
        observacoes="Contrato padrão do sistema.",
    )
    contrato_pontual = Contrato(
        empresa_id=empresa.id,
        descricao="Serviço pontual de limpeza",
        categoria="pontual",
        tipo_servico="Limpeza",
        data_assinatura=date(2026, 2, 10),
        vigencia="A combinar",
        valor_atual=650.0,
        responsavel="Maria Financeiro",
        status="ativo",
        observacoes="Serviço pontual cadastrado como exemplo.",
    )
    db.add_all([contato, contrato_fixo, contrato_pontual])
    db.commit()

    db.add_all([
        NF(
            contrato_id=contrato_fixo.id,
            razao_nf="Empresa Demo",
            data_nf=date(2026, 8, 5),
            numero_nf="NF-001",
            valor_nf=850.0,
            data_pagamento=date(2026, 8, 25),
            forma_pagamento="PIX",
            descricao_item="Manutenção mensal",
            observacoes="NF pendente de pagamento.",
            status_conferencia="pendente",
        ),
        NF(
            contrato_id=contrato_fixo.id,
            razao_nf="Empresa Demo",
            data_nf=date(2026, 7, 18),
            numero_nf="NF-002",
            valor_nf=950.0,
            data_pagamento=date(2026, 7, 22),
            forma_pagamento="TED",
            descricao_item="Manutenção mensal",
            observacoes="NF já paga.",
            status_conferencia="confirmada",
        ),
        NF(
            contrato_id=contrato_pontual.id,
            razao_nf="Empresa Demo",
            data_nf=date(2026, 7, 2),
            numero_nf="NF-003",
            valor_nf=620.0,
            data_pagamento=date(2026, 6, 28),
            forma_pagamento="BOLETO",
            descricao_item="Serviço pontual",
            observacoes="NF vencida.",
            status_conferencia="pendente",
        ),
    ])
    db.commit()
    db.close()


class Empresa(Base):
    __tablename__ = "empresas"

    id = Column(Integer, primary_key=True, index=True)
    nome = Column(String(255), nullable=False)
    cnpj = Column(String(50), nullable=True)
    tipo = Column(String(100), nullable=True)
    status = Column(String(50), default="ativo")
    observacoes = Column(Text, nullable=True)
    data_cadastro = Column(DateTime, default=datetime.utcnow)

    contratos = relationship("Contrato", back_populates="empresa")
    contatos = relationship("Contato", back_populates="empresa")
    anexos = relationship("Anexo", back_populates="empresa")


class Contato(Base):
    __tablename__ = "contatos"

    id = Column(Integer, primary_key=True, index=True)
    empresa_id = Column(Integer, ForeignKey("empresas.id"), nullable=False)
    nome = Column(String(255), nullable=False)
    cargo = Column(String(150), nullable=True)
    email = Column(String(255), nullable=True)
    telefone = Column(String(80), nullable=True)
    observacoes = Column(Text, nullable=True)
    data_cadastro = Column(DateTime, default=datetime.utcnow)

    empresa = relationship("Empresa", back_populates="contatos")


class Contrato(Base):
    __tablename__ = "contratos"

    id = Column(Integer, primary_key=True, index=True)
    empresa_id = Column(Integer, ForeignKey("empresas.id"), nullable=False)
    descricao = Column(String(255), nullable=False)
    categoria = Column(String(50), default="fixo")
    tipo_servico = Column(String(255), nullable=True)
    data_assinatura = Column(Date, nullable=True)
    vigencia = Column(String(100), nullable=True)
    renovacao = Column(String(255), nullable=True)
    denuncias = Column(String(255), nullable=True)
    valor_atual = Column(Float, nullable=True)
    responsavel = Column(String(255), nullable=True)
    status = Column(String(50), default="ativo")
    observacoes = Column(Text, nullable=True)
    data_cadastro = Column(DateTime, default=datetime.utcnow)

    empresa = relationship("Empresa", back_populates="contratos")
    nfs = relationship("NF", back_populates="contrato")
    anexos = relationship("Anexo", back_populates="contrato")


class NF(Base):
    __tablename__ = "nfs"

    id = Column(Integer, primary_key=True, index=True)
    contrato_id = Column(Integer, ForeignKey("contratos.id"), nullable=False)
    razao_nf = Column(String(255), nullable=True)
    data_nf = Column(Date, nullable=True)
    numero_nf = Column(String(100), nullable=True)
    valor_nf = Column(Float, nullable=True)
    data_pagamento = Column(Date, nullable=True)
    forma_pagamento = Column(String(100), nullable=True)
    descricao_item = Column(Text, nullable=True)
    observacoes = Column(Text, nullable=True)
    status_conferencia = Column(String(50), default="pendente")
    data_cadastro = Column(DateTime, default=datetime.utcnow)

    contrato = relationship("Contrato", back_populates="nfs")
    anexos = relationship("Anexo", back_populates="nf")


class Anexo(Base):
    __tablename__ = "anexos"

    id = Column(Integer, primary_key=True, index=True)
    empresa_id = Column(Integer, ForeignKey("empresas.id"), nullable=False)
    contrato_id = Column(Integer, ForeignKey("contratos.id"), nullable=True)
    nf_id = Column(Integer, ForeignKey("nfs.id"), nullable=True)
    arquivo_nome = Column(String(255), nullable=False)
    mime_type = Column(String(100), nullable=True)
    tamanho = Column(Integer, nullable=True)
    conteudo = Column(LargeBinary, nullable=False)
    uploaded_at = Column(DateTime, default=datetime.utcnow)

    empresa = relationship("Empresa", back_populates="anexos")
    contrato = relationship("Contrato", back_populates="anexos")
    nf = relationship("NF", back_populates="anexos")


initialize_database()
