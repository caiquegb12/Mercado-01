from app.main import parse_invoice_from_text

sample = """NF-e N° 1065444
ITEM 1 R$ 500,00
ITEM 2 R$ 230,00
VALOR TOTAL DA NOTA 220,00
FRETE R$ 10,00
"""

result = parse_invoice_from_text(sample, "nf-220.pdf")
print(result["valor_nf"])
