from pathlib import Path


def escape_pdf_text(value: str) -> str:
    return value.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')

out_path = Path('C:/Users/caiqu/OneDrive/Desktop/Projeto Mercado Central/arquivos/teste_nf_mercado_central.pdf')
out_path.parent.mkdir(parents=True, exist_ok=True)

content = (
    "BT\n"
    "/F1 18 Tf\n"
    "72 760 Td\n"
    f"({escape_pdf_text('NF-e')}) Tj\n"
    "0 -28 Td\n"
    "/F1 12 Tf\n"
    f"({escape_pdf_text('N° 1065444')}) Tj\n"
    "0 -22 Td\n"
    f"({escape_pdf_text('Série 1')}) Tj\n"
    "0 -24 Td\n"
    f"({escape_pdf_text('DATA DE RECEBIMENTO: 24/12/2024')}) Tj\n"
    "0 -24 Td\n"
    f"({escape_pdf_text('PRINCIPAL VAREJO DE COSMÉTICOS')}) Tj\n"
    "0 -24 Td\n"
    f"({escape_pdf_text('CNPJ: 00.000.000/0000-00')}) Tj\n"
    "0 -24 Td\n"
    f"({escape_pdf_text('TOTAL DA NOTA: R$ 67,90')}) Tj\n"
    "ET\n"
)

objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
    f"<< /Length {len(content.encode('latin-1'))} >>\nstream\n{content}endstream",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
]

pdf = bytearray(b'%PDF-1.4\n')
offsets = []
for index, obj in enumerate(objects, start=1):
    offsets.append(len(pdf))
    pdf.extend(f"{index} 0 obj\n".encode('latin-1'))
    pdf.extend(obj.encode('latin-1'))
    pdf.extend(b"\nendobj\n")

xref_pos = len(pdf)
pdf.extend(f"xref\n0 {len(objects) + 1}\n".encode('latin-1'))
pdf.extend(b"0000000000 65535 f \n")
for offset in offsets:
    pdf.extend(f"{offset:010d} 00000 n \n".encode('latin-1'))

pdf.extend(f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref_pos}\n%%EOF\n".encode('latin-1'))

out_path.write_bytes(pdf)
print(f"PDF gerado em: {out_path}")
print(f"Tamanho: {out_path.stat().st_size} bytes")
