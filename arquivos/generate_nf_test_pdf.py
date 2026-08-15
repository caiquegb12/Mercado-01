from pathlib import Path


def esc(text: str) -> str:
    return text.replace('\\', '\\\\').replace('(', '\\(').replace(')', '\\)')

out_path = Path(r'C:\Users\caiqu\OneDrive\Desktop\Projeto Mercado Central\arquivos\teste_nf_mercado_central.pdf')
out_path.parent.mkdir(parents=True, exist_ok=True)

stream = (
    "BT\n"
    "/F1 18 Tf\n"
    "72 760 Td\n"
    "(NF-e) Tj\n"
    "0 -28 Td\n"
    "/F1 12 Tf\n"
    "(N° 1065444) Tj\n"
    "0 -22 Td\n"
    "(Série 1) Tj\n"
    "0 -24 Td\n"
    "(DATA DE RECEBIMENTO: 24/12/2024) Tj\n"
    "0 -24 Td\n"
    "(PRINCIPAL VAREJO DE COSMÉTICOS) Tj\n"
    "0 -24 Td\n"
    "(CNPJ: 00.000.000/0000-00) Tj\n"
    "0 -24 Td\n"
    "(TOTAL DA NOTA: R$ 67,90) Tj\n"
    "ET\n"
)

stream = esc(stream)
# keep content as plain text content stream with a simpler escape map
stream = stream.replace('\\n', '\n')

objects = [
    "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
    "2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n",
    "3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n",
    f"4 0 obj\n<< /Length {len(stream.encode('latin-1'))} >>\nstream\n{stream}endstream\nendobj\n",
    "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n",
]

pdf = '%PDF-1.4\n'.encode('latin-1')
offsets = []
for obj in objects:
    offsets.append(len(pdf))
    pdf += obj.encode('latin-1')

xref_position = len(pdf)
pdf += f"xref\n0 {len(objects)+1}\n".encode('latin-1')
pdf += b"0000000000 65535 f \n"
for off in offsets:
    pdf += f"{off:010d} 00000 n \n".encode('latin-1')

pdf += f"trailer\n<< /Size {len(objects)+1} /Root 1 0 R >>\nstartxref\n{xref_position}\n%%EOF\n".encode('latin-1')

out_path.write_bytes(pdf)
print(out_path)
print(f"Tamanho: {out_path.stat().st_size} bytes")
