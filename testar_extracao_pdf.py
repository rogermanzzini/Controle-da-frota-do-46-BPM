"""
testar_extracao_pdf.py
Diagnóstico: mostra linha a linha o que pdfplumber extrai de um PDF de nota fiscal.
Uso: python testar_extracao_pdf.py caminho/para/nota.pdf
"""
import sys
import re

def main():
    if len(sys.argv) < 2:
        print("Uso: python testar_extracao_pdf.py caminho/para/nota.pdf")
        sys.exit(1)

    caminho = sys.argv[1]
    try:
        import pdfplumber
    except ImportError:
        print("ERRO: pdfplumber não instalado. Execute: python -m pip install pdfplumber --user")
        sys.exit(1)

    print(f"\n{'='*70}")
    print(f"  DIAGNÓSTICO DE EXTRAÇÃO PDF: {caminho}")
    print(f"{'='*70}\n")

    linhas_pdf = []
    with pdfplumber.open(caminho) as pdf:
        for n_pag, page in enumerate(pdf.pages, 1):
            texto = page.extract_text() or ''
            linhas_pagina = texto.split('\n')
            linhas_pdf.extend(linhas_pagina)
            print(f"--- PÁGINA {n_pag} ({len(linhas_pagina)} linhas) ---")
            for i, linha in enumerate(linhas_pagina):
                print(f"  [{i:03d}] {repr(linha)}")

    print(f"\n{'='*70}")
    print("  ANÁLISE AUTOMÁTICA")
    print(f"{'='*70}")

    texto_pdf = '\n'.join(linhas_pdf)

    # CNPJ
    cnpj_pattern = re.compile(r'\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}')
    cnpjs = cnpj_pattern.findall(texto_pdf)
    print(f"\nCNPJs encontrados: {cnpjs}")

    # Número 9 dígitos (DANFE)
    danfe_nums = re.findall(r'\b\d{3}\.\d{3}\.\d{3}\b', texto_pdf)
    print(f"Números 000.000.000 (DANFE): {danfe_nums}")

    # Valores R$
    valores = re.findall(r'R\$\s*([\d\.]+,\d{2})', texto_pdf)
    print(f"Valores R$ encontrados: {valores}")

    # Bloco emitente
    for i, linha in enumerate(linhas_pdf):
        if 'IDENTIFICAÇ' in linha.upper() and 'EMITENTE' in linha.upper():
            print(f"\nBloco IDENTIFICAÇÃO DO EMITENTE na linha [{i}]: {repr(linha)}")
            print("  Próximas 10 linhas:")
            for j in range(i+1, min(i+11, len(linhas_pdf))):
                print(f"    [{j}] {repr(linhas_pdf[j])}")
            break

    print(f"\n{'='*70}\n")

if __name__ == '__main__':
    main()
