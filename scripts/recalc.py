#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
recalc.py
=========
Recalcula fórmulas de um .xlsx com LibreOffice headless, preservando a logo
(diferente de reabrir e salvar com openpyxl, que remove o drawing). Ver
SKILL.md passo 6, este script é o caminho preferido para esse passo.

Uso:
  python3 recalc.py ARQUIVO.xlsx

Saída: uma linha JSON com status ("success"/"error") e total_errors (nº de
células com erro de fórmula, ex. #REF!, #VALUE!) no arquivo recalculado.
Código de saída 0 = sucesso e sem erros, 2 = falha de conversão ou erros
encontrados.
"""
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

ERRO_TOKENS = ('#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#NULL!', '#NUM!', '#N/A')


def encontrar_binario():
    """'libreoffice' nem sempre está no PATH (ex.: instalação via Homebrew no
    macOS só expõe 'soffice'), tenta os dois nomes."""
    for nome in ('libreoffice', 'soffice'):
        caminho = shutil.which(nome)
        if caminho:
            return caminho
    return None


def contar_erros(caminho):
    import openpyxl
    wb = openpyxl.load_workbook(caminho, data_only=True)
    total = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value in ERRO_TOKENS:
                    total += 1
    return total


# Tamanho da janela que o Excel abre, gravado em xl/workbook.xml. O LibreOffice
# escreve o default dele, 16384x8192 twips (1/20 de ponto), que dá uma janela de
# meia tela: a planilha abre mostrando ate a coluna C e o usuario maximiza toda
# vez. Estes valores equivalem a 2560x1440 px a 96 dpi, maiores que a tela, e o
# Excel encolhe para o que couber, ou seja, abre maximizado em qualquer monitor.
JANELA = {'windowWidth': 38400, 'windowHeight': 21600}


def normalizar_janela(caminho):
    """Reescreve a geometria da janela em xl/workbook.xml, dentro do zip.

    Feito por cirurgia no zip, e nao com openpyxl, pela mesma razao do passo 4:
    reabrir e salvar com openpyxl apagaria a logo. Devolve True se mudou algo."""
    with zipfile.ZipFile(caminho) as z:
        entradas = [(i, z.read(i.filename)) for i in z.infolist()]

    mudou = False
    for pos, (info, dados) in enumerate(entradas):
        if info.filename != 'xl/workbook.xml':
            continue
        xml = dados.decode('utf-8')
        for atributo, valor in JANELA.items():
            xml, n = re.subn(r'(<workbookView[^>]*?\b%s=")\d+(")' % atributo,
                             r'\g<1>%d\g<2>' % valor, xml, count=1)
            mudou = mudou or bool(n)
        entradas[pos] = (info, xml.encode('utf-8'))

    if not mudou:
        print('AVISO: <workbookView> sem windowWidth/windowHeight, janela nao normalizada.',
              file=sys.stderr)
        return False

    tmp = caminho + '.janela-tmp'
    with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z:
        for info, dados in entradas:
            z.writestr(info, dados)
    os.replace(tmp, caminho)
    return True


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(1)
    arquivo = os.path.abspath(sys.argv[1])
    if not os.path.isfile(arquivo):
        print(json.dumps({"status": "error", "mensagem": "arquivo não encontrado: %s" % arquivo}))
        sys.exit(2)

    binario = encontrar_binario()
    if not binario:
        print(json.dumps({"status": "error", "mensagem": "nem 'libreoffice' nem 'soffice' encontrados no PATH"}))
        sys.exit(2)

    tmp = tempfile.mkdtemp(prefix="recalc_")
    try:
        # Perfil de usuário próprio por execução: o LibreOffice recusa rodar duas
        # instâncias headless sobre o mesmo perfil, e a segunda falha calada. Como
        # o fechamento roda cliente a cliente (e às vezes duas sessões em paralelo),
        # sem isto o recalc falha por colisão e não por defeito da planilha.
        perfil = "-env:UserInstallation=file://%s" % os.path.join(tmp, "profile")
        cmd = [binario, perfil, "--headless", "--convert-to", "xlsx", "--outdir", tmp, arquivo]
        resultado = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        saida = os.path.join(tmp, os.path.basename(arquivo))
        if resultado.returncode != 0 or not os.path.isfile(saida):
            print(json.dumps({
                "status": "error",
                "mensagem": "conversão falhou",
                "stdout": resultado.stdout,
                "stderr": resultado.stderr,
            }))
            sys.exit(2)

        shutil.copyfile(saida, arquivo)
        normalizar_janela(arquivo)
        total_errors = contar_erros(arquivo)
        status = "success" if total_errors == 0 else "error"
        print(json.dumps({"status": status, "total_errors": total_errors, "arquivo": arquivo}))
        sys.exit(0 if total_errors == 0 else 2)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    main()
