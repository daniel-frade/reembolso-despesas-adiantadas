#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_relatorio.py
===================
Gera o .xlsx mensal de adiantamentos de um cliente a partir de uma exportação
CSV do ClickUp, seguindo o layout e a formatação documentados em SKILL.md:
fonte Calibri, bordas thin, cores por método de pagamento e por órgão, total
somado em fórmula, e a logo do escritório reinjetada via zipfile a partir de
um arquivo de referência (mês anterior do mesmo cliente, ou o mais recente de
outro cliente do mesmo mês).

Fonte de dados: --csv (exportação manual do ClickUp) ou --manifest (saída de
`clickup_adiantamentos_fetch.py`, a Frente A). Com --manifest, também grava um
sidecar "{saida}.itens.json" (task_id + caminhos dos arquivos baixados por
demanda, na mesma ordem/numeração das linhas do xlsx), consumido por
organizar_recibos.py pra renomear/organizar os comprovantes.

Uso:
  python3 build_relatorio.py --cliente "<cliente>" --csv export.csv \
      --mes Abril --ano 2026 --ref SuperNosso-Marco-2026.xlsx \
      --saida SuperNosso-Abril-2026.xlsx

  python3 build_relatorio.py --cliente "<cliente>" --manifest _raw/manifest.json \
      --mes Julho --ano 2026 --pasta-referencia ../06-Junho --publicar-drive

Tamanho de fonte: SKILL.md não fixa um valor numérico exato (só documenta
Calibri em tudo). Os defaults abaixo (14 no título, 12 no resto) refletem o
que já foi usado nos relatórios anteriores; confira visualmente contra o
--ref antes de considerar definitivo, ver checklist de verificação.

Saída: gera o .xlsx e imprime um resumo JSON com total, nº de itens e lista
de pendências (ex.: CASO novo) para revisão humana, não trava a geração.
"""
import argparse
import csv
import json
import os
import re
import sys
import zipfile
from datetime import datetime

from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill

THIN = Side(style='thin', color='FF000000')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CALIBRI = 'Calibri'
TITLE_SIZE = 14
HEADER_SIZE = 14  # linha 3 (cabeçalho) também é 14, ver SKILL.md > Formatação geral
BODY_SIZE = 12

COR_METODO = {
    'BOLETO': 'FF1BBC9C', 'PIX': 'FF3082B7', 'CRÉDITO': 'FF02BCD4', 'DINHEIRO': 'FFE91E63',
}
COR_ORGAO = {
    'RI': 'FF6A1B9A', 'JUCEMG': 'FF1BBC9C', 'NOTAS': 'FFFF9800', 'PREFEITURA': 'FFE65100',
    'CARTÓRIO': 'FF607D8B', 'MOTOBOY': 'FF795548', 'JUNTA': 'FF1BBC9C', 'RTD': 'FF3F51B5',
    'OUTROS': 'FF1BBC9C', 'TJ': 'FF1BBC9C', 'ONR': 'FF00796B',
}
COR_DESCONHECIDA = 'FFBDBDBD'

# B é a coluna do logo. 27.83203125 veio da largura que o Daniel ajustou à mão num
# relatório de agosto/2026 e foi replicada nos demais clientes do mês.
WIDTHS = {'B': 27.83203125, 'C': 96.0, 'D': 16.0, 'E': 30.51, 'F': 26.5, 'G': 21.83, 'H': 24.0}
ALTURA_TITULO = 66.0
ALTURA_CABECALHO = 42.0

# Geometria da logo em B2, em EMU. Sem isso a logo herda o tamanho que o arquivo
# de --ref tiver, e o de cada cliente tem o seu, todos achatados: o valor que o
# script usava antes dava razão 2,40.
#
# A altura é derivada, não medida: cx dividido por 3,5597, a razão do PNG
# original de 566x159, o que zera o achatamento. O resto sobre a linha 2, que tem
# 66 pt = 838200 EMU, é dividido em dois para centralizar na vertical.
#
# A largura NÃO é a que o Daniel usou à mão (1948780): aquela âncora termina em
# 2044700, fora da coluna B, e o recalc do LibreOffice corta o excesso, devolvendo
# razão 3,41. Aqui ela termina em 1940000, dentro da coluna, e sobrevive ao passo 5
# do pipeline. Se a largura da coluna B mudar, este limite muda junto.
LOGO_ANCORA = {
    'from_col_off': 95920, 'from_row_off': 160081,
    'to_col_off': 1940000, 'to_row_off': 678118,
    'cx': 1844080, 'cy': 518037,
}


def normalizar_ancora_logo(drawing_xml):
    """Reescreve posição e tamanho da logo no drawing1.xml para LOGO_ANCORA.

    O drawing vem copiado inteiro do arquivo de referência, então sem esta
    normalização cada cliente carrega para sempre a proporção que a imagem
    tinha no mês anterior dele."""
    xml = drawing_xml.decode('utf-8')
    pares = [
        (r'(<xdr:from>.*?<xdr:colOff>)\d+(</xdr:colOff>)', LOGO_ANCORA['from_col_off']),
        (r'(<xdr:from>.*?<xdr:rowOff>)\d+(</xdr:rowOff>)', LOGO_ANCORA['from_row_off']),
        (r'(<xdr:to>.*?<xdr:colOff>)\d+(</xdr:colOff>)', LOGO_ANCORA['to_col_off']),
        (r'(<xdr:to>.*?<xdr:rowOff>)\d+(</xdr:rowOff>)', LOGO_ANCORA['to_row_off']),
    ]
    for padrao, valor in pares:
        xml, n = re.subn(padrao, r'\g<1>%d\g<2>' % valor, xml, count=1, flags=re.S)
        if not n:
            print('AVISO: âncora da logo não normalizada, padrão ausente no drawing1.xml.',
                  file=sys.stderr)
            return drawing_xml
    xml, n = re.subn(r'<a:ext cx="\d+" cy="\d+"/>',
                     '<a:ext cx="%d" cy="%d"/>' % (LOGO_ANCORA['cx'], LOGO_ANCORA['cy']),
                     xml, count=1)
    if not n:
        print('AVISO: tamanho da logo não normalizado, <a:ext> ausente no drawing1.xml.',
              file=sys.stderr)
        return drawing_xml
    return xml.encode('utf-8')


def carregar_config(caminho, cliente):
    with open(caminho, encoding='utf-8') as f:
        cfg = json.load(f)
    cfg.pop('_comentario', None)
    if cliente not in cfg:
        sys.exit('ERRO: cliente "%s" não está em %s. Clientes conhecidos: %s'
                  % (cliente, caminho, ', '.join(sorted(cfg))))
    return cfg[cliente]


def parse_data(valor):
    valor = (valor or '').strip()
    if not valor:
        return None
    for fmt in ('%m/%d/%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(valor, fmt).date()
        except ValueError:
            continue
    return None


def parse_valor(valor):
    if isinstance(valor, (int, float)):
        return float(valor)
    valor = (valor or '').strip().replace('R$', '').replace(' ', '')
    if not valor:
        return 0.0
    if ',' in valor and '.' in valor:
        valor = valor.replace('.', '').replace(',', '.')
    elif ',' in valor:
        valor = valor.replace(',', '.')
    try:
        return float(valor)
    except ValueError:
        return 0.0


def ler_linhas_xlsx(xlsx_path, cliente_cfg):
    """Lê um .xlsx JÁ ENTREGUE (mês anterior) e devolve as linhas no mesmo
    formato de dict das outras fontes, para cobrança consolidada: quando o
    cliente não reembolsou um mês, o mês seguinte reapresenta os dois juntos
    (ver SKILL.md > Relatório consolidado de cobrança não paga).

    A data volta em ISO de propósito. No .xlsx entregue ela está como
    dd/mm/aaaa, e parse_data() tenta %m/%d/%Y ANTES de %d/%m/%Y (a ordem certa
    para o CSV do ClickUp, que exporta no formato americano) -- "08/06/2026"
    viraria 6 de agosto em vez de 8 de junho, e a consolidação sairia com a
    ordem cronológica e a numeração trocadas.
    """
    wb = load_workbook(xlsx_path)
    ws = wb.active
    colunas_extra = cliente_cfg.get('colunas_extra', [])

    linhas = []
    for r in range(4, ws.max_row + 1):
        valor = ws.cell(r, 4).value
        if valor is None:
            continue
        bruto = ws.cell(r, 6).value
        if isinstance(bruto, datetime):
            data = bruto.date().isoformat()
        elif hasattr(bruto, 'isoformat'):
            data = bruto.isoformat()
        else:
            data = ''
            for fmt in ('%d/%m/%Y', '%Y-%m-%d'):
                try:
                    data = datetime.strptime(str(bruto).strip(), fmt).date().isoformat()
                    break
                except ValueError:
                    continue
            if not data:
                print('AVISO: data ilegível na linha %d de %s: %r'
                      % (r, xlsx_path, bruto), file=sys.stderr)

        linha = {
            'CLIENTES (drop down)': '',
            'Task Name': str(ws.cell(r, 2).value or '').strip(),
            'Task Content': str(ws.cell(r, 3).value or '').strip(),
            'VALOR (currency)': valor,
            'Metodo de pagamento (drop down)': str(ws.cell(r, 5).value or '').strip(),
            'Payment Date (date)': data,
            'ÓRGÃO (drop down)': str(ws.cell(r, 7).value or '').strip(),
            'Status': 'approved',
        }
        h = ws.cell(r, 8).value
        if h and 'Empresa' in colunas_extra:
            # nome completo da PJ: empresa_por_cliente_clickup não tem essa chave,
            # então montar_dados repassa o próprio valor, que já é o final
            linha['CLIENTES (drop down)'] = str(h).strip()
        elif h and 'CASO' in colunas_extra:
            linha['_caso_pronto'] = str(h).strip()
        linhas.append(linha)

    if not linhas:
        print('AVISO: nenhuma linha lida de %s' % xlsx_path, file=sys.stderr)
    return linhas


# Ciclo de vida da tarefa no ClickUp, levantado nas listas de abril, junho e
# julho/2026: to do -> approved -> solicitado reembolso -> pago/reembolsado
# pelo cliente. "approved" é estado TRANSITÓRIO: filtrar por ele fotografa o
# mês num instante e perde tudo que já avançou (junho/2026 não tinha uma única
# tarefa em approved, e o relatório saía vazio). O que separa o que entra do
# que não entra é outra coisa: "to do" é o que o escritório ainda não pagou.
STATUS_NAO_PAGO = {'to do'}


def status_entra(status, status_ok):
    status = (status or '').strip().lower()
    if status_ok:                       # --status explícito manda
        return status in status_ok
    return status not in STATUS_NAO_PAGO


def ler_linhas(csv_path, nomes_clickup, status_ok):
    linhas = []
    with open(csv_path, encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            cliente = (row.get('CLIENTES (drop down)') or '').strip()
            if cliente not in nomes_clickup:
                continue
            if not status_entra(row.get('Status'), status_ok):
                continue
            linhas.append(row)
    if not linhas:
        print('AVISO: nenhuma linha encontrada para %s (status=%s) em %s'
              % (nomes_clickup, status_ok, csv_path), file=sys.stderr)
    return linhas


def ler_linhas_manifest(manifest_path, nomes_clickup, status_ok):
    """Lê o manifest.json da Frente A (clickup_adiantamentos_fetch.py) e
    devolve linhas no mesmo formato de dict que ler_linhas() produz pro CSV,
    mais os campos internos (_task_id, _arquivos_*) usados pra montar o
    sidecar itens.json em main()."""
    with open(manifest_path, encoding='utf-8') as f:
        entradas = json.load(f)
    linhas = []
    for e in entradas:
        cliente = (e.get('cliente') or '').strip()
        if cliente not in nomes_clickup:
            continue
        if not status_entra(e.get('status'), status_ok):
            continue
        linhas.append({
            'CLIENTES (drop down)': cliente,
            'Task Name': e.get('task_name') or '',
            'Task Content': e.get('task_content') or '',
            'VALOR (currency)': e.get('valor'),
            'Metodo de pagamento (drop down)': e.get('metodo_pagamento') or '',
            'Payment Date (date)': e.get('data_pagamento') or '',
            'ÓRGÃO (drop down)': e.get('orgao') or '',
            'Status': e.get('status') or '',
            '_task_id': e.get('task_id'),
            '_arquivos_invoice_photo': e.get('arquivos_invoice_photo') or [],
            '_arquivos_comentarios': e.get('arquivos_comentarios') or [],
            # Arquivos da bandeja da tarefa que nao vieram nem pelo formulario
            # nem por resposta de comentario: existem, mas ninguem declarou o
            # papel deles. Seguem adiante para o organizar_recibos mandar a
            # demanda inteira para revisao em vez de fingir que nao ha nada.
            '_arquivos_tarefa': e.get('arquivos_tarefa') or [],
        })
    if not linhas:
        print('AVISO: nenhuma linha encontrada para %s (status=%s) em %s'
              % (nomes_clickup, status_ok, manifest_path), file=sys.stderr)
    return linhas


_MARCADORES_BANCARIOS = (
    'dados bancários:', 'dados bancarios:', 'dados da conta', 'dados:',
    'chave pix', 'pix copia e cola', 'pix copia-e-cola', 'copia e cola',
    '00020101',   # início do payload EMV do Pix copia-e-cola, colado cru na descrição
)


def limpar_descricao(texto):
    """Corta da descrição o bloco de dados bancários do cartório/terceiro.

    O Task Content do ClickUp serve a quem PAGA, e por isso vem com agência,
    conta e chave Pix do favorecido no fim. Isso não é prestação de contas: o
    cliente recebe a descrição do serviço, não o roteiro de pagamento. O corte
    é sempre do marcador até o fim do texto, e devolve o trecho removido para
    virar pendência auditável -- nada some em silêncio.
    """
    if not texto:
        return texto, None
    minusculo = texto.lower()
    corte = None
    for marcador in _MARCADORES_BANCARIOS:
        i = minusculo.find('\n' + marcador)
        if i == -1 and minusculo.startswith(marcador):
            i = 0
        if i != -1 and (corte is None or i < corte):
            corte = i
    if corte is None or corte == 0:
        return texto, None
    return texto[:corte].rstrip(), texto[corte:].strip()


def montar_dados(cliente_cfg, linhas, casos_conhecidos):
    colunas_extra = cliente_cfg.get('colunas_extra', [])
    empresa_map = cliente_cfg.get('empresa_por_cliente_clickup', {})

    linhas_ordenadas = sorted(
        linhas, key=lambda r: parse_data(r.get('Payment Date (date)')) or datetime.min.date())

    dados, pendencias = [], []
    for i, row in enumerate(linhas_ordenadas, start=1):
        # Coluna B é SEMPRE o Task Name do ClickUp, em todo cliente, sem exceção
        # (ver SKILL.md > Passo 1). Numeração sequencial não é aceita.
        task_name = (row.get('Task Name') or '').strip()
        if not task_name:
            pendencias.append('Linha %d sem Task Name: coluna B ficaria vazia (valor %s, data %s)'
                              % (i, row.get('VALOR (currency)'), row.get('Payment Date (date)')))
        descricao, bloco_removido = limpar_descricao((row.get('Task Content') or '').strip())
        if bloco_removido:
            pendencias.append('Dados bancários removidos da descrição de "%s": %s'
                              % (task_name or '(sem nome)', ' / '.join(bloco_removido.split('\n'))[:110]))
        item = {
            'B': task_name,
            'C': descricao,
            'D': parse_valor(row.get('VALOR (currency)')),
            'E': (row.get('Metodo de pagamento (drop down)') or '').strip().upper(),
            'F': parse_data(row.get('Payment Date (date)')),
            'G': (row.get('ÓRGÃO (drop down)') or '').strip().upper(),
        }
        if 'Empresa' in colunas_extra:
            cliente_raw = (row.get('CLIENTES (drop down)') or '').strip()
            item['H'] = empresa_map.get(cliente_raw, cliente_raw)
        elif 'CASO' in colunas_extra:
            # A chave do mapa é o task_id, não o nome: o mesmo nome de tarefa
            # aparece em casos diferentes ("Certidão de matrícula" sai tanto no
            # imóvel de Cotia quanto na execução fiscal). Nome só é aceito como
            # fallback de mapa antigo.
            mapa = casos_conhecidos or {}
            caso = (row.get('_caso_pronto')
                    or mapa.get(row.get('_task_id'))
                    or mapa.get(task_name))
            if isinstance(caso, dict):
                caso = caso.get('caso')
            if caso:
                item['H'] = caso
            else:
                descricao = (row.get('Task Content') or task_name)[:40]
                item['H'] = 'REVISAR: %s' % descricao
                pendencias.append('CASO novo em "%s": %s' % (item['B'], descricao))
        if row.get('_task_id'):
            item['_task_id'] = row['_task_id']
            item['_arquivos_invoice_photo'] = row.get('_arquivos_invoice_photo', [])
            item['_arquivos_comentarios'] = row.get('_arquivos_comentarios', [])
            item['_arquivos_tarefa'] = row.get('_arquivos_tarefa', [])
        dados.append(item)
    return dados, pendencias


def cor_para(valor, tabela):
    return tabela.get((valor or '').strip().upper(), COR_DESCONHECIDA)


def montar_planilha(dados, titulo, rotulo_h, arquivo_referencia=None):
    """Se `arquivo_referencia` for dado, parte dele (não de um Workbook() vazio):
    é o que faz o objeto de imagem sobreviver ao save() do openpyxl com toda a
    parafernália (drawing1.xml, rels, Content_Types) intacta, ainda que com
    posição/tamanho imprecisos, o que injetar_logo() corrige depois. Um
    Workbook() vazio nunca teve imagem, então save() não gera esses arquivos."""
    if arquivo_referencia:
        wb = load_workbook(arquivo_referencia)
        ws = wb.active
        for rng in list(ws.merged_cells.ranges):
            ws.unmerge_cells(str(rng))
    else:
        wb = Workbook()
        ws = wb.active
    ws.title = 'Tasks'

    ultima_linha = 3 + len(dados)
    cols = ['B', 'C', 'D', 'E', 'F', 'G'] + (['H'] if rotulo_h else [])

    if ws.max_row > ultima_linha:
        ws.delete_rows(ultima_linha + 1, ws.max_row - ultima_linha)
    for row in ws.iter_rows(min_row=1, max_row=max(ws.max_row, ultima_linha)):
        for cell in row:
            if isinstance(cell, MergedCell):
                continue
            cell.value = None

    for letra in cols:
        ws.column_dimensions[letra].width = WIDTHS[letra]

    # Linha 2: logo (inserida depois via zipfile) + título + total.
    # A linha 2 tem que terminar exatamente onde a tabela termina: o total vai
    # SEMPRE na última coluna e o rótulo na penúltima, e o título mescla de C
    # até a coluna anterior a essas duas. Com a coluna H (Empresa/CASO), fixar
    # o total em G deixava H2 vazia e sem borda, e a linha 2 saía "quebrada",
    # mais curta que as linhas de dados.
    col_total = cols[-1]
    col_rotulo = cols[-2]
    col_fim_titulo = cols[-3]
    ws.merge_cells('C2:%s2' % col_fim_titulo)
    ws['C2'] = titulo
    ws['%s2' % col_rotulo] = 'Valor total:'
    ws['%s2' % col_total] = '=SUM(D4:D%d)' % ultima_linha
    ws.row_dimensions[2].height = ALTURA_TITULO
    for letra in cols:
        cell = ws['%s2' % letra]
        # estilo vai também nas MergedCell (só `.value` é somente-leitura nelas):
        # numa faixa mesclada o Excel desenha a borda a partir de cada célula
        # de baixo, então pular as mescladas deixaria a caixa do título aberta
        # à direita. Borda explícita aqui porque herdar do --ref deixava sem
        # borda justamente a última coluna, que lá estava dentro de outro merge,
        # e a linha 2 saía truncada em relação às linhas de dados.
        cell.font = Font(name=CALIBRI, size=TITLE_SIZE, bold=True)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Linha 3: cabeçalho
    # H é a coluna extra do cliente e o rótulo muda com ela: "Empresa" no
    # um cliente, "CASO" noutro. Vem de colunas_extra
    # em clientes_config.json, nunca fixo aqui.
    cabecalhos = {'B': 'Demanda', 'C': 'Descrição', 'D': 'VALOR',
                  'E': 'Metodo de pagamento', 'F': 'Data de pagamento', 'G': 'ÓRGÃO',
                  'H': rotulo_h}
    ws.row_dimensions[3].height = ALTURA_CABECALHO
    for letra in cols:
        cell = ws['%s3' % letra]
        cell.value = cabecalhos[letra]
        cell.font = Font(name=CALIBRI, size=HEADER_SIZE, bold=True)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Linhas 4+: dados
    for idx, item in enumerate(dados):
        r = 4 + idx
        for letra in cols:
            cell = ws['%s%d' % (letra, r)]
            cell.value = item.get(letra)
            cell.font = Font(name=CALIBRI, size=BODY_SIZE)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        ws['D%d' % r].number_format = 'R$ #,##0.00'
        if item.get('F'):
            ws['F%d' % r].number_format = 'DD/MM/YYYY'

        cor_e = cor_para(item['E'], COR_METODO)
        ws['E%d' % r].fill = PatternFill(start_color=cor_e, end_color=cor_e, fill_type='solid')
        ws['E%d' % r].font = Font(name=CALIBRI, size=BODY_SIZE, color='FFFFFFFF')

        cor_g = cor_para(item['G'], COR_ORGAO)
        ws['G%d' % r].fill = PatternFill(start_color=cor_g, end_color=cor_g, fill_type='solid')
        ws['G%d' % r].font = Font(name=CALIBRI, size=BODY_SIZE, color='FFFFFFFF')

        ws.row_dimensions[r].height = None

    return wb


def injetar_logo(arquivo_referencia, arquivo_alvo):
    """Copia o XML de drawing + imagem do arquivo de referência para o alvo e
    religa manualmente sheet1.xml / [Content_Types].xml / rels, método
    documentado em SKILL.md linhas 136-172.

    Importante (achado ao validar contra dados reais): o pressuposto de que
    `load_workbook(referencia)` + `save()` preservaria o drawing (comentado
    antes aqui) é falso, testei contra um arquivo já entregue e confirmado
    correto, e mesmo esse mostra 0 imagens no modelo de objeto do openpyxl
    (`ws._images`); o save() do openpyxl sempre descarta drawings que ele não
    representa como Image() nativo, venha ou não de uma referência. Por isso
    este método reescreve os 4 arquivos (drawing1.xml, seu rels, o rels do
    sheet1 e a imagem) incondicionalmente, sem assumir que algo sobreviveu."""
    with zipfile.ZipFile(arquivo_referencia, 'r') as z:
        nomes = z.namelist()
        if 'xl/drawings/drawing1.xml' not in nomes:
            print('AVISO: referência sem logo (xl/drawings/drawing1.xml ausente), pulando injeção.',
                  file=sys.stderr)
            return False
        drawing_xml = normalizar_ancora_logo(z.read('xl/drawings/drawing1.xml'))
        drawing_rels = z.read('xl/drawings/_rels/drawing1.xml.rels')
        img_name = next(n for n in nomes if n.startswith('xl/media/image'))
        img_bytes = z.read(img_name)

    sheet_rels = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        b'<Relationship Id="rId1" '
        b'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
        b'Target="../drawings/drawing1.xml"/></Relationships>'
    )

    tmp = arquivo_alvo + '.tmp'
    with zipfile.ZipFile(arquivo_alvo, 'r') as z_in, \
            zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as z_out:
        for item in z_in.infolist():
            name = item.filename
            if name == 'xl/worksheets/sheet1.xml':
                xml = z_in.read(name).decode('utf-8')
                if 'xmlns:r=' not in xml:
                    xml = xml.replace(
                        '<worksheet ',
                        '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
                        1)
                if '<drawing ' not in xml:
                    xml = xml.replace('</worksheet>', '<drawing r:id="rId1"/></worksheet>')
                z_out.writestr(item, xml)
            elif name == '[Content_Types].xml':
                xml = z_in.read(name).decode('utf-8')
                if 'Extension="png"' not in xml:
                    xml = xml.replace('</Types>', '<Default Extension="png" ContentType="image/png"/></Types>')
                if 'drawing1.xml' not in xml:
                    xml = xml.replace(
                        '</Types>',
                        '<Override PartName="/xl/drawings/drawing1.xml" '
                        'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
                z_out.writestr(item, xml)
            elif name in ('xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels',
                          'xl/worksheets/_rels/sheet1.xml.rels', 'xl/media/image1.png'):
                continue  # regravados abaixo, incondicionalmente
            else:
                z_out.writestr(item, z_in.read(name))

        z_out.writestr('xl/drawings/drawing1.xml', drawing_xml)
        z_out.writestr('xl/drawings/_rels/drawing1.xml.rels', drawing_rels)
        z_out.writestr('xl/worksheets/_rels/sheet1.xml.rels', sheet_rels)
        z_out.writestr('xl/media/image1.png', img_bytes)

    os.replace(tmp, arquivo_alvo)
    # Não reabrir com openpyxl + save() aqui: SKILL.md (seção "Corrigir valor
    # após geração") documenta que isso remove o drawing recém-reinjetado.
    return True


def _normalizar_nome(s):
    return ''.join(ch.lower() for ch in s if ch.isalnum())


def descobrir_referencia(pasta, nome_entrega):
    """Acha o .xlsx do cliente numa pasta (mês anterior local, ou a pasta do
    Drive já resolvida por --pasta-referencia). Compara nomes normalizados
    (sem espaço/hífen/caixa) porque a convenção varia entre onde o arquivo
    foi salvo -- ex.: 'NomeDoCliente...' localmente vs.
    'Nome-Do-Cliente-...' na pasta oficial do Drive (achado ao validar
    contra dados reais, um prefixo literal não batia com o outro)."""
    if not pasta or not os.path.isdir(pasta):
        return None
    alvo = _normalizar_nome(nome_entrega)
    candidatos = [os.path.join(pasta, f) for f in os.listdir(pasta)
                  if f.endswith('.xlsx') and _normalizar_nome(os.path.splitext(f)[0]).startswith(alvo)]
    if not candidatos:
        return None
    return max(candidatos, key=os.path.getmtime)


_MESES_NUM = {
    'janeiro': '01', 'fevereiro': '02', 'março': '03', 'marco': '03', 'abril': '04',
    'maio': '05', 'junho': '06', 'julho': '07', 'agosto': '08', 'setembro': '09',
    'outubro': '10', 'novembro': '11', 'dezembro': '12',
}


_MESES_NOME = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
               'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']


def meses_cobertos(dados):
    """Meses distintos presentes nos dados, em ordem cronológica, como
    [(nome_do_mes, ano), ...]. É o que define o período real da cobrança
    quando ela consolida mais de um mês em aberto."""
    vistos = []
    for item in sorted(dados, key=lambda d: d['F'] or datetime.min.date()):
        data = item.get('F')
        if not data:
            continue
        chave = (_MESES_NOME[data.month - 1], data.year)
        if chave not in vistos:
            vistos.append(chave)
    return vistos


def rotular_periodo(meses, para_arquivo=False):
    """Descreve o período coberto. Cobrança consolidada tem que se declarar
    tanto no título (célula C2) quanto no NOME DO ARQUIVO -- entregar dois
    meses num arquivo chamado só "Julho" esconde do cliente o que ele está
    pagando (ver SKILL.md > Relatório consolidado de cobrança não paga).

      1 mês:            "julho de 2026"          / "Julho-2026"
      2 meses:          "junho e julho de 2026"  / "Junho-e-Julho-2026"
      3 ou mais:        "maio a julho de 2026"   / "Maio-a-Julho-2026"

    O rótulo NÃO começa com "de": quem consome é o título, que já escreve
    "Reembolsos de %s". Devolver "de maio a julho" produzia "Reembolsos de de
    maio a julho de 2026", erro que ficou latente porque o único relatório de
    3+ meses entregue até agosto/2026 usou --titulo manual.
    """
    if not meses:
        return None
    anos = {ano for _, ano in meses}
    sufixo_ano = str(max(anos))

    def nome(m):
        return m.capitalize() if para_arquivo else m

    if len(meses) == 1:
        corpo = nome(meses[0][0])
    elif len(meses) == 2:
        lig = '-e-' if para_arquivo else ' e '
        corpo = lig.join(nome(m) for m, _ in meses)
    else:
        lig = '-a-' if para_arquivo else ' a '
        corpo = lig.join([nome(meses[0][0]), nome(meses[-1][0])])

    if len(anos) > 1:
        # meses de anos diferentes: carimba o ano em cada ponta
        pontas = ['%s%s%s' % (nome(m), '-' if para_arquivo else ' de ', a) for m, a in (meses[0], meses[-1])]
        lig = ('-a-' if para_arquivo else ' a ')
        return lig.join(pontas)
    return '%s%s%s' % (corpo, '-' if para_arquivo else ' de ', sufixo_ano)


def resolver_drive_root(explicito=None):
    """Acha a raiz do mount local do Google Drive Desktop (Shared drives),
    ex.: ~/Library/CloudStorage/GoogleDrive-fulano@empresa.com/Shared drives.
    Ver estudo em workspace/nlr/clickup/output/estudo-clickup-google-drive.md, é
    um mount de arquivo normal (DriveFS), escrita/rename/delete comuns
    funcionam sem nenhuma integração especial com a API do ClickUp/Drive."""
    if explicito:
        return explicito
    import glob
    candidatos = glob.glob(os.path.expanduser('~/Library/CloudStorage/GoogleDrive-*/Shared drives'))
    if not candidatos:
        return None
    # O Drive Desktop pode deixar mounts antigos/duplicados pra tras (ex. apos
    # reconectar a conta), com um sufixo tipo " (26-05-26 16:32)" no nome da
    # pasta da conta -- confirmado neste workspace: existem os dois ao mesmo
    # tempo, e so o sem sufixo e o mount realmente sincronizado (achado ao
    # validar a publicacao no Drive contra dados reais de junho/2026, onde
    # glob() sem esse filtro escolhia o mount errado e nunca achava a pasta
    # de referencia do mes anterior). Preferir o(s) sem parenteses no nome.
    canonicos = [c for c in candidatos if '(' not in os.path.basename(os.path.dirname(c))]
    return sorted(canonicos or candidatos)[0]


def publicar_no_drive(arquivo_local, cliente_cfg, mes, ano, drive_root=None):
    """Copia o .xlsx já pronto (com logo) pro caminho oficial do cliente no
    Google Drive, resolvido a partir de `destino_drive` em clientes_config.json.
    Só copia, não reescreve o arquivo já sincronizado várias vezes; o
    arquivo é montado inteiro localmente antes (ver main()) e só entra na
    pasta sincronizada numa única cópia atômica no fim."""
    template = cliente_cfg.get('destino_drive')
    if not template:
        return None, 'destino_drive não configurado pra este cliente em clientes_config.json'

    root = resolver_drive_root(drive_root)
    if not root or not os.path.isdir(root):
        return None, 'raiz do Google Drive não encontrada (Drive Desktop instalado/sincronizado?)'

    mm = _MESES_NUM.get(mes.strip().lower(), '00')
    caminho_relativo = template.format(ano=ano, mm=mm, mes_nome=mes)
    destino_dir = os.path.join(root, caminho_relativo)
    os.makedirs(destino_dir, exist_ok=True)
    destino_arquivo = os.path.join(destino_dir, os.path.basename(arquivo_local))

    import shutil
    shutil.copy2(arquivo_local, destino_arquivo)
    return destino_arquivo, None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--cliente', required=True, help='Nome do cliente em clientes_config.json')
    fonte = ap.add_mutually_exclusive_group(required=True)
    fonte.add_argument('--csv', help='CSV exportado do ClickUp')
    fonte.add_argument('--manifest', help='manifest.json gerado por clickup_adiantamentos_fetch.py (Frente A)')
    ap.add_argument('--incluir-xlsx', action='append', default=[],
                    help='.xlsx já entregue de outro mês, cujas linhas entram nesta cobrança '
                         '(cobrança consolidada; pode repetir). Tudo é reordenado por data e '
                         'renumerado 1..N. Use --titulo para declarar o período coberto.')
    ap.add_argument('--mes', required=True, help='Ex: Abril')
    ap.add_argument('--ano', required=True, help='Ex: 2026')
    ap.add_argument('--config', default=os.path.join(os.path.dirname(os.path.abspath(__file__)), 'clientes_config.json'))
    ap.add_argument('--ref', help='Arquivo .xlsx de referência para a logo (mês anterior do cliente)')
    ap.add_argument('--pasta-referencia', help='Pasta onde procurar automaticamente o .xlsx do mês anterior')
    ap.add_argument('--saida', help='Caminho do .xlsx de saída (default: {NomeEntrega}-{Mês}-{Ano}.xlsx)')
    ap.add_argument('--titulo', help='Título da linha 2 (default: "{NomeEntrega} - Reembolsos de {Mês} de {Ano}")')
    ap.add_argument('--status', default=None,
                    help='Status aceitos, separados por vírgula. Sem isto, entra tudo que não '
                         'esteja em "to do" (= o escritório ainda não pagou), que é o critério '
                         'certo: "approved" é estado transitório e some conforme o mês avança.')
    ap.add_argument('--casos-conhecidos',
                    help='JSON opcional {task_id: caso} para clientes com coluna CASO; o valor '
                         'pode ser a string do caso ou {"caso": ..., "evidencia": ...}')
    ap.add_argument('--sem-logo', action='store_true', help='Não injeta logo (uso em teste, sem --ref)')
    ap.add_argument('--publicar-drive', action='store_true',
                     help='Depois de gerar, copia o .xlsx pra pasta oficial do cliente no Google Drive '
                          '(destino_drive em clientes_config.json), ver workspace/nlr/clickup/output/estudo-clickup-google-drive.md')
    ap.add_argument('--drive-root', help='Raiz do Google Drive (default: autodetectar ~/Library/CloudStorage/GoogleDrive-*/Shared drives)')
    args = ap.parse_args()

    cliente_cfg = carregar_config(args.config, args.cliente)
    nome_entrega = cliente_cfg.get('nome_entrega', args.cliente)
    extras = cliente_cfg.get('colunas_extra') or []
    rotulo_h = extras[0] if extras else None

    status_ok = {s.strip().lower() for s in args.status.split(',')} if args.status else None
    if args.csv:
        linhas = ler_linhas(args.csv, cliente_cfg.get('nome_clickup', [args.cliente]), status_ok)
    else:
        linhas = ler_linhas_manifest(args.manifest, cliente_cfg.get('nome_clickup', [args.cliente]), status_ok)

    for extra in args.incluir_xlsx:
        linhas.extend(ler_linhas_xlsx(extra, cliente_cfg))

    casos_conhecidos = None
    if args.casos_conhecidos:
        with open(args.casos_conhecidos, encoding='utf-8') as f:
            casos_conhecidos = json.load(f)

    dados, pendencias = montar_dados(cliente_cfg, linhas, casos_conhecidos)

    referencia = None
    if not args.sem_logo:
        referencia = args.ref or descobrir_referencia(args.pasta_referencia, nome_entrega)
        if not referencia:
            pendencias.append('Logo não injetada: nenhum --ref e nenhum template encontrado em --pasta-referencia.')

    # Título e nome do arquivo saem do período REAL dos dados, não do --mes:
    # numa cobrança consolidada o --mes é só o mês de entrega, e os dois
    # precisam declarar todos os meses que estão sendo cobrados.
    # Cobrança consolidada é um ATO EXPLÍCITO (--incluir-xlsx), não uma inferência
    # a partir das datas. Um mês normal pode conter uma despesa antiga: um cliente, em
    # junho/2026, tinha um estacionamento pago em 22/05, e rotular pelas datas
    # batizou o arquivo de "Maio-e-Junho", como se fosse fatura de dois meses.
    meses = meses_cobertos(dados)
    if args.incluir_xlsx:
        periodo_titulo = rotular_periodo(meses) or '%s de %s' % (args.mes.lower(), args.ano)
        periodo_arquivo = rotular_periodo(meses, para_arquivo=True) or '%s-%s' % (args.mes, args.ano)
        if len(meses) > 1:
            pendencias.append('Cobrança consolidada: %d meses (%s). Confirme com o cliente que '
                              'o mês em aberto entra nesta fatura.'
                              % (len(meses), ', '.join('%s/%s' % (m, a) for m, a in meses)))
    else:
        periodo_titulo = '%s de %s' % (args.mes.lower(), args.ano)
        periodo_arquivo = '%s-%s' % (args.mes, args.ano)
        fora = [m for m in meses if m[0].lower() != args.mes.strip().lower()]
        if fora:
            pendencias.append('Despesa(s) de outro mês dentro da lista de %s: %s. O relatório '
                              'segue sendo de %s; confira se a data no ClickUp está certa.'
                              % (args.mes, ', '.join('%s/%s' % (m, a) for m, a in fora), args.mes))

    titulo = args.titulo or '%s - Reembolsos de %s' % (nome_entrega, periodo_titulo)
    wb = montar_planilha(dados, titulo, rotulo_h, arquivo_referencia=referencia)

    # Espaco vira hifen, nao some: 'Alfa e Beta' sai 'Alfa-e-Beta'. O script colava
    # as palavras e o Daniel renomeou a mao o arquivo de agosto/2026; o hifen e o
    # padrao que os meses anteriores daquele cliente ja usavam.
    saida = args.saida or '%s-%s.xlsx' % (nome_entrega.replace(' ', '-'), periodo_arquivo)
    os.makedirs(os.path.dirname(os.path.abspath(saida)) or '.', exist_ok=True)
    wb.save(saida)

    logo_ok = None
    if referencia:
        logo_ok = injetar_logo(referencia, saida)
        if not logo_ok:
            pendencias.append('Logo não injetada: falha ao processar --ref "%s" (ver stderr).' % referencia)

    itens_path = None
    itens_com_arquivos = [d for d in dados if d.get('_task_id')]
    if itens_com_arquivos:
        itens_path = saida + '.itens.json'
        base_sidecar = os.path.dirname(os.path.abspath(itens_path))

        def relativo(caminho):
            """Caminho relativo ao proprio sidecar, para a frente inteira poder
            mudar de lugar sem invalidar as referencias. Absoluto quebrou em
            10/09/2026, quando input/ e output/ foram reorganizados e os 54
            caminhos de tres sidecars tiveram que ser reescritos a mao."""
            try:
                return os.path.relpath(caminho, base_sidecar)
            except ValueError:
                return caminho  # volumes diferentes no Windows, mantem absoluto

        with open(itens_path, 'w', encoding='utf-8') as f:
            json.dump([
                {
                    'demanda': i + 1,
                    'task_id': d['_task_id'],
                    'arquivos_invoice_photo': [relativo(c) for c in d.get('_arquivos_invoice_photo', [])],
                    'arquivos_comentarios': [relativo(c) for c in d.get('_arquivos_comentarios', [])],
                    'arquivos_tarefa': [relativo(c) for c in d.get('_arquivos_tarefa', [])],
                }
                for i, d in enumerate(dados) if d.get('_task_id')
            ], f, ensure_ascii=False, indent=2)

    publicado_em, erro_drive = None, None
    if args.publicar_drive:
        publicado_em, erro_drive = publicar_no_drive(saida, cliente_cfg, args.mes, args.ano, args.drive_root)
        if erro_drive:
            pendencias.append('Publicação no Drive falhou: %s' % erro_drive)

    total = sum(item['D'] for item in dados)
    resumo = {
        'arquivo': saida,
        'itens_json': itens_path,
        'publicado_em': publicado_em,
        'cliente': args.cliente,
        'itens': len(dados),
        'total': round(total, 2),
        'logo_injetada': logo_ok,
        'pendencias': pendencias,
    }
    print(json.dumps(resumo, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
