#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
renumerar_recibos.py
====================
Recompacta a numeração XX.YY de uma pasta "Recibos e comprovantes" depois
que o Daniel poda ou acrescenta arquivos na mão.

Por que existe: organizar_recibos.py numera a partir do que o ClickUp
entregou, mas parte do que vem de lá não é documento de entrega (o
"brink" -- pedido de compra de crédito do ONR/JUCEMG, print de tela com
dados bancários, QR Pix). Quem decide o que sai é o Daniel, olhando
arquivo por arquivo. Depois da poda sobram buracos (05.02, 05.03 sem
05.01) e às vezes arquivos novos com o nome original do banco/cartório.

O que este script faz:
  - agrupa os arquivos já nomeados por demanda (prefixo XX)
  - recompacta cada demanda para .01, .02, .03... preservando a ordem
    relativa que já estava lá (que é a ordem revisada: documento/recibo
    primeiro, comprovante de pagamento por último)
  - lista os arquivos que ainda não seguem o padrão, para atribuição
    manual via --atribuir

O que ele NÃO faz: decidir de qual demanda é um arquivo solto. Isso exige
ler o conteúdo (valor, data, favorecido) e continua sendo trabalho de
quem revisa -- ver SKILL.md > Identificação e renomeação dos comprovantes.

Uso:
  # 1) ver o que seria feito, sem tocar em nada
  python3 renumerar_recibos.py --pasta ".../<Cliente>/Recibos e comprovantes" --dry-run

  # 2) atribuir os arquivos soltos e recompactar de uma vez
  python3 renumerar_recibos.py --pasta ".../<Cliente>/Recibos e comprovantes" \
      --atribuir "Recibo_e-protocolo_20260708_0000000000.pdf=01" \
      --atribuir "Imagem JPEG-0000-0000-00-0.jpeg=01"

Arquivos soltos atribuídos à mesma demanda entram na ordem em que foram
passados na linha de comando, depois dos que já estavam numerados.
"""
import argparse
import os
import re
import sys

PADRAO = re.compile(r'^(\d{2})\.(\d{2})(\.[A-Za-z0-9]+)$')
IGNORAR = {'.DS_Store', 'Thumbs.db'}


def levantar(pasta, atribuicoes):
    """Devolve (por_demanda, soltos). por_demanda: {'01': [nome, ...]} na ordem final."""
    nomes = sorted(n for n in os.listdir(pasta)
                   if n not in IGNORAR and os.path.isfile(os.path.join(pasta, n)))

    por_demanda = {}
    soltos = []
    for nome in nomes:
        m = PADRAO.match(nome)
        if m:
            por_demanda.setdefault(m.group(1), []).append((int(m.group(2)), nome))
        elif nome not in atribuicoes:
            soltos.append(nome)

    # ordem relativa preservada pela sequência atual
    for dem in por_demanda:
        por_demanda[dem] = [nome for _, nome in sorted(por_demanda[dem])]

    # atribuídos entram depois dos já numerados, na ordem da linha de comando
    for nome, dem in atribuicoes.items():
        if not os.path.exists(os.path.join(pasta, nome)):
            sys.exit('ERRO: arquivo atribuído não existe na pasta: %s' % nome)
        por_demanda.setdefault(dem, []).append(nome)

    return por_demanda, soltos


def planejar(por_demanda):
    """Devolve [(origem, destino), ...] só do que muda de nome."""
    plano = []
    for dem in sorted(por_demanda):
        for i, nome in enumerate(por_demanda[dem], 1):
            ext = os.path.splitext(nome)[1].lower()
            destino = '%s.%02d%s' % (dem, i, ext)
            if nome != destino:
                plano.append((nome, destino))
    return plano


def aplicar(pasta, plano):
    """Renomeia em duas fases para não colidir com nomes ainda ocupados."""
    for origem, _ in plano:
        os.rename(os.path.join(pasta, origem), os.path.join(pasta, origem + '.tmp-renum'))
    for origem, destino in plano:
        os.rename(os.path.join(pasta, origem + '.tmp-renum'), os.path.join(pasta, destino))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pasta', required=True, help='Pasta "Recibos e comprovantes" do cliente')
    ap.add_argument('--atribuir', action='append', default=[],
                    help='"nome_do_arquivo=NN" para atribuir um arquivo solto à demanda NN (pode repetir)')
    ap.add_argument('--dry-run', action='store_true', help='Só mostra o plano, não renomeia')
    ap.add_argument('--limpar-lixo', action='store_true',
                    help='Remove .DS_Store / Thumbs.db da pasta')
    args = ap.parse_args()

    if not os.path.isdir(args.pasta):
        sys.exit('ERRO: pasta não encontrada: %s' % args.pasta)

    atribuicoes = {}
    for item in args.atribuir:
        if '=' not in item:
            sys.exit('ERRO: --atribuir espera "arquivo=NN", recebi: %s' % item)
        nome, dem = item.rsplit('=', 1)
        dem = dem.strip().zfill(2)
        if not dem.isdigit():
            sys.exit('ERRO: demanda deve ser numérica em: %s' % item)
        atribuicoes[nome.strip()] = dem

    if args.limpar_lixo and not args.dry_run:
        for lixo in IGNORAR:
            caminho = os.path.join(args.pasta, lixo)
            if os.path.exists(caminho):
                os.remove(caminho)

    por_demanda, soltos = levantar(args.pasta, atribuicoes)
    plano = planejar(por_demanda)

    print('Demandas: %d | arquivos: %d | renomeações: %d'
          % (len(por_demanda), sum(len(v) for v in por_demanda.values()), len(plano)))
    for origem, destino in plano:
        print('  %s -> %s' % (origem, destino))

    faltando = [d for d in sorted(por_demanda) if not por_demanda[d]]
    if faltando:
        print('AVISO: demandas sem nenhum arquivo: %s' % ', '.join(faltando))
    if soltos:
        print('PENDENTE: %d arquivo(s) fora do padrão, sem atribuição '
              '(identifique pelo conteúdo e passe --atribuir):' % len(soltos))
        for nome in soltos:
            print('  %s' % nome)

    if args.dry_run:
        print('(dry-run, nada foi renomeado)')
        return

    aplicar(args.pasta, plano)
    print('ok')


if __name__ == '__main__':
    main()
