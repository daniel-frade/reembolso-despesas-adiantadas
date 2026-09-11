#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
organizar_recibos.py
=====================
Organiza os comprovantes baixados pela Frente A (clickup_adiantamentos_fetch.py)
na estrutura documentada em SKILL.md ("Estrutura de pastas final" e
"Identificação e renomeação dos comprovantes"), a partir do sidecar
"{saida}.itens.json" que build_relatorio.py grava quando roda com --manifest.

Regra aplicada automaticamente (o caso mais comum, visto nesta sessão):
  exatamente 1 arquivo em "Invoice Photo" + exatamente 1 anexo de comentário
  -> XX.01.ext (documento) e XX.02.ext (comprovante de pagamento), seguindo
  a ordem "documento antes do comprovante" de SKILL.md.

Qualquer outro caso (0 ou 2+ anexos de comentário, 0 ou 2+ invoice photos,
ou seja: multi-boleto, "brink" + comprovante + recibo de terceiros, etc.)
NÃO é adivinhado, os arquivos originais vão pra "_revisar/XX/" e a demanda
entra na lista de pendências impressa no final, pra decisão humana (ver
SKILL.md, seção "Relação comprovante <-> demanda não é sempre 1:1").

Uso:
  python3 organizar_recibos.py --itens SuperNosso-Julho-2026.xlsx.itens.json \
      --destino "<Cliente>/Recibos e comprovantes"
"""
import argparse
import json
import os
import shutil
import sys


def extensao(caminho):
    return os.path.splitext(caminho)[1].lower() or '.bin'


def resolver(caminho, base):
    """Sidecar novo grava caminho relativo a si mesmo; sidecar antigo, absoluto.

    Aceitar os dois nao e flexibilidade especulativa: existem sidecars dos dois
    formatos no disco, entregues a cliente, e o organizador precisa reabrir os
    antigos quando um mes e reaproveitado numa cobranca consolidada."""
    return caminho if os.path.isabs(caminho) else os.path.normpath(os.path.join(base, caminho))


def organizar(itens, destino, mover=False, base=''):
    os.makedirs(destino, exist_ok=True)
    organizados, revisar = [], []

    for item in itens:
        demanda = item['demanda']
        prefixo = '%02d' % demanda
        invoices = [resolver(c, base) for c in item.get('arquivos_invoice_photo', [])]
        comentarios = [resolver(c, base) for c in item.get('arquivos_comentarios', [])]
        # Anexo largado direto na tarefa nao tem papel declarado: pode ser o
        # comprovante que faltava ou pode ser um print de tela. Sua presenca
        # tira a demanda do caso automatico, mesmo que o 1+1 esteja formado,
        # porque numerar 1+1 e deixar o terceiro arquivo para tras foi
        # exatamente o que entregou o comprovante errado ao cliente.
        soltos = [resolver(c, base) for c in item.get('arquivos_tarefa', [])]
        transferir = shutil.move if mover else shutil.copy2

        if len(invoices) == 1 and len(comentarios) == 1 and not soltos:
            destino_doc = os.path.join(destino, '%s.01%s' % (prefixo, extensao(invoices[0])))
            destino_comp = os.path.join(destino, '%s.02%s' % (prefixo, extensao(comentarios[0])))
            transferir(invoices[0], destino_doc)
            transferir(comentarios[0], destino_comp)
            organizados.append({'demanda': demanda, 'arquivos': [destino_doc, destino_comp]})
            continue

        todos = invoices + comentarios + soltos
        if not todos:
            revisar.append({'demanda': demanda, 'motivo': 'nenhum arquivo encontrado (invoice_photo e comentários vazios)'})
            continue

        pasta_revisar = os.path.join(destino, '_revisar', prefixo)
        os.makedirs(pasta_revisar, exist_ok=True)
        copiados = []
        for caminho in todos:
            alvo = os.path.join(pasta_revisar, os.path.basename(caminho))
            transferir(caminho, alvo)
            copiados.append(alvo)
        revisar.append({
            'demanda': demanda,
            'motivo': '%d invoice_photo + %d comentários%s (esperado 1+1), arquivos em %s'
                      % (len(invoices), len(comentarios),
                         ' + %d solto(s) na bandeja da tarefa' % len(soltos) if soltos else '',
                         pasta_revisar),
            'arquivos': copiados,
        })

    return organizados, revisar


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--itens', required=True, help='Sidecar {saida}.itens.json gerado por build_relatorio.py --manifest')
    ap.add_argument('--destino', required=True, help='Pasta "Recibos e comprovantes" de destino (criada se não existir)')
    ap.add_argument('--mover', action='store_true', help='Mover em vez de copiar os arquivos de origem (default: copiar, não destrutivo)')
    args = ap.parse_args()

    if not os.path.isfile(args.itens):
        print(json.dumps({'status': 'error', 'mensagem': 'itens.json não encontrado: %s' % args.itens}))
        sys.exit(2)

    with open(args.itens, encoding='utf-8') as f:
        itens = json.load(f)

    base = os.path.dirname(os.path.abspath(args.itens))
    organizados, revisar = organizar(itens, args.destino, mover=args.mover, base=base)

    resumo = {
        'destino': args.destino,
        'organizados_automaticamente': len(organizados),
        'para_revisar': len(revisar),
        'revisar': revisar,
    }
    print(json.dumps(resumo, ensure_ascii=False, indent=2))
    sys.exit(0 if not revisar else 1)


if __name__ == '__main__':
    main()
