#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
varrer_cliente.py
=================
Varre TODAS as listas de adiantamento de um cliente e diz o que nunca foi
cobrado, comparando com os sidecars .itens.json das planilhas já entregues.

Por que existe: o fluxo do mês olha a lista daquele mês, e uma tarefa na lista
errada é invisível para todo mundo. Em agosto/2026 uma corrida de aplicativo de
um cliente acumulativo foi criada em 26/08 dentro da lista de junho: não entrou
no relatório de julho, não entrou no de agosto, e só apareceu porque o Daniel
desconfiou do total. Uma demanda de R$ 15,94 achada por intuição é sorte, não
processo.

A comparação é por task_id, que é a única chave estável: o nome da tarefa se
repete entre meses e entre casos.

Uso:
  export CLICKUP_TOKEN="pk_..."    # ou carregue o .env do CRM, que traz os tres
  export CLICKUP_SPACE_ID="..."    # espaco dos adiantamentos; dispensavel com --folder-id

  # tudo que existe do cliente, contra o que já foi entregue
  python3 varrer_cliente.py --cliente "{Nome no ClickUp}" \\
      --entregas ".../output/{Cliente}"

  # um ano fechado, quando a cobrança é retroativa. O status de ano fechado
  # não é mantido, então o resultado sai como aviso e fora do total.
  python3 varrer_cliente.py --cliente "{Nome no ClickUp}" --ano 2025 \\
      --entregas ".../output/{Cliente}"

Saída: JSON com "nao_cobradas" (o que interessa), "ja_cobradas" e "ignoradas"
(as em `to do`, que o escritório ainda não pagou). Código de saída 2 quando há
demanda não cobrada, para o orquestrador poder avisar sem parsear texto.
"""
import argparse
import datetime
import glob
import json
import os
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)

from clickup_adiantamentos_fetch import (  # noqa: E402
    api_get, extract_custom_fields, fetch_all_tasks, task_status,
)


# A escada documentada em SKILL.md. Fora dela nao se deduz nada.
LIQUIDADOS = {'pago pelo cliente', 'reembolsado pelo cliente'}
CONHECIDOS = LIQUIDADOS | {'to do', 'approved', 'solicitado reembolso'}
# Id do espaco no ClickUp. Mesmo motivo do folder no fetch: identifica o
# workspace de quem opera, entao mora no .env, nao aqui.
ESPACO_ADIANTAMENTOS = os.environ.get('CLICKUP_SPACE_ID')


def e_arquivo(nome_pasta):
    """Pasta de ano encerrado, tipo "2025 - ARQUIVO".

    O status das tarefas ali nao e mantido: o Daniel confirmou em 10/09/2026 que
    o 2025 de um cliente estava desatualizado. Entao demanda dessas pastas
    nunca entra no total em aberto, sai so como aviso, para o caso de haver de
    fato algo esquecido. Somar status velho a uma cobranca e cobrar do cliente o
    que a planilha nao consegue provar."""
    return 'arquivo' in nome_pasta.lower()


def pasta_do_ano(token, space_id, ano):
    """A pasta daquele ano, achada pelo ano no nome.

    Varre-se um ano por vez, e nao o espaco inteiro: 2026 tem 9 listas, o espaco
    tem 42, e o saldo de um cliente e uma conta anual. Ano fechado so se varre
    quando o Daniel pedir, e com a ressalva do `e_arquivo`."""
    data = api_get(token, '/space/%s/folder' % space_id)
    pastas = data.get('folders', [])
    for f in pastas:
        if str(ano) in f['name']:
            return [(f['id'], f['name'])]
    disponiveis = [f['name'] for f in pastas]
    sys.exit('ERRO: nenhuma pasta do ano %s no espaco de adiantamentos. '
             'Disponiveis: %s' % (ano, disponiveis))


def listas_da_pasta(token, folder_id):
    data = api_get(token, '/folder/%s/list' % folder_id)
    return [(lst['id'], lst['name']) for lst in data.get('lists', [])]


def task_ids_entregues(pastas):
    """Lê os sidecars .itens.json das entregas e devolve {task_id: arquivo}."""
    entregues = {}
    for pasta in pastas:
        if os.path.isfile(pasta) and pasta.endswith('.itens.json'):
            arquivos = [pasta]
        else:
            arquivos = glob.glob(os.path.join(pasta, '**', '*.itens.json'), recursive=True)
        for caminho in arquivos:
            try:
                with open(caminho, encoding='utf-8') as f:
                    itens = json.load(f)
            except (OSError, ValueError):
                continue
            for item in itens:
                if item.get('task_id'):
                    entregues.setdefault(item['task_id'], os.path.basename(caminho))
    return entregues


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--cliente', required=True, action='append',
                    help='Nome do cliente no dropdown CLIENTES do ClickUp (pode repetir)')
    ap.add_argument('--ano', default=str(datetime.date.today().year),
                    help='Ano a varrer (default: o corrente). Ano fechado tem status nao mantido')
    ap.add_argument('--folder-id', action='append', default=None,
                    help='Pasta de listas, quando o --ano nao resolver. Pode repetir')
    ap.add_argument('--entregas', action='append', default=[],
                    help='Pasta de entregas do cliente ou um .itens.json (pode repetir)')
    args = ap.parse_args()

    token = os.environ.get('CLICKUP_TOKEN')
    if not token:
        sys.exit('ERRO: CLICKUP_TOKEN nao esta no ambiente. Carregue o .env do CRM antes.')

    if args.folder_id:
        pastas = [(fid, '') for fid in args.folder_id]
    else:
        if not ESPACO_ADIANTAMENTOS:
            sys.exit('ERRO: CLICKUP_SPACE_ID nao esta no ambiente. Carregue o .env '
                     'do CRM ou passe --folder-id.')
        pastas = pasta_do_ano(token, ESPACO_ADIANTAMENTOS, args.ano)
    alvo = {c.strip().lower() for c in args.cliente}
    entregues = task_ids_entregues(args.entregas)

    nao_cobradas, ja_cobradas, ignoradas = [], [], []
    liquidadas, status_estranho, em_arquivo = [], [], []
    for folder_id, folder_nome in pastas:
        arquivo = e_arquivo(folder_nome)
        for list_id, list_name in listas_da_pasta(token, folder_id):
            for task in fetch_all_tasks(token, list_id):
                campos = extract_custom_fields(task)
                if (campos.get('cliente') or '').strip().lower() not in alvo:
                    continue
                registro = {
                    'task_id': task['id'],
                    'task_name': task.get('name'),
                    'lista': list_name,
                    'pasta': folder_nome,
                    'status': task_status(task),
                    'valor': campos.get('valor'),
                    'data_pagamento': campos.get('data_pagamento'),
                    'url': task.get('url'),
                }
                status = registro['status']
                if arquivo:
                    em_arquivo.append(registro)
                elif status == 'to do':
                    # `to do` significa que o escritorio ainda nao pagou: nao ha
                    # o que reembolsar, e cobrar isso e cobrar o que nao saiu.
                    ignoradas.append(registro)
                elif status in LIQUIDADOS:
                    # Ja pago pelo cliente. Nao ter sidecar local so quer dizer
                    # que a cobranca e anterior a esta automacao, nao que esta
                    # em aberto. Cobrar de novo o que ja foi pago e pior do que
                    # deixar de cobrar.
                    liquidadas.append(registro)
                elif task['id'] in entregues:
                    registro['entregue_em'] = entregues[task['id']]
                    ja_cobradas.append(registro)
                elif status not in CONHECIDOS:
                    # Status fora da escada documentada. Nao da para deduzir se
                    # significa pago, arquivado ou esquecido: quem decide e o
                    # Daniel, entao sai separado e fora do total.
                    status_estranho.append(registro)
                else:
                    nao_cobradas.append(registro)

    def por_data(r):
        return r.get('data_pagamento') or ''

    nao_cobradas.sort(key=por_data)
    total = sum(float(r['valor'] or 0) for r in nao_cobradas)
    print(json.dumps({
        'cliente': args.cliente,
        'pastas_varridas': [n or i for i, n in pastas],
        'aviso_arquivo': 'Demandas em pasta ARQUIVO ficam fora do total: o status delas '
                         'nao e mantido depois que o ano fecha.',
        'total_nao_cobrado': round(total, 2),
        'nao_cobradas': nao_cobradas,
        'status_desconhecido': sorted(status_estranho, key=por_data),
        'total_status_desconhecido': round(sum(float(r['valor'] or 0) for r in status_estranho), 2),
        'em_arquivo_ano_anterior': sorted(em_arquivo, key=por_data),
        'ja_cobradas': sorted(ja_cobradas, key=por_data),
        'liquidadas': sorted(liquidadas, key=por_data),
        'ignoradas': sorted(ignoradas, key=por_data),
    }, ensure_ascii=False, indent=2))
    sys.exit(2 if nao_cobradas else 0)


if __name__ == '__main__':
    main()
