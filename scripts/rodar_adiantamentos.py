#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rodar_adiantamentos.py
=======================
Orquestrador da automação de adiantamentos: fetch -> gerar .xlsx por cliente
-> recalcular -> organizar comprovantes -> (opcional) publicar no Google
Drive oficial do cliente. Amarra clickup_adiantamentos_fetch.py (Frente A),
build_relatorio.py, recalc.py e organizar_recibos.py, um comando por mês.

Uso:
  python3 rodar_adiantamentos.py --mes-clickup "07 Julho 2026" \
      --mes-nome Julho --ano 2026 \
      --out ~/Desktop/marvin-01/workspace/nlr/automacoes-financeiras/reembolso-despesas-adiantadas \
      --publicar-drive

  # Só um cliente, sem tocar no Drive (ex.: teste/depuração):
  python3 rodar_adiantamentos.py --mes-clickup "06 Junho 2026" --mes-nome Junho \
      --ano 2026 --out /tmp/teste-junho --clientes "<cliente>"

Ao final, imprime um resumo JSON: clientes processados, total por cliente, e
a lista de pendências agregada (recibos ambíguos pra revisar, CASO novo,
cliente sem entrada em clientes_config.json, falha ao publicar no Drive),
os 4 casos que por natureza sempre exigem decisão humana (ver SKILL.md,
seção "O que continua sendo decisão humana"). Não trava no meio:
um cliente com erro não impede os demais de rodar.
"""
import argparse
import json
import os
import subprocess
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
from build_relatorio import publicar_no_drive, resolver_drive_root  # noqa: E402

_MESES_ORDEM = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho',
                'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']


def mes_anterior(mes_nome, ano):
    idx = _MESES_ORDEM.index(mes_nome.strip().lower())
    if idx == 0:
        return _MESES_ORDEM[-1].capitalize(), int(ano) - 1
    return _MESES_ORDEM[idx - 1].capitalize(), int(ano)


def rodar(cmd, **kw):
    return subprocess.run([sys.executable] + cmd, capture_output=True, text=True, **kw)


def processar_cliente(nome_cliente, cfg, args, manifest_path, out_dir):
    nome_entrega = cfg.get('nome_entrega', nome_cliente)
    # Cliente mensal entrega em output/<Cliente>/<ano-mm>/; o acumulativo, cuja
    # planilha cobre varios meses de uma vez, entrega direto em output/<Cliente>/.
    pasta_cliente = os.path.join(out_dir, cfg.get('pasta', nome_entrega))
    if not cfg.get('pasta_fora_calendario'):
        pasta_cliente = os.path.join(pasta_cliente, args.mes_pasta)
    os.makedirs(pasta_cliente, exist_ok=True)
    # Espaco vira hifen no nome do arquivo, ver build_relatorio.py.
    saida = os.path.join(pasta_cliente, '%s-%s-%s.xlsx' % (nome_entrega.replace(' ', '-'), args.mes_nome, args.ano))

    # Referência pra logo: se o cliente é acumulativo, o próprio destino_drive
    # (sem subpasta de mês) já guarda o ultimo arquivo entregue -- usar como
    # --pasta-referencia. Se é mensal, a referencia e o mes anterior dentro
    # do destino_drive (o Drive ja e o "arquivo" de historico, nao precisamos
    # manter copia local de template).
    ref_args = []
    drive_root = resolver_drive_root(args.drive_root)
    destino_tpl = cfg.get('destino_drive')
    if drive_root and destino_tpl:
        if cfg.get('periodo') == 'acumulativo':
            pasta_ref = os.path.join(drive_root, destino_tpl.format(ano=args.ano, mm='', mes_nome=''))
        else:
            mes_ant, ano_ant = mes_anterior(args.mes_nome, args.ano)
            from build_relatorio import _MESES_NUM
            pasta_ref = os.path.join(drive_root, destino_tpl.format(
                ano=ano_ant, mm=_MESES_NUM.get(mes_ant.lower(), '00'), mes_nome=mes_ant))
        if os.path.isdir(pasta_ref):
            ref_args = ['--pasta-referencia', pasta_ref]

    build_cmd = [
        os.path.join(SCRIPT_DIR, 'build_relatorio.py'),
        '--cliente', nome_cliente, '--manifest', manifest_path,
        '--mes', args.mes_nome, '--ano', str(args.ano),
        '--config', args.config, '--saida', saida,
    ] + ref_args + (['--status', args.status] if args.status else [])
    r = rodar(build_cmd)
    if r.returncode != 0:
        return {'cliente': nome_cliente, 'status': 'erro', 'etapa': 'build_relatorio',
                'stdout': r.stdout, 'stderr': r.stderr}
    resultado = json.loads(r.stdout)
    pendencias = list(resultado.get('pendencias', []))

    if resultado.get('itens') == 0:
        return {'cliente': nome_cliente, 'status': 'sem_itens', 'total': 0, 'pendencias': pendencias}

    recalc_r = rodar([os.path.join(SCRIPT_DIR, 'recalc.py'), saida])
    try:
        recalc_info = json.loads(recalc_r.stdout)
    except json.JSONDecodeError:
        recalc_info = {'status': 'error', 'mensagem': recalc_r.stderr or recalc_r.stdout}
    if recalc_info.get('status') != 'success':
        pendencias.append('recalc.py: %s' % recalc_info.get('mensagem', recalc_info))

    # Conferir o que saiu, e nao so o que entrou. Todo defeito de formato de
    # 2026 (logo achatada, janela pela metade, nome do arquivo) atravessou o
    # pipeline inteiro sem ninguem reclamar e foi descoberto pelo Daniel
    # abrindo o arquivo. O conferidor roda depois do recalc porque e o recalc
    # que grava o cache do SUM e a geometria da janela.
    conf_r = rodar([os.path.join(SCRIPT_DIR, 'conferir_planilha.py'), saida,
                    '--itens', str(resultado.get('itens') or 0),
                    '--nome-entrega', nome_entrega])
    try:
        conf_info = json.loads(conf_r.stdout)
        for falha in conf_info.get('falhas', []):
            pendencias.append('Conferencia da planilha (%s): %s' % (nome_cliente, falha))
    except json.JSONDecodeError:
        pendencias.append('conferir_planilha.py falhou: %s' % (conf_r.stderr or conf_r.stdout))

    # Cliente acumulativo se varre sempre, sem depender de alguem lembrar.
    # A instrucao escrita ("varra quando o cliente estiver atrasado") falhou no
    # teste de 10/09/2026: quem gerava agosto de um cliente acumulativo nao ligava
    # "fechar agosto" com "esse cliente nao paga desde janeiro", e uma demanda que
    # morava na lista de junho sumiu de novo. Acumulativo ja
    # significa atrasado, entao a condicao esta no config e nao na memoria.
    if cfg.get('periodo') == 'acumulativo' or cfg.get('pasta_fora_calendario'):
        var_cmd = [os.path.join(SCRIPT_DIR, 'varrer_cliente.py'), '--entregas', pasta_cliente]
        for nome_ck in cfg.get('nome_clickup', [nome_cliente]):
            var_cmd += ['--cliente', nome_ck]
        var_r = rodar(var_cmd)
        try:
            var_info = json.loads(var_r.stdout)
            for d in var_info.get('nao_cobradas', []):
                pendencias.append(
                    'Demanda fora deste relatorio, achada na lista "%s": %s (R$ %s, %s)'
                    % (d.get('lista'), d.get('task_name'), d.get('valor'), d.get('url')))
        except json.JSONDecodeError:
            pendencias.append('varrer_cliente.py falhou: %s' % (var_r.stderr or var_r.stdout))

    destino_recibos = os.path.join(pasta_cliente, 'Recibos e comprovantes')
    itens_json = resultado.get('itens_json')
    if itens_json and os.path.isfile(itens_json):
        org_r = rodar([os.path.join(SCRIPT_DIR, 'organizar_recibos.py'),
                        '--itens', itens_json, '--destino', destino_recibos])
        try:
            org_info = json.loads(org_r.stdout)
            for item in org_info.get('revisar', []):
                pendencias.append('Recibos p/ revisar (demanda %s) de %s: %s'
                                   % (item['demanda'], nome_cliente, item['motivo']))
        except json.JSONDecodeError:
            pendencias.append('organizar_recibos.py falhou: %s' % (org_r.stderr or org_r.stdout))

    publicado_em = None
    if args.publicar_drive:
        publicado_em, erro = publicar_no_drive(saida, cfg, args.mes_nome, args.ano, args.drive_root)
        if erro:
            pendencias.append('Publicação no Drive (%s): %s' % (nome_cliente, erro))
        elif os.path.isdir(destino_recibos):
            destino_drive_recibos = os.path.join(os.path.dirname(publicado_em), 'Recibos e comprovantes')
            os.makedirs(destino_drive_recibos, exist_ok=True)
            import shutil
            for fn in os.listdir(destino_recibos):
                origem = os.path.join(destino_recibos, fn)
                if os.path.isfile(origem):
                    shutil.copy2(origem, os.path.join(destino_drive_recibos, fn))

    # Quais tarefas o Daniel precisa mover depois de enviar. O status e
    # declaracao de fato dele e eu nunca escrevo nele, mas nao dizer quais sao
    # deixa o ClickUp atrasado em relacao ao que o cliente ja recebeu: em
    # setembro/2026 nove tarefas de um cliente estavam paradas em "approved",
    # uma delas desde a cobranca de 11/08.
    mover_status = []
    if itens_json and os.path.isfile(itens_json):
        try:
            with open(itens_json, encoding='utf-8') as f:
                for item in json.load(f):
                    if item.get('task_id'):
                        mover_status.append(item['task_id'])
        except (OSError, ValueError):
            pass

    return {
        'cliente': nome_cliente, 'status': 'ok', 'arquivo': saida,
        'publicado_em': publicado_em, 'itens': resultado.get('itens'),
        'total': resultado.get('total'), 'pendencias': pendencias,
        'mover_para_solicitado_reembolso': mover_status,
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--mes-clickup', required=True, help='Nome exato da lista no ClickUp, ex: "07 Julho 2026"')
    ap.add_argument('--mes-nome', required=True, help='Ex: Julho')
    ap.add_argument('--ano', required=True)
    ap.add_argument('--out', required=True,
                help='Raiz da frente no workspace; o script cria input/<ano-mm>/ e output/<Cliente>/ dentro')
    ap.add_argument('--config', default=os.path.join(SCRIPT_DIR, 'clientes_config.json'))
    ap.add_argument('--clientes', help='Lista separada por vírgula pra restringir (default: todos em clientes_config.json)')
    ap.add_argument('--status', help='Filtro de status pro fetch + build (default: sem filtro no fetch, approved no build)')
    ap.add_argument('--publicar-drive', action='store_true', help='Publica os .xlsx + recibos organizados na pasta oficial de cada cliente no Drive')
    ap.add_argument('--drive-root', help='Raiz do Google Drive (default: autodetectar)')
    ap.add_argument('--dry-run', action='store_true', help='Passa --dry-run pro fetch (não baixa arquivos, só lista tarefas)')
    args = ap.parse_args()

    # Desde 10/09/2026 --out e a raiz da frente, nao a pasta do mes: o bruto do
    # ClickUp cai em input/<ano-mm>/ e a entrega em output/<Cliente>/. Antes os
    # dois moravam juntos dentro da pasta do mes, o que obrigava a pasta a ser
    # ao mesmo tempo insumo e entrega, e nao acomodava cliente acumulativo.
    base_dir = os.path.abspath(args.out)
    mm = args.mes_clickup.split()[0]
    if not (len(mm) == 2 and mm.isdigit()):
        sys.exit('ERRO: --mes-clickup precisa comecar com o mes em dois digitos, '
                 'ex: "09 Setembro 2026". Recebi: %s' % args.mes_clickup)
    args.mes_pasta = '%s-%s' % (args.ano, mm)
    raw_dir = os.path.join(base_dir, 'input', args.mes_pasta)
    out_dir = os.path.join(base_dir, 'output')
    os.makedirs(raw_dir, exist_ok=True)

    fetch_cmd = [os.path.join(SCRIPT_DIR, 'clickup_adiantamentos_fetch.py'),
                 '--month', args.mes_clickup, '--out', raw_dir]
    if args.dry_run:
        fetch_cmd.append('--dry-run')
    print('Buscando dados do ClickUp (%s)...' % args.mes_clickup, file=sys.stderr)
    r = rodar(fetch_cmd)
    print(r.stdout, file=sys.stderr)
    if r.returncode != 0:
        print(json.dumps({'status': 'erro', 'etapa': 'fetch', 'stderr': r.stderr}, ensure_ascii=False))
        sys.exit(2)

    manifest_path = os.path.join(raw_dir, 'manifest.json')
    with open(args.config, encoding='utf-8') as f:
        cfg_todos = json.load(f)
    cfg_todos.pop('_comentario', None)

    clientes = [c.strip() for c in args.clientes.split(',')] if args.clientes else list(cfg_todos)

    resultados = []
    for nome_cliente in clientes:
        if nome_cliente not in cfg_todos:
            resultados.append({'cliente': nome_cliente, 'status': 'erro',
                                'etapa': 'config', 'mensagem': 'cliente não está em clientes_config.json'})
            continue
        print('Processando %s...' % nome_cliente, file=sys.stderr)
        resultados.append(processar_cliente(nome_cliente, cfg_todos[nome_cliente], args, manifest_path, out_dir))

    todas_pendencias = []
    for res in resultados:
        for p in res.get('pendencias', []):
            todas_pendencias.append('[%s] %s' % (res['cliente'], p))

    mover = {r['cliente']: r['mover_para_solicitado_reembolso']
             for r in resultados if r.get('mover_para_solicitado_reembolso')}

    resumo = {
        'mes': '%s/%s' % (args.mes_nome, args.ano),
        'clientes_processados': len([r for r in resultados if r.get('status') == 'ok']),
        'clientes_com_erro': [r['cliente'] for r in resultados if r.get('status') == 'erro'],
        'resultados': resultados,
        'pendencias': todas_pendencias,
        'depois_de_enviar': {
            'instrucao': 'Mover estas tarefas para "solicitado reembolso" no ClickUp. '
                         'O status e declaracao do Daniel, o script nao escreve nele.',
            'tarefas_por_cliente': mover,
        },
    }
    print(json.dumps(resumo, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
