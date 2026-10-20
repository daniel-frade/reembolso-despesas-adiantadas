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
      --out ".../reembolso-despesas-adiantadas" \
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
from build_relatorio import (publicar_no_drive, resolver_drive_root, destino_no_drive,  # noqa: E402
                             divergencias_no_drive, arquivos_da_entrega)

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


def meses_anteriores_em_aberto(mm_atual):
    """Tarefas das listas de meses anteriores do ano que ainda nao viraram
    dinheiro de volta, agrupadas por cliente do ClickUp.

    Em 19/10/2026 o setembro de um cliente saiu sozinho quando agosto estava
    cobrado e nao pago, e quem lembrou foi o Daniel. O sinal estava no status:
    `solicitado reembolso` e cobrado e nao pago. `approved` num mes passado e
    cobranca que nunca saiu, ou saiu e a tarefa nao foi movida; o script nao
    distingue os dois, so avisa. O ClickUp e lido aqui, direto pela API."""
    from clickup_adiantamentos_fetch import (api_get, extract_custom_fields,
                                             fetch_all_tasks, task_status)
    token, pasta = os.environ.get('CLICKUP_TOKEN'), os.environ.get('CLICKUP_FOLDER_ID')
    if not (token and pasta):
        return None
    por_cliente = {}
    for lst in api_get(token, '/folder/%s/list' % pasta).get('lists', []):
        mm = lst['name'].split()[0]
        if not (mm.isdigit() and mm < mm_atual):
            continue
        for task in fetch_all_tasks(token, lst['id']):
            status = task_status(task)
            if status not in ('solicitado reembolso', 'approved'):
                continue
            campos = extract_custom_fields(task)
            cliente = (campos.get('cliente') or '').strip().lower()
            por_cliente.setdefault(cliente, []).append(
                (lst['name'], status, float(campos.get('valor') or 0)))
    return por_cliente


def avisos_em_aberto(cfg, em_aberto):
    """Uma linha por mes e status, do tipo
    'Mes anterior em aberto: 08 Agosto 2026, 4 em "solicitado reembolso", R$ 978,32'."""
    soma = {}
    for nome in cfg.get('nome_clickup', []):
        for lista, status, valor in em_aberto.get(nome.strip().lower(), []):
            qtd, tot = soma.get((lista, status), (0, 0.0))
            soma[(lista, status)] = (qtd + 1, tot + valor)
    avisos = []
    for (lista, status), (qtd, tot) in sorted(soma.items()):
        valor = ('%.2f' % tot).replace('.', ',')
        if status == 'solicitado reembolso':
            avisos.append('Mes anterior cobrado e nao pago: %s, %d tarefa(s), R$ %s. '
                          'Considere consolidar (--incluir-xlsx)' % (lista, qtd, valor))
        else:
            avisos.append('Mes anterior com tarefa em approved: %s, %d tarefa(s), R$ %s. '
                          'Nunca cobrada, ou cobrada e nao movida' % (lista, qtd, valor))
    return avisos


def limpar_publicado(nome, saida, cfg, args):
    # So apaga o que esta identico no Drive; uma divergencia trava o cliente
    # inteiro, e a copia local continua sendo a unica boa.
    destino_dir, erro = destino_no_drive(cfg, args.mes_nome, args.ano, args.drive_root)
    if erro:
        return {'cliente': nome, 'erro': erro}
    falhas = divergencias_no_drive(saida, destino_dir)
    if falhas:
        return {'cliente': nome, 'erro': 'não apaguei nada, diverge do Drive: %s' % ', '.join(falhas)}
    base, rel = arquivos_da_entrega(saida)
    for r in rel:
        os.remove(os.path.join(base, r))
    recibos = os.path.join(base, 'Recibos e comprovantes')
    if os.path.isdir(recibos):
        for fn in os.listdir(recibos):
            if fn.startswith('.'):
                os.remove(os.path.join(recibos, fn))
        os.rmdir(recibos)
    sidecar = saida + '.itens.json'
    if cfg.get('apagar_itens_json_ao_limpar') and os.path.isfile(sidecar):
        os.remove(sidecar)
        rel.append(os.path.basename(sidecar))
    # Pasta que sobrou so com lixo do Finder sai junto, a do mes e a do cliente.
    for pasta in (base, os.path.dirname(base)):
        if os.path.isdir(pasta) and all(fn.startswith('.') for fn in os.listdir(pasta)):
            for fn in os.listdir(pasta):
                os.remove(os.path.join(pasta, fn))
            os.rmdir(pasta)
    resultado = {'cliente': nome, 'apagados': len(rel)}
    if os.path.isfile(sidecar):
        resultado['mantido'] = sidecar
    return resultado


def limpar_input(raw_dir, out_dir):
    # O bruto do mes so serve ate a triagem; depois a entrega mora no Drive e o
    # ClickUp guarda o original. Enquanto houver planilha em output/, de qualquer
    # cliente, alguem ainda pode precisar regerar, entao nada sai.
    import glob
    import shutil
    pendentes = glob.glob(os.path.join(glob.escape(out_dir), '**', '*.xlsx'), recursive=True)
    if pendentes:
        return {'input': raw_dir, 'mantido': 'ainda ha planilha em output/: %s'
                % ', '.join(os.path.relpath(p, out_dir) for p in pendentes)}
    if not os.path.isdir(raw_dir):
        return {'input': raw_dir, 'mantido': 'ja nao existia'}
    shutil.rmtree(raw_dir)
    return {'input': raw_dir, 'apagado': True}


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
    ap.add_argument('--so-publicar', action='store_true',
                    help='Não busca nem gera: publica no Drive o que já está em output/, depois da '
                         'triagem e da aprovação do Daniel. É o caminho normal da etapa 5')
    ap.add_argument('--limpar', action='store_true',
                    help='Última etapa, depois do envio confirmado: apaga de output/ a planilha e os '
                         'comprovantes que estão idênticos no Drive. Mantém o .itens.json, que a '
                         'varredura usa para saber o que já foi cobrado, salvo no cliente com '
                         'apagar_itens_json_ao_limpar no config. Se não sobrar planilha nenhuma em '
                         'output/, apaga também input/<ano-mm>/')
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

    if args.so_publicar or args.limpar:
        # Regerar apagaria a triagem dos comprovantes, entao a publicacao de um
        # mes ja revisado le direto o que esta em output/.
        with open(args.config, encoding='utf-8') as f:
            cfg_todos = json.load(f)
        cfg_todos.pop('_comentario', None)
        nomes = [c.strip() for c in args.clientes.split(',')] if args.clientes else list(cfg_todos)
        publicados = []
        for nome in nomes:
            cfg = cfg_todos[nome]
            entrega = cfg.get('nome_entrega', nome)
            pasta = os.path.join(out_dir, cfg.get('pasta', entrega))
            if not cfg.get('pasta_fora_calendario'):
                pasta = os.path.join(pasta, args.mes_pasta)
            # Pelo nome do mes nao serve: a consolidada se chama Agosto-e-Setembro.
            import glob
            planilhas = glob.glob(os.path.join(glob.escape(pasta), '*.xlsx'))
            if len(planilhas) != 1:
                publicados.append({'cliente': nome, 'erro': 'esperava uma planilha em %s, achei %d'
                                   % (pasta, len(planilhas))})
                continue
            saida = planilhas[0]
            if args.limpar:
                publicados.append(limpar_publicado(nome, saida, cfg, args))
                continue
            destino, erro = publicar_no_drive(saida, cfg, args.mes_nome, args.ano, args.drive_root)
            publicados.append({'cliente': nome, 'publicado_em': destino, 'erro': erro})
        if args.limpar:
            publicados.append(limpar_input(raw_dir, out_dir))
        print(json.dumps(publicados, ensure_ascii=False, indent=2))
        return

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

    em_aberto = meses_anteriores_em_aberto(mm)
    if em_aberto is None:
        print('Aviso: sem CLICKUP_FOLDER_ID, meses anteriores em aberto nao foram conferidos',
              file=sys.stderr)

    resultados = []
    for nome_cliente in clientes:
        if nome_cliente not in cfg_todos:
            resultados.append({'cliente': nome_cliente, 'status': 'erro',
                                'etapa': 'config', 'mensagem': 'cliente não está em clientes_config.json'})
            continue
        print('Processando %s...' % nome_cliente, file=sys.stderr)
        res = processar_cliente(nome_cliente, cfg_todos[nome_cliente], args, manifest_path, out_dir)
        # Acumulativo ja junta os meses e tem a varredura propria.
        if em_aberto is not None and not cfg_todos[nome_cliente].get('pasta_fora_calendario'):
            res.setdefault('pendencias', []).extend(avisos_em_aberto(cfg_todos[nome_cliente], em_aberto))
        resultados.append(res)

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
