#!/usr/bin/env python3
"""
Puxa dados e anexos do ClickUp para o relatorio mensal de adiantamentos de despesas,
substituindo o export manual de CSV + o download tarefa-a-tarefa.

So biblioteca padrao do Python (sem dependencias externas).

Uso:
    export CLICKUP_TOKEN="pk_..."
    export CLICKUP_FOLDER_ID="..."   # pasta do ano; ou passe --folder-id
    python3 clickup_adiantamentos_fetch.py --month "07 Julho 2026" --out ../2026-07/_raw

O que faz, por tarefa da lista do mes:
    - Le os custom fields (CLIENTES, VALOR, Metodo de pagamento, Payment Date, ORGAO,
      Task Content) e resolve os dropdowns pro nome legivel.
    - Baixa o(s) arquivo(s) do campo "Invoice Photo" (documento inicial: boleto ou "brink").
    - Le a bandeja da propria tarefa (attachments), que e o inventario completo, e
      baixa como "tarefa_NN_" o que nao veio por nenhuma das outras duas portas.
    - Percorre comentarios e respostas da tarefa e baixa todo anexo de arquivo encontrado
      (comprovante de pagamento e, quando existir, recibo oficial de terceiros). Isso exige
      chamada direta a API do ClickUp -- a ferramenta MCP clickup_get_threaded_comments nao
      expoe a URL do anexo, so o nome do arquivo dentro do texto.

Saida em --out:
    manifest.json   -- uma entrada por tarefa, com todos os campos + caminhos dos arquivos
    manifest.csv    -- mesma coisa em CSV, equivalente ao export manual do ClickUp
    {task_id}/...   -- arquivos baixados (Invoice Photo + anexos de comentario)

Depois disso, a Etapa 2 em diante da skill reembolso-despesas-adiantadas (filtrar por cliente,
gerar o .xlsx, nomear XX.YY e organizar em "Recibos e comprovantes") continua igual.
"""

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

API_BASE = "https://api.clickup.com/api/v2"
# Id da pasta do ano no ClickUp. Mora no .env junto do token, nao no codigo,
# porque identifica o workspace de quem opera; --folder-id continua sobrepondo.
DEFAULT_FOLDER_ID = os.environ.get("CLICKUP_FOLDER_ID")


def api_get(token, path, params=None):
    url = f"{API_BASE}{path}"
    if params:
        url += "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": token})
    last_err = None
    for attempt in range(5):
        try:
            with urllib.request.urlopen(req, timeout=30) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429:
                time.sleep(2**attempt)
                last_err = e
                continue
            body = e.read().decode("utf-8", "ignore")
            raise RuntimeError(f"ClickUp API {e.code} em {path}: {body}") from e
    raise RuntimeError(f"ClickUp API: excedeu tentativas em {path}") from last_err


def download_file(url, dest_path):
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": "adiantamentos-fetch/1.0"})
    with urllib.request.urlopen(req, timeout=60) as resp, open(dest_path, "wb") as f:
        f.write(resp.read())


def find_list_id(token, folder_id, month_name):
    data = api_get(token, f"/folder/{folder_id}/list")
    for lst in data.get("lists", []):
        if lst["name"].strip().lower() == month_name.strip().lower():
            return lst["id"]
    available = [lst["name"] for lst in data.get("lists", [])]
    raise SystemExit(
        f"Lista '{month_name}' nao encontrada na pasta {folder_id}. "
        f"Listas disponiveis: {available}"
    )


def resolve_dropdown(field, value):
    if value is None:
        return None
    options = field.get("type_config", {}).get("options", [])
    for opt in options:
        if opt.get("orderindex") == value:
            return opt.get("name")
    if isinstance(value, int) and 0 <= value < len(options):
        return options[value].get("name")
    return value


def epoch_ms_to_date(value):
    if not value:
        return None
    try:
        return datetime.fromtimestamp(int(value) / 1000, tz=timezone.utc).strftime("%Y-%m-%d")
    except (ValueError, TypeError):
        return None


def extract_custom_fields(task):
    out = {
        "cliente": None,
        "metodo_pagamento": None,
        "orgao": None,
        "data_pagamento": None,
        "valor": None,
        "invoice_photo": [],
    }
    for field in task.get("custom_fields", []):
        name = (field.get("name") or "").strip()
        value = field.get("value")
        if name == "CLIENTES":
            out["cliente"] = resolve_dropdown(field, value)
        elif name == "Metodo de pagamento":
            out["metodo_pagamento"] = resolve_dropdown(field, value)
        elif name == "ÓRGÃO":
            out["orgao"] = resolve_dropdown(field, value)
        elif name == "Payment Date":
            out["data_pagamento"] = epoch_ms_to_date(value)
        elif name == "VALOR":
            out["valor"] = value
        elif name == "Invoice Photo" and value:
            out["invoice_photo"] = value
    return out


def _extract_attachments_from_comment(comment_obj):
    found = []
    for block in comment_obj.get("comment", []):
        if block.get("type") == "attachment" and "attachment" in block:
            att = block["attachment"]
            found.append(
                {
                    # O id e obrigatorio: e por ele que process_task deduplica
                    # o anexo de comentario contra a bandeja da tarefa, que
                    # contem os mesmos arquivos. Sem ele o dedup compara None
                    # e baixa tudo duas vezes.
                    "id": att.get("id"),
                    "url": att.get("url"),
                    "title": att.get("title"),
                    "extension": att.get("extension"),
                    "source_comment_id": comment_obj.get("id"),
                    "uploaded_by": (comment_obj.get("user") or {}).get("username"),
                    "date": comment_obj.get("date"),
                }
            )
    return found


def collect_task_attachments(token, task_id):
    """Anexos da bandeja da propria tarefa (GET /task/{id}, campo attachments).

    Esta e a lista COMPLETA do que existe na tarefa: conferido em 10/09/2026,
    ela e superconjunto do campo "Invoice Photo" e dos anexos de comentario, com
    os mesmos ids. As outras duas gavetas nao dizem o que existe, dizem o PAPEL
    de cada arquivo (documento inicial x comprovante de pagamento), e e por isso
    que continuam sendo lidas separadamente.

    O que aparece so aqui e arquivo que alguem arrastou para a tarefa sem
    classificar. Ate setembro/2026 o fetch nao lia esta gaveta, porque o plano
    de julho a deu como sempre vazia a partir de uma amostra, e cada arquivo
    solto sumia calado: em setembro/2026 isso custou tres entregas, dois recibos
    que nunca chegaram a planilha e um comprovante que passou dois meses na
    demanda errada, e saiu ao cliente assim duas vezes."""
    task = api_get(token, f"/task/{task_id}")
    return task.get("attachments") or []


def collect_comment_attachments(token, task_id):
    """Percorre comentarios e respostas da tarefa, devolve lista de anexos (url/title)."""
    attachments = []
    comments = api_get(token, f"/task/{task_id}/comment").get("comments", [])
    for c in comments:
        attachments.extend(_extract_attachments_from_comment(c))
        if c.get("reply_count", 0) > 0:
            replies = api_get(token, f"/comment/{c['id']}/reply").get("comments", [])
            for r in replies:
                attachments.extend(_extract_attachments_from_comment(r))
    return attachments


_SAFE_CHARS = set(
    "-_.() "
    + "".join(chr(c) for c in range(48, 58))
    + "".join(chr(c) for c in range(65, 91))
    + "".join(chr(c) for c in range(97, 123))
)


def safe_filename(name):
    cleaned = "".join(ch if ch in _SAFE_CHARS else "_" for ch in name).strip()
    return cleaned or "arquivo"


def task_status(task):
    status = task.get("status")
    if isinstance(status, dict):
        return status.get("status")
    return status


def process_task(token, task, out_dir, skip_downloads=False):
    task_id = task["id"]
    fields = extract_custom_fields(task)
    task_dir = out_dir / task_id
    downloaded = {"invoice_photo": [], "comentarios": [], "tarefa": []}

    if not skip_downloads:
        vistos = set()
        for i, item in enumerate(fields["invoice_photo"], start=1):
            url = item.get("url")
            if not url:
                continue
            vistos.add(item.get("id"))
            fname = f"invoice_{i:02d}_{safe_filename(item.get('title') or 'arquivo')}"
            dest = task_dir / fname
            download_file(url, dest)
            downloaded["invoice_photo"].append(str(dest))

        comment_attachments = collect_comment_attachments(token, task_id)
        for i, att in enumerate(comment_attachments, start=1):
            url = att.get("url")
            if not url:
                continue
            vistos.add(att.get("id"))
            fname = f"comentario_{i:02d}_{safe_filename(att.get('title') or 'arquivo')}"
            dest = task_dir / fname
            download_file(url, dest)
            downloaded["comentarios"].append(str(dest))

        # O que sobra da bandeja da tarefa e arquivo sem papel declarado. Baixar
        # com prefixo proprio, para quem revisa ver de onde veio: ele nao passou
        # nem pelo formulario nem por resposta de comentario, entao nada diz se
        # e documento ou comprovante, e adivinhar isso ja custou entrega errada.
        for att in collect_task_attachments(token, task_id):
            if att.get("id") in vistos or not att.get("url"):
                continue
            vistos.add(att.get("id"))
            i = len(downloaded["tarefa"]) + 1
            fname = f"tarefa_{i:02d}_{safe_filename(att.get('title') or 'arquivo')}"
            dest = task_dir / fname
            download_file(att["url"], dest)
            downloaded["tarefa"].append(str(dest))
    else:
        downloaded["invoice_photo"] = [item.get("title") for item in fields["invoice_photo"]]
        comentarios = collect_comment_attachments(token, task_id)
        downloaded["comentarios"] = [a.get("title") for a in comentarios]
        vistos = {i.get("id") for i in fields["invoice_photo"]} | {a.get("id") for a in comentarios}
        downloaded["tarefa"] = [a.get("title") for a in collect_task_attachments(token, task_id)
                                if a.get("id") not in vistos]

    return {
        "task_id": task_id,
        "task_name": task.get("name"),
        "task_content": task.get("text_content") or task.get("description"),
        "status": task_status(task),
        "cliente": fields["cliente"],
        "valor": fields["valor"],
        "metodo_pagamento": fields["metodo_pagamento"],
        "data_pagamento": fields["data_pagamento"],
        "orgao": fields["orgao"],
        "url_tarefa": task.get("url"),
        "arquivos_invoice_photo": downloaded["invoice_photo"],
        "arquivos_comentarios": downloaded["comentarios"],
        "arquivos_tarefa": downloaded["tarefa"],
        "total_arquivos": (len(downloaded["invoice_photo"]) + len(downloaded["comentarios"])
                           + len(downloaded["tarefa"])),
    }


def fetch_all_tasks(token, list_id):
    """Busca tarefas da lista, combinando ativas e arquivadas.

    A API do ClickUp esconde tarefas arquivadas por padrao (mesmo com
    include_closed=true) -- e o escritorio parece arquivar a tarefa assim que ela
    fica "reembolsado pelo cliente" / "done", entao sem isso a maior parte
    do mes desaparece silenciosamente. Buscamos os dois grupos e juntamos.
    """
    seen = {}
    for archived_flag in ("false", "true"):
        page = 0
        while True:
            params = {"page": page, "include_closed": "true", "archived": archived_flag}
            data = api_get(token, f"/list/{list_id}/task", params=params)
            batch = data.get("tasks", [])
            for t in batch:
                seen[t["id"]] = t
            if data.get("last_page", True) or not batch:
                break
            page += 1
    return list(seen.values())


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--month", required=True, help='Nome exato da lista no ClickUp, ex: "07 Julho 2026"')
    parser.add_argument("--out", required=True, help="Pasta de saida para manifest + arquivos baixados")
    parser.add_argument("--folder-id", default=DEFAULT_FOLDER_ID, required=not DEFAULT_FOLDER_ID,
                        help="ID da pasta do ano no ClickUp (default: $CLICKUP_FOLDER_ID)")
    parser.add_argument("--status", action="append", help="Filtrar por status (pode repetir). Default: todos.")
    parser.add_argument("--limit", type=int, default=None, help="Limitar numero de tarefas processadas (uso em teste)")
    parser.add_argument("--dry-run", action="store_true", help="Nao baixa arquivos, so lista o que baixaria (teste rapido)")
    args = parser.parse_args()

    token = os.environ.get("CLICKUP_TOKEN")
    if not token:
        sys.exit("Defina a variavel de ambiente CLICKUP_TOKEN antes de rodar.")

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    list_id = find_list_id(token, args.folder_id, args.month)
    print(f"Lista '{args.month}' -> list_id={list_id}")

    tasks = fetch_all_tasks(token, list_id)

    if args.status:
        wanted = {s.lower() for s in args.status}
        tasks = [t for t in tasks if (task_status(t) or "").lower() in wanted]

    if args.limit:
        tasks = tasks[: args.limit]

    print(f"{len(tasks)} tarefa(s) para processar")

    manifest = []
    for i, task in enumerate(tasks, start=1):
        print(f"[{i}/{len(tasks)}] {task.get('name')}")
        try:
            manifest.append(process_task(token, task, out_dir, skip_downloads=args.dry_run))
        except Exception as e:  # noqa: BLE001 - queremos seguir processando as outras tarefas
            print(f"  ERRO em {task.get('id')}: {e}", file=sys.stderr)
            manifest.append({"task_id": task.get("id"), "task_name": task.get("name"), "erro": str(e)})

    manifest_json = out_dir / "manifest.json"
    manifest_json.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    manifest_csv = out_dir / "manifest.csv"
    cols = [
        "task_id",
        "task_name",
        "task_content",
        "status",
        "cliente",
        "valor",
        "metodo_pagamento",
        "data_pagamento",
        "orgao",
        "total_arquivos",
        "url_tarefa",
    ]
    with open(manifest_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        writer.writeheader()
        for row in manifest:
            writer.writerow(row)

    print(f"\nManifest gravado em {manifest_json} e {manifest_csv}")
    if not args.dry_run:
        print(f"Arquivos baixados em subpastas por task_id dentro de {out_dir}")


if __name__ == "__main__":
    main()
