# Armadilhas conhecidas

Cada linha aqui custou um erro real. Consulte quando algo quebrar ou sair estranho, e
acrescente quando um mês novo ensinar algo. As de formatação vivem em
[formato-planilha.md](formato-planilha.md).

## Datas

| Problema | Causa | Solução |
|---|---|---|
| Consolidação sai com ordem e numeração trocadas | `parse_data()` tenta `%m/%d/%Y` **antes** de `%d/%m/%Y`, porque o CSV do ClickUp exporta no formato americano. Um `.xlsx` já entregue traz dd/mm, então `08/06/2026` vira 6 de agosto | `ler_linhas_xlsx()` converte para ISO explicitamente antes de passar adiante. Qualquer fonte nova de dados precisa declarar seu formato, nunca confiar na ordem de tentativa |
| Demanda na posição errada, com data anterior ao pagamento | O `Payment Date` do ClickUp às vezes é o **vencimento**. Julho/2026: campo 02/07, comprovante bancário e recibo ONR 08/07 | Nas demandas que têm brink, comparar o campo com o "Data do Vencimento" impresso nele. Se baterem, a data boa é a do comprovante. Corrigir no ClickUp e regerar |

Automatizar a segunda foi testado e descartado em julho/2026: comparar o campo com a data
embutida no nome dos arquivos (`Recibo_Certidao_20260720_...`) cobre 12 de 26 tarefas e produziu
1 alerta, legítimo, porque o recibo do ONR é emitido dias depois do pagamento por natureza.
Alerta que erra 1 em 12 treina a ignorar alertas.

## ClickUp

| Problema | Causa | Solução |
|---|---|---|
| A maior parte do mês some | A API esconde tarefas arquivadas por padrão, mesmo com `include_closed=true`, e o escritório arquiva a tarefa assim que vira "reembolsado pelo cliente" | Buscar `archived=false` **e** `archived=true` e juntar |
| Relatório sai vazio, ou com uma fração das demandas | Filtro `--status approved`. O status caminha `to do` → `approved` → `solicitado reembolso` → `pago pelo cliente` / `reembolsado pelo cliente`, então `approved` é transitório: em abril/2026 sobrou 0 tarefa nele, em junho 0, em julho 3 de 26 | O default do `build_relatorio.py` agora é por exclusão: entra tudo que não esteja em `to do`, que é o único status que significa "o escritório ainda não pagou" |
| Anexo de comentário vem sem URL | A ferramenta MCP `clickup_get_threaded_comments` descarta o campo e devolve só o nome do arquivo dentro do texto | Chamada HTTP direta: `GET /comment/{id}/reply` traz `comment[].attachment.url` |
| Demanda sem o recibo que existe no ClickUp | O fetch lia só o campo "Invoice Photo" e os anexos de resposta de comentário, e o plano de julho/2026 deu o array `attachments` da tarefa como sempre vazio a partir de uma amostra. Ele é, na verdade, o inventário completo | Corrigido em 10/09/2026: o fetch lê as três portas e deduplica por id do anexo. O que só existe na bandeja vira `tarefa_NN_` e manda a demanda para `_revisar/`. Custou três entregas antes disso, em três clientes diferentes |
| JSON gigante por tarefa | O dropdown "CLIENTES" carrega as ~246 opções do workspace inteiro a cada tarefa | Resolver o dropdown contra as `options` do próprio payload e gravar só o nome no manifest |
| Uma tarefa vira várias linhas no relatório | Ex.: "Notificações extrajudiciais RTDPJ" agrupa 4 destinatários numa tarefa só | O script gera 1 linha por tarefa e não tenta dividir: a divisão exige ler os recibos individuais. Continua manual |

## Arquivos e pastas

| Problema | Causa | Solução |
|---|---|---|
| 26 pastas de nome ilegível junto das entregas | `--out` do fetch apontado para `output/` em vez de `input/{ano-mm}` | Ver a estrutura de pastas na SKILL.md. `rodar_adiantamentos.py` já faz certo sozinho |
| Caminhos quebrados no `.xlsx.itens.json` | Sidecar antigo grava caminho absoluto para o bruto; mover a pasta invalida todos, o que aconteceu em 10/09/2026 com 54 caminhos | Desde 10/09/2026 o sidecar grava caminho relativo a si mesmo, e `organizar_recibos.py` aceita os dois formatos. Sidecar antigo ainda precisa de reescrita à mão |
| `--pasta-referencia` não acha o mês anterior | A comparação exigia prefixo literal (`nome.replace(' ','')`), mas a pasta do Drive usa hífens (`Nome-Do-Cliente-...`) | Normalizar os dois lados (minúsculo, só alfanumérico) antes de comparar |
| `rm` devolve `Operation not permitted` | Arquivos no mount do Google Drive | `mcp__cowork__allow_cowork_file_delete` com um dos arquivos libera a pasta toda |
| EXIF vazio nas imagens | Screenshots e downloads não preservam metadados | Ler visualmente com a ferramenta `Read`: o comprovante bancário mostra valor, data e favorecido |

## Identificação de comprovantes

| Problema | Causa | Solução |
|---|---|---|
| Duas demandas com o mesmo valor | Nada nos custom fields diferencia | Usar o horário da sessão de pagamento no comprovante bancário |
| Não dá para saber de qual pedido é o recibo |, | O campo "Mensagem" do comprovante bancário traz um protocolo que corresponde ao "pedido nº" do recibo ONR. É o cruzamento mais confiável quando existe |
| Descartei o comprovante e entreguei o brink | Decisão tomada pelo nome do arquivo | Julho/2026: `ONR - <nome do caso>.pdf` é comprovante Pix de R$ 39,16; `Pagametnto CCIR pix.pdf` é PagTesouro "Aguardando realização do pagamento". Abrir e olhar, sempre |
