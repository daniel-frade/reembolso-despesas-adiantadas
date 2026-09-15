---
name: reembolso-despesas-adiantadas
description: >
  Gera as planilhas mensais de adiantamentos/reembolsos de despesas dos clientes a
  partir do ClickUp e organiza os comprovantes por demanda. Use sempre que o Daniel disser
  "roda os adiantamentos de <mês>", "relatório de adiantamentos", "planilha de adiantamentos
  do <cliente>", "reembolsos do mês", "adiantamento de despesas", ou nomear um cliente do
  processo (os nomes ficam em scripts/clientes_config.json, fora do versionado)
  no contexto de custas pagas pelo caixa do escritório. Use também quando ele pedir para
  continuar de onde parou num mês em andamento ("segue pro proximo cliente", "faz o do fulano"),
  para juntar num relatório só um mês que o cliente não pagou com o mês corrente (cobrança
  consolidada), para renomear ou organizar recibos e comprovantes no padrão XX.YY, ou para
  decidir quais anexos do ClickUp vão para o cliente e quais são só pedido de pagamento
  ("brink"). Vale igualmente quando houver um CSV exportado do ClickUp mais arquivos de
  recibos soltos para transformar em relatório. Na dúvida entre esta skill e trabalho
  genérico de planilha, prefira esta: o formato, o filtro de status e a triagem dos
  comprovantes têm regras próprias que erram silenciosamente se improvisadas.
---

# Relatório Mensal de Adiantamentos

Custas de clientes pagas pelo caixa do escritório, reembolsadas mensalmente. Para cada cliente,
a entrega é um `.xlsx` formatado com uma linha por demanda, mais uma pasta de comprovantes
numerados por demanda.

## A regra que organiza todo o resto

**`scripts/build_relatorio.py` é a única implementação do formato. Não reimplemente.**

Toda a formatação (fonte, tamanhos, bordas, cores, larguras, geometria da linha 2, logo, fórmula
do total) está no script e é aplicada sozinha. Se você escrever um script próprio para "só esse
caso", vai reencontrar em uma tarde os bugs que custaram semanas: cabeçalho no tamanho errado,
rótulo fixo na coluna extra, data brasileira lida como americana.

Corolário que importa mais do que parece: **quando a saída estiver errada, o defeito é do
script.** Conserte lá e regere. Corrigir só o arquivo do mês faz o erro voltar no mês seguinte,
e você não vai lembrar.

A especificação completa do formato está em
[references/formato-planilha.md](references/formato-planilha.md), para consultar ao mexer no
script ou investigar uma célula estranha. Para rodar o mês, você não precisa dela.

## Scripts

| Script | Papel |
|---|---|
| `rodar_adiantamentos.py` | **Orquestrador, um comando por mês.** Amarra todos os outros e imprime totais por cliente + pendências. Comece por ele |
| `clickup_adiantamentos_fetch.py` | Busca tarefas, custom fields e anexos (inclusive de resposta de comentário) na API do ClickUp; grava `manifest.json` |
| `build_relatorio.py` | Gera o `.xlsx` a partir do `manifest.json` ou de um CSV, mais o sidecar `{saida}.itens.json` |
| `organizar_recibos.py` | Cria `Recibos e comprovantes/` e nomeia `XX.YY` quando a demanda tem 1 documento + 1 comprovante |
| `renumerar_recibos.py` | Recompacta a numeração depois da poda manual |
| `recalc.py` | Recalcula fórmulas via LibreOffice headless, preservando a logo, e grava a geometria da janela |
| `conferir_planilha.py` | Confere o `.xlsx` pronto contra o contrato de formato e devolve as violações |
| `varrer_cliente.py` | Varre todas as listas do cliente e diz o que nunca foi cobrado |
| `clientes_config.json` | Registro de clientes: nome no ClickUp vs. nome de entrega, combos, colunas extras, período, `destino_drive`. **Cliente novo = nova entrada aqui**, sem tocar em código |

## Antes de tudo: o pedido é do mês inteiro ou de um cliente?

O Daniel pede das duas formas, e a diferença de custo é enorme.

| Ele diz | Escopo | Como rodar |
|---|---|---|
| "roda os adiantamentos de julho" | mês inteiro | `rodar_adiantamentos.py` sem `--clientes` |
| "roda o adiantamento do <cliente>", "segue pro <cliente>" | **um cliente só** | `rodar_adiantamentos.py --clientes "<cliente>"` |

Se ele nomeou um cliente, entregue **aquele cliente**. Processar o mês inteiro "já que estou
aqui" gasta quinze minutos e produz quatro planilhas que ninguém pediu, cada uma exigindo
revisão dele. Isso aconteceu de verdade numa avaliação: o pedido era de um cliente e saíram
cinco relatórios.

A etapa 0 abaixo continua valendo mesmo para um cliente só, porque a lista pode ter mudado. O
que muda é o que você **gera**, não o que você **confere**.

## Fluxo do mês

### 0. Reler a lista, imediatamente antes de gerar

A lista do ClickUp **se move enquanto você trabalha**. Em julho/2026 a primeira leitura deu 31
tarefas e 8 clientes; depois que o Daniel reorganizou, virou 26 e 6, com dois clientes saindo e
um entrando. Levantamento do começo da conversa não vale na hora de gerar.

- Buscar `archived=false` **e** `archived=true` e juntar: o escritório arquiva a tarefa assim que vira
  "reembolsado pelo cliente", e a API esconde arquivadas mesmo com `include_closed=true`.
- Confrontar os clientes encontrados com `clientes_config.json`. Cliente da lista que não está
  no config para o processo **só para ele**; os demais seguem normalmente.
- Cliente do config que não apareceu no mês é normal: em julho/2026, num cliente de entrega
  conjunta, um dos dois nomes do ClickUp teve zero tarefas.
- Despesa do próprio escritório (ele aparece como cliente no dropdown) não gera relatório: não há quem reembolse.
- **Status: não filtre por `approved`.** A tarefa caminha de `to do` → `approved` →
  `solicitado reembolso` → `pago pelo cliente` / `reembolsado pelo cliente`, então `approved` é
  estado transitório e some conforme o mês envelhece: junho/2026 não tinha uma única tarefa em
  `approved`, e filtrar por ele devolvia relatório vazio. O que separa é `to do`, que significa
  que o escritório ainda não pagou. O default do `build_relatorio.py` já é esse (entra tudo
  menos `to do`); use `--status` só para desviar dele, e diga ao Daniel o que desviou.

### 1. Rodar

O `CLICKUP_TOKEN` não está no ambiente e nunca vai estar: cada chamada de shell nasce limpa. Ele
mora no `.env` do projeto do CRM, fora do versionado, junto do `CLICKUP_FOLDER_ID` (a pasta do
ano) e do `CLICKUP_SPACE_ID` (o espaço dos adiantamentos), que saíram do código em 12/09/2026
para a skill poder ser pública. Quem sabe onde esse `.env` está é o `scripts/ambiente.sh`, que
também não é versionado, porque o caminho identifica o workspace de quem opera. Os três se
carregam de uma vez, a cada comando, da raiz da skill:

```bash
set -a && . scripts/ambiente.sh && set +a
```

Se o `ambiente.sh` não existir, copie o `ambiente.sh.exemplo` ao lado dele e ajuste o caminho.

Se a API responder `401 OAUTH_025` ("Token invalid"), o token foi revogado e nenhum ajuste de
script resolve: só Daniel gera outro em Settings > Apps > API Token e reescreve esse `.env`.

```bash
# mês inteiro
python3 scripts/rodar_adiantamentos.py --mes-clickup "07 Julho 2026" \
    --mes-nome Julho --ano 2026 --out ".../reembolso-despesas-adiantadas"

# um cliente só (o nome é a chave em clientes_config.json)
python3 scripts/rodar_adiantamentos.py --mes-clickup "07 Julho 2026" \
    --mes-nome Julho --ano 2026 --out ".../reembolso-despesas-adiantadas" \
    --clientes "<cliente>"
```

`--out` é a **raiz da frente**, não a pasta do mês: o orquestrador cria `input/{ano-mm}/` para o
bruto e `output/{Cliente}/` para a entrega, e tira o `{ano-mm}` do primeiro campo de
`--mes-clickup`. Chamando o fetch na mão, aponte `--out` para `input/{ano-mm}`, nunca para
`output/`, ou dezenas de pastas com nome de `task_id` se misturam às pastas de entrega.

### 1b. O que o orquestrador confere sozinho

Duas verificações rodam automaticamente e viram pendência quando falham, então leia as
pendências antes de olhar a planilha:

- **`conferir_planilha.py`** compara o `.xlsx` produzido com o contrato de
  [formato-planilha.md](references/formato-planilha.md): proporção da logo, âncora dentro da
  coluna B, alturas das linhas 2 e 3, larguras, fonte por linha, ordem cronológica, `SUM` com
  valor em cache, tamanho da janela e nome do arquivo. Existe porque **todo** defeito de formato
  de 2026 atravessou o pipeline calado e foi o Daniel quem viu, abrindo o arquivo.
- **Anexo solto na bandeja da tarefa.** O fetch agora lê as três portas de anexo e baixa como
  `tarefa_NN_` o que não veio nem pelo formulário nem por resposta de comentário. Arquivo assim
  não tem papel declarado, então a demanda inteira vai para `_revisar/`, mesmo que o par
  documento e comprovante já esteja formado.

Rodar o conferidor à mão, quando você editou a planilha depois de gerar:

```bash
python3 scripts/conferir_planilha.py ".../{Cliente}-{Mês}-{Ano}.xlsx" \
    --itens 15 --total 977.67 --nome-entrega "{Cliente}"
```

### 1c. Cliente acumulativo: a varredura é automática

Para cliente marcado `acumulativo` ou `pasta_fora_calendario` no `clientes_config.json`, o
orquestrador varre sozinho e transforma cada demanda não cobrada em pendência. Acumulativo já
quer dizer atrasado, então a condição mora no config, não na sua memória: num teste de
10/09/2026 a instrução condicional falhou, porque quem estava fechando agosto de um cliente
acumulativo não ligou aquilo com "esse cliente não paga desde janeiro", e a demanda de um Uber
sumiu de novo. Rode à mão quando o total parecer pequeno demais em qualquer outro cliente:

```bash
# ano corrente, que e o default
python3 scripts/varrer_cliente.py --cliente "{Nome no ClickUp}" \
    --entregas ".../output/{Cliente}"

# um ano fechado, so quando a cobranca for retroativa
python3 scripts/varrer_cliente.py --cliente "{Nome no ClickUp}" --ano 2025 \
    --entregas ".../output/{Cliente}"
```

Varre-se **um ano por vez**: o saldo de um cliente é uma conta anual, e o espaço inteiro tem 42
listas contra 9 do ano corrente. Ano fechado tem o status **não mantido**, então o que vier de
pasta `ARQUIVO` sai em `em_arquivo_ano_anterior`, como aviso e fora do total.

Ele compara por `task_id` contra os sidecars `.itens.json` já entregues e devolve
`nao_cobradas`. Em agosto/2026 uma demanda de R$ 15,94 estava na lista de junho e não entrou em
nenhum dos dois relatórios do cliente; foi achada por desconfiança, não por processo.

### 2. Ler as pendências e conferir a planilha

O resumo JSON traz uma lista de `pendencias`. Ela existe porque o script **sinaliza em vez de
adivinhar**: leia item por item, não é decoração. Depois confira o que só olho humano pega:
descrição que não faz sentido, valor que destoa, demanda faltando.

### 3. Podar os comprovantes

O Daniel tira da pasta o que não é prestação de contas (ver [O que não vai para o
cliente](#o-que-não-vai-para-o-cliente)) e acrescenta o que faltava. Isso é etapa fixa do
processo, não exceção.

### 4. Renumerar

A poda deixa buracos (`05.02` sem `05.01`) e os arquivos acrescentados vêm com o nome original
do banco:

```bash
# mostra o plano e lista os arquivos soltos que faltam atribuir
python3 scripts/renumerar_recibos.py --pasta ".../{Cliente}/Recibos e comprovantes" --dry-run

python3 scripts/renumerar_recibos.py --pasta ".../{Cliente}/Recibos e comprovantes" --limpar-lixo \
    --atribuir "Recibo_e-protocolo_20260708.pdf=01"
```

O script não adivinha de qual demanda é um arquivo solto: ele lista os pendentes e você
identifica pelo conteúdo, nunca pelo nome antigo.

### 5. Entregar

O resumo traz `depois_de_enviar.tarefas_por_cliente`, a lista de `task_id` que precisa sair de
`approved` para `solicitado reembolso` depois que a cobrança for enviada. **O status é
declaração de fato do Daniel e nunca se escreve nele por script**; o que se faz é dizer quais
são. Não dizer deixa o ClickUp atrasado em relação ao que o cliente já recebeu: em setembro/2026
nove tarefas de um cliente estavam paradas, uma desde a cobrança de 11/08.

Só depois da aprovação do Daniel. `--publicar-drive` copia para a pasta oficial do cliente
(`destino_drive` no config); cliente com `destino_drive: null` fica só local.

## O que continua sendo decisão humana

O script sinaliza estes casos como pendência e segue em frente. Nenhum deles é inferível dos
dados, e fingir que são é como se erra bonito:

**Brink ou comprovante.** Ver a seção abaixo. É a mais frequente e a mais fácil de errar.

**Valor de CASO** (clientes com essa coluna extra, ver `colunas_extra` no config). Não vem do ClickUp.
O script marca `REVISAR:` e pergunta.

**`Payment Date` que é vencimento, não pagamento.** Nas demandas que têm brink, compare o campo
do ClickUp com o "Data do Vencimento" impresso nele. Se baterem, a data boa é a do comprovante
bancário. Isso não afeta só a coluna F: a data define a ordem cronológica e a numeração de todas
as demandas, e arrasta os comprovantes junto.

**Tarefa que vira várias linhas.** Uma tarefa do ClickUp pode agrupar várias sub-cobranças (ex.:
uma notificação extrajudicial com 4 destinatários). O script gera 1 linha por tarefa; dividir
exige ler os recibos individuais.

**Cliente novo.** Sem entrada no `clientes_config.json`, pare para esse cliente e pergunte ao
Daniel o nome de entrega, as colunas extras e a pasta no Drive. Não invente configuração.

## O que não vai para o cliente

Nem todo arquivo que o ClickUp devolve é documento de entrega.

**Sai da pasta:** o **brink**, que é o pedido de compra de crédito ou prenotação (ONR, JUCEMG,
PagTesouro) com status "Aguardando pagamento" e QR Pix; e **prints de tela** com dados bancários
do cartório ou conversa de WhatsApp combinando valor. O brink é o pedido, não a prova: o
documento válido é o recibo oficial que a plataforma emite depois, geralmente anexado como
resposta de comentário dias depois.

**Fica na pasta:** boleto, recibo oficial, certidão, nota, e o comprovante de pagamento
(transferência ou Pix).

> ⚠️ **O caso automático é o mais perigoso, não o `_revisar`.** `organizar_recibos.py` numera
> sozinho a demanda que tem "1 documento + 1 comprovante", e num teste com julho/2026 ele
> numerou um brink como `10.01` exatamente por encaixar nesse molde: o pedido não pago no lugar
> do documento, e nada sinalizado. O que cai em `_revisar` você vai abrir de qualquer jeito; o
> que ele resolve sozinho é o que passa sem ninguém olhar. **Confira também os automáticos.**

> ⚠️ **Não decida pelo nome do arquivo. Abra e olhe.** Em julho/2026, no mesmo cliente,
> `ONR - <nome do caso>.pdf` era comprovante Pix de R$ 39,16 (fica) e `Pagametnto CCIR pix.pdf`
> era PagTesouro "Aguardando realização do pagamento" (sai). Filtrar por nome teria mandado o
> brink ao cliente e descartado a prova verdadeira.

A descrição também não vai crua: o `Task Content` costuma terminar com agência, conta e chave
Pix do favorecido, porque serve a quem paga. `build_relatorio.py` corta esse bloco e reporta o
corte em `pendencias`. Menção no meio da frase ("(dados: matrícula 00.000)") não é tocada.

## Comprovantes: nomenclatura e ordem

Formato `XX.YY.ext`, onde `XX` é a demanda com 2 dígitos e `YY` o arquivo dentro dela:
`01.01.jpeg`, `06.03.pdf`.

Dentro de cada demanda, documento primeiro e comprovante de pagamento por último. Com mais de
dois arquivos, todos os documentos e recibos primeiro, todos os comprovantes depois:

```
07.01.pdf   ← certidão
07.02.pdf   ← boleto ou recibo oficial
07.03.jpeg  ← comprovante de pagamento
```

A relação demanda ↔ arquivo **não é 1:1**. Um recibo único pode cobrir duas demandas (recibo
RCPN de R$ 47,24 = R$ 23,90 + R$ 23,34): duplique o arquivo nas duas. Uma demanda pode ter
vários recibos que somam o total (3 recibos ONR de R$ 65,31 = R$ 195,93).

Expectativa realista de esforço: em julho/2026, **19 das 22 demandas** caíram em `_revisar`,
porque `organizar_recibos.py` só resolve sozinho o caso de 1 documento + 1 comprovante. Isso não
é defeito do script, é a natureza do material. Reserve tempo para ler arquivo por arquivo, e
consulte [references/armadilhas.md](references/armadilhas.md) para as regras de casamento
(protocolo do comprovante bancário × pedido ONR, horário da sessão quando dois valores são iguais).

## Casos especiais

### Cliente com duas empresas no ClickUp

Duas empresas do mesmo dono, entregues juntas num único relatório. O config já cuida: filtra os
dois valores do ClickUp e acrescenta a coluna **Empresa** com o nome completo da pessoa
jurídica. Vale **somente** para esse par.

### Cobrança consolidada: o cliente não pagou um mês

Acontece com qualquer cliente. Passe `--incluir-xlsx` para cada mês em aberto; tudo é reordenado
por data e renumerado de 1 a N.

- **Nome do arquivo e título declaram o período, os dois.** Entregar dois meses num arquivo
  chamado só `...-Julho-2026.xlsx` esconde do cliente o que ele está pagando. O script deriva
  ambos das datas reais: `...-Junho-e-Julho-2026.xlsx` para dois meses, `...-Maio-a-Julho-2026`
  para três ou mais.
- ⚠️ **Os comprovantes do mês reaproveitado precisam ser remapeados.** Renumerar a planilha e
  esquecer a pasta manda o recibo para a demanda errada. A reordenação cronológica embaralha: em
  julho/2026 a demanda 1 de junho virou a 3, a 2 virou 1, a 3 virou 2. Monte o de-para pelo
  **valor**, que é único dentro do mês.
- Se o `.xlsx` reaproveitado tiver numeração antiga na coluna B, recupere os Task Names
  verdadeiros na lista daquele mês no ClickUp e aplique numa **cópia** em `_raw/`.
- O arquivo do mês não pago continua onde está e não é reescrito: ele é insumo, não entrega.

## Estrutura de pastas

Desde 10/09/2026 o insumo e a entrega vivem em ramos separados, `input/` por mês e `output/`
por cliente. O corte é esse porque o ClickUp entrega por mês e o cliente recebe por cliente: os
acumulativos têm uma planilha só cobrindo vários meses e nunca couberam numa pasta de mês.

```
reembolso-despesas-adiantadas/
  input/
    2026-07/                    ← download bruto do mês; não apagar antes de fechar os clientes
      manifest.json
      <task_id>/                ← uma pasta por task_id
    <Cliente>/                  ← quando o fetch é por cliente, caso retroativo
  output/
    <Cliente>/                  ← mensal: uma subpasta por mês
      2026-07/
        <Cliente>-Julho-2026.xlsx
        Recibos e comprovantes/
          01.01.pdf
          01.02.jpeg
    <Cliente acumulativo>/      ← sem subpasta de mês (pasta_fora_calendario no config)
      <Cliente>-Janeiro-a-Julho-2026.xlsx
      Recibos e comprovantes/
  refs/                         ← logo-referencia.xlsx, o `--ref` de quem não tem mês anterior
  scripts/                      ← scripts avulsos de mês fechado, não a skill
```

A pasta do cliente contém **apenas** o `.xlsx` e a subpasta de comprovantes, sem arquivo solto.
`input/` preserva os brinks descartados e os arquivos ainda não usados; os caminhos dentro dos
`.xlsx.itens.json` apontam para lá, então movê-la depois de gerar exige reescrever esses
caminhos.

## Checklist de entrega

O que o script já garante sozinho não está aqui de propósito: conferir o que o código impõe é
como as duas fontes divergem sem ninguém notar. Confira o que depende de julgamento:

- [ ] Lista relida imediatamente antes de gerar, arquivadas incluídas, clientes confrontados com o config
- [ ] Escopo respeitado: gerou os clientes pedidos, nem mais nem menos
- [ ] Status: rodou com o default (tudo menos `to do`) ou avisou o Daniel do desvio
- [ ] Todas as `pendencias` do resumo lidas e resolvidas
- [ ] Coluna B preenchida em todas as linhas (nenhuma demanda sem Task Name)
- [ ] Demandas com brink: `Payment Date` conferido contra o vencimento impresso
- [ ] Brinks e prints de tela fora da pasta de entrega
- [ ] Cada comprovante identificado pelo conteúdo, não pelo nome do arquivo
- [ ] Numeração contígua: cada demanda começa em `.01`, sem buracos
- [ ] Relação 1:N e N:1 tratada (arquivo duplicado quando um recibo cobre duas demandas)
- [ ] Consolidada: comprovantes do mês reaproveitado remapeados pelo valor
- [ ] Recalc sem erros e logo presente no arquivo final
- [ ] Publicação no Drive só depois da aprovação do Daniel

## Referências

- [references/formato-planilha.md](references/formato-planilha.md), especificação do formato
  que o `build_relatorio.py` implementa. Consulte ao mexer no script ou investigar uma célula.
- [references/armadilhas.md](references/armadilhas.md), erros já cometidos, com causa e
  solução: datas, API do ClickUp, arquivos, identificação de comprovantes.
