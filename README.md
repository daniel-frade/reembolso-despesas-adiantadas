# reembolso-despesas-adiantadas

Automação em Python do fechamento mensal de reembolsos de despesas em um escritório de
advocacia: do dado bruto no ClickUp até o `.xlsx` formatado e os comprovantes numerados,
publicados na pasta oficial de cada cliente no Google Drive. Um comando por mês, 8 scripts,
cerca de 2.500 linhas.

## O problema

O escritório paga custas de cartório, junta comercial e registro de imóveis com o próprio
caixa, e cobra dos clientes depois. Cada despesa nasce como uma tarefa no ClickUp, aberta
por formulário, com o boleto ou o pedido de pagamento anexado no campo do formulário. O
comprovante bancário chega dias depois, anexado como resposta de um comentário, e o recibo
oficial que algumas plataformas só emitem após o pagamento chega depois disso.

No fim do mês, alguém transforma esse material em prestação de contas por cliente: uma
planilha com uma linha por despesa, no layout que aquele cliente já recebe há meses, mais uma
pasta de comprovantes numerados na mesma ordem da planilha. À mão, isso era abrir tarefa por
tarefa, baixar anexo por anexo, montar a planilha e depois casar cada comprovante com cada
despesa lendo valor, data e favorecido dentro de cada arquivo, com o script de geração
reescrito por cliente, todo mês. Hoje o mês inteiro sai de um comando, e o trabalho humano
que sobra é só o que não é inferível do dado.

## Como a coisa roda

O orquestrador chama os outros sete scripts e nunca trava no meio: cliente com erro não
impede os demais.

```
 rodar_adiantamentos.py --out <raiz da frente> --mes-clickup "09 Setembro 2026"
 │
 ├─ deriva 2026-09 do "09" e monta input/ e output/
 │
 ├─(1) clickup_adiantamentos_fetch.py ─────────►  input/2026-09/
 │     API direta, sem MCP                          ├ manifest.json
 │     archived=false + archived=true               └ <task_id>/arquivos
 │     lê as 3 gavetas de anexo, dedup por id
 │
 └─ POR CLIENTE
    │
    ├─(2) build_relatorio.py ─────────────────►  <Cliente>-Setembro-2026.xlsx
    │     filtra cliente e status                  <...>.xlsx.itens.json
    │     monta dados e planilha
    │     wb.save()        ◄── destrói a logo
    │     injetar_logo()   ◄── cirurgia no zip, razão 3,5597
    │
    ├─(3) recalc.py ──────────────────────────►  o mesmo .xlsx
    │     LibreOffice: grava o SUM em cache,
    │     preserva o drawing, abre maximizado
    │
    ├─(4) conferir_planilha.py                    violação vira pendência
    │     compara com references/formato-planilha.md
    │
    ├─(5) varrer_cliente.py                       demanda não cobrada vira
    │     só cliente acumulativo, ano corrente     pendência
    │
    ├─(6) organizar_recibos.py ───────────────►  Recibos e comprovantes/
    │     1 documento + 1 comprovante → XX.01, XX.02
    │     qualquer outro caso        → _revisar/XX/
    │
    └─(7) --publicar-drive (opcional) ────────►  pasta oficial do cliente

 Depois, o que depende de julgamento humano:
   ler as pendências · podar os pedidos de pagamento · renumerar_recibos.py ·
   enviar · mover as tarefas para o status seguinte
```

## As decisões de projeto

**A ordem entre 2 e 3 é a única possível.** A logo do escritório não entra pela API de
imagens do openpyxl, que ignora o posicionamento original e redimensiona pelo tamanho bruto
do PNG. Ela é copiada como drawing cru, reescrevendo o zip do `.xlsx` por fora. Só que o
openpyxl destrói qualquer drawing que ele não represente como objeto nativo, então a logo
tem que entrar depois do último `save()`. E ela também não pode ser a última etapa, porque o
arquivo sairia com a fórmula de soma sem valor em cache, e quem abrisse a planilha pelo
navegador veria a célula do total vazia. Quem grava esse cache é o LibreOffice em modo
headless, que preserva o drawing. Sobra uma ordem só, salvar, injetar, recalcular, e é por
isso que o `recalc.py` é o lugar de qualquer ajuste final, inclusive a reescrita do tamanho
de janela que faz a planilha abrir maximizada.

O mesmo passo controla a geometria: a âncora é reescrita para uma constante em EMU derivada
da razão real do PNG, 3,5597, e termina com folga dentro da coluna B, porque o Excel tolera
transbordo de âncora e o LibreOffice corta, achatando a imagem sem avisar.

**A publicação no Drive não usa integração nenhuma.** O passo 7 escreve no lugar certo do
disco local e o cliente do Google Drive sincroniza sozinho. Sem OAuth, sem escopo de API,
sem token para renovar, sem tratar retomada de upload. O custo é a dependência de uma pasta
montada, que falha de um jeito visível e local, não de um jeito silencioso e remoto.

**Sinalizar em vez de adivinhar.** A relação entre comprovante e despesa não é um para um:
um recibo único pode cobrir duas despesas, uma despesa pode ter três recibos que somam o
total, e uma tarefa pode agrupar várias cobranças. O `organizar_recibos.py` só numera sozinho
o caso inequívoco de um documento mais um comprovante; todo o resto vai para `_revisar/` e
vira pendência impressa no resumo. Num mês real, 19 das 22 despesas caíram em `_revisar`,
e isso é a natureza do material, não lacuna do código. A alternativa foi testada e
descartada: casar a data do campo com a data embutida no nome dos arquivos cobria 12 de 26
tarefas e gerava um alerta falso, porque certos recibos são emitidos dias depois do
pagamento. Um alerta que erra 1 em 12 treina a ignorar alertas, e um script que acerta quase
sempre treina a confiar sem conferir.

A parte contraintuitiva é que o caso automático é mais perigoso que o `_revisar`: o que cai
em revisão vai ser aberto de qualquer jeito, e o que o script resolve sozinho é o que passa
sem ninguém olhar. Num teste, um pedido de pagamento não quitado encaixou no molde de um
documento mais um comprovante e foi numerado como se fosse o documento oficial.

**Cliente novo é configuração, não código.** Nome no ClickUp, nome de entrega, colunas
extras, periodicidade e pasta de destino no Drive vivem em um `clientes_config.json` fora do
versionado, hoje com seis entradas. Cliente com duas empresas do mesmo dono entra como um
combo que sai num relatório só; cliente que não paga mensalmente entra com pasta fora do
calendário e ganha, de brinde, uma varredura automática do ano atrás de despesa que nunca foi
cobrada. Nenhum desses casos toca no código. O rótulo da coluna extra vem do config pelo
mesmo motivo: quando era constante no script, um relatório saiu com o cabeçalho de outro
cliente.

**Um conferidor que valida o arquivo produzido contra o contrato de formato.** O formato da
planilha está escrito em `references/formato-planilha.md`, e o `conferir_planilha.py` compara
o `.xlsx` pronto com quase tudo o que está lá: proporção da logo, âncora dentro da coluna B,
alturas de linha, larguras, fonte, ordem cronológica, total em `SUM` com valor em cache, nome
do arquivo sem espaço. São vinte e poucas asserções rodando depois do recalc, e a violação
vira pendência, não exceção. Ele existe porque todo defeito de formato do ano atravessou o
pipeline calado e foi descoberto por um humano abrindo o arquivo. Regra nova na especificação
que possa ser verificada em código ganha uma linha no conferidor, senão a especificação vira
decoração e diverge do script, o que já aconteceu uma vez.

## De onde vêm os arquivos no ClickUp

Esta seção existe porque a forma do pipeline é consequência direta da forma do dado. Foi
mapeada contra tarefas reais, não contra a documentação da API.

Uma despesa quase nunca tem um arquivo só. Um boleto de verdade dá dois arquivos, o documento
e o comprovante de pagamento. Um pedido de crédito ou QR Pix, o caso frequente, dá três: o
pedido, o comprovante, e o recibo oficial que a plataforma só emite depois de compensar. Uma
tarefa com vários boletos dá quatro ou mais. É por isso que a ordem dentro da despesa importa,
documento primeiro e comprovante por último, e por que o fetch entrega os arquivos já
agrupados por tarefa.

E eles chegam por três portas diferentes, o que é a razão de o script falar HTTP direto com a
API em vez de usar o servidor MCP. A primeira é o campo customizado do formulário. A segunda
é o anexo de resposta de comentário, o caminho normal do comprovante: `GET /task/{id}/comment`
traz o comentário-pai sem o anexo, e quem traz a URL é `GET /comment/{id}/reply`, em
`comment[].attachment.url`. A ferramenta MCP equivalente descarta esse campo e devolve só o
nome do arquivo dentro do texto, o que inviabiliza o download. A terceira é o array
`attachments` da própria tarefa.

A terceira não é uma porta paralela às outras duas, e entender isso é o que faz o fetch ficar
simples: ela é o inventário completo, superconjunto das outras, com os mesmos ids de anexo.
As duas primeiras não dizem o que existe, dizem o papel de cada arquivo, documento inicial
contra comprovante de pagamento, e é esse papel que determina a numeração. O que aparece só
na terceira é arquivo que alguém largou na tarefa sem classificar. O fetch lê as três e
deduplica por id: o que sobra sem papel declarado manda a despesa inteira para `_revisar/`,
porque arquivo sem papel não se numera por adivinhação.

Essa correção custou três entregas erradas antes de existir. O plano original registrou o
array de anexos como "vazio em toda tarefa amostrada", e a amostra virou premissa: dois
clientes receberam a prestação de contas sem um recibo que estava no ClickUp o tempo todo, e
num terceiro um comprovante ficou dois meses pendurado na despesa errada.

## Estrutura

Os oito scripts ficam em `scripts/`, junto do registro de clientes. O `SKILL.md` é o
procedimento de operação, `references/formato-planilha.md` é a especificação do formato que o
`build_relatorio.py` implementa, `references/armadilhas.md` guarda os erros já cometidos com
causa e solução, e `evals/` tem os casos de teste. A regra que organiza os três documentos:
`build_relatorio.py` é a única implementação do formato, então saída errada é defeito do
script e se corrige lá, porque consertar só o arquivo do mês faz o erro voltar no mês
seguinte.

Depende de Python 3 com openpyxl e de um LibreOffice para o recalc. O token da API vive em
variável de ambiente, nunca em arquivo versionado.

## Privacidade e licença

O pipeline lida com dados financeiros de clientes: valores, métodos de pagamento e
comprovantes com nomes e contas. Nada disso é versionado. O `.gitignore` bloqueia CSVs,
planilhas geradas, recibos e o próprio registro de clientes, e os diretórios de trabalho
vivem fora do repositório.

Código privado, de uso interno, todos os direitos reservados.
