# reembolso-despesas-adiantadas

Automação em Python do fechamento mensal de reembolsos de despesas em um escritório de
advocacia: do dado bruto no ClickUp até o `.xlsx` formatado e os comprovantes numerados,
salvos na pasta de cada cliente no Google Drive, prontos para envio. O mês inteiro sai de um
comando, e o trabalho humano que sobra é só o que não é inferível do dado. São 8 scripts e
2.300 linhas.

## O problema

O escritório adianta custas de cartório, junta comercial e registro de imóveis com o próprio
caixa e cobra o reembolso dos clientes depois. Todo mês isso vira uma prestação de contas por
cliente: uma planilha com uma linha por despesa, no layout padrão do escritório, e uma pasta de
comprovantes numerados na mesma ordem. É documento que sai do escritório e chega a quem paga a
conta, então erro ali não é defeito de software, é valor cobrado errado.

O volume é modesto e o trabalho não. Um mês típico tem em torno de 40 despesas espalhadas por
dez clientes, e essas 40 despesas carregam cerca de 110 arquivos. Nada disso chega junto nem
pelo mesmo caminho. O boleto vem no campo do formulário da tarefa, o comprovante bancário
aparece dias depois como resposta de comentário, e o recibo oficial depois disso.

Montar a entrega à mão era isto: abrir a planilha do mês anterior como modelo, exportar o CSV do
ClickUp e ir colando de um para o outro, cliente por cliente. Depois baixar cada recibo e cada
comprovante, um a um, e renomear um a um, na ordem da planilha. Dois a três dias de trabalho,
todo mês, por mais de três anos.

## Um comando, o mês inteiro

Um comando cobre o mês inteiro: lê o ClickUp, monta a planilha de cada cliente no layout do
escritório, confere o arquivo produzido, numera os comprovantes e deixa cada pasta pronta para
enviar. Dez clientes, 40 despesas, 110 arquivos, sem ninguém abrir tarefa.

O ganho é tempo, e é o grosso daqueles dois a três dias. Ele nunca esteve na planilha: estava em
abrir as 40 tarefas, baixar os 110 arquivos e descobrir, um por um, qual pertence a qual despesa.
A prestação de contas sempre saiu correta, e saiu correta porque alguém conferia esses 110
arquivos todo mês para que ela saísse. O que a automação eliminou foi essa conferência, não o
cuidado dela.

Automatizar, por outro lado, cria um risco que o trabalho manual não tinha: errar rápido e em
silêncio. Daí o pipeline levar junto um conferidor que valida cada arquivo produzido contra a
especificação escrita, e daí a violação virar pendência em vez de exceção. Há também uma coisa
que a máquina enxerga e a conferência humana não enxergava, porque ninguém olha doze meses ao
mesmo tempo: a despesa paga e arquivada na lista do mês errado, que a varredura anual encontrou
valendo R$ 15,94 e dois meses de atraso.

A escala vem de graça: cliente novo é uma entrada em arquivo de configuração, não uma variante
do script, então o próximo custa o mesmo que o anterior.

## Como a skill roda

O orquestrador chama cinco scripts em toda execução, um sexto só nos clientes de cobrança
acumulativa, e publica no Drive por função importada. O oitavo, o renumerador, é manual de
propósito. Nada disso trava no meio: cliente com erro não impede os demais.

```
 rodar_adiantamentos.py --out <raiz da frente> --mes-clickup "09 Setembro 2026"
 │
 ├─ deriva 2026-09 do "09" e monta input/ e output/
 │
 ├─(1) clickup_adiantamentos_fetch.py ─────────►  input/2026-09/
 │     API direta, sem conector pronto              ├ manifest.json
 │     archived=false + archived=true               └ <task_id>/arquivos
 │     lê as 3 gavetas de anexo, dedup por id
 │
 └─ POR CLIENTE
    │
    ├─(2) build_relatorio.py ─────────────────►  <Cliente>-Setembro-2026.xlsx
    │     filtra cliente e status                  <...>.xlsx.itens.json
    │     monta dados e planilha
    │     wb.save()        ◄── destrói a logo
    │     injetar_logo()   ◄── cirurgia no zip, por fora do openpyxl
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

## Cinco decisões e o porquê de cada uma

Quase todas nasceram de algo que já tinha dado errado: uma alternativa que parecia boa e cobria
menos da metade dos casos, um relatório entregue com o cabeçalho de outro cliente, defeitos de
formato que atravessaram o ano inteiro sem ninguém ver.

**Sinalizar em vez de adivinhar.** A relação entre comprovante e despesa não é um para um:
um recibo pode cobrir duas despesas, uma despesa pode ter três recibos que somam o total, e uma
tarefa pode agrupar várias cobranças. O `organizar_recibos.py` só numera sozinho o caso
inequívoco de um documento mais um comprovante; todo o resto vai para `_revisar/` e vira
pendência impressa no resumo, o que num mês real foi 19 das 22 despesas. A alternativa foi
testada e descartada: casar a data do campo com a data embutida no nome dos arquivos cobria 12
de 26 tarefas e gerava um alerta falso, porque certos recibos são emitidos dias depois do
pagamento. E a parte contraintuitiva é que o caso automático é o mais perigoso, não o `_revisar`:
o que cai em revisão vai ser aberto de qualquer jeito, e o que o script resolve sozinho é o que
passa sem ninguém olhar. Num teste, um pedido de pagamento não quitado encaixou no molde de um
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

**A logo obriga uma ordem, e só uma.** Ela não entra pela API de imagens do openpyxl, que ignora
o posicionamento original e redimensiona pelo tamanho bruto do PNG, então é copiada como drawing
cru, reescrevendo o zip do `.xlsx` por fora. Só que o openpyxl destrói qualquer drawing que ele
não represente como objeto nativo, o que obriga a logo a entrar depois do último `save()`. E ela
também não pode ser a última etapa, porque o arquivo sairia com a fórmula de soma sem valor em
cache e quem abrisse a planilha pelo navegador veria a célula do total vazia. Quem grava esse
cache é o LibreOffice headless, que preserva o drawing. Sobra uma ordem só, salvar, injetar,
recalcular, e é por isso que o `recalc.py` é o lugar de qualquer ajuste final no arquivo.

Esse mesmo passo controla a geometria. A âncora é reescrita para uma constante derivada da razão
real do PNG, 3,5597, e termina com folga dentro da coluna B, porque o Excel tolera transbordo de
âncora e o LibreOffice corta, achatando a imagem sem avisar.

**A publicação no Drive não usa integração nenhuma.** O passo 7 escreve no lugar certo do disco
local e o cliente do Google Drive sincroniza sozinho. Sem OAuth, sem escopo de API, sem token
para renovar, sem tratar retomada de upload. O custo é a dependência de uma pasta montada, que
falha de um jeito visível e local, não de um jeito silencioso e remoto.

## Engenharia reversa da API do ClickUp

A forma do pipeline é consequência direta da forma do dado, mapeada contra tarefas reais e não
contra a documentação da API. Esta seção termina numa falha real, e é por isso que ela existe:
o que a gente achou que sabia sobre esse dado custou três prestações de contas incompletas.

Uma despesa quase nunca tem um arquivo só. Um boleto de verdade dá dois, o documento e o
comprovante. Um pedido de crédito ou QR Pix, o caso frequente, dá três, porque a plataforma só
emite o recibo oficial depois de compensar. Uma tarefa com vários boletos dá quatro ou mais. Daí
a ordem dentro da despesa importar, documento primeiro e comprovante por último, e o fetch
entregar os arquivos já agrupados por tarefa.

E eles chegam por três portas, o que é a razão de o script falar HTTP direto com a API em vez
de usar o servidor MCP, o conector pronto que entrega as ferramentas do ClickUp a um agente: o
campo do formulário, o anexo de resposta de comentário, e o array `attachments` da própria
tarefa. A segunda é o caminho normal do comprovante e exige duas chamadas, porque o
comentário-pai vem sem o anexo e a URL só aparece em `GET /comment/{id}/reply`; a ferramenta
MCP equivalente descarta esse campo e devolve só o nome do arquivo dentro do texto, o que
inviabiliza o download.

A terceira porta não é paralela às outras, e entender isso é o que faz o fetch ficar simples: ela
é o inventário completo, superconjunto das demais, com os mesmos ids. As duas primeiras não dizem
o que existe, dizem o papel de cada arquivo, e é o papel que determina a numeração. O que aparece
só na terceira é arquivo que alguém largou na tarefa sem classificar, então o fetch lê as três,
deduplica por id, e manda a despesa inteira para `_revisar/` quando sobra algo sem papel
declarado. Descobrir isso custou caro: o plano original registrou aquele array como "vazio em
toda tarefa amostrada", a amostra virou premissa, e três clientes receberam a prestação de
contas sem um recibo que estava no ClickUp o tempo todo.

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

Em produção desde julho de 2026, em fechamentos mensais consecutivos. O processo que ela
substituiu era feito à mão desde 2023.

O que o mês virou, no fim, foi isto: a máquina faz a parte que é dedução, e devolve por escrito
a parte que não é. Abrir arquivo e decidir continua sendo trabalho humano. Deixou de ser
trabalho humano procurar o arquivo.

## Privacidade e licença

O pipeline lida com dados financeiros de clientes: valores, métodos de pagamento e
comprovantes com nomes e contas. Nada disso é versionado. O `.gitignore` bloqueia CSVs,
planilhas geradas, recibos e o próprio registro de clientes, e os diretórios de trabalho
vivem fora do repositório.

Código proprietário, publicado como estudo de caso. Todos os direitos reservados: repositório
público não é código aberto, e ler para avaliar o trabalho é a única permissão concedida. Os
termos estão em [LICENSE](LICENSE).
