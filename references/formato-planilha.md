# Formato da planilha: o contrato que `build_relatorio.py` implementa

Este arquivo é a **especificação do formato**, não um roteiro de execução. `build_relatorio.py`
aplica tudo isto sozinho, toda vez. Leia aqui quando for **mexer no script**, conferir se ele
está cumprindo o contrato, ou entender por que uma célula saiu de um jeito.

Se você está só gerando o relatório do mês, não precisa deste arquivo: rode o script.

> **Quem faz valer:** `scripts/conferir_planilha.py` compara o `.xlsx` pronto com quase tudo o
> que está escrito aqui, e o `rodar_adiantamentos.py` o executa depois do `recalc.py`, virando
> pendência quando falha. Regra nova nesta página que possa ser verificada em código deve ganhar
> uma linha lá, senão volta a ser decoração.
>
> ⚠️ Esta especificação e o código já divergiram uma vez. A SKILL.md mandava cabeçalho tamanho
> 14 enquanto o script aplicava 12, e ninguém percebeu por semanas. Quando mudar o formato,
> mude o script **e** este arquivo na mesma passada, ou a documentação vira ficção.

## Índice

- [Estrutura da aba "Tasks"](#estrutura-da-aba-tasks)
- [Geometria da linha 2](#geometria-da-linha-2)
- [Formatação geral](#formatação-geral)
- [Formatação por coluna](#formatação-por-coluna)
- [Cores](#cores)
- [Dimensões de referência](#dimensões-de-referência)
- [Injeção da logo](#injeção-da-logo)
- [Tamanho da janela ao abrir](#tamanho-da-janela-ao-abrir)
- [Corrigir um valor depois de gerar](#corrigir-um-valor-depois-de-gerar)

## Estrutura da aba "Tasks"

| Linha | B | C | D | E | F | G | H *(se o cliente tiver coluna extra)* |
|---|---|---|---|---|---|---|---|
| 2 | *(logo)* | título mesclado, C até a antepenúltima coluna |, |, |, | "Valor total:" *(penúltima)* | `=SUM(D4:D{última})` *(última)* |
| 3 | Demanda | Descrição | VALOR | Metodo de pagamento | Data de pagamento | ÓRGÃO | Empresa \| CASO |
| 4+ | Task Name | Task Content | valor | método | data | órgão | valor da coluna extra |

O rótulo da coluna H vem de `colunas_extra` em `clientes_config.json`, nunca fixo no código:
"Empresa" num cliente, "CASO" noutro. Deixar fixo já produziu um
relatório com o cabeçalho do cliente errado.

## Geometria da linha 2

A linha 2 é cabeçalho e precisa terminar exatamente onde a tabela termina, com borda thin em
todas as células. Linha 2 mais curta que as linhas de dados dá a impressão de planilha quebrada.

| Cliente | Título | Rótulo | Soma |
|---|---|---|---|
| Sem coluna H | `C2:E2` | F2 | G2 |
| Com coluna H | `C2:F2` | G2 | H2 |

Generalizando: a soma vai **sempre na última coluna**, o rótulo na penúltima, e o título mescla
de C até a anterior a essas duas.

A borda tem que ser aplicada explicitamente, inclusive nas células mescladas. Só `.value` é
somente-leitura numa `MergedCell`; estilo funciona. Numa faixa mesclada o Excel desenha a borda
a partir de cada célula de baixo, então pular as mescladas deixa a caixa do título aberta à
direita. E herdar a borda do `--ref` não resolve: no arquivo de referência a última coluna
costuma estar dentro de outro merge, e volta sem borda.

## Formatação geral

- **Calibri em tudo**, mesmo que o `--ref` use Arial.
- **Tamanho:** linha 2 (título) = 14, linha 3 (cabeçalho) = 14, linhas de dados (4+) = 12.
  Fixar explicitamente por linha; herdar do `--ref` foi o que produziu cabeçalho 12.
- **Bordas:** thin `FF000000` em tudo, da linha 2 em diante.
- **Bold:** só nas linhas 2 e 3.
- **Alinhamento dos dados:** `horizontal=center`, `vertical=center`, `wrap_text=True`.
- **Altura da linha 2:** 66.0 pt, obrigatório para a logo caber.
- **Altura da linha 3:** 42.0 pt.
- **Altura das linhas de dados:** `None` (autofit). Herdar alturas fixas do `--ref` gera espaço
  em branco enorme quando o texto novo é mais curto.

## Formatação por coluna

| Coluna | Formato | Observação |
|---|---|---|
| B | texto | Sempre `Task Name`. Nunca numeração sequencial. |
| C | texto centralizado | Sempre `Task Content`, o texto longo. Inverter B e C é o erro clássico. |
| D | `R$ #,##0.00` | Largura mínima 16, senão vira `######` |
| E | texto centralizado + fundo colorido | ver tabela abaixo |
| F | `DD/MM/YYYY` | converter para objeto `date` antes de salvar |
| G | texto centralizado + fundo colorido | ver tabela abaixo |
| H | texto centralizado, sem fundo | header com fill `FFE0E0E0` |

## Cores

Método de pagamento (E), fonte branca em todas:

| BOLETO | PIX | CRÉDITO | DINHEIRO |
|---|---|---|---|
| `FF1BBC9C` | `FF3082B7` | `FF02BCD4` | `FFE91E63` |

Órgão (G), fonte branca em todas:

| RI | JUCEMG | NOTAS | PREFEITURA | CARTÓRIO | MOTOBOY | JUNTA | RTD | OUTROS / TJ | ONR |
|---|---|---|---|---|---|---|---|---|---|
| `FF6A1B9A` | `FF1BBC9C` | `FFFF9800` | `FFE65100` | `FF607D8B` | `FF795548` | `FF1BBC9C` | `FF3F51B5` | `FF1BBC9C` | `FF00796B` |

Valor fora da paleta (RCPN, DIVERSOS) cai num cinza neutro `FFBDBDBD`. Cor é desejável, não
crítica: nunca segure uma entrega por causa de cor de fundo.

## Dimensões de referência

Larguras: B `27.83203125` (coluna do logo, ajustada em 09/09/2026) · C `96.0` · D `16.0` · E `30.51` · F `26.5` · G `21.83` · H `24.0`

## Injeção da logo

A logo é um drawing "cru" copiado do arquivo de referência via `zipfile`, não a API
`openpyxl.drawing.image.Image` (que ignora o posicionamento original e redimensiona pelo
tamanho bruto do PNG).

A parte contraintuitiva: `load_workbook()` + `save()` **destrói** o drawing mesmo quando o
arquivo de origem já tinha a logo. O openpyxl só preserva imagens que ele representa como
objeto `Image()` nativo; um drawing injetado por fora não sobrevive ao roundtrip. Por isso a
reinjeção reescreve `xl/drawings/drawing1.xml`, o rels dele,
`xl/worksheets/_rels/sheet1.xml.rels` e a imagem **incondicionalmente**, sem checar se já
existem, porque não vão existir.

Ordem obrigatória, e ela não perdoa: **openpyxl `save()` → injetar logo → recalcular com
LibreOffice**. Qualquer `save()` do openpyxl depois da injeção apaga o drawing de novo. O
recalc do LibreOffice preserva.

Se o `.xlsx` do próprio cliente não abrir (`Resource deadlock avoided`, `BadZipFile`, típico de
falha temporária do mount do Google Drive), use como `--ref` o relatório mais recente de
qualquer outro cliente: a logo é a mesma do escritório em todos.

**A geometria não vem da referência.** Copiar o drawing inteiro trazia junto o tamanho que a
imagem tinha no arquivo de origem, e cada cliente carregava uma proporção diferente, todas
achatadas. Desde 09/09/2026 a `normalizar_ancora_logo()` reescreve posição e tamanho para a
constante `LOGO_ANCORA`, em EMU: `from` colOff `95920` / rowOff `160081`, `to` colOff
`1940000` / rowOff `678118`, `ext` cx `1844080` / cy `518037`. Qualquer `--ref` serve, a
geometria é sobrescrita.

**A âncora tem que caber na coluna B.** O `recalc.py` do passo 5 reescreve o drawing e corta o
que passa da borda da coluna, o que achata a imagem sem avisar: uma âncora terminando em
`2044700`, como a que o Daniel montou no Excel, volta do recalc com razão `3,41` em vez de
`3,56`. O Excel tolera o transbordo, o LibreOffice não. Por isso `to_col_off` é `1940000`, com
folga sobre a borda. Mudou a largura da coluna B, revisar este limite.

**A regra é a proporção, não o número.** A razão do PNG original, `566x159`, é `3,5597`, e a
altura da âncora é a largura dividida por ela, o que zera o achatamento. A largura veio de um
relatório de agosto/2026 dimensionado à mão pelo Daniel. Quem mudar a largura recalcula a altura pela
mesma conta e recentraliza na linha 2, que tem 66 pt, ou seja 838200 EMU. O valor que o script
usava antes dava razão `2,40`, um terço fora, e era o que deixava a logo espremida em todo
relatório.

## Tamanho da janela ao abrir

O Excel abre a planilha com o tamanho de janela gravado em `xl/workbook.xml`, no atributo
`windowWidth`/`windowHeight` de `<workbookView>`, em twips (1/20 de ponto). Quem escreve esse
valor é o LibreOffice do passo 5, e o default dele é `16384x8192`, meia tela: a planilha abria
mostrando até a coluna C e o usuário maximizava toda vez.

Desde 10/09/2026 o `recalc.py` reescreve os dois para `38400x21600`, o equivalente a 2560x1440
px a 96 dpi. São valores maiores que a tela de propósito: o Excel encolhe a janela para o que
couber, então a planilha abre maximizada em qualquer monitor. A reescrita é cirurgia no zip, e
não `openpyxl`, pelo motivo de sempre: reabrir e salvar com `openpyxl` apagaria a logo.

## Nome do arquivo

`{Nome do cliente}-{Período}-{Ano}.xlsx`, com **espaço virando hífen**, não sumindo: um cliente
chamado "Alfa e Beta" sai `Alfa-e-Beta-Agosto-2026.xlsx`. O script colava as palavras, escrevendo
`AlfaeBeta`, e o Daniel renomeou à mão o arquivo de agosto/2026; o hífen é o padrão que os meses
anteriores daquele cliente já usavam.

## Corrigir um valor depois de gerar

1. `load_workbook`, alterar a célula D (linha 4 = demanda 1, linha 5 = demanda 2, …)
2. `ws.row_dimensions[2].height = 66.0`
3. `wb.save()`, a última chamada ao openpyxl
4. reinjetar a logo via `zipfile`
5. `recalc.py`, que preserva o drawing e grava o total no cache do SUM

Antes de fazer isso, pergunte-se se o dado errado não está no ClickUp. Se estiver, corrija lá e
regere: editar o `.xlsx` à mão cria uma verdade que o próximo `build_relatorio.py` apaga.
