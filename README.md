# Formatador de Planilhas

Programa com janela gráfica para limpar e padronizar arquivos **CSV** e **Excel** (`.xlsx`) sem precisar abrir cada planilha e corrigir tudo na mão.

Ele resolve problemas comuns em planilhas exportadas de sistemas e de outras pessoas:

- espaços sobrando no começo, no fim e entre palavras (`"  Ana   Maria "` → `"Ana Maria"`);
- linhas e colunas vazias espalhadas pela tabela;
- linhas repetidas;
- cabeçalho com caixa inconsistente (`nome completo` → `NOME COMPLETO`);
- CSV com separador `;` (padrão do Excel em português), codificação cp1252 ou UTF-8 com BOM;
- números brasileiros (`1.234,56`) que o Excel trata como texto;
- conversão entre os formatos: CSV → Excel e Excel → CSV.

Você vê o resultado na tela antes de salvar.

---

## Sumário

1. [Requisitos](#requisitos)
2. [Instalação](#instalação)
3. [Como usar](#como-usar)
4. [Regras de formatação](#regras-de-formatação)
5. [Formatos de entrada e saída](#formatos-de-entrada-e-saída)
6. [Segurança dos seus dados](#segurança-dos-seus-dados)
7. [Exemplos](#exemplos)
8. [Solução de problemas](#solução-de-problemas)
9. [Limitações conhecidas](#limitações-conhecidas)
10. [Testes](#testes)
11. [Estrutura do projeto](#estrutura-do-projeto)

---

## Requisitos

- **Python 3.10 ou superior** com suporte a Tk (tkinter). Testado com Python 3.14 e Tk 9.1 no macOS.
- As bibliotecas do arquivo [requirements.txt](requirements.txt):
  - `openpyxl`: lê e grava arquivos `.xlsx`.
  - `tkinterdnd2`: habilita arrastar e soltar o arquivo na janela. É opcional: sem ela, o programa funciona normalmente com o botão **Procurar...**.

> **Atenção no macOS:** o Python que vem com o Xcode usa Tk 8.5, que pode travar ao abrir a janela. Se isso acontecer, use um Python do Homebrew ou do python.org (veja [Solução de problemas](#solução-de-problemas)).

---

## Instalação

### macOS (com Homebrew)

```bash
brew install python-tk@3.14

cd /Users/casmtts/Projetos/formatador-planilhas
/opt/homebrew/bin/python3.14 -m venv .venv
./.venv/bin/pip install -r requirements.txt
```

### Windows

```powershell
cd C:\caminho\para\formatador-planilhas
py -3 -m venv .venv
.\.venv\Scripts\pip install -r requirements.txt
```

### Linux (Debian/Ubuntu)

```bash
sudo apt install python3-tk
cd /caminho/para/formatador-planilhas
python3 -m venv .venv
./.venv/bin/pip install -r requirements.txt
```

O uso de um **ambiente virtual** (`.venv`) mantém as bibliotecas isoladas do restante do sistema. Você só precisa instalar uma vez.

---

## Como usar

Para abrir o programa:

```bash
./.venv/bin/python app.py
```

No Windows, use `.\.venv\Scripts\python app.py`.

A janela tem quatro áreas, de cima para baixo:

### 1. Arquivo de entrada

- **Arraste** um arquivo `.csv`, `.xlsx` ou `.xlsm` para a janela, **ou** clique em **Procurar...**.
- Se for Excel, escolha a **aba** na lista. Por padrão, é a primeira aba.
- Se for CSV, escolha o **separador** ou deixe em **Automático**. O padrão funciona na maioria dos casos.

Assim que o arquivo é carregado, o programa sugere um nome para a saída na mesma pasta, como `vendas_formatado.xlsx`.

### 2. Regras de formatação

Marque o que deseja aplicar. Cada alteração atualiza a pré-visualização na hora. As regras estão detalhadas na seção [Regras de formatação](#regras-de-formatação).

### 3. Arquivo de saída

- **Salvar em:** caminho do arquivo que será criado. Clique em **Escolher...** para mudar a pasta ou o nome.
- **Separador (CSV de saída):** só fica ativo quando a saída é `.csv`.
- **Estilizar planilha Excel:** só fica ativo quando a saída é `.xlsx`.
- **Converter números brasileiros:** só fica ativo quando a saída é `.xlsx`.

A **extensão do arquivo de saída decide o formato**: `.csv` gera um CSV, `.xlsx` gera uma planilha Excel. Para converter um CSV em Excel, basta escolher um nome terminado em `.xlsx`.

### 4. Pré-visualização

Mostra as primeiras **100 linhas** do resultado, já com as regras aplicadas. A barra de status embaixo informa quantas linhas e colunas havia antes e quantas ficaram depois.

Quando estiver satisfeito, clique em **Formatar e salvar**.

---

## Regras de formatação

| Regra | O que faz | Padrão |
|---|---|---|
| Remover espaços extras | Tira espaços do início e do fim de cada célula e reduz vários espaços seguidos a um só. Também trata o espaço não quebrável (`\u00a0`) comum em exportações. Quebras de linha dentro de uma célula são mantidas. | **ligada** |
| Remover linhas totalmente vazias | Apaga linhas em que todas as células estão vazias, depois da limpeza de espaços. | **ligada** |
| Remover colunas totalmente vazias | Apaga colunas em que nenhuma célula tem conteúdo. O cabeçalho também precisa estar vazio. | desligada |
| Remover linhas duplicadas | Mantém a primeira ocorrência de cada linha repetida. Compara todas as colunas. O cabeçalho nunca é comparado. | desligada |
| Maiúsculas/minúsculas do cabeçalho | Altera a caixa **só dos títulos** das colunas: `MAIÚSCULAS`, `minúsculas` ou `Título`. Os dados não mudam. | Não alterar |
| Separador do CSV de saída | Vírgula, ponto e vírgula, tabulação ou barra vertical. | vírgula |
| Estilizar planilha Excel | Cabeçalho em negrito com fundo cinza-azulado, bordas finas, filtro automático, cabeçalho fixo na tela e largura de coluna ajustada ao conteúdo. | **ligada** |
| Converter números brasileiros | Transforma texto como `1.234,56` em número de verdade no Excel. Veja a seção abaixo. | desligada |

### Como o programa entende a tabela

- **A primeira linha é sempre o cabeçalho.** Ela nunca é apagada por "linha vazia" nem por "duplicada".
- Linhas com menos colunas que as outras são completadas com células vazias, para a tabela ficar retangular.
- Espaços são limpos só em células de texto. Números e datas não são alterados por essa regra.

### Converter números brasileiros

Planilhas em português costumam usar `.` como separador de milhar e `,` como separador decimal. Sem conversão, o Excel guarda `1.234,56` como texto, e a coluna não pode ser somada.

Com a opção ligada, o programa converte apenas os textos que seguem exatamente esse padrão:

| Texto no arquivo | Resultado no Excel |
|---|---|
| `1.234,56` | número `1234.56`, exibido como `1.234,56` |
| `3,5` | número `3.5`, exibido como `3,5` |
| `1234` | número `1234` |
| `-0,75` | número `-0.75` |

E **mantém como texto** os casos que poderiam ser códigos:

| Texto | Motivo |
|---|---|
| `01310100` (CEP) | começa com zero |
| `123.456.789-09` (CPF) | tem hífen |
| `11987654321.0` | ponto que não separa milhares |
| `1.234` | ponto sem vírgula é ambíguo |
| números com mais de 15 dígitos | o Excel não guarda com precisão |

Ainda assim, **deixe a opção desligada** se a planilha tiver telefones, códigos ou identificadores sem zero à esquerda que você não queira converter. A conversão só afeta as linhas de dados; o cabeçalho continua texto.

---

## Formatos de entrada e saída

| Tipo | Entrada | Saída |
|---|---|---|
| CSV (`.csv`) | sim | sim |
| Excel (`.xlsx`) | sim | sim |
| Excel com macros (`.xlsm`) | sim | não (salve como `.xlsx`) |
| Excel antigo (`.xls`) | não | não |

Detalhes importantes:

- **Codificação do CSV de entrada:** o programa tenta UTF-8 (com ou sem BOM, que é o que o Excel costuma gravar), depois cp1252 (padrão do Windows em português) e, por último, latin-1.
- **Separador do CSV de entrada:** em **Automático**, o programa testa `,`, `;`, tabulação e `|`. Se a detecção errar, escolha o separador manualmente.
- **CSV de saída:** sempre em UTF-8 com BOM. Isso faz o Excel abrir acentos corretamente sem precisar importar o arquivo.
- **Excel de entrada:** lê os **valores** das células. Fórmulas são substituídas pelo último resultado calculado salvo no arquivo.
- **Excel de saída:** uma única aba, chamada `Dados`.

---

## Segurança dos seus dados

- **O arquivo original nunca é alterado.** A saída é sempre um arquivo novo. O programa bloqueia salvar a saída com o mesmo caminho da entrada.
- **Confirmação antes de sobrescrever.** Se o arquivo de destino já existir, o programa pergunta. A resposta padrão é **Não**, então fechar a caixa de diálogo nunca apaga nada.
- **Texto que parece fórmula fica como texto.** Uma célula com `=1+1` continua sendo o texto `=1+1`, e não uma fórmula executada no Excel. Isso evita que um arquivo recebido de terceiros execute cálculos ou comandos ao ser aberto.
- **Erros aparecem em janelas simples.** Se algo der errado, a mensagem diz o que aconteceu. O programa não mostra rastros técnicos.

---

## Exemplos

### Exemplo 1: limpar uma lista de clientes

Entrada (`clientes.csv`, separador `;`):

```text
  nome ; cidade ; 
 Ana   Maria ; São Paulo ;
;;
 Ana   Maria ; São Paulo ;
Bruno;Rio;
```

Com **Remover linhas duplicadas**, **Remover colunas totalmente vazias** e **Cabeçalho em MAIÚSCULAS** marcados, a saída `clientes_formatado.csv` fica:

```text
NOME,CIDADE
Ana Maria,São Paulo
Bruno,Rio
```

A terceira coluna sumiu porque não tinha nome nem dados. A linha `;;` foi removida como vazia, e a linha repetida da Ana foi removida como duplicada.

### Exemplo 2: transformar um CSV do sistema em planilha Excel

1. Arraste `vendas.csv` para a janela.
2. Em **Salvar em**, escolha `vendas.xlsx`.
3. Marque **Converter números brasileiros**.
4. Clique em **Formatar e salvar**.

Resultado: uma planilha com cabeçalho em negrito, bordas e filtro, em que a coluna de valores pode ser somada direto no Excel.

---

## Solução de problemas

**A janela não abre ou trava ao abrir (macOS).**
O Python do Xcode usa Tk 8.5, que pode travar. Instale um Python com Tk moderno:

```bash
brew install python-tk@3.14
/opt/homebrew/bin/python3.14 -m venv .venv
```

**`ModuleNotFoundError: No module named '_tkinter'`.**
Seu Python não tem Tk. Instale o pacote de Tk da sua distribuição (veja [Instalação](#instalação)) e recrie o ambiente virtual.

**Arrastar e soltar não funciona.**
Confirme que `tkinterdnd2` está instalado (`./.venv/bin/pip show tkinterdnd2`). Se estiver e mesmo assim não funcionar, use o botão **Procurar...**: o programa continua funcionando normalmente.

**"Formato não suportado".**
Só são aceitos `.csv`, `.xlsx` e `.xlsm` na entrada. Para saída, use `.csv` ou `.xlsx`.

**Caracteres estranhos (`Ã©` no lugar de `é`).**
O arquivo de origem provavelmente está em uma codificação que o programa não reconheceu. Salve o CSV original como "CSV UTF-8" no Excel e tente de novo.

**O separador foi detectado errado.**
Troque **Separador (CSV)** de **Automático** para o separador certo.

**"A aba não existe nesta planilha".**
A planilha pode ter sido alterada desde que foi carregada. Escolha a aba de novo ou arraste o arquivo outra vez.

---

## Limitações conhecidas

- **Tudo fica na memória.** O arquivo inteiro é carregado de uma vez. Planilhas com centenas de milhares de linhas podem ficar lentas.
- **Só uma aba de saída.** Em uma planilha com várias abas, a saída tem apenas a aba escolhida.
- **Formatação original não é preservada.** Cores, mesclagens, validações e gráficos do Excel de origem não são copiados. Mesclagens aparecem apenas no valor da célula de cima à esquerda.
- **Fórmulas viram valores.** A saída contém o resultado calculado, e não a fórmula.
- **Sem arquivo `.xls` antigo.** Converta antes para `.xlsx` no Excel.
- **A detecção de separador e de tipo de número é heurística.** Confira a pré-visualização antes de salvar.

---

## Testes

Os testes cobrem as regras de formatação, leitura e gravação de CSV e Excel, a conversão de números e as proteções de segurança:

```bash
./.venv/bin/python -m pytest
```

Para rodar, instale também o pytest:

```bash
./.venv/bin/pip install pytest
```

A interface gráfica **não tem testes automatizados**. Ela é validada manualmente, usando o programa com um arquivo real.

---

## Estrutura do projeto

```text
formatador-planilhas/
├── app.py                  # Janela (tkinter): campos, pré-visualização e botões
├── formatacao.py           # Regras: leitura, formatação e gravação (sem interface)
├── requirements.txt        # Dependências
├── pytest.ini              # Configuração dos testes
├── README.md
└── tests/
    ├── test_formatacao.py     # Testes das regras, CSV e Excel
    └── test_layout_excel.py   # Testes da conversão de números e do estilo Excel
```

A separação é intencional: `formatacao.py` não depende de janela, então as regras podem ser testadas sem abrir o programa e reaproveitadas em outros scripts:

```python
from formatacao import FormatOptions, read_table, format_table, write_table

tabela = read_table("entrada.csv")
tabela = format_table(tabela, FormatOptions(remove_duplicate_rows=True))
write_table("saida.xlsx", tabela, FormatOptions(convert_numbers=True))
```
