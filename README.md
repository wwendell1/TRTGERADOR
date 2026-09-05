📊 Automação de Planilhas de Audiências

Este projeto é um pequeno programa em **Python** que pega informações de uma planilha de origem e coloca essas informações automaticamente em outra planilha, seguindo uma ordem específica.

A ideia é simples:

> **Você tem uma planilha cheia de informações → o programa lê os dados → organiza → encontra os campos na planilha de destino → preenche tudo automaticamente.**

Assim, em vez de copiar e colar dezenas ou centenas de informações manualmente, o Python faz esse trabalho por você. 🤖

---

## 🧠 1. O que este programa faz?

O programa trabalha com **duas planilhas Excel**:

### 📥 Planilha de origem

É onde estão os dados das audiências, por exemplo:

- Data
- Horário
- Reclamante
- Reclamado
- Número do processo
- Turma
- Relator
- Tipo de recurso

### 📤 Planilha de destino

É o modelo que será preenchido automaticamente.

O programa procura textos como:

```text
Nº processo:
Reclamante:
Horário:
Relator:
Reclamado:
Turma:
Local:
```

Quando encontra um desses textos, ele coloca a informação correspondente **na célula imediatamente ao lado**.

---

# 🚀 2. Antes de começar

Você precisa ter:

1. **Python instalado**
2. Os arquivos Excel
3. As bibliotecas necessárias
4. O código deste projeto

As bibliotecas utilizadas são:

- `pandas`
- `openpyxl`

---

# 🐍 3. Instalação do Python

Se você ainda não possui Python, instale uma versão recente do Python 3.

Durante a instalação no Windows, é importante marcar a opção:

```text
Add Python to PATH
```

Depois da instalação, abra o **Prompt de Comando (CMD)** e digite:

```bash
python --version
```

Se aparecer algo parecido com:

```text
Python 3.x.x
```

significa que o Python está instalado.

---

# 📦 4. Instalar as bibliotecas

Abra o CMD ou PowerShell e execute:

```bash
pip install pandas openpyxl
```

Espere a instalação terminar.

Depois disso, o Python já terá as ferramentas necessárias para trabalhar com Excel.

---

# 📁 5. Organização dos arquivos

Uma forma simples de organizar o projeto é:

```text
TESTAT/
│
├── automacao.py
├── RelatóriodeAudiências25.xlsx
└── TRT04.08a08.08.xlsx
```

Onde:

- `automacao.py` → é o programa Python.
- `RelatóriodeAudiências25.xlsx` → é a planilha de origem.
- `TRT04.08a08.08.xlsx` → é a planilha que será preenchida.

> ⚠️ Os nomes dos arquivos precisam ser exatamente iguais aos informados no código, caso você não altere os caminhos.

---

# 🛠️ 6. Configurando os caminhos

No código existe esta parte:

```python
caminho_origem = r"C:\Users\wendel.ferreira\Desktop\TESTAT\RelatóriodeAudiências25.xlsx"

caminho_destino = r"C:\Users\wendel.ferreira\Desktop\TESTAT\TRT04.08a08.08.xlsx"
```

Esses caminhos dizem ao programa:

> "Ei, Python! O arquivo de origem está aqui e o arquivo de destino está aqui."

## 🔴 Se o seu computador tiver outro usuário

Você precisa alterar o caminho.

Por exemplo:

```python
caminho_origem = r"C:\Users\Joao\Desktop\TESTAT\RelatóriodeAudiências25.xlsx"
```

E:

```python
caminho_destino = r"C:\Users\Joao\Desktop\TESTAT\TRT04.08a08.08.xlsx"
```

### 💡 Dica

A letra `r` antes das aspas é importante:

```python
r"C:\Users\..."
```

Ela ajuda o Python a interpretar corretamente as barras `\` usadas nos caminhos do Windows.

---

# 📄 7. Como deve ser a planilha de origem?

O programa espera encontrar uma aba chamada:

```text
Sheet1
```

E algumas colunas com estes nomes:

```text
DATA
HORARIO
RECLAMANTE
RECLAMADO
Nº DO PROCESSO
TURMA
RELATOR
TIPO DE RECURSO
```

O código transforma esses nomes internamente para:

```text
DATA
HORARIO
RECLAMANTE
RECLAMADO
NUM_PROCESSO
TURMA
RELATOR
TIPO_RECURSO
```

Isso é feito através deste mapeamento:

```python
mapeamento_colunas = {
    'DATA': 'DATA',
    'HORARIO': 'HORARIO',
    'RECLAMANTE': 'RECLAMANTE',
    'RECLAMADO': 'RECLAMADO',
    'Nº DO PROCESSO': 'NUM_PROCESSO',
    'TURMA': 'TURMA',
    'RELATOR': 'RELATOR',
    'TIPO DE RECURSO': 'TIPO_RECURSO'
}
```

---

# 🧹 8. O programa também organiza os dados

Depois de ler a planilha, o programa:

### 1️⃣ Remove registros sem data ou horário

```python
df_filtrado = df_filtrado.dropna(subset=['DATA', 'HORARIO'])
```

Ou seja:

> Se uma audiência não tiver data ou horário, ela será ignorada.

### 2️⃣ Converte as datas

```python
df_filtrado['DATA'] = pd.to_datetime(
    df_filtrado['DATA'],
    errors='coerce'
)
```

### 3️⃣ Converte os horários

```python
df_filtrado['HORARIO_TIME'] = pd.to_datetime(
    df_filtrado['HORARIO'],
    format='%H:%M',
    errors='coerce'
).dt.time
```

### 4️⃣ Ordena os registros

Primeiro pela:

```text
DATA
```

Depois pelo:

```text
HORÁRIO
```

Isso significa que as audiências ficam organizadas cronologicamente.

---

# ⚖️ 9. Como o tipo de recurso funciona?

O programa também verifica o campo:

```text
TIPO DE RECURSO
```

Se encontrar:

```text
RECURSO ORDINARIO
```

ele adiciona:

```text
RO
```

antes do número do processo.

Exemplo:

```text
RO 1234567-89.2025.5.04.0001
```

Se encontrar:

```text
AGRAVO DE PETIÇÃO
```

ele adiciona:

```text
AP
```

Exemplo:

```text
AP 1234567-89.2025.5.04.0001
```

Se não encontrar nenhum desses tipos, o número permanece sem prefixo.

---

# 🏛️ 10. Como a Turma e o Local são separados?

Imagine que na planilha de origem esteja:

```text
1ª Turma Porto Alegre
```

O programa divide essa informação em duas partes.

Resultado:

```text
Turma: 1ª
Local: Turma Porto Alegre
```

O código faz isso usando:

```python
partes = turma_texto.split()
```

Depois pega a primeira parte como número da turma:

```python
numero_turma = partes[0]
```

E o restante como local:

```python
local_turma = " ".join(partes[1:])
```

> ⚠️ Isso significa que o código pressupõe que o primeiro "pedaço" do texto representa a turma.

---

# 🔎 11. Como o programa encontra os campos?

Esta é uma das partes mais importantes.

A função:

```python
colar_valor_ao_lado()
```

procura dentro da planilha de destino um texto específico.

Por exemplo:

```text
Nº processo:
```

Quando encontra esse texto, olha para a célula imediatamente à direita e coloca o valor.

Imagine:

| A | B |
|---|---|
| Nº processo: | |
| Reclamante: | |
| Horário: | |

Depois da automação:

| A | B |
|---|---|
| Nº processo: | RO 123456 |
| Reclamante: | João da Silva |
| Horário: | 09:30 |

É como se o programa dissesse:

> "Achei o campo! Agora vou colocar a informação do lado dele."

---

# 🔢 12. O que significa `ocorrencia`?

Essa parte é muito importante.

A planilha de destino pode ter vários blocos iguais.

Por exemplo:

```text
Nº processo:
Reclamante:
Horário:
Relator:
Reclamado:
Turma:
Local:
```

e depois novamente:

```text
Nº processo:
Reclamante:
Horário:
Relator:
Reclamado:
Turma:
Local:
```

O programa precisa saber se está preenchendo o **primeiro bloco, segundo bloco, terceiro bloco**, etc.

Por isso existe:

```python
ocorrencia_global = 1
```

Para o primeiro registro:

```text
ocorrência = 1
```

Para o segundo:

```text
ocorrência = 2
```

Para o terceiro:

```text
ocorrência = 3
```

E assim por diante.

A cada audiência processada:

```python
ocorrencia_global += 1
```

Ou seja:

> "Terminei uma audiência. Agora vou para o próximo bloco."

---

# ▶️ 13. Como executar o programa?

Depois de configurar tudo:

### Passo 1

Abra o CMD ou PowerShell.

### Passo 2

Entre na pasta do projeto.

Por exemplo:

```bash
cd C:\Users\wendel.ferreira\Desktop\TESTAT
```

### Passo 3

Execute:

```bash
python automacao.py
```

Se tudo estiver correto, aparecerá:

```text
Automação concluída com sucesso.
```

🎉 Pronto!

---

# ⚠️ 14. Feche o Excel antes de executar

Antes de executar o programa, é recomendado fechar os arquivos:

```text
RelatóriodeAudiências25.xlsx
TRT04.08a08.08.xlsx
```

Principalmente a planilha de destino.

Por quê?

Porque o Python precisa conseguir abrir e salvar o arquivo.

Se o Excel estiver usando o arquivo, pode acontecer um erro de acesso ou bloqueio.

---

# 🛡️ 15. O programa verifica se os arquivos existem

Antes de começar, o código verifica:

```python
if not os.path.exists(caminho_origem) or not os.path.exists(caminho_destino):
```

Se algum arquivo não existir, aparece:

```text
Erro: Um dos arquivos não foi encontrado.
```

Nesse caso, confira:

- Se o nome está correto;
- Se a pasta está correta;
- Se o arquivo realmente existe;
- Se a extensão é `.xlsx`;
- Se o caminho está apontando para o local correto.

---

# ❌ 16. Principais erros e soluções

## Erro: `No module named pandas`

Significa que o pandas não está instalado.

Execute:

```bash
pip install pandas
```

---

## Erro: `No module named openpyxl`

Execute:

```bash
pip install openpyxl
```

---

## Erro: arquivo não encontrado

Confira:

```python
caminho_origem
caminho_destino
```

---

## Erro relacionado à aba `Sheet1`

A planilha de origem precisa ter uma aba chamada:

```text
Sheet1
```

Se a aba tiver outro nome, altere:

```python
pd.read_excel(caminho_origem, sheet_name='Sheet1')
```

Por exemplo:

```python
pd.read_excel(caminho_origem, sheet_name='Audiências')
```

---

# ⚠️ 17. Muito cuidado com a planilha de destino

O programa **altera a planilha de destino**.

No final existe:

```python
wb_destino.save(caminho_destino)
```

Isso significa que o arquivo será salvo com as informações novas.

### Recomenda-se fazer um backup.

Por exemplo:

```text
TRT04.08a08.08 - ORIGINAL.xlsx
```

E trabalhar com uma cópia:

```text
TRT04.08a08.08.xlsx
```

Assim, se alguma coisa der errado, você ainda terá o original.

---

# 🧩 18. Resumo do funcionamento

O programa segue aproximadamente esta sequência:

```text
INÍCIO
   ↓
Verifica se os arquivos existem
   ↓
Lê a planilha de origem
   ↓
Seleciona as colunas necessárias
   ↓
Remove registros sem data/horário
   ↓
Organiza por data e horário
   ↓
Abre a planilha de destino
   ↓
Procura o primeiro bloco
   ↓
Preenche os dados
   ↓
Vai para o próximo bloco
   ↓
Repete até terminar
   ↓
Salva a planilha
   ↓
FIM
```

---

# 🧑‍💻 19. Estrutura principal do código

O programa possui duas partes principais.

## Função `colar_valor_ao_lado()`

Responsável por:

> Procurar um texto e colocar um valor ao lado dele.

```python
def colar_valor_ao_lado(ws, texto_alvo, valor, ocorrencia=1):
```

---

## Função `processar_planilhas()`

É o "cérebro" da automação.

Ela:

- abre os arquivos;
- lê os dados;
- organiza os registros;
- interpreta o tipo de recurso;
- separa turma/local;
- preenche a planilha;
- salva o resultado.

```python
def processar_planilhas():
```

---

# 🔧 20. Como alterar os campos?

Se no futuro você quiser adicionar outro campo, será necessário fazer algumas alterações.

Por exemplo, suponha que queira adicionar:

```text
Advogado:
```

Primeiro, a informação precisa existir na planilha de origem.

Depois, você adicionaria a coluna ao conjunto de colunas necessárias.

Por fim, adicionaria uma linha semelhante a:

```python
colar_valor_ao_lado(
    ws_destino,
    "Advogado:",
    str(row['ADVOGADO']),
    ocorrencia=ocorrencia_global
)
```

A lógica é sempre:

```text
"Texto que o programa procura"
            ↓
"Valor que será colocado ao lado"
```

---

# 💡 21. Uma dica importante sobre textos

O programa procura o texto de forma que não diferencia letras maiúsculas e minúsculas.

Por exemplo, procurar:

```text
Reclamante:
```

pode encontrar:

```text
RECLAMANTE:
```

ou:

```text
reclamante:
```

Isso acontece porque o código utiliza:

```python
texto_alvo.lower() in cell.value.lower()
```

Porém, é importante manter os textos da planilha de destino consistentes.

Por exemplo:

```text
Reclamante:
```

é melhor do que ter:

```text
Reclamante
```

em um lugar e:

```text
Nome do Reclamante:
```

em outro.

---

# 🚨 22. Limitações atuais

O programa funciona muito bem quando a estrutura das planilhas permanece consistente.

Porém, existem algumas limitações:

### 1. Caminhos estão fixos

Os arquivos estão definidos diretamente no código.

### 2. Nome da aba de origem está fixo

Atualmente:

```text
Sheet1
```

### 3. Nome da aba de destino

O programa procura primeiro:

```text
Julho
```

Se não encontrar, utiliza a aba ativa:

```python
ws_destino = wb_destino["Julho"] if "Julho" in wb_destino.sheetnames else wb_destino.active
```

### 4. A ordem dos blocos precisa estar correta

O programa assume que o primeiro registro deve preencher a primeira ocorrência, o segundo a segunda ocorrência e assim por diante.

### 5. O programa procura os campos pelo texto

Se o modelo da planilha mudar muito, pode ser necessário alterar o código.

---

# ⭐ 23. Exemplo simples

Imagine que a planilha de origem tenha:

```text
DATA: 05/08/2026
HORARIO: 09:30
RECLAMANTE: João
RECLAMADO: Empresa XYZ
Nº DO PROCESSO: 123456
TURMA: 1ª Porto Alegre
RELATOR: Maria
TIPO DE RECURSO: RECURSO ORDINARIO
```

O programa poderá gerar:

```text
Nº processo: → RO 123456
Reclamante: → João
Horário: → 09:30
Relator: → Maria
Reclamado: → Empresa XYZ
Turma: → 1ª
Local: → Porto Alegre
```

Tudo automaticamente.

---

# 🎯 24. Checklist antes de executar

Antes de apertar o botão "rodar", confira:

- [ ] Python está instalado
- [ ] `pandas` está instalado
- [ ] `openpyxl` está instalado
- [ ] O arquivo de origem existe
- [ ] O arquivo de destino existe
- [ ] O caminho dos arquivos está correto
- [ ] A aba `Sheet1` existe na origem
- [ ] A aba `Julho` existe no destino ou a aba ativa é a correta
- [ ] Os nomes das colunas estão corretos
- [ ] A planilha de destino está fechada
- [ ] Foi feito backup da planilha original

---

# 🏁 25. Em uma frase

Se você precisar explicar este projeto para alguém de forma bem simples:

> **"Esse programa pega os dados de uma planilha de audiências, organiza por data e horário e depois preenche automaticamente um modelo de outra planilha, evitando o trabalho de copiar e colar tudo manualmente."**

---

## 📌 Tecnologias utilizadas

- **Python**
- **Pandas** — leitura e organização dos dados
- **OpenPyXL** — leitura e edição dos arquivos Excel
- **OS** — verificação dos arquivos no computador
- **Datetime** — tratamento de datas e horários

---

## 📄 Licença

Este projeto pode ser adaptado conforme a necessidade do usuário e da estrutura das planilhas utilizadas.

---

## 👨‍💻 Autor

Automação desenvolvida em Python para facilitar o preenchimento e organização de planilhas de audiências.
README.md…]()
