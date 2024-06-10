# RenafinderTurbo

## Descrição

RenafinderTurbo é uma aplicação GUI desenvolvida em Python para carregar, buscar e processar dados de planilhas em formatos `.xlsx`, `.csv`, e `.xml`. Ele também permite a visualização dos dados através de diferentes tipos de gráficos.

## Funcionalidades

- **Carregar Planilha:**
  - Suporta arquivos `.xlsx`, `.csv`, e `.xml`.
  - Permite a seleção de múltiplas abas em arquivos Excel.

- **Buscar Dados:**
  - Permite buscas com base em palavras-chave em colunas selecionadas.
  - Suporta critérios de busca como "Contém", "Igual", "Começa com", "Termina com".

- **Aplicar Processamento de Dados:**
  - Filtragem de dados com base em palavras-chave.
  - Operações de agregação como Soma, Média e Contagem.
  - Normalização de dados.
  - Criação de novas colunas com base em expressões fornecidas pelo usuário.

- **Visualização de Dados:**
  - Geração de gráficos de histogramas, dispersão, pizza e barras.
  - Opção para salvar gráficos gerados em arquivos de imagem.

- **Exportar Resultados:**
  - Exportação dos resultados das buscas e processamentos para arquivos `.csv` ou `.xlsx`.

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/RenafinderTurbo.git
   cd RenafinderTurbo

2. Crie um ambiente virtual e instale as dependências:
   ```bash
   python -m venv venv
   source venv/bin/activate  # No Windows, use `venv\Scripts\activate`
   pip install -r requirements.txt

2. Crie um ambiente virtual e instale as dependências:
   ```bash
   python -m venv venv
   source venv/bin/activate  # No Windows, use `venv\Scripts\activate`
   pip install -r requirements.txt

## Uso

1. Para executar o programa:
   ```bash
   python RenafinderTurbo.py

## Compilar para Executável

1. Para criar um executável, utilize o PyInstaller:
   ```bash
   pip install pyinstaller
   pyinstaller --onefile --windowed RenafinderTurbo.py

O executável será gerado na pasta dist.






