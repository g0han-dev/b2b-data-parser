# DataParser Pro

[🇺🇸 Read in English](#english-version)

Um utilitário em Python criado para transformar relatórios B2B caóticos e desestruturados em tabelas limpas, prontas para injeção no banco de dados.

## O problema que me fez criar isso
No dia a dia da operação, a equipe recebia relatórios de pedidos de fornecedores que eram formatados apenas para leitura humana. Informações cruciais, como CNPJ, endereço e ID do pedido, vinham misturadas em blocos de texto soltos no meio da lista de produtos (SKUs).

Como não havia uma estrutura padrão, era impossível rodar rotinas automatizadas de `INSERT` no ERP. O resultado? A equipe perdia horas todos os dias formatando planilhas na mão, o que era lento e abria margem para erros de digitação na base da empresa.

## Como o script resolve isso
Desenvolvi o DataParser para atuar como um tradutor no meio do caminho:

1. Ele usa a biblioteca `pandas` para ler o arquivo de cima a baixo, identificando "âncoras" (como a palavra "CNPJ:"). 
2. Quando acha essas âncoras, ele guarda a informação na memória e vai replicando (achatando) esses dados de cabeçalho para cada item da lista.
3. No fim, ele limpa as strings de texto desnecessárias e exporta uma "flat table" estruturada.

O que antes levava horas de trabalho manual chato e repetitivo, agora roda em menos de 2 segundos. O arquivo final já sai no formato exato que as rotinas de banco de dados (INT, VARCHAR) exigem.

## Stack
- **Linguagem:** Python 3
- **Processamento:** Pandas e Openpyxl
- **Interface Gráfica:** CustomTkinter (criei uma tela simples para a equipe conseguir visualizar e validar o que está processando)
- **Deploy:** PyInstaller (empacotei tudo em um `.exe` para o pessoal da operação usar direto, sem precisar instalar Python ou configurar ambiente na máquina deles).

---

<a id="english-version"></a>
# English Version

A straightforward Python utility built to turn messy, unstructured B2B reports into clean, database-ready tables.

## The pain point
In our daily operations, the team received supplier order reports that were formatted strictly for human eyes. Crucial data—like Tax IDs, addresses, and Order IDs—were mixed into loose text blocks right in the middle of the product (SKU) lists.

Without a standard structure, it was impossible to run automated `INSERT` routines into the ERP. The result? The team wasted hours every day manually formatting spreadsheets, which was slow and highly prone to data entry errors.

## How it works
I built DataParser to act as a middleman translator:

1. It uses `pandas` to read the file top-to-bottom, looking for "anchors" (like the string "CNPJ:").
2. When it finds these anchors, it holds the info in memory and maps (flattens) this header data to every single item on the list.
3. Finally, it strips away unnecessary text strings and outputs a structured flat table.

What used to take hours of tedious, repetitive manual work now runs in under 2 seconds. The output file is ready for direct database consumption with the proper data types (INT, VARCHAR).

## Stack
- **Language:** Python 3
- **Data Processing:** Pandas and Openpyxl
- **GUI:** CustomTkinter (added a simple interface so the ops team can visually validate the batches)
- **Deployment:** PyInstaller (packaged into a standalone `.exe` so end-users can run it without needing to install Python or set up environments on their machines).
