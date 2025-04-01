# DODF scraper
Script (MVP) para a extração de informações acerca de Extrato de Notas de Empenho publicadas no Diário Oficial dp Distrito Federal(DODF) utilizando Python, Selenium e Openpyxl.

## Objetivo

- Extrair diariamente as publicações de extratos de notas de empenho.
- Salvar os dados extraídos em um arquivo Excel.

## Instalação
1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/dodf_scraper.git
   ```

2. Criando e Ativando o Ambiente Virtual:<br><br>
    Recomenda-se o uso de ambientes virtuais para isolar as dependências do projeto.<br><br>
    No Windows:
    ```bash
    python -m venv venv
    venv\Scripts\activate
    ```

    No Linux/macOS:
    ```bash
    python3 -m venv venv
    source venv/bin/activate
    ```
3. Instalando as Dependências:<br><br>
    Com o ambiente virtual ativado, instale as dependências listadas no arquivo requirements.txt.<br>
    ```bash
    pip install -r requirements.txt

## Execução
Para executar o scraper e gerar o arquivo Excel com os dados extraídos, basta executar:<br>
    
    python main.py

O script realizará as seguintes atividades:

1. Acessará a página de listagem dos extratos de nota de empenho para a data atual;

2. Extrairá os links dos extratos disponíveis;

3. Para cada link, fará a extração dos dados utilizando _Selenium_ e _Regular Expressions_;

4. Gerará um arquivo Excel na pasta do projeto com o nome no formato "YYYYMMDD_extrato_notas_empenho_dodf.xlsx";

## Considerações

Não há um padrão de publicação comum a todos os órgãos.<br><br>
Portanto, é possível que tais diferenças interfiram nos resultados obtidos após a execução.
