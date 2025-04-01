import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from datetime import datetime
import time

# URL base para consulta no DODF
DODF_URL = "https://dodf.df.gov.br/dodf/jornal/diario?tpSecao=III"


def configurar_driver():
    """Retorna uma instância do driver do Chrome."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("window-size=1920,1080")  # Set your desired window size
    options.add_argument(
        "zoom-factor=0.1"
    )  # Set the desired zoom level (1.0 means 100%)
    driver = webdriver.Chrome(options=options)
    driver.implicitly_wait(5)
    return driver


def baixar_pagina_listagem(data_consulta):
    """
    Acessa a página contendo as publicações do DODF, seleciona a opção "Extrato" e
    extrai os links dos extratos. Filtra apenas os links cujo texto visível contenha
    a expressão "nota de empenho". Percorre todas as páginas disponíveis.
    """
    driver = configurar_driver()
    driver.get(DODF_URL)
    time.sleep(5)

    all_links = []

    try:
        # Seleciona a opção "Extrato" no menu
        tipo_de_materia = driver.find_element(By.XPATH, "//*[@id='tpMateria']")
        time.sleep(1)
        Select(tipo_de_materia).select_by_visible_text("Extrato")
        time.sleep(3)

        while True:
            # Coleta todos os links da página atual que contenham "extrato" no href
            links_elements = driver.find_elements(
                By.XPATH,
                "//a[contains(translate(@href, 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'extrato')]",
            )
            for elem in links_elements:
                # Filtra pelo texto visível: somente se contiver "nota de empenho"
                if "nota de empenho" in elem.text.lower():
                    link = elem.get_attribute("href")
                    if link and link not in all_links:
                        all_links.append(link)

            # Tenta localizar e clicar no botão de próxima página usando CSS selector
            try:
                next_page_element = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, ".page-link.proxima-pagina")
                    )
                )
                driver.execute_script(
                    "arguments[0].scrollIntoView(true);", next_page_element
                )
                time.sleep(1)
                next_page_element.click()
                time.sleep(5)
            except Exception as e:
                print(
                    "Nenhuma próxima página encontrada ou ocorreu um erro ao clicar:", e
                )
                break

    except Exception as e:
        print("Ocorreu um erro durante a coleta de links:", e)

    driver.quit()
    return all_links


def extrair_dados_texto(texto):
    """
    Utiliza expressões regulares para extrair os dados do texto completo do extrato.
    Apenas processa se o texto contiver a expressão "NOTA DE EMPENHO".

    Os campos extraídos são:
      1. Nota de Empenho
      2. Processo
      3. Contratante (Partes, contratante)
      4. Contratado (Partes, contratado)
      5. CNPJ do contratado
      6. Objeto
      7. Contrato/Ata/Dispensa
      8. Valor
      9. Data do empenho
     10. Prazo (se houver)
    """
    dados = {}

    # Filtra apenas textos que realmente se referem a notas de empenho
    if not re.search(r"NOTA DE EMPENHO", texto, re.IGNORECASE):
        return None

    # Nota de Empenho: Exemplo: "EXTRATO DA NOTA DE EMPENHO 2024NE00414"
    m = re.search(
        r"EXTRATO\s+(?:DA|DE)\s+NOTA DE EMPENHO(?:\s*Nº)?\s*([\w\d]+)",
        texto,
        re.IGNORECASE,
    )
    dados["Nota de Empenho"] = m.group(1) if m else None

    # Processo: Exemplo: "Processo: 04008-00000354/2025-83"
    m = re.search(r"Processo:\s*([\d\-/]+)", texto, re.IGNORECASE)
    dados["Processo"] = m.group(1).strip() if m else None

    # Partes: Captura o texto após "Partes:" até ponto ou ponto e vírgula.
    m = re.search(r"Partes:\s*(.*?)(?:[.;])", texto, re.IGNORECASE)
    if m:
        partes = m.group(1).strip()
        partes_split = re.split(r"\s+e\s+", partes, flags=re.IGNORECASE)
        if len(partes_split) >= 2:
            dados["Contratante"] = partes_split[0].strip()
            dados["Contratado"] = partes_split[1].strip()
            # Extração do CNPJ no formato XX.XXX.XXX/XXXX-XX
            m_cnpj = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", partes_split[1])
            dados["CNPJ do contratado"] = m_cnpj.group(0) if m_cnpj else None
        else:
            dados["Contratante"] = partes
            dados["Contratado"] = None
            dados["CNPJ do contratado"] = None
    else:
        dados["Contratante"] = None
        dados["Contratado"] = None
        dados["CNPJ do contratado"] = None

    # Objeto: Exemplo: "Objeto: Despesa com diárias de viagem ..."
    m = re.search(r"Objeto:\s*(.*?)(?:[.;])", texto, re.IGNORECASE)
    dados["Objeto"] = m.group(1).strip() if m else None

    # Contrato/Ata/Dispensa: Procura por palavras-chave
    m = re.search(r"(Ata|Contrato|Dispensa).*?(?:[.;])", texto, re.IGNORECASE)
    dados["Contrato/Ata/Dispensa"] = m.group(0).strip() if m else None

    # Valor: Exemplo: "VALOR: R$ 1.172,46"
    m = re.search(r"VALOR:\s*(R\$[\s\d\.,]+)", texto, re.IGNORECASE)
    dados["Valor"] = m.group(1).strip() if m else None

    # Data do empenho: Exemplo: "Data do Empenho: 28/03/2025"
    m = re.search(
        r"Data\s+(?:do Empenho|da Emissão da Nota de Empenho):\s*([\d/]+)",
        texto,
        re.IGNORECASE,
    )
    dados["Data do empenho"] = m.group(1).strip() if m else None

    # Prazo: Exemplo: "Prazo: 30 dias" ou "PRAZO DE ENTREGA: 100% em 30 dias"
    m = re.search(r"PRAZO(?:\s+DE\s+ENTREGA)?:\s*(.*?)(?:[.;]|$)", texto, re.IGNORECASE)
    dados["Prazo"] = m.group(1).strip() if m else None

    return dados


def extrair_dados_extrato(url):
    """
    Acessa a página do extrato individual, obtém o texto completo e extrai os dados
    utilizando a função 'extrair_dados_texto'. Retorna None se o extrato não for de nota de empenho.
    """
    driver = configurar_driver()
    driver.get(url)
    time.sleep(3)

    try:
        texto_completo = driver.find_element(By.TAG_NAME, "body").text
    except Exception as e:
        print(f"Erro ao obter o texto da página: {e}")
        texto_completo = ""

    driver.quit()
    dados = extrair_dados_texto(texto_completo)
    return dados


def salvar_para_excel(dados, data_execucao):
    """
    Recebe uma lista de dicionários com os dados extraídos e salva em um arquivo Excel.
    O nome do arquivo inicia com a data de execução no formato YYYYMMDD.
    """
    df = pd.DataFrame(dados)
    nome_arquivo = f"{data_execucao.strftime('%Y%m%d')}_extrato_notas_empenho_dodf.xlsx"
    df.to_excel(nome_arquivo, index=False)
    print(f"Arquivo salvo: {nome_arquivo}")


def baixar_e_processar_dados(data_execucao):
    """
    Função orquestradora:
      1. Converte a data para o formato esperado pela consulta (DD-MM-AAAA).
      2. Baixa a página de listagem e extrai os links dos extratos, navegando por todas as páginas se necessário.
      3. Processa cada link, extrai os dados do texto (filtrando apenas notas de empenho) e salva os resultados.
    """
    data_consulta = data_execucao.strftime("%d-%m-%Y")
    links = baixar_pagina_listagem(data_consulta)

    if not links:
        print("Nenhum link de extrato foi encontrado.")
        return

    todos_dados = []
    for link in links:
        if link.startswith("/"):
            link = "https://dodf.df.gov.br" + link
        print(f"Processando: {link}")
        dados = extrair_dados_extrato(link)
        if dados:
            todos_dados.append(dados)

    if todos_dados:
        salvar_para_excel(todos_dados, data_execucao)
    else:
        print("Nenhum dado foi extraído.")
