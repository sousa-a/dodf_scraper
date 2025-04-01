from datetime import datetime
from src.scraper import baixar_e_processar_dados

if __name__ == "__main__":
    data_execucao = datetime.now()
    baixar_e_processar_dados(data_execucao)
