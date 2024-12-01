from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

def obter_entidades(driver):
    entidades = []
    try:
        driver.get('https://fas.curitiba.pr.gov.br/formulariodoacao.aspx?fundo=14')

        # Espera o elemento estar presente e clicável
        elemento = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'conteudoMaster_rblFundo_1'))
        )
        # Clica no elemento
        elemento.click()
        
        entidade_select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'conteudoMaster_ddlEntidade'))
        )
        entidade_select = Select(entidade_select_element)

        for option in entidade_select.options:
            entidade_texto = option.text
            entidade_valor = option.get_attribute('value')
            entidades.append((entidade_texto, entidade_valor))

    except Exception as e:
        print(f"Erro ao obter entidades: {e}")

    return entidades

def extrair_detalhes_projeto(driver):
    try:
        contato = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'conteudoMaster_lblInformacoesContato'))
        ).text
        email = driver.find_element(By.ID, 'conteudoMaster_lblInformacoesEmail').text
        endereco = driver.find_element(By.ID, 'conteudoMaster_lblInformacoesEndereco').text
        telefone = driver.find_element(By.ID, 'conteudoMaster_lblInformacoesTelefone').text
        valor_aprovado = driver.find_element(By.ID, 'conteudoMaster_lblInformacoesAprovado').text
        valor_captado = driver.find_element(By.ID, 'conteudoMaster_lblInformacoesRepassado').text
        informacoes = driver.find_element(By.ID, 'conteudoMaster_lblInformacoesProjeto').text
        data_vigencia = driver.find_element(By.ID, 'conteudoMaster_lblDataVigencia').text

        return {
            'Contato': contato,
            'Email': email,
            'Endereço': endereco,
            'Telefone': telefone,
            'Valor Aprovado': valor_aprovado,
            'Valor Captado': valor_captado,
            'Informações': informacoes,
            'Data de Vigência': data_vigencia
        }
    except TimeoutException as e:
        print("Erro ao extrair informações do projeto:", e)
        return None

def capturar_projetos_para_entidades(driver, entidades):
    dados_projetos = []
    try:
        driver.get('https://fas.curitiba.pr.gov.br/formulariodoacao.aspx?fundo=14')

        for entidade_texto, entidade_valor in entidades:
            attempts = 0
            while attempts < 3:
                try:

                    # Espera o elemento estar presente e clicável
                    elemento = WebDriverWait(driver, 10).until(
                        EC.element_to_be_clickable((By.ID, 'conteudoMaster_rblFundo_1'))
                    )
                    # Clica no elemento
                    elemento.click()                    

                    entidade_select_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'conteudoMaster_ddlEntidade'))
                    )
                    entidade_select = Select(entidade_select_element)

                    entidade_select.select_by_value(entidade_valor)
                    print(f'Selecionando entidade: {entidade_texto} (Value: {entidade_valor})')

                    projeto_select_element = WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'conteudoMaster_ddlProjeto'))
                    )
                    projeto_select = Select(projeto_select_element)

                    for projeto_option in projeto_select.options:
                        projeto_texto = projeto_option.text
                        projeto_valor = projeto_option.get_attribute('value')
                        print(f'  Projeto: {projeto_texto} (Value: {projeto_valor})')

                        if projeto_valor != '0':
                            projeto_select.select_by_value(projeto_valor)

                            # Extração de detalhes do projeto
                            detalhes = extrair_detalhes_projeto(driver)
                            if detalhes:
                                detalhes['Entidade'] = entidade_texto
                                detalhes['Projeto'] = projeto_texto.upper()
                                dados_projetos.append(detalhes)

                    break  # Se tudo correr bem, sair do loop de tentativas

                except StaleElementReferenceException:
                    print(f"Stale reference for {entidade_texto}, tentativa {attempts + 1}")
                    attempts += 1
                    WebDriverWait(driver, 10).until(
                        EC.staleness_of(entidade_select_element)
                    )  # Espera pelo DOM ser atualizado

                except (TimeoutException, Exception) as e:
                    print(f"Erro ao acessar elementos para a entidade {entidade_texto}: {str(e)}")
                    break  # Se outro erro, saia do loop de tentativas

    except Exception as e:
        print(f"Erro ao capturar projetos: {e}")

    return dados_projetos

def exportar_para_excel(dados_projetos):
    if dados_projetos:
        df = pd.DataFrame(dados_projetos)
        df.to_excel('dados_projetos.xlsx', index=False)
        print("Dados exportados para 'dados_projetos.xlsx'")
    else:
        print("Nenhum dado foi coletado para exportação.")

def main():
    # driver = webdriver.Chrome(executable_path='./chromedriver/chromedriver.exe')

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)

    try:
        lista_de_entidades = obter_entidades(driver)
        dados_projetos = capturar_projetos_para_entidades(driver, lista_de_entidades)
        exportar_para_excel(dados_projetos)

    except Exception as e:
        print(f"Ocorreu um erro: {e}")

    finally:
        driver.quit()

if __name__ == "__main__":
    main()