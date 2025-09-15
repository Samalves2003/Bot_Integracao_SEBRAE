import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.options import Options
import time

# --- 1. PROCESSO DE EXECUÇÃO: IMPORTANTE! ---
# Para que este script funcione corretamente, siga estes passos:
#
# 1. Feche todas as janelas do Chrome.
# 2. Abra o terminal (Prompt de Comando no Windows) e execute o seguinte comando:
#    (No Windows) "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\temp\chrome_dev_profile"
# 3. Uma nova janela do Chrome será aberta. Faça o login no sistema da empresa e passe pela verificação de duas etapas.
# 4. **Mantenha esta janela aberta.**
# 5. Execute este script Python. Ele se conectará à janela que você acabou de abrir.

# --- 2. CONFIGURAÇÃO DO NAVEGADOR ---
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

# Inicializa o navegador conectando-se à sessão existente.
try:
    driver = webdriver.Chrome(options=chrome_options)
except WebDriverException as e:
    print("Erro ao tentar conectar ao navegador.")
    print("Verifique se o Google Chrome foi iniciado manualmente com a porta 9222 ativada.")
    print(f"Detalhes do erro: {e}")
    exit()

# Mensagens e Constantes
MENSAGEM_SUCESSO_ESPERADA = "Processo de Cadastro do Usuário Concluído com Sucesso."
MENSAGEM_ERRO_CONSULTA = "O campo Usuário Consultor não está preenchido."
TEMPO_ESPERA_TOAST = 20
TEMPO_ESPERA = 10
XPATH_BOTAO_INTEGRAR = '//*[@id="syncUserSASButton"]'
XPATH_PESQUISAR = "/html/body/div[4]/div[1]/section/header/div[2]/div[2]/div/div/button"
XPATH_CAMPO_PESQUISA = '//input[@placeholder="Pesquisar..."]'
XPATH_LINK_PESSOAS = '//a[text()="Pessoas"]'
XPATH_PERFIL = '//div[@class="name"]//a[contains(@class, "outputLookupLink")]'
XPATH_MENSAGEM = '//span[contains(@class, "toastMessage") and contains(@class, "forceActionsText")]'
XPATH_EDITAR = '//*[@id="brandBand_3"]/div[2]/div/div[1]/div/div[1]/div[1]/header/div[2]/div/div[1]/div[3]/div[2]/ul/li[1]/a/div'
XPATH_CAMPO_EDITAR_CPF = "//label[span[text()='CPF do Usuário']]/following-sibling::input"
XPATH_CAMPO_CONSULTOR = "/html/body/div[4]/div[2]/div[2]/div[2]/div/div[2]/div/div/div[1]/div/article/div[4]/div/fieldset[1]/div/div/div[6]/div[2]/div/div/div/div/div/div/div/a"
XPATH_OPCAO_NAO = "//a[@title='Não' and text()='Não']"
XPATH_OPCAO_SIM = "//a[@title='Sim' and text()='Sim']"
XPATH_BOTAO_SALVAR = "/html/body/div[4]/div[2]/div[2]/div[2]/div/div[2]/div/div/div[2]/div/div/div[2]/button[2]/span"
XPATH_TIPO_USUARIO = "//span[text()='Tipo de Usuário']/ancestor::div[contains(@class,'slds-form-element')]/div[@class='slds-form-element__control slds-grid itemBody']//span[1]"
XPATH_NOME_USUARIO = "//span[text()='Nome completo']/ancestor::div[@class='slds-form-element']/div[@class='slds-form-element__control slds-grid itemBody']//span[contains(@class, 'test-id__field-value')]"
# --- 3. DEFINIÇÃO DAS FUNÇÕES DE AUTOMAÇÃO ---


def capturar_toast(driver):
    """
    Aguarda e captura a mensagem de toast (popup) na tela.

    Args:
        driver (WebDriver): O objeto do Selenium WebDriver.

    Returns:
        str: O texto da mensagem de toast.
    """
    print("Aguardando mensagem de resultado (toast)...")
    wait_toast = WebDriverWait(driver, TEMPO_ESPERA_TOAST)
    toast_mensagem = wait_toast.until(EC.visibility_of_element_located((By.XPATH, XPATH_MENSAGEM))).text

    return toast_mensagem


def integrar_usuario(driver, df, index, cpf_usuario):
    """
    Realiza a automação para integrar um usuário no SAS e atualiza a planilha em caso de erro.

    Args:
        driver (WebDriver): O objeto do Selenium WebDriver.
        df (DataFrame): O DataFrame a ser atualizado.
        index (int): O índice da linha atual.
        cpf_usuario (str): O CPF do usuário a ser integrado.

    Returns:
        str: A mensagem de toast (sucesso ou erro) que aparece na tela, ou uma mensagem de erro genérica.
    """
    try:

        # Passo 1: Esperar e clicar no botão de pesquisa.
        print("Buscando e clicando no botão de pesquisa...")
        botao_pesquisar = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_PESQUISAR)))
        botao_pesquisar.click()

        # Passo 2: Esperar, limpar e enviar o CPF para o campo de pesquisa.
        print("Buscando o campo de pesquisa e inserindo o CPF...")
        campo_pesquisa = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_PESQUISA)))
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(cpf_usuario)
        campo_pesquisa.send_keys(Keys.ENTER)

        # Clicar no link "Pessoas" para abrir a seção correta
        print("Aguardando e clicando no link 'Pessoas'...")
        link_pessoas = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_LINK_PESSOAS)))
        link_pessoas.click()

        # Clicar no perfil dentro da seção "Pessoas".
        print("Buscando o link do perfil...")
        perfil_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_PERFIL)))
        perfil_link.click()
        
    
       
        # Puxa o texto do elemento
     
       
        

        # Clicar em "Integrar usuário SAS".
        print("Buscando e clicando no botão 'Integrar usuário SAS'...")
        botao_integrar = wait.until(
            EC.element_to_be_clickable((By.XPATH, XPATH_BOTAO_INTEGRAR)))
        botao_integrar.click()

        # Chama a função para capturar a mensagem de resultado
        return capturar_toast(driver)

    except (NoSuchElementException, TimeoutException) as e:
        df.at[index, 'Status de Integração'] = 'Erro de Automação'
        df.at[index,
              'Console'] = f'Erro: Elemento não encontrado ou tempo limite excedido. Detalhe: {e}'
        print(f"Erro de Automação: {e}")
        return "Erro de Automação"

    except Exception as e:
        df.at[index, 'Status de Integração'] = 'Erro Geral'
        df.at[index, 'Console'] = f'Erro inesperado: {str(e)}'
        print(f"Erro inesperado: {str(e)}")
        return "Erro Geral"


def tentar_corrigir_consultor(driver, toast_mensagem_inicial, df, index):
    """
    Executa a lógica de correção do campo "Usuário é Consultor?".

    Args:
        driver (WebDriver): O objeto do Selenium WebDriver.
        toast_mensagem_inicial (str): A mensagem de toast que causou a falha.
        df (DataFrame): O DataFrame a ser atualizado.
        index (int): O índice da linha atual.
    """
    wait = WebDriverWait(driver, TEMPO_ESPERA)

    # --- Passo 1: Lógica de correção para o caso 'O campo Usuário Consultor não está preenchido' ---
    if "O campo Usuário Consultor não está preenchido." in toast_mensagem_inicial:
        print("Erro 'Usuário Consultor' detectado. Tentando correção...")

        # Verificar o tipo de usuário

        elemento_tipo_usuario = wait.until(EC.presence_of_element_located((By.XPATH, XPATH_TIPO_USUARIO)))
        texto_tipo_usuario = elemento_tipo_usuario.text
       
        # Puxa o texto do elemento
     
       
        print("TIPO USUARIO " + texto_tipo_usuario)
        
       

        if texto_tipo_usuario == "Credenciado":
            
            # Mudar para 'Sim'
            botao_editar = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_EDITAR)))
            botao_editar.click()

            campo_cpf = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_EDITAR_CPF)))
            campo_cpf.clear()
            campo_cpf.send_keys(cpf_usuario)
            
            print("Usuario é credenciado: Alterando para 'Sim'...")
            campo_consultor = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_CONSULTOR)))
            campo_consultor.click()

            opcao_sim = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_OPCAO_SIM)))
            opcao_sim.click()

            botao_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_BOTAO_SALVAR)))
            botao_salvar.click()
            return True
        
        elif texto_tipo_usuario == "":
             print("TIPO DE USUARIO NÃO CONSTA NO PERFIL")
             df.at[index, 'Status de Integração'] = 'Falha'
             df.at[index, 'Console'] = f'Falha após tentativa de correção. Mensagem final: TIPO DE USUARIO NÃO CONSTA NO PERFIL'
             return False
             
        
        else:
            print(texto_tipo_usuario)
            # Mudar para 'Não'
            botao_editar = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_EDITAR)))
            botao_editar.click()

            campo_cpf = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_EDITAR_CPF)))
            campo_cpf.clear()
            campo_cpf.send_keys(cpf_usuario)
            
            print("Usuario é credenciado: Alterando para 'Sim'...")
            campo_consultor = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_CAMPO_CONSULTOR)))
            campo_consultor.click()

            opcao_nao = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_OPCAO_NAO)))
            opcao_nao.click()

            botao_salvar = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_BOTAO_SALVAR)))
            botao_salvar.click()
            return True
        
        
        


def tentar_integrar_e_verificar(driver, df, index):
    """
    Tenta integrar o usuário novamente e verifica o resultado para decidir a próxima ação.

    Args:
        driver (WebDriver): O objeto do Selenium WebDriver.
        df (DataFrame): O DataFrame a ser atualizado.
        index (int): O índice da linha atual.
    """
    wait = WebDriverWait(driver, TEMPO_ESPERA)

    # --- Tenta integrar novamente ---
    botao_integrar = wait.until(EC.element_to_be_clickable((By.XPATH, XPATH_BOTAO_INTEGRAR)))
    botao_integrar.click()

    # Captura a nova mensagem de toast
    toast_mensagem = capturar_toast(driver)

    # --- Verificação do resultado da primeira correção ---
    if MENSAGEM_SUCESSO_ESPERADA in toast_mensagem:
        df.at[index, 'Status de Integração'] = 'Integrado (Após Correção 1)'
        df.at[index, 'Console'] = 'Sucesso: Usuário integrado no SAS após correção para "Não".'
        print("Sucesso: Usuário integrado após correção para 'Não'.")
       

    else:

        # Captura a mensagem final
        toast_mensagem_final = capturar_toast(driver)

        # --- Verificação final do resultado ---
        if MENSAGEM_SUCESSO_ESPERADA in toast_mensagem_final:
            df.at[index,
                  'Status de Integração'] = 'Integrado (Após Correção 2)'
            df.at[index, 'Console'] = 'Sucesso: Usuário integrado no SAS após correção para "Sim".'
            print("Sucesso: Usuário integrado após correção para 'Sim'.")
        else:
            df.at[index, 'Status de Integração'] = 'Falha Total'
            df.at[index,
                  'Console'] = f'Falha após tentativa de correção. Mensagem final: {toast_mensagem_final}'
            print("Falha total após tentativa de correção.")


# --- 4. CARREGAMENTO E PREPARAÇÃO DA PLANILHA ---
try:
    df = pd.read_excel('teste.xlsx')
except FileNotFoundError:
    print("Erro: O arquivo 'DadosPIntegrar.xlsx' não foi encontrado.")
    driver.quit()
    exit()

if 'Status de Integração' not in df.columns:
    df['Status de Integração'] = ''
if 'Console' not in df.columns:
    df['Console'] = ''

# --- 5. LOOP PRINCIPAL DE AUTOMAÇÃO E COLETA DE DADOS ---
for index, row in df.iterrows():
    cpf_usuario = str(row['CPF do Usuário'])
    print(f"\n--- Processando CPF: {cpf_usuario} ---")
    driver.get("https://sebraecrm.lightning.force.com/lightning/page/home")
    wait = WebDriverWait(driver, TEMPO_ESPERA)

    # Chama a função principal de automação e recebe a mensagem de volta
    toast_mensagem = integrar_usuario(driver, df, index, cpf_usuario)

    # Lógica de decisão baseada na mensagem retornada
    if MENSAGEM_SUCESSO_ESPERADA in toast_mensagem:
        df.at[index, 'Status de Integração'] = 'Integrado'
        df.at[index, 'Console'] = 'Sucesso: Usuário integrado no SAS.'
        print("Sucesso: Usuário integrado.")

    elif MENSAGEM_ERRO_CONSULTA in toast_mensagem:

        sucesso = tentar_corrigir_consultor(driver, toast_mensagem, df, index)
        if sucesso == True:
           time.sleep(7)
           tentar_integrar_e_verificar(driver, df, index)
        #Se o erro do PJ permanacer, verificar com retrun true ou false na função de reintegrar!
        
        
       

    # Esta verificação lida com mensagens inesperadas, mas que não são erros de automação
    elif "Erro" not in toast_mensagem:
        df.at[index, 'Status de Integração'] = 'Falha'
        df.at[index, 'Console'] = f'Falha na integração: {toast_mensagem}'
        print(f'Falha na integração. Mensagem: {toast_mensagem}')
        
    

# --- 6. SALVAR RESULTADOS E GERAR DASHBOARD DE ANÁLISE ---
nome_arquivo = 'DadosPIntegrar_resultados.xlsx'
df.to_excel(nome_arquivo, index=False)
print("\nProcesso de automação concluído. A planilha de resultados foi salva.")

# Gera a análise de erros
print("Gerando dashboard de análise de erros...")
contagem_erros = df['Console'].value_counts()
contagem_apenas_erros = contagem_erros[~contagem_erros.index.isin(
    ['Sucesso: Usuário integrado no SAS.', ''])]
porcentagem_erros = (contagem_apenas_erros / contagem_apenas_erros.sum()) * 100

df_dashboard = pd.DataFrame({
    'Ocorrências': contagem_apenas_erros,
    'Porcentagem (%)': porcentagem_erros
})

with pd.ExcelWriter(nome_arquivo, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    df_dashboard.to_excel(writer, sheet_name='Análise de Erros')

print("Análise de erros salva na aba 'Análise de Erros'.")
