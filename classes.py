import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select


class Clientes:
    def __init__(self):
        self.cnpj = []
        self.razaoSocial = []
        self.valor = []
        self.mes = None
        self.ano = None
        self.aliquota = None

    def criar_cliente(self):
        arquivo = pd.read_excel('notas fiscais.xlsx', dtype=str)

        for cnpj in arquivo['CNPJ']:
            self.cnpj.append(cnpj)

        for razao in arquivo['RAZÃO SOCIAL']:
            self.razaoSocial.append(razao)

        for preco in arquivo['VALOR']:
            self.valor.append(int(preco))

    def mostrar_clientes(self):
        print(self.cnpj)
        print(self.razaoSocial)
        print(self.valor)

    def selecionar_mes(self):
        while True:
            print('Para qual mês será emitida a nota fiscal?')
            print('Digite 1 para Janeiro')
            print('Digite 2 para Fevereiro')
            print('Digite 3 para Março')
            print('Digite 4 para Abril')
            print('Digite 5 para Maio')
            print('Digite 6 para Junho')
            print('Digite 7 para Julho')
            print('Digite 8 para Agosto')
            print('Digite 9 para Setembro')
            print('Digite 10 para Outubro')
            print('Digite 11 para Novembro')
            print('Digite 12 para Dezembro')
            try:
                self.mes = int(input())
            except Exception:
                print(f'Valor inválido. Valor digitado deve ser um número de 1 a 12.')

            if self.mes == 1:
                print('As notas serão emitidas com o mês de janeiro.')
                break
            elif self.mes == 2:
                print('As notas serão emitidas com o mês de feveiro.')
                break
            elif self.mes == 3:
                print('As notas serão emitidas com o mês de março.')
                break
            elif self.mes == 4:
                print('As notas serão emitidas com o mês de abril.')
                break
            elif self.mes == 5:
                print('As notas serão emitidas com o mês de maio.')
                break
            elif self.mes == 6:
                print('As notas serão emitidas com o mês de junho.')
                break
            elif self.mes == 7:
                print('As notas serão emitidas com o mês de julho.')
                break
            elif self.mes == 8:
                print('As notas serão emitidas com o mês de agosto.')
                break
            elif self.mes == 9:
                print('As notas serão emitidas com o mês de setembro.')
                break
            elif self.mes == 10:
                print('As notas serão emitidas com o mês de outubro.')
                break
            elif self.mes == 11:
                print('As notas serão emitidas com o mês de novembro.')
                break
            elif self.mes == 12:
                print('As notas serão emitidas com o mês de dezembro.')
                break
            else:
                print(f'Valor inválido. Valor digitado deve ser um número de 1 a 12.')

    def selecionar_ano(self):
        print('Para qual ano será emitidas as notas fiscais ?')
        self.ano = int(input())

    def selecionar_aliquota(self):

        print('Digite o valor da alíquota:')
        aliquotaInicial = input()

        listaSemVirgula = []

        for x in aliquotaInicial:
            if x == ',':
                listaSemVirgula.append('.')
            else:
                listaSemVirgula.append(x)

        aliquotaFinal = ""
        for valor in listaSemVirgula:
            aliquotaFinal += valor

        aliquotaFinal = float(aliquotaFinal)
        aliquotaFinal = aliquotaFinal * 100
        aliquotaFinal = round(aliquotaFinal)
        self.aliquota = aliquotaFinal

    def listarClientes(self):
        # Mostrando os dados importandos
        input('Pressione ENTER para ver a lista de clientes:\n')

        posicao = 0
        for _ in self.cnpj:
            print(f'CNPJ:{self.cnpj[posicao]} -> Razão: {self.razaoSocial[posicao]} -> Valor: R$ {self.valor[posicao]}')
            posicao += 1

        totalValores = 0
        for valor in self.valor:
            totalValores += valor

        print('\nAcima está a lista de clientes importados.')
        print(f'Serão emitidas {len(self.cnpj)} NF-e, ao todo R$ {totalValores}.')
        input('Pressione ENTER para prosseguir')


class ChromeAuto:
    def __init__(self):
        self.login = '013.681.763-72'
        self.senha = 'BAE32066'
        self.url = 'https://iss.fortaleza.ce.gov.br/grpfor/login.seam?cid=36452'
        self.driver = webdriver.Chrome()

    def abrir_site(self):
        self.driver.get(self.url)
        fieldLogin = self.driver.find_element(By.XPATH, '//*[@id="login:username"]')
        fieldPassword = self.driver.find_element(By.XPATH, '//*[@id="login:password"]')

        if fieldLogin.is_displayed() and fieldLogin.is_enabled():
            fieldLogin.send_keys(self.login)

        if fieldPassword.is_displayed() and fieldPassword.is_enabled():
            fieldPassword.send_keys(self.senha)

        input('Faça login, depois volte aqui e pressione ENTER para o programa prosseguir:')

    @staticmethod
    def baixar_nota():
        global verificador

        print('\nVocê deseja baixar as notas fiscais (PDF) para seu computador quando elas forem emitidas?')
        print('Digite 1 para SIM ou digite 2 para NÃO')
        verificador = int(input())
        print()

        if verificador > 2 or verificador < 1:
            print('\nVALOR INVÁLIDO\n')
            ChromeAuto.baixar_nota()

    def gerar_notas(self, dadosClientes):

        # OBTENDO O USUÁRIO DO COMPUTADOR
        usuario = os.getlogin()

        # Reload na página
        self.driver.find_element(By.XPATH, '//*[@id="j_id7"]/img').click()

        posicao = 0
        for _ in dadosClientes.cnpj:
            # Clica em emitir notas
            self.driver.find_element(By.XPATH, '//*[@id="homeForm:divHotLinks"]/div[1]/a/h4').click()

            # Clica em CNPJ
            self.driver.find_element(By.XPATH, '//*[@id="emitirnfseForm:tipoPesquisaTomadorRb:1"]').click()
            time.sleep(2)

            # Digita o CNPJ
            self.driver.find_element(By.XPATH, '//*[@id="emitirnfseForm:cpfPesquisaTomador"]').send_keys(
                dadosClientes.cnpj[posicao])
            time.sleep(2)

            # Seleciona o cliente
            self.driver.find_element(By.XPATH, '//*[@id="emitirnfseForm:cpfPesquisaTomador"]').send_keys('\ue007')
            time.sleep(2)

            nome_denominacao = self.driver.find_element(By.XPATH, '//*[@id="emitirnfseForm:idNome"]').get_attribute(
                'value')

            if len(nome_denominacao) > 0:
                # Clica em serviço
                self.driver.find_element(By.XPATH, '//*[@id="emitirnfseForm:abaServico_lbl"]').click()

                # Selecionando o ano
                seletor_ano = self.driver.find_element(By.ID, 'emitirnfseForm:comboEscolherAnoCompetencia')
                seletor_ano_dd = Select(seletor_ano)
                seletor_ano_dd.select_by_value(str(dadosClientes.ano))

                # Selecionando o mês
                seletor_mes = self.driver.find_element(By.ID, 'emitirnfseForm:comboEscolherMesCompetencia')
                seletor_mes_dd = Select(seletor_mes)
                seletor_mes_dd.select_by_value(str(dadosClientes.mes))
                time.sleep(2)

                # DISCRIMINAÇÃO DO SERVIÇO
                seletor_cnae = self.driver.find_element(By.ID, 'emitirnfseForm:comboEscolherAtividadeCpbs')
                seletor_cnae_dd = Select(seletor_cnae)
                seletor_cnae_dd.select_by_visible_text(
                    'ATIVIDADES DE INTERMEDIAÇÃO E AGENCIAMENTO DE SERVIÇOS E NEGÓCIOS EM GERAL, EXCETO IMOBILIÁRIOS')
                time.sleep(2)

                # ALIQUOTA
                self.driver.find_element(By.ID, 'emitirnfseForm:idAliquota').send_keys(dadosClientes.aliquota)

                # Descrição do Serviço
                fieldDescricaoServico = self.driver.find_element(By.XPATH,
                                                                 '//*[@id="emitirnfseForm:idDescricaoServico"]')
                if dadosClientes.mes == 1:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Janeiro.')
                elif dadosClientes.mes == 2:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Fevereiro.')
                elif dadosClientes.mes == 3:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Março.')
                elif dadosClientes.mes == 4:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Abril.')
                elif dadosClientes.mes == 5:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Maio.')
                elif dadosClientes.mes == 6:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Junho.')
                elif dadosClientes.mes == 7:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Julho.')
                elif dadosClientes.mes == 8:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Agosto.')
                elif dadosClientes.mes == 9:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Setembro.')
                elif dadosClientes.mes == 10:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Outubro.')
                elif dadosClientes.mes == 11:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Novembro.')
                elif dadosClientes.mes == 12:
                    fieldDescricaoServico.send_keys('Serviços prestados no mês de Dezembro.')

                # VALOR DO SERVIÇO
                self.driver.find_element(By.ID, 'emitirnfseForm:idValorServicoPrestado').send_keys(
                    dadosClientes.valor[posicao] * 100)

                # VALIDAR CAMPOS OBRIGATÓRIOS DA NF-E
                return
                self.driver.find_element(By.XPATH, '//*[@id="emitirnfseForm:btnCalcular"]').click()

                # CONFIRMAÇÃO DA EMISSÃO DA NOTA
                self.driver.find_element(By.XPATH,
                                         '/html/body/div[2]/div[2]/form/table[2]/tbody/tr/td[1]/input').click()

                # RELOAD DA PÁGINA
                time.sleep(2)

                # MOSTRANDO A NOTA QUE FOI FEITA
                print(f'Nota fiscal do CNPJ {dadosClientes.cnpj[posicao]} emitida com sucesso.')

                # BAIXAR NOTA
                if verificador == 1:
                    # CLICA EM BAIXAR NOTA PDF
                    self.driver.find_element(By.XPATH, '//*[@id="j_id160"]/table[2]/tbody/tr/td[2]/input').click()

                    # AGUARDANDO DOWNLOAD
                    time.sleep(2)

                    # RENOMEANDO ARQUIVO
                    oldName = f'C:\\Users\\{usuario}\\Downloads\\relatorio.pdf'
                    newName = f'C:\\Users\\{usuario}\\Downloads\\' \
                              f'{dadosClientes.razao[posicao]} {dadosClientes.cnpj[posicao]}.pdf'
                    os.rename(oldName, newName)

                # VOLTA AO INÍCIO
                self.driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/div[1]/a/img').click()

                # MUDANDO AS INFORMAÇÕES PARA O PRÓXIMO CLIENTE
                posicao = posicao + 1
            else:
                print(f'Cliente {dadosClientes.cnpj[posicao]} não cadastrado')
                # MUDANDO AS INFORMAÇÕES PARA O PRÓXIMO CLIENTE
                posicao = posicao + 1

                # VOLTA AO INÍCIO
                self.driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/div/div/div[1]/a/img').click()

    def encerrar(self):
        self.driver.quit()


class OutrasFuncoes:
    def __init__(self):
        # Créditos ao desenvolvedor
        print('DESENVOLVIDO POR GABRIEL FELIPE\n')

        # Como utilizar o programa da maneira correta.
        print('PARA O BOM FUNCIONAMENTO DO PROGRAMA:')
        print('Na mesma pasta que você está executando esse programa, tenha uma planilha chamada "notas fiscais".')
        print('A planilha deve conter apenas 3 colunas: CNPJ, Razão Social e VALOR.')
        print('A coluna CNPJ serão as lojas que serão emitidas.')
        print('A coluna VALOR serão os valores que serão emitidos para cada loja/CNPJ.\n')
        print('Lembre-se de REMOVER os seguintes clientes:')
        print('1. Clientes que não pagaram a mensalidade do mês;')
        print('2. Clientes que nunca pagaram;')
        print('3. Clientes que pagam por outro CNPJ;')
        print('4. Clientes que pagam com depósito;')
        print('5. Clientes de adesão;')
        print('6. Clientes da Mary;')
        input('\nSe você já possuir uma planilha com esse modelo, aperte enter para prosseguir.\n')
