import time
import pyautogui

import pandas as pd


tabela = pd.read_excel(r"C:\BotPerdCompPrevidencia\PerdCompPrevidencia_dados.xlsx", sheet_name="PerdCompPrev_Robot02").astype(str)
#total_rows=len(tabela.axes[0]) #===> Axes of 0 is for a row
#print(total_rows)
#print(tabela)

A_01 = 	tabela['DataCriacao']
A_02 = 	tabela['Contribuiente']
A_03 = 	tabela['CNPJ_CPF']
A_04 = 	tabela['Tipo_de_Credito']
A_05 = 	tabela['Numero_de_Indetificacao_do_Trabalhador']
A_06 = 	tabela['AnoCompetencia']
A_07 = 	tabela['MesCompetencia']
A_08 = 	tabela['Nome_do_Trabalhador']
A_09 = 	tabela['DataNascimento']
A_10 = 	tabela['Tipo_de_Conta']
A_11 = 	tabela['Banco']
A_12 = 	tabela['Agencia']
A_13 = 	tabela['ContaCorrente']
A_14 = 	tabela['DV']
A_15 = 	tabela['TipoCreditos']
A_16 = 	tabela['Categoria_Segurado']
A_17 = 	tabela['Justificativa_do_Pedido']
A_18 = 	tabela['DDD']
A_19 = 	tabela['Telefone']
A_20 = 	tabela['Valor_do_Pedido_Restituico']
A_21 = 	tabela['Ano_01']
A_22 = 	tabela['Mes_01']
A_23 = 	tabela['CNPJ_CEI_01']
A_24 = 	tabela['Nome_Pessoa_Fisica_01']
A_25 = 	tabela['Remuneracao_Recebida_01']
A_26 = 	tabela['Valor_Contribuicao_Descontada_01']
A_27 = 	tabela['Ano_02']
A_28 = 	tabela['Mes_02']
A_29 = 	tabela['CNPJ_CEI_02']
A_30 = 	tabela['Nome_Pessoa_Fisica_02']
A_31 = 	tabela['Remuneracao_Recebida_02']
A_32 = 	tabela['Valor_Contribuicao_Descontada_02']
A_33 = 	tabela['Ano_03']
A_34 = 	tabela['Mes_03']
A_35 = 	tabela['CNPJ_CEI_03']
A_36 = 	tabela['Nome_Pessoa_Fisica_03']
A_37 = 	tabela['Remuneracao_Recebida_03']
A_38 = 	tabela['Valor_Contribuicao_Descontada_03']
A_39 = 	tabela['Ano_04']
A_40 = 	tabela['Mes_04']
A_41 = 	tabela['CNPJ_CEI_04']
A_42 = 	tabela['Nome_Pessoa_Fisica_04']
A_43 = 	tabela['Remuneracao_Recebida_04']
A_44 = 	tabela['Valor_Contribuicao_Descontada_04']
A_45 = 	tabela['Ano_05']
A_46 = 	tabela['Mes_05']
A_47 = 	tabela['CNPJ_CEI_05']
A_48 = 	tabela['Nome_Pessoa_Fisica_05']
A_49 = 	tabela['Remuneracao_Recebida_05']
A_50 = 	tabela['Valor_Contribuicao_Descontada_05']
A_51 = 	tabela['Ano_06']
A_52 = 	tabela['Mes_06']
A_53 = 	tabela['CNPJ_CEI_06']
A_54 = 	tabela['Nome_Pessoa_Fisica_06']
A_55 = 	tabela['Remuneracao_Recebida_06']
A_56 = 	tabela['Valor_Contribuicao_Descontada_06']
A_57 = 	tabela['Ano_07']
A_58 = 	tabela['Mes_07']
A_59 = 	tabela['CNPJ_CEI_07']
A_60 = 	tabela['Nome_Pessoa_Fisica_07']
A_61 = 	tabela['Remuneracao_Recebida_07']
A_62 = 	tabela['Valor_Contribuicao_Descontada_07']
A_63 = 	tabela['Ano_08']
A_64 = 	tabela['Mes_08']
A_65 = 	tabela['CNPJ_CEI_08']
A_66 = 	tabela['Nome_Pessoa_Fisica_08']
A_67 = 	tabela['Remuneracao_Recebida_08']
A_68 = 	tabela['Valor_Contribuicao_Descontada_08']
A_69 = 	tabela['Ano_09']
A_70 = 	tabela['Mes_09']
A_71 = 	tabela['CNPJ_CEI_09']
A_72 = 	tabela['Nome_Pessoa_Fisica_09']
A_73 = 	tabela['Remuneracao_Recebida_09']
A_74 = 	tabela['Valor_Contribuicao_Descontada_09']
A_75 = 	tabela['Ano_10']
A_76 = 	tabela['Mes_10']
A_77 = 	tabela['CNPJ_CEI_10']
A_78 = 	tabela['Nome_Pessoa_Fisica_10']
A_79 = 	tabela['Remuneracao_Recebida_10']
A_80 = 	tabela['Valor_Contribuicao_Descontada_10']
A_81 = 	tabela['Ano_11']
A_82 = 	tabela['Mes_11']
A_83 = 	tabela['CNPJ_CEI_11']
A_84 = 	tabela['Nome_Pessoa_Fisica_11']
A_85 = 	tabela['Remuneracao_Recebida_11']
A_86 = 	tabela['Valor_Contribuicao_Descontada_11']
A_87 = 	tabela['Ano_12']
A_88 = 	tabela['Mes_12']
A_89 = 	tabela['CNPJ_CEI_12']
A_90 = 	tabela['Nome_Pessoa_Fisica_12']
A_91 = 	tabela['Remuneracao_Recebida_12']
A_92 = 	tabela['Valor_Contribuicao_Descontada_12']




pyautogui.hotkey('winleft','r')
pyautogui.write(r"C:\Arquivos de Programas RFB\PERDCOMP69\PERDCOMP69.exe")
pyautogui.press('enter')
time.sleep(4)


for DataCriacao,Contribuiente,CNPJ_CPF,Tipo_de_Credito,Numero_de_Indetificacao_do_Trabalhador,AnoCompetencia,MesCompetencia,Nome_do_Trabalhador\
,DataNascimento,Tipo_de_Conta,Banco,Agencia,ContaCorrente,DV,TipoCreditos,Categoria_Segurado,Justificativa_do_Pedido,DDD,Telefone,Valor_do_Pedido_Restituico\
,Ano_01,Mes_01,CNPJ_CEI_01,Nome_Pessoa_Fisica_01,Remuneracao_Recebida_01,Valor_Contribuicao_Descontada_01,Ano_02,Mes_02,CNPJ_CEI_02,Nome_Pessoa_Fisica_02,Remuneracao_Recebida_02\
,Valor_Contribuicao_Descontada_02,Ano_03,Mes_03,CNPJ_CEI_03,Nome_Pessoa_Fisica_03,Remuneracao_Recebida_03,Valor_Contribuicao_Descontada_03,Ano_04,Mes_04,CNPJ_CEI_04,Nome_Pessoa_Fisica_04\
,Remuneracao_Recebida_04,Valor_Contribuicao_Descontada_04,Ano_05,Mes_05,CNPJ_CEI_05,Nome_Pessoa_Fisica_05,Remuneracao_Recebida_05,Valor_Contribuicao_Descontada_05,Ano_06,Mes_06,CNPJ_CEI_06,Nome_Pessoa_Fisica_06,Remuneracao_Recebida_06\
,Valor_Contribuicao_Descontada_06,Ano_07,Mes_07,CNPJ_CEI_07,Nome_Pessoa_Fisica_07,Remuneracao_Recebida_07,Valor_Contribuicao_Descontada_07,Ano_08,Mes_08,CNPJ_CEI_08,Nome_Pessoa_Fisica_08\
,Remuneracao_Recebida_08,Valor_Contribuicao_Descontada_08,Ano_09,Mes_09,CNPJ_CEI_09,Nome_Pessoa_Fisica_09,Remuneracao_Recebida_09,Valor_Contribuicao_Descontada_09,Ano_10,Mes_10\
,CNPJ_CEI_10,Nome_Pessoa_Fisica_10,Remuneracao_Recebida_10,Valor_Contribuicao_Descontada_10,Ano_11,Mes_11,CNPJ_CEI_11,Nome_Pessoa_Fisica_11,Remuneracao_Recebida_11\
,Valor_Contribuicao_Descontada_11,Ano_12,Mes_12,CNPJ_CEI_12,Nome_Pessoa_Fisica_12,Remuneracao_Recebida_12,Valor_Contribuicao_Descontada_12 in zip(A_01,A_02,A_03,A_04,A_05,A_06,A_07,A_08,A_09,A_10,A_11,A_12,A_13,A_14,A_15,A_16,A_17,A_18,A_19,A_20,A_21,A_22,A_23,A_24,A_25,A_26,A_27,A_28,A_29,A_30,A_31,A_32,A_33,A_34,A_35,A_36,A_37,A_38,A_39,A_40,A_41,A_42,A_43,A_44,A_45,A_46,A_47,A_48,A_49,A_50,A_51,A_52,A_53,A_54,A_55,A_56,A_57,A_58,A_59,A_60,A_61,A_62,A_63,A_64,A_65,A_66,A_67,A_68,A_69,A_70,A_71,A_72,A_73,A_74,A_75,A_76,A_77,A_78,A_79,A_80,A_81,A_82,A_83,A_84,A_85,A_86,A_87,A_88,A_89,A_90,A_91,A_92):

    def DeslocaparaBaixoPrimeitaTela():
        DescolocarAbaixo = pyautogui.position(524, 583)
        pyautogui.moveTo(DescolocarAbaixo)
        pyautogui.click()
        pyautogui.click()
        pyautogui.click()
        pyautogui.click()
        pyautogui.click()


    def DeslocaparaBaixoUltimaTela():
        DescolocarAbaixoUltimatelas = pyautogui.position(286, 264)
        pyautogui.moveTo(DescolocarAbaixoUltimatelas)
        pyautogui.click()
        pyautogui.click()
        pyautogui.click()
        pyautogui.click()
        pyautogui.click()

    def CompetenciaPrimeiraTela(Ano):
       if Ano == '2022':
           Ano2022 = pyautogui.position(349, 483)
           pyautogui.moveTo(Ano2022)
           pyautogui.click()
       if Ano == '2021':
           Ano2021 = pyautogui.position(349, 498)
           pyautogui.moveTo(Ano2021)
           pyautogui.click()
       if Ano == '2020':
           Ano2020 = pyautogui.position(349, 514)
           pyautogui.moveTo(Ano2020)
           pyautogui.click()
       if Ano == '2019':
           Ano2019 = pyautogui.position(349, 526)
           pyautogui.moveTo(Ano2019)
           pyautogui.click()
       if Ano == '2018':
           Ano2018 = pyautogui.position(349, 540)
           pyautogui.moveTo(Ano2018)
           pyautogui.click()
       if Ano == '2017':
           Ano2017 = pyautogui.position(349, 554)
           pyautogui.moveTo(Ano2017)
           pyautogui.click()
       if Ano == '2016':
           Ano2016 = pyautogui.position(349, 568)
           pyautogui.moveTo(Ano2016)
           pyautogui.click()
       if Ano == '2015':
           Ano2015 = pyautogui.position(349, 583)
           pyautogui.moveTo(Ano2015)
           pyautogui.click()

    def MESPrimeiraTela(Mes):

       if Mes == '1':
           MES01 = pyautogui.position(447, 485)
           pyautogui.moveTo(MES01)
           print(MES01)
           pyautogui.click()

       if Mes == '2':
           MES02 = pyautogui.position(447, 499)
           pyautogui.moveTo(MES02)
           pyautogui.click()

       if Mes == '3':
           MES03 = pyautogui.position(447, 513)
           pyautogui.moveTo(MES03)
           pyautogui.click()

       if Mes == '4':
           MES04 = pyautogui.position(447, 525)
           pyautogui.moveTo(MES04)
           pyautogui.click()

       if Mes == '5':
           MES05 = pyautogui.position(447, 541)
           pyautogui.moveTo(MES05)
           pyautogui.click()

       if Mes == '6':
           MES06 = pyautogui.position(447, 554)
           pyautogui.moveTo(MES06)
           pyautogui.click()

       if Mes == '7':
           MES07 = pyautogui.position(447, 568)
           pyautogui.moveTo(MES07)
           pyautogui.click()

       if Mes == '8':
           MES08 = pyautogui.position(447, 583)
           pyautogui.moveTo(MES08)
           pyautogui.click()

       if Mes == '9':
           DeslocaparaBaixoPrimeitaTela()
           time.sleep(2)
           MES09 = pyautogui.position(447, 527)
           pyautogui.moveTo(MES09)
           pyautogui.click()

       if Mes == '10':
           DeslocaparaBaixoPrimeitaTela()
           time.sleep(2)
           MES10 = pyautogui.position(447, 540)
           pyautogui.moveTo(MES10)
           pyautogui.click()

       if Mes == '11':
           DeslocaparaBaixoPrimeitaTela()
           time.sleep(2)
           MES11 = pyautogui.position(447, 554)
           pyautogui.moveTo(MES11)
           pyautogui.click()
       if Mes == '12':
           DeslocaparaBaixoPrimeitaTela()
           time.sleep(2)
           MES12 = pyautogui.position(447, 568)
           pyautogui.moveTo(MES12)
           pyautogui.click()
       if Mes == '13':
           DeslocaparaBaixoPrimeitaTela()
           time.sleep(2)
           MES13 = pyautogui.position(447, 581)
           pyautogui.moveTo(MES13)
           pyautogui.click()


##########################FUNÇAO PARA TRATAR OS COMBOX DA ULTIMA TELA DE PREENCHIMENTO

    def CompetenciaUltimaTela(AnoUltimaTela):
       if AnoUltimaTela == '2022':
           Ano2022 = pyautogui.position(148, 168)
           pyautogui.moveTo(Ano2022)
           pyautogui.click()
       if AnoUltimaTela == '2021':
           Ano2021 = pyautogui.position(148, 182)
           pyautogui.moveTo(Ano2021)
           pyautogui.click()
       if AnoUltimaTela == '2020':
           Ano2020 = pyautogui.position(148, 195)
           pyautogui.moveTo(Ano2020)
           pyautogui.click()
       if AnoUltimaTela == '2019':
           Ano2019 = pyautogui.position(148, 210)
           pyautogui.moveTo(Ano2019)
           pyautogui.click()
       if AnoUltimaTela == '2018':
           Ano2018 = pyautogui.position(148, 225)
           pyautogui.moveTo(Ano2018)
           pyautogui.click()
       if AnoUltimaTela == '2017':
           Ano2017 = pyautogui.position(148, 238)
           pyautogui.moveTo(Ano2017)
           pyautogui.click()
       if AnoUltimaTela == '2016':
           Ano2016 = pyautogui.position(148, 252)
           pyautogui.moveTo(Ano2016)
           pyautogui.click()
       if AnoUltimaTela == '2015':
           Ano2015 = pyautogui.position(148, 266)
           pyautogui.moveTo(Ano2015)
           pyautogui.click()

    def MESUltimaTela(MesUltimaTela):

       if MesUltimaTela == '1':
           MES01 = pyautogui.position(245, 167)
           pyautogui.moveTo(MES01)
           print(MES01)
           pyautogui.click()

       if MesUltimaTela == '2':
           MES02 = pyautogui.position(245, 182)
           pyautogui.moveTo(MES02)
           pyautogui.click()

       if MesUltimaTela == '3':
           MES03 = pyautogui.position(245, 195)
           pyautogui.moveTo(MES03)
           pyautogui.click()

       if MesUltimaTela == '4':
           MES04 = pyautogui.position(245, 210)
           pyautogui.moveTo(MES04)
           pyautogui.click()

       if MesUltimaTela == '5':
           MES05 = pyautogui.position(245, 223)
           pyautogui.moveTo(MES05)
           pyautogui.click()

       if MesUltimaTela == '6':
           MES06 = pyautogui.position(245, 237)
           pyautogui.moveTo(MES06)
           pyautogui.click()

       if MesUltimaTela == '7':
           MES07 = pyautogui.position(245, 251)
           pyautogui.moveTo(MES07)
           pyautogui.click()

       if MesUltimaTela == '8':
           MES08 = pyautogui.position(245, 267)
           pyautogui.moveTo(MES08)
           pyautogui.click()

       if MesUltimaTela == '9':
           DeslocaparaBaixoUltimaTela()
           time.sleep(2)
           MES09 = pyautogui.position(245, 210)
           pyautogui.moveTo(MES09)
           pyautogui.click()

       if MesUltimaTela == '10':
           DeslocaparaBaixoUltimaTela()
           time.sleep(2)
           MES10 = pyautogui.position(245, 225)
           pyautogui.moveTo(MES10)
           pyautogui.click()

       if MesUltimaTela == '11':
           DeslocaparaBaixoUltimaTela()
           time.sleep(2)
           MES11 = pyautogui.position(245, 239)
           pyautogui.moveTo(MES11)
           pyautogui.click()
       if MesUltimaTela == '12':
           DeslocaparaBaixoUltimaTela()
           time.sleep(2)
           MES12 = pyautogui.position(245, 252)
           pyautogui.moveTo(MES12)
           pyautogui.click()
       if MesUltimaTela == '13':
           DeslocaparaBaixoUltimaTela()
           time.sleep(2)
           MES13 = pyautogui.position(245, 265)
           pyautogui.moveTo(MES13)
           pyautogui.click()


    def FunctionClicaBotaoIncluir():

        ClicarEmIncluirUltimaTela = pyautogui.position(744, 121)
        pyautogui.moveTo(ClicarEmIncluirUltimaTela)
        pyautogui.click()


    NovoDocumento = pyautogui.position(21,59)
    pyautogui.moveTo(NovoDocumento)
    pyautogui.click()

    DtCriacao = pyautogui.position(322, 135)
    pyautogui.moveTo(DtCriacao)
    pyautogui.click()
    pyautogui.write(DataCriacao.replace('/',''))

    SelecionaContribuiente = pyautogui.position(493, 133)
    pyautogui.moveTo(SelecionaContribuiente)
    pyautogui.click()

    SelecionaPF = pyautogui.position(455, 166)
    pyautogui.moveTo(SelecionaPF)
    pyautogui.click()

    PreencherCNPJCPF = pyautogui.position(512, 135)
    pyautogui.moveTo(PreencherCNPJCPF)
    pyautogui.click()
    pyautogui.write(CNPJ_CPF.replace('/',''))

    time.sleep(2)

    SelecionaTipoDocumento = pyautogui.position(455, 185)
    pyautogui.moveTo(SelecionaTipoDocumento)
    pyautogui.click()

    SelecionarPedidoRestituicao = pyautogui.position(375, 218)
    pyautogui.moveTo(SelecionarPedidoRestituicao)
    pyautogui.click()

    time.sleep(2)

    SelecionarTipoCredito = pyautogui.position(943, 185)
    pyautogui.moveTo(SelecionarTipoCredito)
    pyautogui.click()
    time.sleep(2)


    SelecionarContriPrevIndevida= pyautogui.position(693, 218)
    pyautogui.moveTo(SelecionarContriPrevIndevida)
    pyautogui.click()
    time.sleep(2)

    InfoNumeroTrabalhador= pyautogui.position(333, 402)
    pyautogui.moveTo(InfoNumeroTrabalhador)
    pyautogui.click()
    pyautogui.write(Numero_de_Indetificacao_do_Trabalhador)
    time.sleep(2)

    SelecionaCompetenciaTela1 = pyautogui.position(422, 465)
    pyautogui.moveTo(SelecionaCompetenciaTela1)
    pyautogui.click()

    #print(AnoCompetencia)


    CompetenciaPrimeiraTela(str(AnoCompetencia))
    time.sleep(2)
    #print(MesCompetencia)


    SelecionaMesTela1 = pyautogui.position(522, 467)
    pyautogui.moveTo(SelecionaMesTela1)
    pyautogui.click()

    MESPrimeiraTela(str(MesCompetencia))

    time.sleep(2)

    ConfirmarDadosExistentes = pyautogui.position(751, 449)
    pyautogui.moveTo(ConfirmarDadosExistentes)
    pyautogui.click()

    ConfirmarDadosExistentesP = pyautogui.position(745, 439)
    pyautogui.moveTo(ConfirmarDadosExistentesP)
    pyautogui.click()

    ConfirmarDadosExistentesP = pyautogui.position(754, 449)
    pyautogui.moveTo(ConfirmarDadosExistentesP)
    pyautogui.click()

    ConfirmarDadosExistentesP = pyautogui.position(748, 443)
    pyautogui.moveTo(ConfirmarDadosExistentesP)
    pyautogui.click()

    ConfirmarDataEmissaoTransport = pyautogui.position(752, 449)
    pyautogui.moveTo(ConfirmarDataEmissaoTransport)
    pyautogui.click()

    time.sleep(4)
    ClicarEmOKPrimeiraTela = pyautogui.position(584, 556)
    pyautogui.moveTo(ClicarEmOKPrimeiraTela)
    pyautogui.click()

    time.sleep(2)

    ConfirmaraDataCriacao = pyautogui.position(752, 451)
    pyautogui.moveTo(ConfirmaraDataCriacao)
    pyautogui.click()
    ConfirmarDadosExistentes = pyautogui.position(751, 449)
    pyautogui.moveTo(ConfirmarDadosExistentes)
    pyautogui.click()

    ConfirmarDadosExistentesP = pyautogui.position(745, 439)
    pyautogui.moveTo(ConfirmarDadosExistentesP)
    pyautogui.click()

    ConfirmarDadosExistentesP = pyautogui.position(754, 449)
    pyautogui.moveTo(ConfirmarDadosExistentesP)
    pyautogui.click()

    ConfirmarDadosExistentesP = pyautogui.position(748, 443)
    pyautogui.moveTo(ConfirmarDadosExistentesP)
    pyautogui.click()

    ConfirmarDataEmissaoTransport = pyautogui.position(752, 449)
    pyautogui.moveTo(ConfirmarDataEmissaoTransport)
    pyautogui.click()

    ##Final da programaçao da primeira Tela

    time.sleep(2)

    # Inicio da Tela dos Dados Iniciais
    InfoNomeTrabalhador = pyautogui.position(140, 144)
    pyautogui.moveTo(InfoNomeTrabalhador)
    pyautogui.click()
    time.sleep(2)#Default before equals 4
    pyautogui.write(str(Nome_do_Trabalhador.upper()))
    time.sleep(2)

    InfoDataNascimento = pyautogui.position(673, 144)
    pyautogui.moveTo(InfoDataNascimento)
    pyautogui.click()
    time.sleep(1)
    pyautogui.write(str(DataNascimento.replace('/','')))
    time.sleep(2)

    InfoTipoconta = pyautogui.position(446, 327)
    pyautogui.moveTo(InfoTipoconta)
    pyautogui.click()

    SelecionaTipoContaCorrente = pyautogui.position(389, 365)
    pyautogui.moveTo(SelecionaTipoContaCorrente)
    pyautogui.click()

    InfoCodigoBanco = pyautogui.position(464, 334)
    pyautogui.moveTo(InfoCodigoBanco)
    pyautogui.click()
    pyautogui.write(str(Banco.replace('/','')))
    pyautogui.press('tab')
    time.sleep(2)

    ConfirmaroBancoSelecionado = pyautogui.position(656, 439)
    pyautogui.moveTo(ConfirmaroBancoSelecionado)
    pyautogui.click()
    time.sleep(2)

    InfoAgenciaBanco = pyautogui.position(509, 334)
    pyautogui.moveTo(InfoAgenciaBanco)
    pyautogui.click()
    time.sleep(2)
    pyautogui.write(str(Agencia))

    InfoContaCorrente = pyautogui.position(567, 334)
    pyautogui.moveTo(InfoContaCorrente)
    pyautogui.click()
    pyautogui.write(str(ContaCorrente))

    InfoDigitoVerificadorConta = pyautogui.position(645, 334)
    pyautogui.moveTo(InfoDigitoVerificadorConta)
    pyautogui.click()
    pyautogui.write(str(DV))

    SelecioneoUltimoComboComoNao = pyautogui.position(472, 513)
    pyautogui.moveTo(SelecioneoUltimoComboComoNao)
    pyautogui.click()

    MarcaUltimoComboComoNao = pyautogui.position(407, 547)
    pyautogui.moveTo(MarcaUltimoComboComoNao)
    pyautogui.click()

#Final da Tela dos Dados Iniciais
    time.sleep(2)

#Seleciona a Tela referente  a Creditos

    SelecionaTelaCreditoss = pyautogui.position(54, 716)
    pyautogui.moveTo(SelecionaTelaCreditoss)
    pyautogui.click()

    SelecionaCategoriaDoSegurado = pyautogui.position(323, 419)
    pyautogui.moveTo(SelecionaCategoriaDoSegurado)
    pyautogui.click()

    SelecionaCategoriaDoSeguradoIndividual = pyautogui.position(178, 483)
    pyautogui.moveTo(SelecionaCategoriaDoSeguradoIndividual)
    pyautogui.click()

    SelecionaJustificativaDoPedido = pyautogui.position(649, 419)
    pyautogui.moveTo(SelecionaJustificativaDoPedido)
    pyautogui.click()

    SelecionaOpcaoContriAcimaValor = pyautogui.position(390, 456)
    pyautogui.moveTo(SelecionaOpcaoContriAcimaValor)
    pyautogui.click()

    InfoDDD = pyautogui.position(306, 463)
    pyautogui.moveTo(InfoDDD)
    pyautogui.click()
    pyautogui.write(str(DDD))
    time.sleep(2)

    InfoTelefone = pyautogui.position(391, 466)
    pyautogui.moveTo(InfoTelefone)
    pyautogui.click()
    pyautogui.write(str(Telefone))

    Tamanho = len(str(Valor_do_Pedido_Restituico))

    ValorInteiro = (str(Valor_do_Pedido_Restituico)[:(Tamanho - 3)])
    ValorDecimal = (str(Valor_do_Pedido_Restituico)[(Tamanho - 3) + 1:])

    InfoValorInteiro = pyautogui.position(751, 467)
    pyautogui.moveTo(InfoValorInteiro)
    pyautogui.click()
    pyautogui.write(str(ValorInteiro))

    InfoValorDecimal = pyautogui.position(772, 467)
    pyautogui.moveTo(InfoValorDecimal)
    pyautogui.click()
    pyautogui.press('left')
    pyautogui.press('left')
    pyautogui.write(str(ValorDecimal))


    SelecionarContribuicaoDescontada = pyautogui.position(74, 158)
    pyautogui.moveTo(SelecionarContribuicaoDescontada)
    pyautogui.click()
    time.sleep(2)

    #Final da Tela Contribuiçao Previdenciaria

    #Inicio da Contribuiçao Descontada


    ClicarnoBotaoIncluir = pyautogui.position(746, 123)
    pyautogui.moveTo(ClicarnoBotaoIncluir)
    pyautogui.click()


    ClicaComboUltimaTelaAno = pyautogui.position(216, 150)
    pyautogui.moveTo(ClicaComboUltimaTelaAno)
    pyautogui.click()


    CompetenciaUltimaTela(Ano_01)
    time.sleep(2)

    ClicaComboUltimaTelaMes = pyautogui.position(286, 150)
    pyautogui.moveTo(ClicaComboUltimaTelaMes)
    pyautogui.click()
    MESUltimaTela(Mes_01)

    time.sleep(2)
    InfoCNPJ_CEI_01 = pyautogui.position(366, 150)
    pyautogui.moveTo(InfoCNPJ_CEI_01)
    pyautogui.click()
    pyautogui.write(CNPJ_CEI_01)

    InfoNome_Pessoa_Fisica_01 = pyautogui.position(132, 200)
    pyautogui.moveTo(InfoNome_Pessoa_Fisica_01)
    pyautogui.click()
    pyautogui.write(str(Nome_Pessoa_Fisica_01.upper()))

    time.sleep(2)
    InfoRemuneracao_Recebida_01 = pyautogui.position(232, 257)
    pyautogui.moveTo(InfoRemuneracao_Recebida_01)
    pyautogui.click()
    pyautogui.write(str(Remuneracao_Recebida_01))
    time.sleep(2)

    InfoValor_Contribuicao_Descontada_01 = pyautogui.position(674, 255)
    pyautogui.moveTo(InfoValor_Contribuicao_Descontada_01)
    pyautogui.click()
    pyautogui.write(str(Valor_Contribuicao_Descontada_01))

    time.sleep(2)
    ClicaremOkParaInserirRegistro = pyautogui.position(748, 124)
    pyautogui.moveTo(ClicaremOkParaInserirRegistro)
    pyautogui.click()

    ConfirmarValoresInseridos = pyautogui.position(834, 442)
    pyautogui.moveTo(ConfirmarValoresInseridos)
    pyautogui.click()

    FunctionClicaBotaoIncluir()
    #Segunda informações a serem inseridas

    ClicaComboUltimaTelaAno = pyautogui.position(216, 150)
    pyautogui.moveTo(ClicaComboUltimaTelaAno)
    pyautogui.click()

    CompetenciaUltimaTela(Ano_02)
    time.sleep(2)

    ClicaComboUltimaTelaMes = pyautogui.position(286, 150)
    pyautogui.moveTo(ClicaComboUltimaTelaMes)
    pyautogui.click()
    MESUltimaTela(Mes_02)

    time.sleep(2)
    InfoCNPJ_CEI_02 = pyautogui.position(366, 150)
    pyautogui.moveTo(InfoCNPJ_CEI_02)
    pyautogui.click()
    pyautogui.write(CNPJ_CEI_02)

    InfoNome_Pessoa_Fisica_02 = pyautogui.position(132, 200)
    pyautogui.moveTo(InfoNome_Pessoa_Fisica_02)
    pyautogui.click()
    pyautogui.write(str(Nome_Pessoa_Fisica_02.upper()))

    time.sleep(2)
    InfoRemuneracao_Recebida_02 = pyautogui.position(232, 257)
    pyautogui.moveTo(InfoRemuneracao_Recebida_02)
    pyautogui.click()
    pyautogui.write(str(Remuneracao_Recebida_02))
    time.sleep(2)

    InfoValor_Contribuicao_Descontada_02 = pyautogui.position(674, 255)
    pyautogui.moveTo(InfoValor_Contribuicao_Descontada_02)
    pyautogui.click()
    pyautogui.write(str(Valor_Contribuicao_Descontada_02))

    time.sleep(2)
    ClicaremOkParaInserirRegistro = pyautogui.position(748, 124)
    pyautogui.moveTo(ClicaremOkParaInserirRegistro)
    pyautogui.click()

    ConfirmarValoresInseridos = pyautogui.position(834, 442)
    pyautogui.moveTo(ConfirmarValoresInseridos)
    pyautogui.click()

    fecharApplication = pyautogui.position(70, 58)
    pyautogui.moveTo(fecharApplication)
    pyautogui.click()










