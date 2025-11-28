import pandas as pd
from unidecode import unidecode
import numpy as np
import time
import pyodbc
import warnings
from viabilipy.NFV import extracao_teia, tratamento_gaia
from PIL import Image
from sqlalchemy import create_engine
import customtkinter as ctk

from selenium import webdriver
from selenium.webdriver.chrome.service import Service 
from selenium.webdriver.support.ui import Select, WebDriverWait
import chromedriver_autoinstaller
service = Service(chromedriver_autoinstaller.install())

pd.set_option('display.max_columns', None)
pd.set_option('future.no_silent_downcasting', True)
pd.set_option('display.max_columns', None)
warnings.filterwarnings('ignore')

engine = create_engine('mysql+pymysql://viabilidade:senha_segura123#@10.0.15.243:3306/desenvolvimento_viabilidade')



global fechamento_teia

try:
    # conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ=C:\Users\F257064\Documents\MODELO NFV\NFV LIMPO\Base_Terceiros.accdb;PWD=maderomcbk')
    # conn = pyodbc.connect(conn_str)
    # provedor_ethernet = pd.read_sql('SELECT * FROM provedorxlocalidade_ethernet', conn)
    valores_ethernet_ = pd.read_sql('SELECT * FROM valores_terceiros_eth_filtered', engine)
    status = pd.read_sql('SELECT * FROM status', engine)
    # conn.close()
    status = status.fillna('OK')
except:
    # provedor_ethernet = pd.read_excel('./arquivos/provedorxlocalidade_ethernet.xlsx')
    # valores_ethernet = pd.read_excel('./arquivos/valores_ethernet.xlsx')
    # status = pd.read_excel('./arquivos/status.xlsx')
    pass

capacity = pd.read_excel('./arquivos/capacity.xlsx')
municipio_estacao = pd.read_excel('./arquivos/bb_municipio_estacao.xlsx')
estacoes_uf = pd.read_excel('./arquivos/estacoes_entregas.xlsx')
cidades = pd.read_excel('./arquivos/cidades_bbip.xlsx')
custos_acesso = pd.read_excel('./arquivos/custos_nfv.xlsx')
tecnologia_facilidade = pd.read_excel('./arquivos/id_tecnologia_facilidades.xlsx')
id_provedores = pd.read_excel('./arquivos/id_provedores.xlsx')
capacity_fixa = pd.read_excel('./arquivos/capacity_fixa.xlsx')
tecnologia_capacity = pd.read_excel('./arquivos/tecnologia_capacity.xlsx')
capacity_funil = pd.read_excel('./arquivos/capacity_funil.xlsx')
estacoes_newteia = pd.read_excel('./arquivos/lista_estacoes_newteia.xlsx')
municipio_localidade = pd.read_excel('./arquivos/municipio_localidade.xlsx')
consulta_banda = pd.read_excel('./arquivos/consulta_banda.xlsx')


fechamento_teia = pd.DataFrame(columns=['sequencial','latitude','longitude','uf','cnl','facilidade','id_facilidade','provedor','id_provedor','entrega',
                                        'abordado','custo_de_acesso_proprio','instalacao_terceiros','mensalidade_terceiros','tipo_terceiros','id_da_sev','prazo',
                                        'bb_ip','hp_bsod','codigo_spe','sinalizacao_sip','protocolo_gaia','obs','justificativa','ID Justificativa','status','tecnologia'])

# facilidades = pd.read_excel('./arquivos/facilidades.xlsx').sort_values('PRIORIDADE',ascending=True).fillna('')
facilidades = pd.read_excel('./arquivos/facilidades_tecnologia_prioridade.xlsx').sort_values('PRIORIDADE',ascending=True).fillna('')



def arquivo_teia():
    global teia, removidas, sevs_tratar
    arquivo_entrada_sevs_teia = ctk.filedialog.askopenfilename(title='Abrir arquivo de extração do TEIA')
    teia = pd.read_csv(arquivo_entrada_sevs_teia, sep=';')
    tratamento = extracao_teia.extracaoTeia(teia)
    removidas =  tratamento.tratar_modelo_gaia(removed_sevs='S')
    sevs_tratar = teia.drop(index=teia[teia.SEV.isin(removidas)].index).reset_index(drop=True)

    sevs_tratar['PONTA_A'] = sevs_tratar['PONTA_A'].fillna('')

    for index, value in sevs_tratar.iterrows():
        aux = municipio_localidade[municipio_localidade.SIGLA_LOC == value.CNL]
        if len(aux) > 0:
            sevs_tratar.at[index,'CNL'] = aux.CNL.values[0]
        else:
            print(f'{value.SEV} nao encontrou CNL')

    button_browse = ctk.CTkButton(janela,text="Arquivo TEIA",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=arquivo_teia)
    button_browse.place(x=10,y=100)


def inclui_restricao():

    if check_restricao.get() == 'S':

        button_restricao.place(x=245, y=200)

    else:
        button_restricao.place(x=1000, y=1000)

def selecionar_resumosoe():
    global resumosoe
    arquivo_resumosoe = ctk.filedialog.askopenfilename(title='Abrir arquivo resumoSoE')
    resumosoe = tratamento_gaia.tratamentoResumosoe(f"{arquivo_resumosoe}").trata_resumosoe()
    
    for index, value in resumosoe.iterrows():
        if value.TERCEIROS_ETH == 'Viável':
            if 'MOBWIRE' in value.TERCEIROS_ETH_INFORMACAO:
                aux_res = value.TERCEIROS_ETH_INFORMACAO
                aux_res = aux_res.replace('/ MOBWIRE /','').replace(' MOBWIRE /','').replace(' MOBWIRE','')
                if aux_res.split(':')[-1] == '':
                    resumosoe.at[index,'TERCEIROS_ETH'] = 'Inviável'
                    resumosoe.at[index,'TERCEIROS_ETH_INFORMACAO'] = ''
                else:
                    resumosoe.at[index,'TERCEIROS_ETH_INFORMACAO'] = aux_res

    resumosoe.to_excel('arquivo_resumosoe.xlsx',index=False)
    button_resumosoe = ctk.CTkButton(janela,text="ResumoSoE",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_resumosoe)
    button_resumosoe.place(x=10,y=200)

def selecionar_nuvens():
    global nuvens
    arquivo_nuvens = ctk.filedialog.askopenfilename(title='Abrir arquivo nuvens')
    nuvens = tratamento_gaia.tratamentoNuvens(f"{arquivo_nuvens}").trata_nuvens()
    nuvens.to_excel('arquivo_nuvens.xlsx',index=False)
    button_nuvens = ctk.CTkButton(janela,text="Nuvens",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_nuvens)
    button_nuvens.place(x=100,y=200)

def selecionar_resultado():
    global resultado
    arquivo_resultado = ctk.filedialog.askopenfilename(title='Abrir arquivo resultado')
    resultado = tratamento_gaia.tratamentoResultado(f"{arquivo_resultado}").trata_resultado()
    resultado.to_excel('arquivo_resultado.xlsx',index=False)
    button_resultado = ctk.CTkButton(janela,text="Resultado",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_resultado)
    button_resultado.place(x=165,y=200)

def selecionar_restricao():
    global restricao
    arquivo_restricao = ctk.filedialog.askopenfilename(title='Abrir arquivo restricão')
    restricao = tratamento_gaia.tratamentoRestricao(f"{arquivo_restricao}").trata_restricao()
    restricao.to_excel('arquivo_restricao.xlsx',index=False)
    button_restricao = ctk.CTkButton(janela,text="Restrição",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=selecionar_restricao)
    button_restricao.place(x=245, y=200)

def tratar_sevs():
    global nuvens, valores_ethernet_, sevs_tratar,status
    sevs_tratar = teia.drop(index=teia[teia.SEV.isin(removidas)].index).reset_index(drop=True)


    acum_nuvens = pd.DataFrame(columns=nuvens.columns)

    for i, v in nuvens.iterrows():
        try:
            tec = v.TECNOLOGIA.split(' / ')
        except:
            pass
        
        for value in tec:
            aux = nuvens.loc[i].to_dict()
            aux.update(TECNOLOGIA=value)
            acum_nuvens.loc[len(acum_nuvens)] = aux

    nuvens = acum_nuvens.drop_duplicates()


    for i, v in nuvens.iterrows():
        aux_capcity = tecnologia_capacity[tecnologia_capacity.TECNOLOGIA == v.TECNOLOGIA]
        if len(aux_capcity) > 0:
            if aux_capcity.CAPACITY.values[0] not in ['SIGLA_ESTACAO_CLARO','ESTACAO_ENTREGA']:
                if v.TECNOLOGIA == 'VIRTUA':
                    if sevs_tratar[sevs_tratar.SEV == v.SEV].SERVICO.values[0] == 'VPE - VIP BSOD LIGHT':
                        nuvens.at[i,'CAPACITY_NUVEM'] = 1000
                    else:
                        nuvens.at[i,'CAPACITY_NUVEM'] = 0
                else:
                    nuvens.at[i,'CAPACITY_NUVEM'] = int(aux_capcity.CAPACITY.values[0])
            else:
                if aux_capcity.CAPACITY.values[0] == 'ESTACAO_ENTREGA':
                # CONSULTAR UTILIZANDO AS COLUNAS DA TABELA TECNOLOGIA_CAPACITY, DO FUNIL E TABELA DE CAPACITYS.
                    if v.TECNOLOGIA == 'FO EDD NET':
                        if len(capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)]) > 0:
                            nuvens.at[i,'CAPACITY_NUVEM'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0]
                            nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CENTRO_ROTEAMENTO.values[0]
                    elif v.TECNOLOGIA == 'GPON NET':
                        if len(capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)]) > 0:
                            if capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0] > 200:
                                nuvens.at[i,'CAPACITY_NUVEM'] = 200
                                nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CENTRO_ROTEAMENTO.values[0]
                            else:
                                nuvens.at[i,'CAPACITY_NUVEM'] = capacity[(capacity.NUVEM == 'FO EDD NET') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0]
                    elif v.TECNOLOGIA == 'SDH':
                        if len(capacity[(capacity.NUVEM == 'FO SDH') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)]) > 0:
                            nuvens.at[i,'CAPACITY_NUVEM'] = capacity[(capacity.NUVEM == 'FO SDH') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CAPACITY_MB.values[0]
                            nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity[(capacity.NUVEM == 'FO SDH') & (capacity.ESTACAO_ENTREGA == v.ESTACAO_ENTREGA)].CENTRO_ROTEAMENTO.values[0]
                    else:
                        if len(capacity_fixa[(capacity_fixa.TECNOLOGIA == v.TECNOLOGIA) & (capacity_fixa.SIGLA_EMBRATEL == v.ESTACAO_ENTREGA)]) > 0:
                            nuvens.at[i,'CAPACITY_NUVEM'] = capacity_fixa[(capacity_fixa.TECNOLOGIA == v.TECNOLOGIA) & (capacity_fixa.SIGLA_EMBRATEL == v.ESTACAO_ENTREGA)].TOTAL.values[0]
                            nuvens.at[i,'CENTRO_ROTEAMENTO'] = capacity_fixa[(capacity_fixa.TECNOLOGIA == v.TECNOLOGIA) & (capacity_fixa.SIGLA_EMBRATEL == v.ESTACAO_ENTREGA)].ESTACAO_BB.values[0]
                elif aux_capcity.CAPACITY.values[0] == 'SIGLA_ESTACAO_CLARO':
                    aux_funil = capacity_funil[capacity_funil.SITE == v.SIGLA_ESTACAO_CLARO]
                    if len(aux_funil) > 0:
                        nuvens.at[i,'CENTRO_ROTEAMENTO'] = v.ESTACAO_ENTREGA
                        if nuvens.at[i,'REDE'] == 'CORTE CAPACIDADE-BANDA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        elif 'CORTE PLANEJAMENTO REGIONAL' in nuvens.at[i,'REDE']:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'ESGOTADA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'CONCLUIDA':
                            if nuvens.at[i,'MEIO_TRANSMISSAO'] == 'REDE OPTICA':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 100
                            elif nuvens.at[i,'MEIO_TRANSMISSAO'] == 'ENLACE DE RADIO':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        nuvens.at[i,'CAPACITY_NUVEM'] = aux_funil.BANDA.values[0]
                        if (v.TECNOLOGIA == 'GPON MOVEL') & (nuvens.at[i,'CAPACITY_NUVEM'] > 200):
                            nuvens.at[i,'CAPACITY_NUVEM'] = 200

                    else:
                        nuvens.at[i,'CENTRO_ROTEAMENTO'] = v.ESTACAO_ENTREGA
                        if nuvens.at[i,'REDE'] == 'CORTE CAPACIDADE-BANDA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        elif 'CORTE PLANEJAMENTO REGIONAL' in nuvens.at[i,'REDE']:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'ESGOTADA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'CONCLUIDA':
                            if nuvens.at[i,'MEIO_TRANSMISSAO'] == 'REDE OPTICA':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 100
                            elif nuvens.at[i,'MEIO_TRANSMISSAO'] == 'ENLACE DE RADIO':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        else:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        
                    if nuvens.at[i,'CAPACITY_NUVEM'] == 0:
                        nuvens.at[i,'CENTRO_ROTEAMENTO'] = v.ESTACAO_ENTREGA
                        if nuvens.at[i,'REDE'] == 'CORTE CAPACIDADE-BANDA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 10
                        elif 'CORTE PLANEJAMENTO REGIONAL' in nuvens.at[i,'REDE']:
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'ESGOTADA':
                            nuvens.at[i,'CAPACITY_NUVEM'] = 0
                        elif nuvens.at[i,'SITUACAO'] == 'CONCLUIDA':
                            if nuvens.at[i,'MEIO_TRANSMISSAO'] == 'REDE OPTICA':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 100
                            elif nuvens.at[i,'MEIO_TRANSMISSAO'] == 'ENLACE DE RADIO':
                                nuvens.at[i,'CAPACITY_NUVEM'] = 10
                
    nuvens['CAPACITY_NUVEM'] = nuvens['CAPACITY_NUVEM'].fillna(0)

    resumosoe['BANDA_ABORDADO'] = 0
    resumosoe['FACILIDADE_ABORDADO'] = resumosoe['FACILIDADE_ABORDADO'].fillna('')

    for index, value in resumosoe.iterrows():
        aux_capacity = capacity[(capacity.NUVEM_ABORDADO == value.FACILIDADE_ABORDADO) & (capacity.ESTACAO_ENTREGA == value.ESTACAO_ENTREGA_ABORDADO)]
        aux_capacity_fixa = capacity_fixa[(capacity_fixa.FACILIDADE == value.FACILIDADE_ABORDADO) & (capacity_fixa.SIGLA_EMBRATEL == value.ESTACAO_ENTREGA_ABORDADO)]
        if value.FACILIDADE_ABORDADO == 'FOetherNET':
            resumosoe.at[index,'BANDA_ABORDADO'] = 1000
        elif value.FACILIDADE_ABORDADO == 'FO SDH':
            resumosoe.at[index,'BANDA_ABORDADO'] = 0
        elif len(aux_capacity_fixa) > 0:
            resumosoe.at[index,'BANDA_ABORDADO'] = aux_capacity_fixa.TOTAL.values[0]
        elif len(aux_capacity) > 0:
            resumosoe.at[index,'BANDA_ABORDADO'] = aux_capacity.CAPACITY_MB.values[0]


    for index, value in sevs_tratar.iterrows():
        if value.VELOCIDADE[-4] == 'M':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4])
        elif value.VELOCIDADE[-4] == 'G':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) * 1000
        elif value.VELOCIDADE[-4] == 'K':
            sevs_tratar.at[index,'VEL'] = int(value.VELOCIDADE[:-4]) / 1000

    def convert_velocidade(vel):
        try:
            if vel[-1] == 'M':
                return int(vel[:-1])
            elif vel[-1] == 'G':
                return int(vel[:-1]) * 1000
            elif vel[-1] == 'K':
                return int(vel[:-1]) / 1000
            else:
                return 0
        except:
            return 0

    valores_ethernet_['VEL'] = valores_ethernet_['VELOCIDADE'].apply(convert_velocidade)
    
    sevs_tratar['TRATADO'] = ''

    ## ESCOLHA DO ACESSO ABORDADO
    for index, value in sevs_tratar.iterrows():
        aux_resumosoe = resumosoe[resumosoe.SEV == value.SEV]
        if len(aux_resumosoe) > 0:
            if (aux_resumosoe.FACILIDADE_ABORDADO.values[0] == 'FO GPON ETH') & (value.VEL <= 202):
                if (aux_resumosoe.BANDA_ABORDADO.values[0] >= value.VEL):
                    sevs_tratar.at[index,'FACILIDADE_ABORDADO'] = aux_resumosoe.FACILIDADE_ABORDADO.values[0]
                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = 'FO_GPON_ETH'
                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe.ESTACAO_ENTREGA_ABORDADO.values[0]
                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'GPON FIXA'
                    sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_resumosoe.ESTACAO_ENTREGA_ABORDADO.values[0]
                    sevs_tratar.at[index,'TRATADO'] = 'X'
            elif (aux_resumosoe.FACILIDADE_ABORDADO.values[0] == 'FO SDH') & (value.VEL <= 100):
                if (aux_resumosoe.BANDA_ABORDADO.values[0] >= value.VEL):
                    sevs_tratar.at[index,'FACILIDADE_ABORDADO'] = aux_resumosoe.FACILIDADE_ABORDADO.values[0]
                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = 'FO_SDH'
                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe.ESTACAO_ENTREGA_ABORDADO.values[0]
                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'SDH'
                    sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_resumosoe.ESTACAO_ENTREGA_ABORDADO.values[0]
                    sevs_tratar.at[index,'TRATADO'] = 'X'
            elif (aux_resumosoe.FACILIDADE_ABORDADO.values[0] != 'FO GPON ETH') & (aux_resumosoe.FACILIDADE_ABORDADO.values[0] != 'FO SDH'):
                if (aux_resumosoe.BANDA_ABORDADO.values[0] >= value.VEL):
                    sevs_tratar.at[index,'FACILIDADE_ABORDADO'] = aux_resumosoe.FACILIDADE_ABORDADO.values[0]
                    if aux_resumosoe.FACILIDADE_ABORDADO.values[0] == 'FO EDD ETH':
                        sevs_tratar.at[index,'FACILIDADE_ACESSO'] = 'FO_EDD_ETH'
                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'FO EDD FIXA'
                    elif aux_resumosoe.FACILIDADE_ABORDADO.values[0] == 'FOetherNET':
                        sevs_tratar.at[index,'FACILIDADE_ACESSO'] = 'FOETHERNET'
                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'FO EDD NET'
                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe.ESTACAO_ENTREGA_ABORDADO.values[0]
                    sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_resumosoe.ESTACAO_ENTREGA_ABORDADO.values[0]
                    sevs_tratar.at[index,'TRATADO'] = 'X'



    for index, value in sevs_tratar[sevs_tratar.TRATADO == 'X'].iterrows():
        if value.SERVICO in ['LAN - LAN EPL','LAN - LAN EPL MEF']:
            sevs_tratar.at[index,'TRATADO'] = ''
            sevs_tratar.at[index,'FACILIDADE_ABORDADO'] = ''
        elif value.SERVICO in ['EIN - TRANSMUX CIRCUITO', 'EIN - TRANSMUX REDE']:
            if value.FACILIDADE_ABORDADO != 'FO SDH':
                sevs_tratar.at[index,'TRATADO'] = ''
                sevs_tratar.at[index,'FACILIDADE_ABORDADO'] = ''

    nuvens.NOME_NUVEM = nuvens.NOME_NUVEM.replace(' ','')



    for i, v in nuvens.iterrows():
        if v.NOME_NUVEM == '':
            nuvens.at[i,'NOME_NUVEM'] = v.ESTACAO_ENTREGA

    for index,value in resumosoe.iterrows():
        for i,v in facilidades.iterrows():
            if 'NUVEM: /' in value[f'{v.FACILIDADE}_INFORMACAO']:
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = value[f'{v.FACILIDADE}_INFORMACAO'].replace('NUVEM: /','')
            elif 'NUVEM:  / ' in value[f'{v.FACILIDADE}_INFORMACAO']:
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = value[f'{v.FACILIDADE}_INFORMACAO'].replace('NUVEM:  / ','')
            elif 'NUVEM: ' in value[f'{v.FACILIDADE}_INFORMACAO']:
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = value[f'{v.FACILIDADE}_INFORMACAO'].replace('NUVEM: ','')
            
            if resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] == '':
                resumosoe.at[index,f'{v.FACILIDADE}_INFORMACAO'] = resumosoe.at[index,f'{v.FACILIDADE}_ESTACAO_ENTREGA']


    for index, value in sevs_tratar.iterrows():
        aux_resumosoe = resumosoe[resumosoe.SEV == value.SEV]
        if len(aux_resumosoe) >0:
            for i,v in facilidades.iterrows():
                if sevs_tratar.at[index,'TRATADO'] != 'X':
                    if aux_resumosoe[f'{v.FACILIDADE}'].values[0] == 'Viável':
                        if v.VERIFICA_CAPACITY == 'N':
                            if v.FACILIDADE == 'HFC_BSOD':
                                if 'HP GED' in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]:
                                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                    sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]
                                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = ''
                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'VIRTUA HFC'
                                    else:
                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'HFC BSOD'
                                    sevs_tratar.at[index,'TRATADO'] = 'X'
                                    break
                                else:
                                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                    sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0]
                                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].replace('ESTAÇÃO ENTRONCAMENTO:','')
                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'VIRTUA HFC'
                                    else:
                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'HFC BSOD'
                                    sevs_tratar.at[index,'TRATADO'] = 'X'
                                    break
                            elif v.FACILIDADE == '4G':
                                sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0]
                                sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = ''
                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'LTE (4G)'
                                sevs_tratar.at[index,'TRATADO'] = 'X'
                                break
                            elif v.FACILIDADE == 'FO_GPON_RESID_ETH_PRE_VIAVEL':
                                for tec in v.TECNOLOGIA.split('/'):
                                    if sevs_tratar.at[index,'TRATADO'] != 'X':
                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                            if len(aux_nuvem) > 0:
                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                                    sevs_tratar.at[index,'NUVEM_ACESSO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = tec
                                                    sevs_tratar.at[index,'TRATADO'] = 'X'
                                                    break
                            elif v.FACILIDADE == 'FO_XGSPON_RESID_ETH':
                                for tec in v.TECNOLOGIA.split('/'):
                                    if sevs_tratar.at[index,'TRATADO'] != 'X':
                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                            if len(aux_nuvem) > 0:
                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                                    sevs_tratar.at[index,'NUVEM_ACESSO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = tec
                                                    sevs_tratar.at[index,'TRATADO'] = 'X'
                                                    break
                            elif v.FACILIDADE == 'FO_GPON_RESID_ETH':
                                for tec in v.TECNOLOGIA.split('/'):
                                    if sevs_tratar.at[index,'TRATADO'] != 'X':
                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem) & (nuvens.TECNOLOGIA == tec)]
                                            if len(aux_nuvem) > 0:
                                                if value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]:
                                                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                                    sevs_tratar.at[index,'NUVEM_ACESSO'] = f'FABRICANTE {aux_nuvem.FABRICANTE_OLT.values[0]} CONCENTRADOR OLT {aux_nuvem.CONCENTRADOR_OLT.values[0]}'
                                                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe[f'{v.FACILIDADE}_ESTACAO_ENTREGA'].values[0] 
                                                    if value.SERVICO == 'VPE - VIP BSOD LIGHT':
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'VIRTUA GPON'
                                                        sevs_tratar.at[index,'TRATADO'] = 'X'
                                                        break
                                                    else:
                                                        if tec != 'VIRTUA GPON':
                                                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = tec
                                                            sevs_tratar.at[index,'TRATADO'] = 'X'
                                                            break
                                                    
                                                   

                            elif (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                                sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                                sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = 'RJO AM'
                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = v.TECNOLOGIA
                                sevs_tratar.at[index,'TRATADO'] = 'X'
                                break
                        else:
                            if v.FACILIDADE == 'TERCEIROS_ETH':
                                
                                sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0].split('PROPRIETÁRIO ')[-1]
                                sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_resumosoe.TERCEIROS_ETH_ESTACAO_ENTREGA.values[0].split(' / ')[-1]
                                sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = 'TERCEIRO ETH'
                                sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_resumosoe.TERCEIROS_ETH_ESTACAO_ENTREGA.values[0].split(' / ')[-1]
                                sevs_tratar.at[index,'TRATADO'] = 'X'
                                
                            else:
                                for tec in v.TECNOLOGIA.split('/'):
                                    if sevs_tratar.at[index,'TRATADO'] != 'X':
                                        for nome_nuvem in aux_resumosoe[f'{v.FACILIDADE}_INFORMACAO'].values[0].split(' / '):
                                            aux_nuvem = nuvens[(nuvens.SEV == value.SEV) & (nuvens.NOME_NUVEM == nome_nuvem.replace(':','')) & (nuvens.TECNOLOGIA == tec)]
                                            if len(aux_nuvem) > 0:
                                                if (value.SERVICO == 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):
                                                    if 'EPL MEF - NOK' not in aux_nuvem.OBSERVACAO.values[0]:
                                                        sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                                        if tec in v.CONSULTA_FUNIL.split('/'):
                                                            sevs_tratar.at[index,'NUVEM_ACESSO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                        else:
                                                            sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_nuvem.NOME_NUVEM.values[0]
                                                        sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_nuvem.ESTACAO_ENTREGA.values[0]
                                                        sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                        sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = tec
                                                        sevs_tratar.at[index,'TRATADO'] = 'X'
                                                        break

                                                if (value.SERVICO != 'LAN - LAN EPL MEF') & (value.VEL <= aux_nuvem.CAPACITY_NUVEM.values[0]):
                                                    sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                                                    if tec in v.CONSULTA_FUNIL.split('/'):
                                                        sevs_tratar.at[index,'NUVEM_ACESSO'] = f'NUVEM {aux_nuvem.NOME_NUVEM.values[0]} SITE {aux_nuvem.SIGLA_ESTACAO_CLARO.values[0]}'
                                                    else:
                                                        sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_nuvem.NOME_NUVEM.values[0]
                                                    sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = aux_nuvem.ESTACAO_ENTREGA.values[0]
                                                    sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_nuvem.CENTRO_ROTEAMENTO.values[0]
                                                    sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = tec
                                                    sevs_tratar.at[index,'TRATADO'] = 'X'
                                                    break

                            
                    if aux_resumosoe[f'{v.FACILIDADE}'].values[0] == '%Disponibilidade não atende ao desejado':
                        if (v.FACILIDADE == 'SATELITE_BANDA_KA') | (v.FACILIDADE == 'SATELITE_BANDA_KU'):
                            sevs_tratar.at[index,'FACILIDADE_ACESSO'] = v.FACILIDADE
                            sevs_tratar.at[index,'NUVEM_ACESSO'] = aux_resumosoe.TERCEIROS_ETH_INFORMACAO.values[0]
                            sevs_tratar.at[index,'ESTACAO_ENTREGA_ACESSO'] = 'RJO AM'
                            sevs_tratar.at[index,'TECNOLOGIA_ACESSO'] = v.TECNOLOGIA
                            sevs_tratar.at[index,'TRATADO'] = 'X'

    sevs_tratar.FACILIDADE_ACESSO = sevs_tratar.FACILIDADE_ACESSO.fillna('')
    sevs_tratar.CENTRO_ROTEAMENTO = sevs_tratar.CENTRO_ROTEAMENTO.fillna('')

    valores_ethernet_ = valores_ethernet_.fillna(0)
    print('PRECIFICANDO TERCEIROS')
    for index, value in sevs_tratar[(sevs_tratar.FACILIDADE_ACESSO == 'TERCEIROS_ETH')].iterrows():
        print(round((index/len(sevs_tratar)*100),2),end="\r")
        aux_provedores = valores_ethernet_[(valores_ethernet_.SIGLA_MUNICIPIO == value.CNL) & (valores_ethernet_.UF == value.UF)]
        melhor_provedor = ''
        melhor_instal = 0
        melhor_mensal = 0
        custo_mensalizado = 0
        for p in value.NUVEM_ACESSO.split(' / '):
            aux_provedor = aux_provedores[(aux_provedores.PROVEDOR == p)]
            # aux_provedor = provedor_ethernet[(provedor_ethernet.PROVEDOR == p) & (provedor_ethernet.SIGLA_LOC == value.CNL) & (provedor_ethernet.UF == value.UF)]
            if len(aux_provedor) > 0:
                aux_status = status[(status.PROVEDOR == aux_provedor.PROVEDOR.values[0]) & (status.UF == aux_provedor.UF.values[0])]
                if len(aux_status) > 0:
                    if (aux_status.STATUS.values[0] != 'BLOQUEADO') & (aux_status.STATUS.values[0] != 'AEROPORTO'):
                        aux_valores = aux_provedor[(aux_provedor.VEL == value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                        
                        if len(aux_valores) > 0:
                            if melhor_provedor == '':
                                melhor_provedor = p
                                melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                            else:
                                if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                    melhor_provedor = p
                                    melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                    melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                    custo_mensalizado = (melhor_instal / 24) + melhor_mensal
                        else:
                            aux_valores = aux_provedor[(aux_provedor.VEL >= value.VEL) & (aux_provedor.PRAZO == '24 MESES')].sort_values(by='VEL', ascending=True)
                            if len(aux_valores) > 0:
                                if melhor_provedor == '':
                                    melhor_provedor = p
                                    melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                    melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                    custo_mensalizado = (melhor_instal / 24) + melhor_mensal 
                                else:
                                    if custo_mensalizado > ((aux_valores.TAXA_INSTALACAO.values[0] / 24) + aux_valores.CUSTO_MENSAL.values[0]):
                                        melhor_provedor = p
                                        melhor_instal = aux_valores.TAXA_INSTALACAO.values[0]
                                        melhor_mensal = aux_valores.CUSTO_MENSAL.values[0]
                                        custo_mensalizado = (melhor_instal / 24) + melhor_mensal
                    
        if melhor_provedor != '':
            sevs_tratar.at[index,'PROVEDOR_ESCOLHA'] = melhor_provedor
            sevs_tratar.at[index,'PROVEDOR_INSTALACAO'] = melhor_instal
            sevs_tratar.at[index,'PROVEDOR_MENSALIDADE'] = melhor_mensal
        else:
            sevs_tratar.loc[index, ['TRATADO','FACILIDADE_ACESSO','NUVEM_ACESSO','ESTACAO_ENTREGA_ACESSO','TECNOLOGIA_ACESSO','CENTRO_ROTEAMENTO']] = ''

    sevs_tratar = sevs_tratar.fillna('')

    for index, value in sevs_tratar.iterrows():
        if value.SERVICO != 'VPE - VIP BSOD LIGHT':
            if value.TECNOLOGIA_ACESSO != '':
                if consulta_banda[consulta_banda.TECNOLOGIA == value.TECNOLOGIA_ACESSO].VEL_BBIP.values[0] < value.VEL:
                    if value.FACILIDADE_ACESSO in facilidades[facilidades.VERIFICA_CAPACITY == 'S'].FACILIDADE.tolist():

                        aux_municipio_estacao = municipio_estacao[municipio_estacao.ESTACAO == value.CENTRO_ROTEAMENTO]
                        if len(aux_municipio_estacao) > 0:
                            sevs_tratar.at[index,'BBIP'] = f'{aux_municipio_estacao.MUNICIPIO.values[0]}|{aux_municipio_estacao.ESTACAO.values[0]}'
                        else:
                            aux_estacao_uf= estacoes_uf[estacoes_uf.UF == value.UF]
                            sevs_tratar.at[index,'BBIP'] = f'{aux_estacao_uf.MUNICIPIO.values[0]}|{aux_estacao_uf.ESTACAO.values[0]}'
        
    try:
        sevs_tratar.BBIP = sevs_tratar.BBIP.fillna('')
    except:
        sevs_tratar['BBIP'] = ''

    for index, value in sevs_tratar.iterrows():
        if len(resultado[resultado.SEV == value.SEV]) > 0:
            sevs_tratar.at[index,'PROTOCOLO_ATENDIMENTO'] = resultado[resultado.SEV == value.SEV].PROTOCOLO.values[0]

    #################### REMOVE AS SEVS COM RESTRICAO ##########################
    if check_restricao.get() == 'S':
        for index, value in sevs_tratar[sevs_tratar.TRATADO == 'X'].iterrows():
            aux_restricao = restricao[restricao.SEV == value.SEV]
            if len(aux_restricao) > 0:

                if len(aux_restricao) > 1:
                    if 'TOTAL' in aux_restricao.TIPO_DE_IMPACTO.tolist():
                        sevs_tratar.at[index,'RESTRICAO'] = aux_restricao.OBSERVACAO.values[0]
                        sevs_tratar.at[index,'TRATADO'] = ''

                else:
                    if aux_restricao.TIPO_DE_IMPACTO.values[0] == 'TOTAL':
                        sevs_tratar.at[index,'RESTRICAO'] = aux_restricao.OBSERVACAO.values[0]
                        sevs_tratar.at[index,'TRATADO'] = ''

    button_fase_1 = ctk.CTkButton(janela,text="Fase 1",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=tratar_sevs)
    button_fase_1.place(x=10,y=250)


def rodar_bbip():
    navegador = webdriver.Chrome(service=service)
    navegador.implicitly_wait(5)
    navegador.get('http://10.100.1.30/admredes/admredes/RPVB_BLD_Cadastrar_2.asp')
    time.sleep(25)

    for index, value in sevs_tratar[(sevs_tratar.BBIP != '') & (sevs_tratar.TRATADO == 'X')].iterrows():
        try:
            if consulta_banda[consulta_banda.TECNOLOGIA == value.TECNOLOGIA_ACESSO].VEL_BBIP.values[0] < value.VEL:
                navegador.get('http://10.100.1.30/admredes/admredes/RPVB_BLD_Cadastrar_2.asp')
                time.sleep(1)
                if (value.BBIP != None) & (value.BBIP != 'BBIP') & ("ID" not in str(value.BBIP)):
                    navegador.find_element('name','cliente').send_keys(value.CLIENTE)
                    navegador.find_element('name','sev').send_keys(value.SEV)
                    municipio, estacao = value.BBIP.split('|')
                    municipio = unidecode(municipio).upper()
                    id_municipio = str(cidades[cidades.CIDADE == municipio].ID.values[0])
                    Select(navegador.find_element('id', 'combo1')).select_by_value(id_municipio)
                    Select(navegador.find_element('id', 'combo2')).select_by_visible_text(estacao)
                    if 'E-ACCESS' in value.SERVICO or 'EPL' in value.SERVICO:
                        Select(navegador.find_element('name', 'servico')).select_by_value('EPL')
                    else:
                        Select(navegador.find_element('name', 'servico')).select_by_value('Internet')
                    navegador.find_element('name', 'velocidade').send_keys(int(value.VEL))
                    navegador.find_element('xpath', '/html/body/div[6]/font/form/div/table[1]/tbody/tr[2]/td/table[5]/tbody/tr/td/input').click()
                    time.sleep(4)
                    id_bbip = navegador.find_element('xpath','/html/body/div[6]/font/form/div/table[1]/tbody/tr[2]/td/table[1]/tbody/tr[1]/td').text[:-4]
                    status_bbip = navegador.find_element('xpath','/html/body/div[6]/font/form/div/table[1]/tbody/tr[2]/td/table[4]/tbody/tr/td[2]/font/b').text
                    sevs_tratar.at[index,'BBIP'] = f'{id_bbip} / {status_bbip}'
        except:
            continue

    navegador.close()

    for index, value in sevs_tratar[sevs_tratar.BBIP != ''].iterrows():
        if 'Indeferido' in value.BBIP:
            sevs_tratar.at[index,'TRATADO'] = ''

    button_bbip = ctk.CTkButton(janela,text="Rodar BBIP",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=rodar_bbip)
    button_bbip.place(x=70,y=250)

def gerar_arquivos_finais():
    global sevs_tratar, fechamento_teia
    # sevs_tratar.to_excel('validar_sevs_finais.xlsx',index=False)
    for index, value in sevs_tratar[(sevs_tratar.TRATADO == 'X') & (sevs_tratar.FACILIDADE_ACESSO != 'TERCEIROS_ETH')].iterrows():
        aux_custo = custos_acesso[custos_acesso.FACILIDADE == value.FACILIDADE_ACESSO]
        
        try:
            if value.FACILIDADE_ABORDADO != '':
                sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo.ACESSO_ABORDADO.values[0]
            else:
                if value.VEL <= 202:
                    sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo.NOVO_ACESSO_200M.values[0]
                elif value.VEL <= 502:
                    sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo.NOVO_ACESSO_500M.values[0]
                elif value.VEL <= 1000:
                    sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo.NOVO_ACESSO_1G.values[0]
        except:
            if value.VEL <= 202:
                sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo.NOVO_ACESSO_200M.values[0]
            elif value.VEL <= 502:
                sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo.NOVO_ACESSO_500M.values[0]
            elif value.VEL <= 1000:
                sevs_tratar.at[index,'CUSTO_ACESSO_PROPRIO'] = aux_custo.NOVO_ACESSO_1G.values[0]

    for index, value in sevs_tratar[sevs_tratar.TRATADO == 'X'].iterrows():
        if value.FACILIDADE_ACESSO != 'TERCEIROS_ETH':
            aux_facilidade = tecnologia_facilidade[tecnologia_facilidade.FACILIDADE == value.FACILIDADE_ACESSO]
            if len(aux_facilidade) > 0:
                sevs_tratar.at[index,'FACILIDADE_FECHAMENTO'] = aux_facilidade.FACILIDADE_FECHAMENTO.values[0]
                sevs_tratar.at[index,'ID_FACILIDADE'] = aux_facilidade.ID.values[0]
                sevs_tratar.at[index,'EMPRESA_FACILIDADE'] = aux_facilidade.EMPRESA.values[0]
                sevs_tratar.at[index,'ID_EMPRESA_FACILIDADE'] = aux_facilidade.ID_EMPRESA.values[0]
        else:
            aux_facilidade = tecnologia_facilidade[tecnologia_facilidade.FACILIDADE == value.FACILIDADE_ACESSO]
            aux_empresas = id_provedores[id_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_ESCOLHA)]
            if len(aux_empresas) > 1:
                aux_empresas = id_provedores[id_provedores.PROVEDOR_TEIA.str.contains(f'{value.PROVEDOR_ESCOLHA} - {value.UF}')]
                if len(aux_empresas) > 0:
                    sevs_tratar.at[index,'FACILIDADE_FECHAMENTO'] = aux_facilidade.FACILIDADE_FECHAMENTO.values[0]
                    sevs_tratar.at[index,'ID_FACILIDADE'] = aux_facilidade.ID.values[0]
                    sevs_tratar.at[index,'EMPRESA_FACILIDADE'] = aux_empresas.PROVEDOR_TEIA.values[0]
                    sevs_tratar.at[index,'ID_EMPRESA_FACILIDADE'] = aux_empresas.ID.values[0]
                else:
                    aux_empresas = id_provedores[id_provedores.PROVEDOR_TEIA.str.contains(value.PROVEDOR_ESCOLHA)]
                    sevs_tratar.at[index,'FACILIDADE_FECHAMENTO'] = aux_facilidade.FACILIDADE_FECHAMENTO.values[0]
                    sevs_tratar.at[index,'ID_FACILIDADE'] = aux_facilidade.ID.values[0]
                    sevs_tratar.at[index,'EMPRESA_FACILIDADE'] = aux_empresas.PROVEDOR_TEIA.values[0]
                    sevs_tratar.at[index,'ID_EMPRESA_FACILIDADE'] = aux_empresas.ID.values[0]

            elif len(aux_empresas) == 1:
                sevs_tratar.at[index,'FACILIDADE_FECHAMENTO'] = aux_facilidade.FACILIDADE_FECHAMENTO.values[0]
                sevs_tratar.at[index,'ID_FACILIDADE'] = aux_facilidade.ID.values[0]
                sevs_tratar.at[index,'EMPRESA_FACILIDADE'] = aux_empresas.PROVEDOR_TEIA.values[0]
                sevs_tratar.at[index,'ID_EMPRESA_FACILIDADE'] = aux_empresas.ID.values[0]

    sevs_tratar = sevs_tratar.fillna('')

    for i,v in sevs_tratar.iterrows():
        if v.EMPRESA_FACILIDADE == '':
            sevs_tratar.at[i,'TRATADO'] = ''

        if v.CENTRO_ROTEAMENTO == '':
            sevs_tratar.at[i,'CENTRO_ROTEAMENTO'] = estacoes_uf[estacoes_uf.UF == v.UF].ESTACAO.values[0]

    for index, value in sevs_tratar[sevs_tratar.TRATADO == 'X'].iterrows():
        aux_estacao_teia = estacoes_newteia[estacoes_newteia.old_id == value.CENTRO_ROTEAMENTO]
        if len(aux_estacao_teia) > 0:
            sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_estacao_teia.estacao.values[0]

        else:
            aux_estacao_teia = estacoes_newteia[estacoes_newteia.estacao == value.CENTRO_ROTEAMENTO]
            if len(aux_estacao_teia) > 0:
                sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = aux_estacao_teia.estacao.values[0]
            else:
                
                sevs_tratar.at[index,'CENTRO_ROTEAMENTO'] = estacoes_uf[estacoes_uf.UF == value.UF].ESTACAO.values[0]

    for index, value in sevs_tratar[sevs_tratar.TRATADO == 'X'].iterrows():
        if len(fechamento_teia) == 0:
            fechamento_teia.at[0,'sequencial'] = value.SEV
            fechamento_teia.loc[len(fechamento_teia) - 1,['latitude','longitude','uf','cnl','facilidade','id_facilidade',
            'provedor','id_provedor','entrega','id_da_sev','prazo','codigo_spe','sinalizacao_sip',
            'protocolo_gaia','status']] = [str(value.LATITUDE).replace('.',','), str(value.LONGITUDE).replace('.',','), value.UF, value.CNL, value.FACILIDADE_FECHAMENTO, str(value.ID_FACILIDADE).split('.')[0],
            value.EMPRESA_FACILIDADE,str(value.ID_EMPRESA_FACILIDADE).split('.')[0], value.CENTRO_ROTEAMENTO, value.ID_ANALISE, '', '','',value.PROTOCOLO_ATENDIMENTO,'1']
            try:
                if value.FACILIDADE_ABORDADO != '':
                    fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'SIM'
                else:
                    fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'NAO'
            except:
                fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'NAO'
            if value.FACILIDADE_ACESSO != 'TERCEIROS_ETH':
                fechamento_teia.at[len(fechamento_teia) - 1,'custo_de_acesso_proprio'] = str(round(value.CUSTO_ACESSO_PROPRIO,2)).replace('.',',')
                fechamento_teia.at[len(fechamento_teia) - 1,'obs'] = f'FECHAMENTO NFV FASE 1/ ESTACAO ENTREGA {value.ESTACAO_ENTREGA_ACESSO}/{value.NUVEM_ACESSO}'
                fechamento_teia.at[len(fechamento_teia) - 1,'tecnologia'] = value.TECNOLOGIA_ACESSO

            else:
                fechamento_teia.at[len(fechamento_teia) - 1,'instalacao_terceiros'] = str(round(value.PROVEDOR_INSTALACAO,2)).replace('.',',')
                fechamento_teia.at[len(fechamento_teia) - 1,'mensalidade_terceiros'] = str(round(value.PROVEDOR_MENSALIDADE,2)).replace('.',',')
                fechamento_teia.at[len(fechamento_teia) - 1,'tipo_terceiros'] = 3
                fechamento_teia.at[len(fechamento_teia) - 1,'justificativa'] = 'FORA DE REDE'
                fechamento_teia.at[len(fechamento_teia) - 1,'ID Justificativa'] = 1
                fechamento_teia.at[len(fechamento_teia) - 1,'obs'] = f'FECHAMENTO NFV FASE 1/PRECO PADRAO/ ESTACAO ENTREGA {value.ESTACAO_ENTREGA_ACESSO}'
                fechamento_teia.at[len(fechamento_teia) - 1,'tecnologia'] = 'TERCEIRO ETH'

            if value.BBIP != '':
                fechamento_teia.at[len(fechamento_teia) - 1,'bb_ip'] = value.BBIP.split(' / ')[0].split(': ')[-1]
        
        else:
            fechamento_teia.at[len(fechamento_teia),'sequencial'] = value.SEV
            fechamento_teia.loc[len(fechamento_teia) - 1,['latitude','longitude','uf','cnl','facilidade','id_facilidade',
            'provedor','id_provedor','entrega','id_da_sev','prazo','codigo_spe','sinalizacao_sip',
            'protocolo_gaia','status']] = [str(value.LATITUDE).replace('.',','), str(value.LONGITUDE).replace('.',','), value.UF, value.CNL, value.FACILIDADE_FECHAMENTO, str(value.ID_FACILIDADE).split('.')[0],
            value.EMPRESA_FACILIDADE,str(value.ID_EMPRESA_FACILIDADE).split('.')[0], value.CENTRO_ROTEAMENTO, value.ID_ANALISE, '', '','',value.PROTOCOLO_ATENDIMENTO,'1']
            try:
                if value.FACILIDADE_ABORDADO != '':
                    fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'SIM'
                else:
                    fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'NAO'
            except:
                fechamento_teia.at[len(fechamento_teia) - 1,'abordado'] = 'NAO'
            if value.FACILIDADE_ACESSO != 'TERCEIROS_ETH':
                fechamento_teia.at[len(fechamento_teia) - 1,'custo_de_acesso_proprio'] = str(round(value.CUSTO_ACESSO_PROPRIO,2)).replace('.',',')
                fechamento_teia.at[len(fechamento_teia) - 1,'obs'] = f'FECHAMENTO NFV FASE 1/ ESTACAO ENTREGA {value.ESTACAO_ENTREGA_ACESSO}/{value.NUVEM_ACESSO}'
                fechamento_teia.at[len(fechamento_teia) - 1,'tecnologia'] = value.TECNOLOGIA_ACESSO

            else:
                fechamento_teia.at[len(fechamento_teia) - 1,'instalacao_terceiros'] = str(round(value.PROVEDOR_INSTALACAO,2)).replace('.',',')
                fechamento_teia.at[len(fechamento_teia) - 1,'mensalidade_terceiros'] = str(round(value.PROVEDOR_MENSALIDADE,2)).replace('.',',')
                fechamento_teia.at[len(fechamento_teia) - 1,'tipo_terceiros'] = 3
                fechamento_teia.at[len(fechamento_teia) - 1,'justificativa'] = 'FORA DE REDE'
                fechamento_teia.at[len(fechamento_teia) - 1,'ID Justificativa'] = 1
                fechamento_teia.at[len(fechamento_teia) - 1,'obs'] = f'FECHAMENTO NFV FASE 1/PRECO PADRAO/ ESTACAO ENTREGA {value.ESTACAO_ENTREGA_ACESSO}'
                fechamento_teia.at[len(fechamento_teia) - 1,'tecnologia'] = 'TERCEIRO ETH'

            if value.BBIP != '':
                fechamento_teia.at[len(fechamento_teia) - 1,'bb_ip'] = value.BBIP.split(' / ')[0].split(': ')[-1]

            if value.FACILIDADE_ACESSO == 'HFC_BSOD':
                if 'HP GED' in value.NUVEM_ACESSO:
                    fechamento_teia.at[len(fechamento_teia) - 1,'hp_bsod'] = value.NUVEM_ACESSO.split('HP GED ')[-1]

    fechamento_teia = fechamento_teia.fillna('')

    fechamento_teia.to_csv('fechamento_lote_nfv.csv',sep=';',index=False)
    sevs_tratar[sevs_tratar.TRATADO == ''].to_excel('fechamento_nfv_lote_manual.xlsx',index=False)

    button_finalizar = ctk.CTkButton(janela,text="Gerar arquivos finais",height=20,width=35,corner_radius=8,fg_color='green',hover_color='blue', command=gerar_arquivos_finais)
    button_finalizar.place(x=160,y=250)

janela = ctk.CTk()

janela.title("NFV")
janela.geometry("500x400")
janela.resizable(width=False, height=False)

ctk.CTkLabel(janela,text='NFV',font=('Arial',20)).place(x=220,y=20)

img = ctk.CTkImage(dark_image=Image.open('./claro.png'),light_image=Image.open('./claro.png'), size=(75,75))
ctk.CTkLabel(janela,text='',image=img).place(x=5,y=0)

titulo_teia= ctk.CTkLabel(janela,text='Selecione o arquivo de entrada (TEIA)')
titulo_teia.place(x=10,y=70)

button_browse = ctk.CTkButton(janela,text="Arquivo TEIA",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=arquivo_teia)
button_browse.place(x=10,y=100)


# button_gera_mapinfo = ctk.CTkButton(janela,text="Gerar Arquivo para mapinfo",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue')
# button_gera_mapinfo.place(x=10,y=200)

titulo_arquivo_gaia= ctk.CTkLabel(janela,text='Selecione os arquivo de retorno do GAIA')
titulo_arquivo_gaia.place(x=10,y=130)

button_resumosoe = ctk.CTkButton(janela,text="ResumoSoE",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_resumosoe)
button_resumosoe.place(x=10,y=200)

button_nuvens = ctk.CTkButton(janela,text="Nuvens",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_nuvens)
button_nuvens.place(x=100,y=200)

button_resultado = ctk.CTkButton(janela,text="Resultado",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_resultado)
button_resultado.place(x=165,y=200)

button_restricao = ctk.CTkButton(janela,text="Restrição",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=selecionar_restricao)
button_restricao.place(x=1000, y=1000)

check_restricao = ctk.StringVar(value='N')

checkbox = ctk.CTkCheckBox(janela, text="Possuí arquivo de restrição?",command=inclui_restricao,
                                    variable=check_restricao, onvalue="S", offvalue="N")
checkbox.place(x=10, y=160)

titulo_tratar_sev= ctk.CTkLabel(janela,text='Tratar SEVs')
titulo_tratar_sev.place(x=10,y=230)

button_fase_1 = ctk.CTkButton(janela,text="Fase 1",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=tratar_sevs)
button_fase_1.place(x=10,y=250)

button_bbip = ctk.CTkButton(janela,text="Rodar BBIP",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=rodar_bbip)
button_bbip.place(x=70,y=250)

button_finalizar = ctk.CTkButton(janela,text="Gerar arquivos finais",height=20,width=35,corner_radius=8,fg_color='grey',hover_color='blue', command=gerar_arquivos_finais)
button_finalizar.place(x=160,y=250)

janela.mainloop()