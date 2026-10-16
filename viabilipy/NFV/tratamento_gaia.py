import pandas as pd
import numpy as np
from unidecode import unidecode

facilidades = pd.read_excel('./arquivos/facilidades.xlsx').FACILIDADE.tolist()

class tratamentoResumosoe():

    def __init__(self,arq):
        self.__resumo_soe = pd.read_csv(arq, sep = r'\t')



    def trata_resumosoe(self):

        colunas_resumosoe = ['ID',
        'SEV',
        'GEOCODE_OK',
        'LATITUDE',
        'LONGITUDE',
        'DEGRAU_PARA_TERCEIROS',
        'RESTRICAO',
        '4G',
        '4G_INFORMACAO',
        '4G_ESTACAO_ENTREGA',
        '4G_TECNOLOGIA',
        'FO_GPON_RESID_ETH',
        'FO_GPON_RESID_ETH_INFORMACAO',
        'FO_GPON_RESID_ETH_ESTACAO_ENTREGA',
        'FO_GPON_RESID_ETH_TECNOLOGIA',
        'FO_XGSPON_RESID_ETH',
        'FO_XGSPON_RESID_ETH_INFORMACAO',
        'FO_XGSPON_RESID_ETH_ESTACAO_ENTREGA',
        'FO_XGSPON_RESID_ETH_TECNOLOGIA',
        'HFC_BSOD',
        'HFC_BSOD_INFORMACAO',
        'HFC_BSOD_ESTACAO_ENTREGA',
        'HFC_BSOD_TECNOLOGIA',
        'FO_GPON_RESID_ETH_PRE_VIAVEL',
        'FO_GPON_RESID_ETH_PRE_VIAVEL_INFORMACAO',
        'FO_GPON_RESID_ETH_PRE_VIAVEL_ESTACAO_ENTREGA',
        'FO_GPON_RESID_ETH_PRE_VIAVEL_TECNOLOGIA',
        'FO_GPON_ETH',
        'FO_GPON_ETH_INFORMACAO',
        'FO_GPON_ETH_ESTACAO_ENTREGA',
        'FO_GPON_ETH_TECNOLOGIA',
        'FO_EDD_DIRETO_CLARO',
        'FO_EDD_DIRETO_CLARO_INFORMACAO',
        'FO_EDD_DIRETO_CLARO_ESTACAO_ENTREGA',
        'FO_EDD_DIRETO_CLARO_TECNOLOGIA',
        'FO_EDD_CLARO_PRE_VIAVEL',
        'FO_EDD_CLARO_PRE_VIAVEL_INFORMACAO',
        'FO_EDD_CLARO_PRE_VIAVEL_ESTACAO_ENTREGA',
        'FO_EDD_CLARO_PRE_VIAVEL_TECNOLOGIA',
        'FO_EDD_ETH',
        'FO_EDD_ETH_INFORMACAO',
        'FO_EDD_ETH_ESTACAO_ENTREGA',
        'FO_EDD_ETH_TECNOLOGIA',
        'FOETHERNET',
        'FOETHERNET_INFORMACAO',
        'FOETHERNET_ESTACAO_ENTREGA',
        'FOETHERNET_TECNOLOGIA',
        'FO_SDH',
        'FO_SDH_INFORMACAO',
        'FO_SDH_ESTACAO_ENTREGA',
        'FO_SDH_TECNOLOGIA',
        'ADE_SDH',
        'ADE_SDH_INFORMACAO',
        'ADE_SDH_ESTACAO_ENTREGA',
        'ADE_SDH_TECNOLOGIA',
        'SATELITE_BANDA_KA',
        'SATELITE_BANDA_KA_INFORMACAO',
        'SATELITE_BANDA_KA_ESTACAO_ENTREGA',
        'SATELITE_BANDA_KA_TECNOLOGIA',
        'SATELITE_BANDA_KU',
        'SATELITE_BANDA_KU_INFORMACAO',
        'SATELITE_BANDA_KU_ESTACAO_ENTREGA',
        'SATELITE_BANDA_KU_TECNOLOGIA',
        'TERCEIROS_ETH',
        'TERCEIROS_ETH_INFORMACAO',
        'TERCEIROS_ETH_ESTACAO_ENTREGA',
        'TERCEIROS_ETH_TECNOLOGIA',
        'RADIO_IP',
        'RADIO_IP_INFORMACAO',
        'RADIO_IP_ESTACAO_ENTREGA',
        'RADIO_IP_TECNOLOGIA',
        'ID_ACESSO_ABORDADO',
        'FACILIDADE_ABORDADO',
        'TECNOLOGIA_ABORDADO',
        'ESTACAO_ENTREGA_ABORDADO']

        # for i in range(len(self.__resumo_soe)):
        #     self.__resumo_soe[i] = self.__resumo_soe[i].removesuffix('\n')
        #     self.__resumo_soe[i] = self.__resumo_soe[i].split("\t")

        # acum = 0
        # while acum < len(self.__resumo_soe[0]):
        #     self.__resumo_soe[0][acum] = self.__resumo_soe[0][acum].replace(" ","_").upper()
        #     if self.__resumo_soe[0][acum] == '%DISPONIBILIDADE':
        #         self.__resumo_soe[0][acum] = self.__resumo_soe[0][acum].replace('%','')
        #     acum += 1
        # acum_resumo = 0
        # acum_facilidade = 0
        # while acum_resumo < len(self.__resumo_soe[0]):
        #     if self.__resumo_soe[0][acum_resumo] == facilidades[acum_facilidade]:
        #         self.__resumo_soe[0][acum_resumo + 1] = f'{self.__resumo_soe[0][acum_resumo]}_{unidecode(self.__resumo_soe[0][acum_resumo + 1])}'
        #         self.__resumo_soe[0][acum_resumo + 2] = f'{self.__resumo_soe[0][acum_resumo]}_{unidecode(self.__resumo_soe[0][acum_resumo + 2])}'
        #         self.__resumo_soe[0][acum_resumo + 3] = f'{self.__resumo_soe[0][acum_resumo]}_{unidecode(self.__resumo_soe[0][acum_resumo + 3])}'
        #         if acum_facilidade < len(facilidades)-1:
        #             acum_facilidade += 1
        #     acum_resumo += 1

        # df = pd.DataFrame(np.array(self.__resumo_soe[1:]), columns=self.__resumo_soe[0])
        # df[['ID', 'SEV']] = df[['ID', 'SEV']].apply(pd.to_numeric)
        self.__resumo_soe.columns = colunas_resumosoe
        self.__resumo_soe = self.__resumo_soe.fillna('')

        for i,v in self.__resumo_soe.iterrows():
            if ' / ' in v.FO_EDD_ETH_INFORMACAO:

                self.__resumo_soe.at[i,'FO_EDD_ETH_INFORMACAO'] = v.FO_EDD_ETH_INFORMACAO.replace(' / ',f"{'NUVEM FO EDD ETH / '*v.FO_EDD_ETH_INFORMACAO.count('/')}")
                if self.__resumo_soe.at[i,'FO_EDD_ETH_INFORMACAO'][-2] == '/':
                    self.__resumo_soe.at[i,'FO_EDD_ETH_INFORMACAO'] = self.__resumo_soe.at[i,'FO_EDD_ETH_INFORMACAO'][:-3]

        self.__resumo_soe['FO_EDD_ETH_INFORMACAO'] = self.__resumo_soe['FO_EDD_ETH_INFORMACAO'].replace('','NUVEM: NUVEM FO EDD ETH')

        return self.__resumo_soe
    
class tratamentoResultado():

    def __init__(self,arq):
        file = open(arq, 'r', encoding="utf8")
        self.__resultado = file.readlines()
        file.close()
    
    def trata_resultado(self):
        
        for i in range(len(self.__resultado)):
            self.__resultado[i] = self.__resultado[i].removesuffix('\n')
            self.__resultado[i] = self.__resultado[i].split("\t")

        df = pd.DataFrame(np.array(self.__resultado[1:]),columns=self.__resultado[0])
        df[['ID', 'SEV']] = df[['ID', 'SEV']].apply(pd.to_numeric)

        return df
    
# Siglas de UF (ancora confiavel para detectar deslocamento no inicio da linha)
UFS = {'AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA',
       'PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'}


class tratamentoRestricao():

    def __init__(self, arq):
    
        with open(arq, 'r', encoding="utf8") as file:
            linhas = file.readlines()

        # Remove a linha de titulo ("Restricao") se existir e separa o cabecalho
        if linhas and linhas[0].strip().lower() in ('restricao', 'restrição', ''):
            linhas.pop(0)

        header = linhas[0].rstrip('\n').split('\t')
        # Normaliza os nomes das colunas (sem acento, sem espaco, maiusculo)
        self.colunas = [unidecode(h).replace(" ", "_").upper() for h in header]

        self.ncols = len(self.colunas)              # 17 (agora com CAMADA)
        self.idx_uf = self.colunas.index('ESTADO')  # ancora de deslocamento

        self.dados = [l.rstrip('\n').split('\t') for l in linhas[1:] if l.strip()]
        self.descartadas = []  # linhas que nao deu para realinhar com seguranca

    def _realinhar(self, r):
        """Padroniza UMA linha para self.ncols colunas de forma deterministica.

        Ancoras: POSICAO e sempre o ultimo campo; ESTADO (UF) e sempre sigla de
        2 letras e deve ficar no indice self.idx_uf. Se a linha ja esta completa,
        retorna como esta.
        """
        if len(r) == self.ncols:
            return r
        if len(r) > self.ncols:
            # Nunca observado; sinaliza para conferencia manual.
            self.descartadas.append(('mais_colunas', r))
            return None

        posicao = r[-1]      # POSICAO sempre no fim
        corpo = r[:-1]       # restante da linha

        # Corrige deslocamento inicial: se a UF caiu uma posicao antes, o campo
        # anterior a ESTADO veio vazio e foi descartado -> repoe em branco.
        if len(corpo) > self.idx_uf and corpo[self.idx_uf] in UFS:
            pass                                    # ja alinhado
        elif len(corpo) > self.idx_uf - 1 and corpo[self.idx_uf - 1] in UFS:
            corpo.insert(self.idx_uf - 1, ' ')      # repoe campo vazio antes da UF

        # Completa as colunas finais vazias que o export descartou
        if len(corpo) > self.ncols - 1:
            self.descartadas.append(('corpo_grande', r))
            return None
        corpo += [' '] * (self.ncols - 1 - len(corpo))
        corpo.append(posicao)                       # POSICAO volta para o fim
        return corpo

    def trata_restricao(self):
        data = [self._realinhar(r) for r in self.dados]
        data = [r for r in data if r is not None]

        df = pd.DataFrame(np.array(data), columns=self.colunas)
        df[['ID', 'SEV']] = df[['ID', 'SEV']].apply(pd.to_numeric, errors='coerce')

        return df
    

# Siglas de UF (ancora confiavel para detectar deslocamento no inicio da linha)
UFS = {'AC','AL','AP','AM','BA','CE','DF','ES','GO','MA','MT','MS','MG','PA',
       'PB','PR','PE','PI','RJ','RN','RS','RO','RR','SC','SP','SE','TO'}

COLUNAS = ['ID','SEV','Camada','OBJECTID','ESTACAO_ENTREGA','UF','SIGLA_LOC',
           'NOME_NUVEM','REDE','TECNOLOGIA','MEIO_TRANSMISSAO','SITUACAO',
           'ALT_NUVEM','PADRAO_PON','PROPRIETARIO','OBSERVACAO','ORIGEM',
           'DATA_ATUALIZACAO','DATA_PREVISAO','VEL_MAX_VIABILIDADE',
           'VEL_MAX_SEV_AUTOMATICA','MOTIVO','OPERADORA','NUMERO_OPERACIONAL',
           'SIGLA_ESTACAO_CLARO','CODIGO_DESCARGA_CSL','SIGLA_ESTACAO_RESID',
           'TIPO_INFRA','TIPO','STATUS','DONO','ROTEADORES_QTD','TX',
           'FABRICANTE_OLT','ABRANGENCIA','CONCENTRADOR_OLT','POSICAO']

NCOLS = len(COLUNAS)          # 37
IDX_UF = COLUNAS.index('UF')  # 5


class tratamentoNuvens():

    def __init__(self, arq):
        self.df = pd.read_csv(arq,sep='\t',skiprows=1)
       
    def trata_nuvens(self):
        nuvens_df = self.df.drop(index=self.df[self.df.Camada == 'Nuvens Terceiros Rádio IP'].index.to_list(),columns=['POSICAO','Camada','OBJECTID']).drop_duplicates().reset_index(drop=True)

        nuvens_df['NOME_NUVEM'] = nuvens_df['NOME_NUVEM'].replace(' ','NUVEM FO EDD ETH')

        return nuvens_df
    
class tratamentoNuvensTerceiros():

    def __init__(self,arq):
        file = open(arq, 'r', encoding="utf8")
        self.nuvens_terceiros = file.readlines()
        file.close()

    def trata_nuvens(self):
        self.nuvens_terceiros.pop(0)
        for i in range(len(self.nuvens_terceiros)):
            self.nuvens_terceiros[i] = self.nuvens_terceiros[i].removesuffix('\n')
            self.nuvens_terceiros[i] = self.nuvens_terceiros[i].split("\t")

        self.nuvens_terceiros[0].remove('Camada')

        data = self.nuvens_terceiros[1:]


        i = 0
        tam = len(data)
        while i < tam:
            if i != (len(data) - 1):
                sev_atual = data[i][1]
                tam_nuvem_atual = len(data[i])
                if data[i+1][1] == sev_atual:
                    if len(data[i+1]) > tam_nuvem_atual:
                        data.pop(i)
                        tam -= 1
                    else:
                        i += 1
                else:
                    i += 1
            else: i += 1

        i = 0
        while i < len(data):
            x = 1
            while x < len(data[i]):
                if x == 1:
                    sev = data[i][x]
                    id = data[i][x-1]
                    x += 20
                else:
                    data[i].insert(x, sev)
                    data[i].insert(x, id)
                    x += 21
            i += 1
            
        i = 0
        new_data = []
        while i < len(data):
            x = 0
            while x < len(data[i]):
                new_data.append(data[i][x:x+21])
                x += 21
            i += 1
        data = new_data

        nuvens_terceiros_df = pd.DataFrame(data, columns=self.nuvens_terceiros[0])

        nuvens_terceiros_df['SEV'] = pd.to_numeric(nuvens_terceiros_df['SEV'],downcast='signed', errors='coerce')
        nuvens_terceiros_df = nuvens_terceiros_df.drop(columns=['ID','POSICAO'])
        nuvens_terceiros_df = nuvens_terceiros_df.drop_duplicates()

        return nuvens_terceiros_df