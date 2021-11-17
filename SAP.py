import time
import math
from collections import Counter
import numpy as np
from numpy.testing._private.utils import assert_almost_equal
import pandas as pd

class Portfolio():
    """Analizza un portafoglio"""
    PATH = r'C:\Users\Administrator\Desktop\Sbwkrq\SAP' # Percorso che porta al portafogio

    def __init__(self, file_portafoglio):
        """
        Initialize the class.

        Parameters:
        file_portafoglio(str) = file da analizzare

        """
        self.file_portafoglio = file_portafoglio
        self.df_portfolio = pd.read_excel(self.file_portafoglio, sheet_name='portfolio_valori', index_col=None)
        self.df_mappatura = pd.read_excel(self.file_portafoglio, sheet_name='mappatura', index_col=None)
        self.micro_asset_class = ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute', 'Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Corporate', 'Obbligazionario Euro High Yield',
            'Obbligazionario Globale Aggregate', 'Obbligazionario Paesi Emergenti', 'Obbligazionario Globale High Yield', 'Azionario Europa', 'Azionario North America', 'Azionario Pacific',
            'Azionario Emerging Markets', 'Commodities']
        self.dict_macro = {'Monetario' : ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute'], 
            'Obbligazionario' : ['Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Corporate', 'Obbligazionario Euro High Yield', 'Obbligazionario Globale Aggregate', 'Obbligazionario Paesi Emergenti', 'Obbligazionario Globale High Yield'],
            'Azionario' : ['Azionario Europa', 'Azionario North America', 'Azionario Pacific', 'Azionario Emerging Markets'],
            'Commodities' : ['Commodities'],
            }
        self.macro_asset_class = ['Monetario', 'Obbligazionario', 'Azionario', 'Commodities']
        self.strumenti = ['cash', 'gov_bond', 'corp_bond', 'equity', 'certificate', 'etf', 'fund', 'real_estate', 'hedge_fund', 'private_equity', 'venture_capital', 'private_debt', 'insurance', 'gp', 'pip']
        self.dict_str_fig = {'cash' : 'Conto corrente', 'gov_bond' : 'Obbligazioni', 'corp_bond' : 'Obbligazioni', 'certificate' : 'Obbligazioni strutturate / Certificates', 'equity' : 'Azioni',
            'etf' : 'ETF/ETC', 'fund' : 'Fondi comuni/Sicav', 'real_estate' : 'Real Estate', 'hedge_fund' : 'Hedge funds', 'private_equity' : 'Private Equity', 'venture_capital' : 'Venture Capital',
            'private_debt' : 'Private Debt', 'insurance' : 'Polizze', 'gp' : 'Gestioni patrimoniali',
            'pip' : 'Fondi pensione'}
        self.dict_str_comm = {'cash' : 'liquidità', 'gov_bond' : 'obbligazioni governative', 'corp_bond' : 'obbligazioni societarie', 'certificate' : 'certificati', 'equity' : 'azioni',
            'etf' : 'etf', 'fund' : 'fondi', 'real_estate' : 'real estate', 'hedge_fund' : 'fondi hedge', 'private_equity' : 'private equity', 'venture_capital' : 'venture capital',
            'private_debt' : 'private debt', 'insurance' : 'polizze', 'gp' : 'gestioni patrimoniali', 'pip' : 'fondi pensione'}
        self.valute = ['EUR', 'USD', 'YEN', 'CHF', 'GBP', 'AUD', 'ALTRO']
        self.amministrato = ['cash', 'gov_bond', 'corp_bond', 'equity', 'certificate']
     
    def peso_micro(self):
        """
        Calcola il peso delle micro asset class.
        
        Returns a dictionary.
        """
        vector_peso_prodotti = (self.df_portfolio['controvalore_in_euro'] / self.df_portfolio['controvalore_in_euro'].sum()).to_numpy()
        matrix_mappatura = self.df_mappatura.loc[:, self.micro_asset_class].to_numpy()
        matrix_mappatura = np.nan_to_num(matrix_mappatura, nan=0.0)
        vector_peso_micro = matrix_mappatura.T @ vector_peso_prodotti
        series_peso_micro = pd.Series(vector_peso_micro, index=self.micro_asset_class)
        dict_peso_micro = series_peso_micro.to_dict()
        assert_almost_equal(actual=sum(dict_peso_micro.values()), desired=1.00, decimal=3, err_msg='la somma delle micro categorie non fa cento', verbose=True)
        return dict_peso_micro

    def peso_macro(self):
        """
        Calcola il peso delle macro categorie.

        Returns a dictionary.
        """
        dict_peso_micro = self.peso_micro()
        dict_peso_macro = {macro : sum(dict_peso_micro[item] for item in micro) for macro, micro in self.dict_macro.items()}
        assert_almost_equal(actual=sum(dict_peso_macro.values()), desired=1.00, decimal=3, err_msg='la somma delle macro categorie non fa cento', verbose=True)
        return dict_peso_macro
        
    def peso_strumenti(self):
        """
        Calcola il peso degli strumenti.
        
        Returns 2 dictionaries.
        """
        dict_strumenti = {strumento : self.df_portfolio.loc[self.df_portfolio['strumento']==strumento,'controvalore_in_euro'].sum() / self.df_portfolio['controvalore_in_euro'].sum() for strumento in self.strumenti}
        df_peso_strumenti = pd.DataFrame.from_dict(dict_strumenti, orient='index', columns=['peso_strumento'])
        df_peso_strumenti.rename(self.dict_str_fig, inplace=True)
        df_peso_strumenti = df_peso_strumenti.groupby(df_peso_strumenti.index, sort=False).agg({'peso_strumento' : sum})
        series_peso_strumenti = df_peso_strumenti['peso_strumento'].squeeze()
        dict_peso_strumenti = series_peso_strumenti.to_dict()
        dict_strumenti_attivi = {k : v * 100 for k, v in dict_strumenti.items() if v!=0}
        dict_strumenti_attivi = {k: v for k, v in sorted(dict_strumenti_attivi.items(), key=lambda item: item[1], reverse=True)}
        dict_peso_strumenti_attivi = {self.dict_str_comm[k] : v for k, v in dict_strumenti_attivi.items()}
        assert_almost_equal(actual=sum(dict_peso_strumenti.values()), desired=1.00, decimal=3, err_msg='la somma degli strumenti non fa cento', verbose=True )
        return {'strumenti_figure' : dict_peso_strumenti, 'strumenti_commento' : dict_peso_strumenti_attivi}
    
    def peso_valuta_per_denominazione(self, dataframe=''):
        """
        Calcola il peso delle valute considerando la loro denominazione.
        
        Parameters
        df(str) : nome del dataframe

        Returns a dictionary.
        """
        df = dataframe if isinstance(dataframe, pd.DataFrame)==True else self.df_portfolio
        dict_valute = {valuta : df.loc[df['divisa']==valuta, 'controvalore_in_euro'].sum() / df['controvalore_in_euro'].sum() for valuta in self.valute[:-1]}
        dict_valute['ALTRO'] = df.loc[~df['divisa'].isin(self.valute[:-1]), 'controvalore_in_euro'].sum() / df['controvalore_in_euro'].sum()
        assert_almost_equal(actual=sum(dict_valute.values()), desired=1.00, decimal=3, err_msg='la somma delle valute per denominazione non fa cento', verbose=True)
        return dict_valute
    
    def peso_valuta_per_composizione(self, dataframe=''):
        """
        Calcola il peso delle valute considerando la loro scomposizione in mercati.

        - Mappa ogni mercato in valuta così da ricondurre tutti i mercati ad una sola valuta, e poi moltiplica la matrice ottenuta per il vettore dei pesi del prodotto.
        
        Parameters
        dataframe(str) : nome del dataframe

        Returns a dictionary.
        """
        df_p = dataframe if isinstance(dataframe, pd.DataFrame)==True else self.df_portfolio
        vector_peso_prodotti = (df_p['controvalore_in_euro'] / df_p['controvalore_in_euro'].sum()).to_numpy()
        df_m = self.df_mappatura.loc[(self.df_mappatura['ISIN'].isin(list(dataframe['ISIN']))) & (self.df_mappatura['nome'].isin(list(dataframe['nome']))), self.micro_asset_class] if isinstance(dataframe, pd.DataFrame)==True else self.df_mappatura.loc[:, self.micro_asset_class]
        dict_valute = {'Monetario Euro' : 'EUR', 'Monetario USD' : 'USD', 'Monetario Altre Valute' : 'ALTRO', 'Obbligazionario Euro Governativo All Maturities' : 'EUR', 'Obbligazionario Euro Corporate' : 'EUR', 'Obbligazionario Euro High Yield' : 'EUR',
            'Obbligazionario Globale Aggregate' : 'ALTRO', 'Obbligazionario Paesi Emergenti' : 'ALTRO', 'Obbligazionario Globale High Yield' : 'ALTRO', 'Azionario Europa' : 'EUR', 'Azionario North America' : 'USD', 'Azionario Pacific' : 'ALTRO',
            'Azionario Emerging Markets' : 'ALTRO', 'Commodities' : 'USD'}
        df_m.columns = df_m.columns.map(dict_valute) # assegna ad ogni mercato una valuta
        df_m = df_m.groupby(df_m.columns, axis=1).sum() # raggruppa per valuta
        name_order = ['EUR', 'USD', 'ALTRO']
        df_m = df_m[name_order]
        matrix_mappatura_valute = df_m.T.to_numpy()
        vector_valute = matrix_mappatura_valute @ vector_peso_prodotti
        dict_valute = {name_order[_] : vector_valute[_] for _ in range(len(name_order))}
        assert_almost_equal(actual=np.sum(vector_valute), desired=1.00, decimal=2, err_msg='la somma delle valute per composizione non fa cento', verbose=True)
        return dict_valute

    def peso_valuta_ibrido(self):
        """
        Calcola il peso delle valute considerando la loro scomposizione in mercati per l'amministrato, e la loro scomposizione in mercati per il gestito.
        
        Returns a dictionary.
        """
        df_amministrato = self.df_portfolio.loc[self.df_portfolio['strumento'].isin(self.amministrato)]
        dict_valute_amministrato = self.peso_valuta_per_denominazione(dataframe=df_amministrato) if df_amministrato.empty==False else {}
        df_gestito = self.df_portfolio.loc[~self.df_portfolio['strumento'].isin(self.amministrato)]
        dict_valute_gestito = self.peso_valuta_per_composizione(dataframe=df_gestito) if df_gestito.empty==False else {}
        dict_amministrato_su_ptf = Counter({key : value*df_amministrato['controvalore_in_euro'].sum()/self.df_portfolio['controvalore_in_euro'].sum() for key, value in dict_valute_amministrato.items()})
        dict_gestito_su_ptf = Counter({key : value*df_gestito['controvalore_in_euro'].sum()/self.df_portfolio['controvalore_in_euro'].sum() for key, value in dict_valute_gestito.items()})
        dict_amministrato_su_ptf.update(dict_gestito_su_ptf) # unione dei due dizionari
        dict_valute = dict(dict_amministrato_su_ptf)
        assert_almost_equal(actual=sum(dict_valute.values()), desired=1.00, decimal=2, err_msg='la somma delle valute non fa cento', verbose=True)
        return dict_valute

    def duration(self):
        """
        Calcola le duration per i comparti obbligazionari, non considera le duration ND.

        Returns a dictionary
        """
        df_p = self.df_portfolio[['ISIN', 'nome', 'controvalore_in_euro']]
        df_m = self.df_mappatura
        df_m['asset_class'] = df_m.apply(lambda x : x[self.micro_asset_class].index[x[self.micro_asset_class] == 1.00].values[0] if any(x[self.micro_asset_class]==1.00) else 'Prodotto multi asset', axis=1)
        df_m = df_m[['ISIN', 'nome', 'asset_class']]
        df_bond = pd.read_excel(self.file_portafoglio, sheet_name='obbligazioni', index_col=None)
        df_bond = df_bond[['ISIN', 'Descrizione', 'Duration']]
        df_bond = df_bond.merge(df_p, how='left', left_on=['ISIN', 'Descrizione'], right_on=['ISIN', 'nome']).drop('nome', axis=1) # aggiungi il controvalore
        df_bond = df_bond.merge(df_m, how='left', left_on=['ISIN', 'Descrizione'], right_on=['ISIN', 'nome']).drop('nome', axis=1) # aggiungi l'asset class
        durations = {}
        classi_obbligazionarie = [micro for micro in self.micro_asset_class if micro.startswith('Obbligazionario')]
        for classe in classi_obbligazionarie:
            df = df_bond.loc[(df_bond['asset_class']==classe) & (df_bond['Duration']!='ND')]
            duration = df['Duration'].to_numpy()
            ctv = df['controvalore_in_euro'].to_numpy()
            durations[classe] = duration @ ctv / sum(ctv) if sum(ctv) > 0 else 0
        return durations

    def risk(self):
        """Calcola la volatilità del portafoglio"""
        vector_micro = np.array(list(self.peso_micro().values()))
        df_benchmark = pd.read_excel(self.file_portafoglio, sheet_name='rischio', index_col=0)
        matrix_benchmark = df_benchmark.to_numpy()
        volatilità = math.sqrt((vector_micro @ matrix_benchmark) @ vector_micro.T)
        return volatilità

if __name__ == "__main__":
    # TODO : crea un'instance classmethod per azimut, con i suoi strumenti, le sue micro, macro, valute etc.
    start = time.perf_counter()
    PTF = 'ptf_20.xlsx'
    PTF_ELABORATO = PTF[:-5] + '_elaborato.xlsx'
    _ = Portfolio(file_portafoglio=PTF)
    _.peso_micro()
    _.peso_macro()
    _.peso_strumenti()
    _.peso_valuta_ibrido()
    _.duration()
    _.risk()
    end = time.perf_counter()
    print("Elapsed time: ", round(end - start, 2), 'seconds')