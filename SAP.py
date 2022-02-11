import math
import os
import sys
import time
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

with os.add_dll_directory('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\_blpapi'):
    import blpapi

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import seaborn as sns
from xbbg import blp


class Portfolio():

    def __init__(self, intermediario, file_portafoglio=None):
        """
        Initialize the class.

        Parameters:
            intermediario {str} = intermediario per cui fare l'analisi
            file_portafoglio {str} = nome del file da analizzare

        Returns:
            Initialized portfolio
        """
        self.intermediario = intermediario
        self.path = self.percorso()['path']
        self.file_portafoglio = file_portafoglio
        # Se non viene specificato il nome del portafoglio da analizzare, prendi il primo nella cartella corrente
        # che ha il nome "ptf_*.xlsx"
        if self.file_portafoglio is None:
            self.file_portafoglio = self.percorso()['file']
        # df_portfolio {str} = nome del foglio del file da analizzare in cui si trovano i dati relativi a:
        # ISIN, nome, intermediario, strumento, quantità, ctv_iniziale€, PMC, divisa, prezzo, rateo, ctv_finale€
        self.df_portfolio = pd.read_excel(self.file_portafoglio, sheet_name='portfolio_valori')
        # df_mappatura {str} = nome del foglio del file da analizzare in cui si trova la mappatura dei prodotti
        self.df_mappatura = pd.read_excel(self.file_portafoglio, sheet_name='mappatura')

        #---Caratteristiche dell'analisi condivise da tutti i clienti---#
        # TODO : posso usare la funzione zip per unire in unico dizionario la lista strumenti e la lista fonts_strumenti
        # Lista degli strumenti possibili presenti nel file di input
        self.strumenti = ['cash', 'gov_bond', 'corp_bond', 'equity', 'certificate', 'etf', 'fund', 'real_estate', 'hedge_fund',
            'private_equity', 'venture_capital', 'private_debt', 'insurance', 'gp', 'pip', 'alternative']
        # Lista contenente i colori da associare a ciascun strumento, in ordine
        self.fonts_strumenti = ['B1A0C7', '93DEFF', 'FFFF66', 'F79646', '00B0F0', '0066FF', 'FF3737', 'FB9FDA', 'BF8F00',
            'C6E0B4', '7030A0', 'FFC000', '92D050', 'BFBFBF', 'FFFFCC']
        # Lista delle valute possibili presenti nel file di input
        self.valute = ['EUR', 'USD', 'YEN', 'CHF', 'GBP', 'AUD', 'ALTRO']
        # Lista contenente i colori da associare a ciascuna valuta, in ordine
        self.fonts_valute = ['3366FF', '339966', 'FF99CC', 'FF6600', 'B7DEE8', 'FF9900', 'FFFF66']
        # Lista degli strumenti del comparto amministrato
        self.amministrato = ['cash', 'gov_bond', 'corp_bond', 'equity', 'certificate']
        # Assegnazione di ciascun mercato con una delle valute possibili
        self.dict_valute_per_composizione = {'Monetario Euro' : 'EUR', 'Monetario USD' : 'USD', 'Monetario Altre Valute' : 'ALTRO',
            'Obbligazionario Euro Governativo All Maturities' : 'EUR', 'Obbligazionario Euro Corporate' : 'EUR',
            'Obbligazionario Euro High Yield' : 'EUR', 'Obbligazionario Globale Aggregate' : 'ALTRO',
            'Obbligazionario Paesi Emergenti' : 'ALTRO', 'Obbligazionario Globale High Yield' : 'ALTRO',
            'Azionario Europa' : 'EUR', 'Azionario North America' : 'USD', 'Azionario Pacific' : 'ALTRO',
            'Azionario Emerging Markets' : 'ALTRO', 'Commodities' : 'USD'} # assegna una valuta ad ogni mercato
        # Lista ordinata delle valute permesse nell'ambito del gestito
        self.valute_per_composizione = ['EUR', 'USD', 'ALTRO']

        # TODO: i warning su micro, strumenti e valute vanno inseriti qui

        #---Caratteristiche dell'analisi specifiche del cliente---#
        if intermediario == 'azimut':
            self.micro_asset_class = ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute',
                'Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Corporate', 'Obbligazionario Euro High Yield',
                'Obbligazionario Globale Aggregate', 'Obbligazionario Paesi Emergenti', 'Obbligazionario Globale High Yield',
                'Azionario Europa', 'Azionario North America', 'Azionario Pacific', 'Azionario Emerging Markets', 'Commodities']
            # Lista contenente i colori da associare a ciascuna micro, in ordine
            self.fonts_micro = ['E4DFEC', 'CCC0DA', 'B1A0C7', '92CDDC', '00B0F0', '0033CC', '0070C0', '1F497D', '000080', 'F79646',
                'FFCC66', 'DA5300', 'F62F00', 'EDF06A']
            self.dict_macro_micro = {
                'Monetario' : ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute'], 
                'Obbligazionario' : ['Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Corporate',
                    'Obbligazionario Euro High Yield', 'Obbligazionario Globale Aggregate', 'Obbligazionario Paesi Emergenti', 
                    'Obbligazionario Globale High Yield'],
                'Azionario' : ['Azionario Europa', 'Azionario North America', 'Azionario Pacific', 'Azionario Emerging Markets'],
                'Commodities' : ['Commodities'],
                }
            self.macro_asset_class = ['Monetario', 'Obbligazionario', 'Azionario', 'Commodities']
            # Lista contenente i colori da associare a ciascuna macro, in ordine
            self.fonts_macro = ['B1A0C7', '92CDDC', 'F79646', 'EDF06A']

        elif intermediario == 'copernico':
            self.micro_asset_class = ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute',
                'Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Corporate', 'Obbligazionario Euro High Yield',
                'Obbligazionario Globale Aggregate', 'Obbligazionario Paesi Emergenti', 'Obbligazionario Globale High Yield',
                'Azionario Europa', 'Azionario North America', 'Azionario Pacific', 'Azionario Emerging Markets', 'Commodities']
            # Lista contenente i colori da associare a ciascuna micro, in ordine
            self.fonts_micro = ['E4DFEC', 'CCC0DA', 'B1A0C7', '92CDDC', '00B0F0', '0033CC', '0070C0', '1F497D', '000080', 'F79646',
                'FFCC66', 'DA5300', 'F62F00', 'EDF06A']
            self.dict_macro_micro = {
                'Monetario' : ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute'], 
                'Obbligazionario' : ['Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Corporate',
                    'Obbligazionario Euro High Yield', 'Obbligazionario Globale Aggregate', 'Obbligazionario Paesi Emergenti', 
                    'Obbligazionario Globale High Yield'],
                'Azionario' : ['Azionario Europa', 'Azionario North America', 'Azionario Pacific', 'Azionario Emerging Markets'],
                'Commodities' : ['Commodities'],
                }
            self.macro_asset_class = ['Monetario', 'Obbligazionario', 'Azionario', 'Commodities']
            # Lista contenente i colori da associare a ciascuna macro, in ordine
            self.fonts_macro = ['B1A0C7', '92CDDC', 'F79646', 'EDF06A']

    @classmethod
    def azimut(cls, file_portafoglio):
        """
        Classe Portfolio con i parametri di azimut
        
        Returns:
            Portfolio Azimut istance
        """
        micro_asset_class = ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute', 'Obbligazionario Euro Governativo All Maturities',
            'Obbligazionario Euro Corporate', 'Obbligazionario Euro High Yield', 'Obbligazionario Globale Aggregate',
            'Obbligazionario Paesi Emergenti', 'Obbligazionario Globale High Yield', 'Azionario Europa', 'Azionario North America',
            'Azionario Pacific', 'Azionario Emerging Markets', 'Commodities']
        # Lista contenente i colori da associare a ciascuna micro, in ordine
        fonts_micro = ['E4DFEC', 'CCC0DA', 'B1A0C7', '92CDDC', '00B0F0', '0033CC', '0070C0', '1F497D', '000080', 'F79646', 'FFCC66', 'DA5300', 'F62F00', 'EDF06A']
        dict_macro_micro = {
            'Monetario' : ['Monetario Euro', 'Monetario USD', 'Monetario Altre Valute'], 
            'Obbligazionario' : ['Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Corporate',
                'Obbligazionario Euro High Yield', 'Obbligazionario Globale Aggregate', 'Obbligazionario Paesi Emergenti',
                'Obbligazionario Globale High Yield'],
            'Azionario' : ['Azionario Europa', 'Azionario North America', 'Azionario Pacific', 'Azionario Emerging Markets'],
            'Commodities' : ['Commodities'],
            }
        macro_asset_class = ['Monetario', 'Obbligazionario', 'Azionario', 'Commodities']
        # Lista contenente i colori da associare a ciascuna macro, in ordine
        fonts_macro = ['B1A0C7', '92CDDC', 'F79646', 'EDF06A']
        return cls(file_portafoglio, micro_asset_class, fonts_micro, dict_macro_micro, macro_asset_class, fonts_macro)

    @staticmethod
    def percorso():
        """
        Trova il percorso della cartella di lavoro e del file da analizzare

        Raises:
            Exception = Non esistono file excel del tipo 'ptf_*.xlsx' nella cartella

        Returns:
            path {str} = percorso della cartella corrente
            file {str} = percorso del file da analizzare
        """
        path = Path().cwd()
        path_object = path.glob('[ptf_]*[0-9].xlsx')
        list_of_files = list(path_object)
        # Se la lista è vuota
        if not list_of_files:
            sys.tracebacklimit = 0 # silenzia il traceback
            raise Exception(f"Non esistono file excel del tipo 'ptf_*.xlsx' nella cartella {path}")
        file = list(list_of_files)[0]
        # Se la lista contiene più file con la stessa traccia
        if len(list_of_files) > 1:
            print(f"Ci sono più files del tipo 'ptf_*.xlsx' nella cartella {path}:\n{list_of_files}.\
            \nViene analizzato il primo: {file}.")
        return {'path' : path, 'file' : file}

    def test(self):
        # TODO: testa tutto nell'__init__ della classe
        """
        Test dei parametri del portafoglio
        
        Raises:
            AssertionError = Il numero dei prodotti nel foglio 'portfolio_valori' e quelli nel foglio 'mappatura' non corrispondono
            AssertionError = Almeno una riga della matrice di mappatura non somma ad 1
            
        Returns:
            None
        """
        # TODO: Verifica che non ci siano ISIN doppi, né nomi doppi
        # TODO: Verifica che nel foglio mappatura siano presenti tutte le micro categorie
        # proposte in self.micro_asset_class
        # np.testing.assert_equal(actual=)
        # TODO: Verifica che la numerosità dei font sia pari alla numerosità degli oggetti a cui si riferiscono
        # Verifica che nel foglio di mappatura ci siano tante righe quanti sono gli strumenti
        # presenti nel foglio portfolio_valori
        num_assets = len(self.df_portfolio.index)
        np.testing.assert_equal(actual=len(self.df_mappatura.index), desired=num_assets,
            err_msg="Il numero dei prodotti nel foglio 'portfolio_valori' e quelli nel foglio 'mappatura' non corrispondono")
        # Verifica che la somma delle righe nella matrice di mappatura sia sempre pari ad 1
        list_sum_of_rows = self.df_mappatura.loc[:, self.micro_asset_class].sum(axis=1)
        np.testing.assert_equal(actual=sum(list_sum_of_rows), desired=num_assets,
            err_msg="Almeno una riga della matrice di mappatura non somma ad 1")

    def peso_micro(self):
        """
        Calcola il peso delle micro asset class di un portafoglio.

        Raises:
            AssertionError = La somma delle micro categorie non fa cento

        Returns:
            dict_peso_micro {dict} = dizionario che associa ad ogni micro il peso relativo
        """
        vector_peso_prodotti = (self.df_portfolio['controvalore_in_euro'] / self.df_portfolio['controvalore_in_euro'].sum()).to_numpy()
        matrix_mappatura = self.df_mappatura.loc[:, self.micro_asset_class].to_numpy()
        matrix_mappatura = np.nan_to_num(matrix_mappatura, nan=0.0)
        vector_peso_micro = matrix_mappatura.T @ vector_peso_prodotti
        series_peso_micro = pd.Series(vector_peso_micro, index=self.micro_asset_class)
        dict_peso_micro = series_peso_micro.to_dict()
        np.testing.assert_almost_equal(actual=sum(dict_peso_micro.values()), desired=1.00, decimal=3, err_msg="La somma delle micro categorie non fa cento", verbose=True)
        return dict_peso_micro

    def peso_macro(self):
        """
        Calcola il peso delle macro asset class di un portafoglio.

        Raises:
            AssertionError = La somma delle macro categorie non fa cento

        Returns:
            dict_peso_macro {dict} = dizionario che associa ad ogni macro il peso relativo
        """
        dict_peso_micro = self.peso_micro()
        dict_peso_macro = {macro : sum(dict_peso_micro[item] for item in micro) for macro, micro in self.dict_macro_micro.items()}
        np.testing.assert_almost_equal(actual=sum(dict_peso_macro.values()), desired=1.00, decimal=3, err_msg="La somma delle macro categorie non fa cento", verbose=True)
        return dict_peso_macro
        
    def peso_strumenti(self):
        """
        Calcola il peso degli strumenti di un portafoglio.

        Raises:
            AssertError = La somma dei pesi degli strumenti non da cento
        
        Returns:
            dict_strumenti {dict} = dizionario che associa ad ogni strumento il peso relativo
        """
        dict_strumenti = {strumento : self.df_portfolio.loc[self.df_portfolio['strumento']==strumento,'controvalore_in_euro'].sum() / self.df_portfolio['controvalore_in_euro'].sum() for strumento in self.strumenti}
        np.testing.assert_almost_equal(actual=sum(dict_strumenti.values()), desired=1.00, decimal=3, err_msg="La somma dei pesi degli strumenti non fa cento", verbose=True)
        return dict_strumenti

    def peso_valuta(self):
        """
        Calcola il peso delle valute considerando la loro scomposizione in mercati per l'amministrato,
        e la loro scomposizione in mercati per il gestito.

        Raises:
            AssertionError = La somma dei pesi delle valute non fa cento

        Returns:
            dict_valute {dict} = dizionario che associa ad ogni valuta il peso relativo
        """
        def peso_valuta_per_denominazione(dataframe):
            """
            Calcola il peso delle valute considerando la loro denominazione.
            
            Parameters:
                dataframe {dataframe} = nome del dataframe da analizzare
            
            Raises:
                AssertionError = La somma dei pesi delle valute per denominazione non fa cento

            Returns:
                dict_valute_denominazione {dict} = dizionario che associa ad ogni valuta il peso relativo
            """
            # df = dataframe if isinstance(dataframe, pd.DataFrame)==True else self.df_portfolio
            df = dataframe
            dict_valute_denominazione = {valuta : df.loc[df['divisa']==valuta, 'controvalore_in_euro'].sum() / df['controvalore_in_euro'].sum() for valuta in self.valute[:-1]}
            dict_valute_denominazione['ALTRO'] = df.loc[~df['divisa'].isin(self.valute[:-1]), 'controvalore_in_euro'].sum() / df['controvalore_in_euro'].sum()
            np.testing.assert_almost_equal(actual=sum(dict_valute_denominazione.values()), desired=1.00, decimal=3, err_msg="La somma dei pesi delle valute per denominazione non fa cento", verbose=True)
            return dict_valute_denominazione
        def peso_valuta_per_composizione(dataframe):
            """
            Calcola il peso delle valute considerando la loro scomposizione in mercati.
            Per i fondi hedged, considera la valuta su cui si basa la strategia di copertura.

            Mappa ogni mercato in valuta così da ricondurre tutti i mercati ad una sola valuta,
            e poi moltiplica la matrice ottenuta per il vettore dei pesi del prodotto.
            
            Parameters:
                dataframe {dataframe} = nome del dataframe da analizzare

            Returns:
                dict_valute_composizione {dict} = dizionario che associa ad ogni valuta il peso relativo
            """
            # df_p = dataframe if isinstance(dataframe, pd.DataFrame)==True else self.df_portfolio
            df_p = dataframe
            vector_peso_prodotti = (df_p['controvalore_in_euro'] / df_p['controvalore_in_euro'].sum()).to_numpy()
            df_m = self.df_mappatura.loc[(self.df_mappatura['ISIN'].isin(list(dataframe['ISIN']))) & (self.df_mappatura['nome'].isin(list(dataframe['nome']))), (*self.micro_asset_class, 'Hedging')] if isinstance(dataframe, pd.DataFrame)==True else self.df_mappatura.loc[:, (*self.micro_asset_class, 'Hedging')]
            dict_valute_per_composizione = self.dict_valute_per_composizione
            dict_valute_per_composizione.update({'Hedging' : 'hedging'})
            df_m.columns = df_m.columns.map(dict_valute_per_composizione) # assegna ad ogni mercato una valuta
            df_m = df_m.groupby(df_m.columns, axis=1).sum() # raggruppa per valuta
            # Hedging
            for valuta in self.valute_per_composizione:
                df_m.loc[df_m['hedging']!=False, valuta] = df_m.loc[df_m['hedging']!=False, 'hedging'].apply(lambda x: 1 if x == valuta else 0)
            df_m.drop('hedging', axis=1, inplace=True)
            df_m = df_m[self.valute_per_composizione]
            matrix_mappatura_valute = df_m.T.to_numpy()
            vector_valute = matrix_mappatura_valute @ vector_peso_prodotti
            dict_valute_composizione = {self.valute_per_composizione[_] : vector_valute[_] for _ in range(len(self.valute_per_composizione))}
            np.testing.assert_almost_equal(actual=np.sum(vector_valute), desired=1.00, decimal=2, err_msg='la somma delle valute per composizione non fa cento', verbose=True)
            return dict_valute_composizione

        df_amministrato = self.df_portfolio.loc[self.df_portfolio['strumento'].isin(self.amministrato)]
        dict_valute_amministrato = peso_valuta_per_denominazione(dataframe=df_amministrato) if df_amministrato.empty==False else {valuta : 0 for valuta in self.valute} # il dizionario vuoto è necessario altrimenti se non c'è amministrato non compaiono le valute differenti da EUR USD e ALTRO
        df_gestito = self.df_portfolio.loc[~self.df_portfolio['strumento'].isin(self.amministrato)]
        dict_valute_gestito = peso_valuta_per_composizione(dataframe=df_gestito) if df_gestito.empty==False else {valuta : 0 for valuta in self.valute_per_composizione}
        dict_amministrato_su_ptf = Counter({key : value*df_amministrato['controvalore_in_euro'].sum()/self.df_portfolio['controvalore_in_euro'].sum() for key, value in dict_valute_amministrato.items()})
        dict_gestito_su_ptf = Counter({key : value*df_gestito['controvalore_in_euro'].sum()/self.df_portfolio['controvalore_in_euro'].sum() for key, value in dict_valute_gestito.items()})
        dict_amministrato_su_ptf.update(dict_gestito_su_ptf) # unione dei due dizionari
        dict_valute = dict(dict_amministrato_su_ptf)
        np.testing.assert_almost_equal(actual=sum(dict_valute.values()), desired=1.00, decimal=2, err_msg='la somma delle valute non fa cento', verbose=True)
        return dict_valute

    def peso_emittente(self):
 
        """
        Calcola il peso dell'emittente dei prodotti di un portafoglio.
        Per i fondi / etf uso la compagnia dei fondi, per i rimanenti prodotti quotati uso il nome, per i prodotti non quotati
        e attivi quali polizze e gestione inserisco il nome a mano, e per prodotti non quotati non attivi uso l'etichetta "altri".

        Raises:
            AssertError = La somma dei pesi dei singoli fondi sul totale dei fondi non fa cento
        
        Returns:
            dict_emittente {dict} = dizionario che associa ad ogni fondo il peso relativo al controvalore totale dei fondi
        """
        df = self.df_portfolio
        try:
            dict_emittente = {emittente : df.loc[df["emittente"]==emittente, "controvalore_in_euro"].sum() / df["controvalore_in_euro"].sum() for emittente in df["emittente"].unique()}
            np.testing.assert_almost_equal(actual=sum(dict_emittente.values()), desired=1.00, decimal=3, err_msg="La somma dei pesi dei singoli fondi sul totale dei fondi non fa cento", verbose=True)
            # print(dict_emittente)
            return dict_emittente
        except KeyError:
                print(f"La colonna dell'emittente non è presente nel portafoglio:\n{list(df.columns)}")
        # Per soli fondi
        # df_ptf_funds = self.df_portfolio.loc[(self.df_portfolio["strumento"]=="fund") | (self.df_portfolio["strumento"]=="etf")]
        # if not df_ptf_funds.empty:
        #     try:
        #         dict_emittente_fondi = {emittente : df_ptf_funds.loc[df_ptf_funds["emittente"]==emittente, "controvalore_in_euro"].sum() / df_ptf_funds["controvalore_in_euro"].sum() for emittente in df_ptf_funds["emittente"].unique()}
        #         np.testing.assert_almost_equal(actual=sum(dict_emittente_fondi.values()), desired=1.00, decimal=3, err_msg="La somma dei pesi dei singoli fondi sul totale dei fondi non fa cento", verbose=True)
        #         return dict_emittente_fondi
        #     except KeyError:
        #         print(f"La colonna dell'emittente non è presente nel portafoglio:\n{list(df_ptf_funds.columns)}")
        # else:
        #     return None

    def matrice_correlazioni(self):
        """
        Creazione della matrice delle correlazioni tra fondi (attivi e passivi)
        Viene fatta in un metodo a parte perchè coinvolge lo scarico di dati in Bloomberg.
        Ho passato un'ora con un esperto di bloomberg per capire quale fosse il campo più adatto ed abbiamo ottenuto
        "DAY_TO_DAY_TOT_RETURN_GROSS_DVDS". Questo campo ricostruisce la storia del fondo, a partire dalla data di inizio analisi
        (nel nostro caso 52 settimane fa), includendo i dividendi e reinvestendoli ai tassi di rendimenti futuri del fondo.
        Quindi è come se i fondi a distribuzione non staccassero mai dividendo. Su questa serie storica otteniamo i rendimenti mensili.
        Per curiosità, un altro campo "CUST_TRR_RETURN_HOLDING_PER" permette addirittura di decidere a quale tasso reinvestire
        i dividendi.
        Ho creato la heatmap delle correlazioni imponendo due regole: se il fondo non possiede 52 osservazioni 
        (leggi: ha meno di un anno di vita) vedrà la sua colonna e riga all'interno della matrice annullata;
        se un fondo non esiste su Bloomberg (ho provato con un privato di credem) non viene nemmeno inserito nella matrice.
        Come parametri della matrice ho messo il minimo e il massimo valore possibile, -1 e +1, i valori della correlazione
        all'interno delle celle e una palette di colori chiamata "turbo" che spazia dal blu al rosso.
        """
        df_ptf_funds = self.df_portfolio.loc[(self.df_portfolio["strumento"]=="fund") | (self.df_portfolio["strumento"]=="etf")]
        if not df_ptf_funds.empty:
            df_ptf_funds_isin = df_ptf_funds['ISIN'].to_list()
            first_day_of_current_month = date.today().replace(day=1)
            last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
            last_day_of_previous_month_of_previous_year = last_day_of_previous_month.replace(year=last_day_of_previous_month.year - 1)

            serie_storica = blp.bdh(['/isin/' + fondo for fondo in df_ptf_funds_isin], flds="DAY_TO_DAY_TOT_RETURN_GROSS_DVDS",
                start_date=last_day_of_previous_month_of_previous_year, end_date=last_day_of_previous_month, Days="A", Period="W")
            serie_storica.columns = [column[0][6:] for column in serie_storica.columns] # rinomina solo i fondi che esistono
            # serie_storica.to_excel('ahah.xlsx')
            # print(serie_storica)
            # print(len(serie_storica.index)) # = 52
            corr_matrix = serie_storica.corr(min_periods=len(serie_storica.index))
            # print(corr_matrix)
            upper_triangle_corr_matrix = np.triu(corr_matrix)
            lower_triangle_corr_matrix = np.tril(corr_matrix)
            plt.figure(figsize=(19.2, 9.7))
            if len(df_ptf_funds_isin) <= 20:
                corr_matrix = corr_matrix.round(3)
                sns.heatmap(data=corr_matrix, vmin=-1, vmax=+1, annot=True, annot_kws={'fontsize':14}, cmap="turbo")#, mask=upper_triangle_corr_matrix)
                plt.yticks(fontsize=14)
                plt.xticks(rotation=75, fontsize=14)
            elif len(df_ptf_funds_isin) <= 35:
                corr_matrix = corr_matrix.round(2)
                sns.heatmap(data=corr_matrix, vmin=-1, vmax=+1, annot=True, annot_kws={'fontsize':10}, cmap="turbo")#, mask=upper_triangle_corr_matrix)
                plt.yticks(fontsize=12)
                plt.xticks(rotation=75, fontsize=12)
            else:
                sns.heatmap(data=corr_matrix, vmin=-1, vmax=+1, annot=False, cmap="turbo")#, mask=upper_triangle_corr_matrix)
                plt.yticks(fontsize=9)
                plt.xticks(rotation=75, fontsize=9)
            plt.tight_layout()
            plt.savefig('img/matr_corr.png')
            return True
        else:
            return None

    def duration(self):
        """
            Calcola le duration per i comparti obbligazionari, non considera le duration ND.
            Se il portafoglio non possiede titoli obbligazionari restuisce None.

            Returns:
                durations {dict} = dizionario che associa ad ogni classe obbligazionaria la relativa duration
        """
        if self.df_portfolio[(self.df_portfolio['strumento']=='gov_bond') | (self.df_portfolio['strumento']=='corp_bond')].empty:
            return None
        else:
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
    start = time.perf_counter()
    _ = Portfolio(intermediario='azimut')
    _.test()
    _.peso_micro()
    _.peso_macro()
    _.peso_strumenti()
    _.peso_valuta()
    _.peso_emittente()
    _.duration()
    _.risk()
    end = time.perf_counter()
    print("Elapsed time: ", round(end - start, 2), 'seconds')