import time
from collections import Counter
import numpy as np
from numpy.matrixlib.defmatrix import matrix
from numpy.testing._private.utils import assert_almost_equal
import pandas as pd
import matplotlib.pyplot as plt
import excel2img
from openpyxl import Workbook # Per creare un libro
from openpyxl import load_workbook # Per caricare un libro
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font, numbers # Per cambiare lo stile
from openpyxl.styles.numbers import FORMAT_PERCENTAGE_00, FORMAT_NUMBER_00, FORMAT_NUMBER_COMMA_SEPARATED1 # Stili di numeri
from openpyxl.utils import get_column_letter # Per lavorare sulle colonne
from openpyxl.drawing.text import Paragraph, ParagraphProperties, CharacterProperties
from openpyxl.chart import BarChart, LineChart, PieChart, Reference, DoughnutChart
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.text import RichText
from openpyxl.chart.marker import DataPoint
from docx import Document
from docx import shared
from SAP import Portfolio


class Elaborazione(Portfolio):
    """Elabora un portafoglio."""

    def __init__(self, file_elaborato):
        """
        Initialize the class.

        Parameters:
        file_elaborato(str) = file elaborato
        """
        super().__init__(file_portafoglio=PTF, path=PATH)
        self.wb = load_workbook(self.file_portafoglio)
        self.portfolio = self.wb['portfolio']
        self.file_elaborato = file_elaborato
      
    def new_agglomerato(self):
        """
        Crea un agglomerato del portafoglio diviso per tipo di strumento.

        Parameters:
        limite(int) = limite di strumenti per pagina. Dipende dalla lunghezza della pagina word in cui verrà incollata.
        """
        # Dataframe del portfolio
        df = self.df_portfolio
        controvalori = {strumento : df.loc[df['strumento']==strumento, 'controvalore_in_euro'].sum() for strumento in df['strumento'].unique()}
        # Dizionario che associa ai tipi di strumenti presenti in portafoglio un loro nome in italiano.
        strumenti_dict = {'cash' : 'LIQUIDITÀ', 'gov_bond' : 'OBBLIGAZIONI GOVERNATIVE', 'corp_bond' : 'OBBLIGAZIONI CORPORATE', 'certificate' : 'CERTIFICATI', 'equity' : 'AZIONI',
            'etf' : 'ETF', 'fund' : 'FONDI', 'real_estate' : 'REAL_ESTATE', 'hedge_fund' : 'HEDGE FUND', 'insurance' : 'POLIZZE', 'gp' : 'GESTIONI', 'pip' : 'FONDI PENSIONE'}
        # Lista della numerosità degli strumenti in portafoglio
        c = Counter(list(df.loc[:, 'strumento']))
        # print(c)
        # Header
        ws = self.wb.create_sheet('agglomerato')
        ws = self.wb['agglomerato']
        self.wb.active = ws
        header = ['ISIN', 'Descrizione', 'Quantità', 'Controvalore iniziale', 'Prezzo di carico', 'Divisa', 'Prezzo di mercato in euro', 'Rateo', 'Valore di mercato in euro']
        len_header = len(header)
        for col in ws.iter_cols(min_row=1, max_row=1, min_col=1, max_col=len_header):
            ws[col[0].coordinate].value = header[0]
            del header[0]
            ws[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws[col[0].coordinate].font = Font(name='Century Gothic', size=18, color='FFFFFF', bold=True)
            ws[col[0].coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'))
            ws[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
            ws.row_dimensions[col[0].row].height = 92.55
            if ws[col[0].coordinate].value == 'ISIN':
                ws.column_dimensions[col[0].column_letter].width = 25
            elif ws[col[0].coordinate].value == 'Descrizione':
                ws.column_dimensions[col[0].column_letter].width = max([len(nome) for nome in df['nome'].values])*2
            elif ws[col[0].coordinate].value == 'Quantità':
                ws.column_dimensions[col[0].column_letter].width = max([len(str(round(quantità,2))) for quantità in df['quantità'].values])*2.5
            elif ws[col[0].coordinate].value == 'Controvalore iniziale':
                ws.column_dimensions[col[0].column_letter].width = max(23, max([len(str(round(controvalore_iniziale,2))) for controvalore_iniziale in df['controvalore_iniziale'].values])*2.5)
            elif ws[col[0].coordinate].value == 'Prezzo di carico':
                ws.column_dimensions[col[0].column_letter].width = max([len(str(round(prezzo_di_carico,2))) for prezzo_di_carico in df['prezzo_di_carico'].values])*2.5
            elif ws[col[0].coordinate].value == 'Divisa':
                ws.column_dimensions[col[0].column_letter].width = 12
            elif ws[col[0].coordinate].value == 'Prezzo di mercato in euro':
                ws.column_dimensions[col[0].column_letter].width = 21.29
            elif ws[col[0].coordinate].value == 'Rateo':
                ws.column_dimensions[col[0].column_letter].width = 11.29
            elif ws[col[0].coordinate].value == 'Valore di mercato in euro':
                ws.column_dimensions[col[0].column_letter].width = 30.43
        # Body
        min_row = 2
        max_row = 1
        min_col = 1
        max_col = len_header
        strumenti_in_ptf = [strumento for strumento in self.strumenti if c[strumento] > 0]
        for strumento in strumenti_in_ptf:
            max_row = max_row + c[strumento] + 1
            for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                # Etichetta
                if row[0].row == min_row:
                    ws[row[0].coordinate].value = strumenti_dict[strumento]
                    ws[row[0].coordinate].font = Font(name='Century Gothic', size=18, color='808080', bold=True)
                    ws[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='F2F2F2')
                    ws[row[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                    ws[row[0].coordinate].border = Border(top=Side(border_style='double', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=min_col, end_column=max_col-1)
                    ws[row[max_col-1].coordinate].value = controvalori[strumento]
                    ws[row[max_col-1].coordinate].font = Font(name='Century Gothic', size=18, color='808080', bold=True)
                    ws[row[max_col-1].coordinate].fill = PatternFill(fill_type='solid', fgColor='F2F2F2')
                    ws[row[max_col-1].coordinate].alignment = Alignment(horizontal='right', vertical='center')
                    ws[row[max_col-1].coordinate].border = Border(top=Side(border_style='double', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                    ws[row[max_col-1].coordinate].number_format = '€ #,0.00'
                    ws.row_dimensions[row[0].row].height = 27
                # Strumenti
                else:
                    for _ in range(0, c[strumento]):
                        ws[row[0].offset(row=_, column=len_header-9).coordinate].value = df.loc[df['strumento']==strumento, 'ISIN'].values[_]
                        ws[row[0].offset(row=_, column=len_header-9).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-9).coordinate].alignment = Alignment(horizontal='left', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-9).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-8).coordinate].value = df.loc[df['strumento']==strumento, 'nome'].values[_]
                        ws[row[0].offset(row=_, column=len_header-8).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-8).coordinate].alignment = Alignment(horizontal='left', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-8).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-7).coordinate].value = df.loc[df['strumento']==strumento, 'quantità'].values[_]
                        ws[row[0].offset(row=_, column=len_header-7).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-7).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-7).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-7).coordinate].number_format = '#,0.00'
                        ws[row[0].offset(row=_, column=len_header-6).coordinate].value = df.loc[df['strumento']==strumento, 'controvalore_iniziale'].values[_]
                        ws[row[0].offset(row=_, column=len_header-6).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-6).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-6).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-6).coordinate].number_format = '#,0.00'
                        ws[row[0].offset(row=_, column=len_header-5).coordinate].value = df.loc[df['strumento']==strumento, 'prezzo_di_carico'].values[_]
                        ws[row[0].offset(row=_, column=len_header-5).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-5).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-5).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-5).coordinate].number_format = '#,0.00'
                        ws[row[0].offset(row=_, column=len_header-4).coordinate].value = df.loc[df['strumento']==strumento, 'divisa'].values[_]
                        ws[row[0].offset(row=_, column=len_header-4).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-4).coordinate].alignment = Alignment(horizontal='center', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-4).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-3).coordinate].value = df.loc[df['strumento']==strumento, 'prezzo'].values[_]
                        ws[row[0].offset(row=_, column=len_header-3).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-3).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-3).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-3).coordinate].number_format = '#,0.00'
                        ws[row[0].offset(row=_, column=len_header-2).coordinate].value = df.loc[df['strumento']==strumento, 'rateo'].values[_]
                        ws[row[0].offset(row=_, column=len_header-2).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-2).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-2).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-2).coordinate].number_format = '#,0.00'
                        ws[row[0].offset(row=_, column=len_header-1).coordinate].value = df.loc[df['strumento']==strumento, 'controvalore_in_euro'].values[_]
                        ws[row[0].offset(row=_, column=len_header-1).coordinate].font = Font(name='Century Gothic', size=18, color='000000')
                        ws[row[0].offset(row=_, column=len_header-1).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=_, column=len_header-1).coordinate].border = Border(top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), left=Side(border_style='dotted', color='000000'), right=Side(border_style='dotted', color='000000'))
                        ws[row[0].offset(row=_, column=len_header-1).coordinate].number_format = '#,0.00'
                        ws.row_dimensions[row[0].row].height = 23.25
                    min_row = max_row + 1
                    break
        # Footer
        max_row = min_row
        for col in ws.iter_cols(min_row=max_row, max_row=max_row, min_col=1, max_col=len_header):
            if col[0].column == min_col:
                ws[col[0].coordinate].value = 'TOTALE PORTAFOGLIO'
                ws[col[0].coordinate].font = Font(name='Century Gothic', size=18, color='FFFFFF', bold=True)
                ws[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
                ws[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
                ws[col[0].coordinate].border = Border(top=Side(border_style='double', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
            if col[0].column == len_header:
                ws[col[0].coordinate].value = sum(controvalori.values())
                ws[col[0].coordinate].font = Font(name='Century Gothic', size=18, color='FFFFFF', bold=True)
                ws[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
                ws[col[0].coordinate].alignment = Alignment(horizontal='right', vertical='center')
                ws[col[0].coordinate].border = Border(top=Side(border_style='double', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                ws[col[0].coordinate].number_format = '€ #,0.00'
                ws.row_dimensions[col[0].row].height = 27
            ws.merge_cells(start_row=col[0].row, end_row=col[0].row, start_column=1, end_column=len_header-1)

    def old_agglomerato(self, limite=57):
        """
        Crea un agglomerato del portafoglio diviso per tipo di strumento.

        Parameters:
        limite(int) = limite di strumenti per pagina. Dipende dalla lunghezza della pagina word in cui verrà incollata.
        """
        # Dataframe del portfolio
        df = pd.read_excel(self.file_portafoglio, sheet_name='portfolio_valori')
        '''Devo dichiarare il controvalore di ogni tipo di strumento al'inizio perché poi se uno strumento viene diviso in più pagine, il controvalore nella pagina successiva
        sarà inferiore, dato che un pezzo di dataframe è stato eliminato'''
        controvalori = {strumento : df.loc[df['strumento']==strumento, 'controvalore_in_euro'].sum() for strumento in df['strumento'].unique()}
        # Dizionario che associa ai tipi di strumenti presenti in portafoglio un loro nome in italiano.
        strumenti_dict = {'cash' : 'LIQUIDITÀ', 'gov_bond' : 'OBBLIGAZIONI GOVERNATIVE', 'corp_bond' : 'OBBLIGAZIONI CORPORATE', 'certificate' : 'CERTIFICATI', 'equity' : 'AZIONI',
            'etf' : 'ETF', 'fund' : 'FONDI', 'real_estate' : 'REAL_ESTATE', 'hedge_fund' : 'HEDGE FUND', 'insurance' : 'POLIZZE', 'gp' : 'GESTIONI', 'pip' : 'FONDI PENSIONE'}
        # Lista di strumenti unici in portafoglio
        strumenti_in_ptf = list(df.loc[:, 'strumento'].unique())
        # print(f"Il portafoglio possiede i seguenti strumenti: {strumenti_in_ptf}.")
        numero_prodotti = self.portfolio.max_row - 1 # lunghezza del portfolio
        print(f"Il portafoglio contiene {numero_prodotti} prodotti.\n")
        #len_ptf = len(self.portfolio['B']))
        # Numero di fogli da creare
        if (numero_prodotti + len(list(df['strumento'].unique()))) >= limite:
            fogli = (numero_prodotti + len(list(df['strumento'].unique()))) // limite + 1 # perchè avevo messo + 2?
        else:
            fogli = numero_prodotti // limite + 1
        print(f"sto creando {fogli} fogli...")
        for foglio in range(1,fogli+1):
            # Header
            header = ['ISIN', 'Descrizione', 'Quantità V.N.', 'V.ACQ.',	'PREZZO MEDIO ACQ',	'Divisa', 'Corso secco/prezzo di mercato in Euro', 'Rateo', 'Valore di mercato in Euro']
            len_header = len(header)
            # Nome fogli
            ws = 'ws_'+str(foglio)
            # print(ws)
            # Creazione foglio
            ws = self.wb.create_sheet('agglomerato_'+str(foglio))
            ws = self.wb['agglomerato_'+str(foglio)]
            self.wb.active = ws
            # Header
            for col in ws.iter_cols(min_row=1, max_row=1, min_col=1, max_col=len_header):
                ws[col[0].coordinate].value = header[0]
                del header[0]
                ws[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                ws[col[0].coordinate].font = Font(name='Calibri', size=18, color='FFFFFF', bold=True)
                ws[col[0].coordinate].border = Border(top=Side(border_style='medium', color='000000'), bottom=Side(border_style='medium', color='000000'), right=Side(border_style='medium', color='000000'), left=Side(border_style='medium', color='000000'))
                ws[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
                ws.row_dimensions[col[0].row].height = 92.55
                if ws[col[0].coordinate].value == 'ISIN':
                    ws.column_dimensions[col[0].column_letter].width = 23
                elif ws[col[0].coordinate].value == 'Descrizione':
                    ws.column_dimensions[col[0].column_letter].width = 70 # Calcola la lunghezza massima della colonna
                elif ws[col[0].coordinate].value == 'Quantità V.N.':
                    ws.column_dimensions[col[0].column_letter].width = 17
                elif ws[col[0].coordinate].value == 'V.ACQ.':
                    ws.column_dimensions[col[0].column_letter].width = 17
                elif ws[col[0].coordinate].value == 'PREZZO MEDIO ACQ':
                    ws.column_dimensions[col[0].column_letter].width = 17
                elif ws[col[0].coordinate].value == 'Divisa':
                    ws.column_dimensions[col[0].column_letter].width = 9.71
                elif ws[col[0].coordinate].value == 'Corso secco/prezzo di mercato in Euro':
                    ws.column_dimensions[col[0].column_letter].width = 21.29
                elif ws[col[0].coordinate].value == 'Rateo':
                    ws.column_dimensions[col[0].column_letter].width = 11.29
                elif ws[col[0].coordinate].value == 'Valore di mercato in Euro':
                    ws.column_dimensions[col[0].column_letter].width = 30.43
            # Body
            min_row = 2
            max_row = limite
            min_col = 1
            riga = min_row
            # print(f"La riga di partenza è la numero {riga}.")
            for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=len_header):
                for strumento in (strumento for strumento in strumenti_dict if strumento in strumenti_in_ptf):
                    # controlla quanti prodotti strumento ci sono nel portafoglio e se c'è spazio a sufficienza
                    # print(f"Ci sono {df.loc[df['strumento']==strumento, 'nome'].count()} prodotti di {strumento}, e {max_row - riga} spazi disponibili") # l'ultimo è destinato eventualmente al totale
                    # print(f"e {max_row - riga} spazi disponibili.")
                    if df.loc[df['strumento']==strumento, 'nome'].count() < (max_row-riga):
                        # print(f"quindi posso inserire tutti i prodotti {strumento} e la label")
                        # print(f"sto inserendo la label {strumento} che occupa una riga...")
                        ws[row[0].offset(row=riga-min_row).coordinate].value = strumenti_dict[strumento]
                        ws[row[0].offset(row=riga-min_row).coordinate].font = Font(name='Calibri', size=18, color='808080', bold=True)
                        ws[row[0].offset(row=riga-min_row).coordinate].fill = PatternFill(fill_type='solid', fgColor='F2F2F2')
                        ws[row[0].offset(row=riga-min_row).coordinate].alignment = Alignment(horizontal='center', vertical='center')
                        ws[row[0].offset(row=riga-min_row).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                        ws.merge_cells(start_row=row[0].offset(row=riga-min_row).row, end_row=row[0].offset(row=riga-min_row).row, start_column=min_col, end_column=len_header-1)
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].value = controvalori[strumento]
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].font = Font(name='Calibri', size=18, color='808080', bold=True)
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].fill = PatternFill(fill_type='solid', fgColor='F2F2F2')
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].number_format = '€ #,0.00'
                        ws.row_dimensions[row[0].offset(row=riga-min_row).row].height = 27
                        strumenti_in_ptf.remove(strumento)
                        riga = riga + 1
                        # print(f"Riga attuale : {riga}")
                        # print(f"sto inserendo i prodotti {strumento} (numerosità:{df.loc[df['strumento']==strumento, 'nome'].count()})...")
                        for _ in range(0,df.loc[df['strumento']==strumento, 'nome'].count()):
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].value = df.loc[df['strumento']==strumento, 'ISIN'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].alignment = Alignment(horizontal='left', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].value = df.loc[df['strumento']==strumento, 'nome'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].alignment = Alignment(horizontal='left', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].value = df.loc[df['strumento']==strumento, 'quantità'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].value = df.loc[df['strumento']==strumento, 'controvalore_iniziale'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].value = df.loc[df['strumento']==strumento, 'prezzo_di_carico'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].value = df.loc[df['strumento']==strumento, 'divisa'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].alignment = Alignment(horizontal='center', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].value = df.loc[df['strumento']==strumento, 'prezzo'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].value = df.loc[df['strumento']==strumento, 'rateo'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].value = df.loc[df['strumento']==strumento, 'controvalore_in_euro'].values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].number_format = '#,0.00'
                            ws.row_dimensions[row[0].offset(row=riga-min_row).row].height = 23.25
                        riga = riga + df.loc[df['strumento']==strumento, 'nome'].count()
                        # print(f"Riga attuale : {riga}")
                    else:
                        # print(f"quindi posso inserire solo {(max_row-riga)} prodotti oltre la label.")
                        if riga >= max_row-1: # altrimenti inserirebbe la label del prodotto ma senza alcun prodotto sotto di essa
                            # print("ma non c'è spazio per i prodotti sotto la label e quindi vado a pagina nuova")
                            break
                        # print(f"sto inserendo la label {strumento} che occupa una riga...")
                        ws[row[0].offset(row=riga-min_row).coordinate].value = strumenti_dict[strumento]
                        ws[row[0].offset(row=riga-min_row).coordinate].font = Font(name='Calibri', size=18, color='808080', bold=True)
                        ws[row[0].offset(row=riga-min_row).coordinate].fill = PatternFill(fill_type='solid', fgColor='F2F2F2')
                        ws[row[0].offset(row=riga-min_row).coordinate].alignment = Alignment(horizontal='center', vertical='center')
                        ws[row[0].offset(row=riga-min_row).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                        ws.merge_cells(start_row=row[0].offset(row=riga-min_row).row, end_row=row[0].offset(row=riga-min_row).row, start_column=min_col, end_column=len_header-1)
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].value = controvalori[strumento]
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].font = Font(name='Calibri', size=18, color='808080', bold=True)
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].fill = PatternFill(fill_type='solid', fgColor='F2F2F2')
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                        ws[row[0].offset(row=riga-min_row, column=len_header-1).coordinate].number_format = '€ #,0.00'
                        ws.row_dimensions[row[0].offset(row=riga-min_row).row].height = 27
                        riga = riga + 1
                        # print(f"Riga attuale : {riga}")
                        # print(f"sto inserendo i prodotti {strumento}...")
                        for _ in range(0, max_row-riga):
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].value = df.loc[df['strumento']==strumento, 'ISIN'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].alignment = Alignment(horizontal='left', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-9).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].value = df.loc[df['strumento']==strumento, 'nome'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].alignment = Alignment(horizontal='left', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-8).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].value = df.loc[df['strumento']==strumento, 'quantità'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-7).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].value = df.loc[df['strumento']==strumento, 'controvalore_iniziale'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-6).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].value = df.loc[df['strumento']==strumento, 'prezzo_di_carico'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-5).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].value = df.loc[df['strumento']==strumento, 'divisa'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].alignment = Alignment(horizontal='center', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-4).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].value = df.loc[df['strumento']==strumento, 'prezzo'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-3).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].value = df.loc[df['strumento']==strumento, 'rateo'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-2).coordinate].number_format = '#,0.00'
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].value = df.loc[df['strumento']==strumento, 'controvalore_in_euro'].head(max_row-riga).values[_]
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].font = Font(name='Calibri', size=18, color='000000')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].alignment = Alignment(horizontal='right', vertical='center')
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
                            ws[row[0].offset(row=riga-min_row+_, column=len_header-1).coordinate].number_format = '#,0.00'
                            ws.row_dimensions[row[0].offset(row=riga-min_row).row].height = 23.25
                        df.drop(df.loc[df['strumento']==strumento].head(max_row-riga).index, inplace=True)
                        # print(f"prodotti di tipo {strumento} rimanenti:\n {df.loc[df['strumento']==strumento]}")
                        break # esci dal ciclo
                break # cambia foglio
        # # Footer
        # ws = self.wb['agglomerato_'+str(fogli)] # ultimo foglio
        # ultima_riga = ws.max_row # ultima riga
        # for row in ws.iter_rows(min_row=ultima_riga+1, max_row=ultima_riga+1, min_col=1, max_col=len_header):
        #     ws[row[0].coordinate].value = 'TOTALE PORTAFOGLIO'
        #     ws[row[0].coordinate].font = Font(name='Calibri', size=18, color='FFFFFF', bold=True)
        #     ws[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
        #     ws[row[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
        #     ws[row[0].coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
        #     ws.merge_cells(start_row=row[0].row, end_row=row[0].row, start_column=1, end_column=len_header-1)
        #     ws[row[0].offset(column=len_header-1).coordinate].value = sum(controvalori.values())
        #     ws[row[0].offset(column=len_header-1).coordinate].font = Font(name='Calibri', size=18, color='FFFFFF', bold=True)
        #     ws[row[0].offset(column=len_header-1).coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
        #     ws[row[0].offset(column=len_header-1).coordinate].alignment = Alignment(horizontal='right', vertical='center')
        #     ws[row[0].offset(column=len_header-1).coordinate].border = Border(top=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'))
        #     ws[row[0].offset(column=len_header-1).coordinate].number_format = '€ #,0.00'
        #     ws.row_dimensions[row[0].row].height = 27
        
    def figure(self):
        """Crea le tabelle e le figure delle micro categorie, delle macro categorie, degli strumenti e delle valute."""

        # Creazione foglio figure
        ws_figure = self.wb.create_sheet('figure')
        ws_figure = self.wb['figure']
        self.wb.active = ws_figure

        # Macro asset class #
        dict_peso_macro = self.peso_macro()

        #---Tabella macro asset class---#
        # Header
        header_macro = ['MACRO ASSET CLASS', '', 'Peso']
        dim_macro = [3.4, 47, 9.5]
        min_row, max_row, min_col, max_col = 1, 1, 1, 3
        for col in ws_figure.iter_cols(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[col[0].coordinate].value = header_macro[col[0].column-min_col]
            ws_figure[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws_figure[col[0].coordinate].font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
            ws_figure[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
            ws_figure[col[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure.column_dimensions[ws_figure[col[0].coordinate].column_letter].width = dim_macro[col[0].column-min_col]
        ws_figure.merge_cells(start_row=min_row, end_row=max_row, start_column=min_col, end_column=min_col+1)
        # Body
        fonts_macro = ['B1A0C7', '92CDDC', 'F79646', 'EDF06A']
        min_row = min_row + 1
        max_row = min_row + len(self.macro_asset_class) - 1
        for row in ws_figure.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor=fonts_macro[row[0].row-min_row])
            ws_figure[row[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[1].coordinate].value = self.macro_asset_class[row[0].row-min_row]
            ws_figure[row[1].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[2].coordinate].value = dict_peso_macro[ws_figure[row[1].coordinate].value]
            ws_figure[row[2].coordinate].number_format = '0.0%'
            ws_figure[row[2].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[row[2].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        # Footer
        ws_figure.cell(max_row+1, min_col, value='TOTALE')
        ws_figure.cell(max_row+1, min_col).alignment = Alignment(horizontal='center', vertical='center')
        ws_figure.cell(max_row+1, min_col).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row+1, min_col).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row+1, min_col).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_figure.merge_cells(start_row=max_row+1, end_row=max_row+1, start_column=min_col, end_column=max_col-1)
        ws_figure.cell(max_row+1, max_col, value=sum(dict_peso_macro.values()))
        assert sum(ws_figure.cell(i, max_col).value for i in range(min_row, max_row+1)) == ws_figure.cell(max_row+1, max_col).value
        ws_figure.cell(max_row+1, max_col).alignment = Alignment(horizontal='center', vertical='center')
        ws_figure.cell(max_row+1, max_col).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row+1, max_col).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row+1, max_col).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_figure.cell(max_row+1, max_col).number_format = FORMAT_PERCENTAGE_00

        # Grafico macro asset class
        chart = PieChart()
        chart.height = 4.77
        chart.width = 6.77
        labels = Reference(ws_figure, min_col=min_col+1, max_col=min_col+1, min_row=min_row, max_row=max_row)
        data = Reference(ws_figure, min_col=min_col+2, max_col=min_col+2, min_row=min_row, max_row=max_row)
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(labels)
        chart.dataLabels = DataLabelList(dLblPos='bestFit')
        chart.dataLabels.showVal = True
        chart.dataLabels.textProperties = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1100, b=True)), endParaRPr=CharacterProperties(sz=1100, b=True))])
        chart.legend = None
        chart_fonts = ['B1A0C7', '92CDDC', 'F79646', 'EDF06A'] # cambia colori delle fette
        for _ in range(0,4):
            series = chart.series[0]
            pt = DataPoint(idx=_)
            pt.graphicalProperties.solidFill = chart_fonts[_]
            series.dPt.append(pt)
        chart.layout = Layout(manualLayout=ManualLayout(x=0.5, y=0.5, h=1, w=1)) # posizione e dimensione figura
        ws_figure.add_chart(chart, 'D1')
        # Grafico macro
        plt.subplots(figsize=(4,4))
        try:
            plt.pie([dict_peso_macro[_] for _ in self.macro_asset_class], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.01 else '' for value in dict_peso_macro.values()], radius=1.2, colors=['#B1A0C7', '#92CDDC', '#F79646', '#EDF06A'], pctdistance=0.1, labeldistance=0.4, textprops={'fontsize':14}, normalize=False)
        except ValueError:
            plt.pie([dict_peso_macro[_] for _ in self.macro_asset_class], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.01 else '' for value in dict_peso_macro.values()], radius=1.2, colors=['#B1A0C7', '#92CDDC', '#F79646', '#EDF06A'], pctdistance=0.1, labeldistance=0.4, textprops={'fontsize':14}, normalize=True)
        finally:
            plt.savefig('Media/macro_pie.png', bbox_inches='tight', pad_inches=0)

        # Micro asset class #
        dict_peso_micro = self.peso_micro()
        # Durations #
        durations = self.duration()
        
        #---Tabella micro asset class---#
        # Header
        header_micro = ['', 'ASSET CLASS', 'Indice', 'Peso', 'Warning', 'Duration']
        dim_micro = [3.4, 16, 57, 9.5, 9.5, 9.5]
        min_row, max_row, min_col, max_col = 1, 1, 9, 14
        for col in ws_figure.iter_cols(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[col[0].coordinate].value = header_micro[col[0].column-min_col]
            ws_figure[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws_figure[col[0].coordinate].font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
            ws_figure[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
            ws_figure[col[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure.column_dimensions[ws_figure[col[0].coordinate].column_letter].width = dim_micro[col[0].column-min_col]
        # Body
        fonts_micro = ['E4DFEC', 'CCC0DA', 'B1A0C7', '92CDDC', '00B0F0', '0033CC', '0070C0', '1F497D', '000080', 'F79646', 'FFCC66', 'DA5300', 'F62F00', 'EDF06A']
        list_macro = ['Monetario', '', '', 'Obbligazionario', '', '', '', '', '', 'Azionario', '', '', '', 'Commodities']
        indici_micro = ['The BofA Merrill Lynch 0-1 Year Euro Government Index', 'The BofA Merrill Lynch 0-1 Year US Treasury Index', 
            'The BofA Merrill Lynch 0-1 Year G7 Government Index', 'The BofA Merrill Lynch Euro Broad Market Index', 'The BofA Merrill Lynch Euro Large Cap Corporate Index',
            'The BofA Merrill Lynch Euro High Yield Index', 'The BofA Merrill Lynch Global Broad Market Index', 'The BofA Merrill Lynch Global EM Sovereign & Credit Plus Index',
            'The BofA Merrill Lynch Global High Yield Index', 'MSCI Europe', 'MSCI North America', 'MSCI Pacific', 'MSCI Emerging Markets', 'Thomson Reuters/CoreCommodity CRB Commodity Index',
            ]
        min_row = min_row + 1
        max_row = min_row + len(self.micro_asset_class) - 1
        list_peso_micro = list(dict_peso_micro.values())
        for row in ws_figure.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor=fonts_micro[row[0].row-min_row])
            ws_figure[row[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[1].coordinate].value = list_macro[row[0].row-min_row]
            ws_figure[row[1].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws_figure[row[1].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[2].coordinate].value = indici_micro[row[0].row-min_row]
            ws_figure[row[2].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[3].coordinate].value = tuple(dict_peso_micro.values())[row[0].row-min_row]
            ws_figure[row[3].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[3].coordinate].number_format = '0.0%'
            ws_figure[row[3].coordinate].alignment = Alignment(horizontal='center')
            # warnings
            if ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Euro Large Cap Corporate Index':
                if list_peso_micro[4]/dict_peso_macro['Obbligazionario'] > 0.6:
                    ws_figure[row[4].coordinate].value = '!!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[4]/dict_peso_macro['Obbligazionario'] > 0.4:
                    ws_figure[row[4].coordinate].value = '!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[4]/dict_peso_macro['Obbligazionario'] > 0.2:
                    ws_figure[row[4].coordinate].value = '!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Euro High Yield Index' or ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Global High Yield Index':
                if (list_peso_micro[5]+list_peso_micro[8])/dict_peso_macro['Obbligazionario'] > 0.4:
                    ws_figure[row[4].coordinate].value = '!!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif (list_peso_micro[5]+list_peso_micro[8])/dict_peso_macro['Obbligazionario'] > 0.3:
                    ws_figure[row[4].coordinate].value = '!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif (list_peso_micro[5]+list_peso_micro[8])/dict_peso_macro['Obbligazionario'] > 0.2:
                    ws_figure[row[4].coordinate].value = '!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Global EM Sovereign & Credit Plus Index':
                if list_peso_micro[7]/dict_peso_macro['Obbligazionario'] > 0.4:
                    ws_figure[row[4].coordinate].value = '!!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[7]/dict_peso_macro['Obbligazionario'] > 0.3:
                    ws_figure[row[4].coordinate].value = '!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[7]/dict_peso_macro['Obbligazionario'] > 0.2:
                    ws_figure[row[4].coordinate].value = '!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[2].coordinate].value == 'MSCI Europe':
                if list_peso_micro[9]/dict_peso_macro['Azionario'] > 0.8:
                    ws_figure[row[4].coordinate].value = '!!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[9]/dict_peso_macro['Azionario'] > 0.7:
                    ws_figure[row[4].coordinate].value = '!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[9]/dict_peso_macro['Azionario'] > 0.6:
                    ws_figure[row[4].coordinate].value = '!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[2].coordinate].value == 'MSCI North America':
                if list_peso_micro[10]/dict_peso_macro['Azionario'] > 0.8:
                    ws_figure[row[4].coordinate].value = '!!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[10]/dict_peso_macro['Azionario'] > 0.7:
                    ws_figure[row[4].coordinate].value = '!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[10]/dict_peso_macro['Azionario'] > 0.6:
                    ws_figure[row[4].coordinate].value = '!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[2].coordinate].value == 'MSCI Pacific':
                if list_peso_micro[11]/dict_peso_macro['Azionario'] > 0.4:
                    ws_figure[row[4].coordinate].value = '!!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[11]/dict_peso_macro['Azionario'] > 0.3:
                    ws_figure[row[4].coordinate].value = '!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[11]/dict_peso_macro['Azionario'] > 0.2:
                    ws_figure[row[4].coordinate].value = '!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[2].coordinate].value == 'MSCI Emerging Markets':
                if list_peso_micro[11]/dict_peso_macro['Azionario'] > 0.3:
                    ws_figure[row[4].coordinate].value = '!!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[11]/dict_peso_macro['Azionario'] > 0.2:
                    ws_figure[row[4].coordinate].value = '!!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif list_peso_micro[11]/dict_peso_macro['Azionario'] > 0.1:
                    ws_figure[row[4].coordinate].value = '!C'
                    ws_figure[row[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            ws_figure[row[4].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[4].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[row[4].coordinate].font = Font(color='000000', bold=True)
            if ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Euro Broad Market Index':
                ws_figure[row[5].coordinate].value = round(durations['Obbligazionario Euro Governativo All Maturities'], 2) if durations['Obbligazionario Euro Governativo All Maturities'] > 0.00 else None
            elif ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Euro Large Cap Corporate Index':
                ws_figure[row[5].coordinate].value = round(durations['Obbligazionario Euro Corporate'], 2) if durations['Obbligazionario Euro Corporate'] > 0.00 else None
            elif ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Euro High Yield Index':
                ws_figure[row[5].coordinate].value = round(durations['Obbligazionario Euro High Yield'], 2) if durations['Obbligazionario Euro High Yield'] > 0.00 else None
            elif ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Global Broad Market Index':
                ws_figure[row[5].coordinate].value = round(durations['Obbligazionario Globale Aggregate'], 2) if durations['Obbligazionario Globale Aggregate'] > 0.00 else None
            elif ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Global EM Sovereign & Credit Plus Index':
                ws_figure[row[5].coordinate].value = round(durations['Obbligazionario Paesi Emergenti'], 2) if durations['Obbligazionario Paesi Emergenti'] > 0.00 else None
            elif ws_figure[row[2].coordinate].value == 'The BofA Merrill Lynch Global High Yield Index':
                ws_figure[row[5].coordinate].value = round(durations['Obbligazionario Globale High Yield'], 2) if durations['Obbligazionario Globale High Yield'] > 0.00 else None
            ws_figure[row[5].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[row[5].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        # hard coding : deve ripeterlo tante volte quanti sono gli oggetti in self.dict_macro
        start_row = min_row
        end_row = min_row + len(self.dict_macro[ws_figure.cell(row=start_row, column=min_col+1).value]) - 1
        ws_figure.merge_cells(start_row=start_row, end_row=end_row, start_column=min_col+1, end_column=min_col+1)
        start_row = end_row + 1
        end_row = start_row + len(self.dict_macro[ws_figure.cell(row=start_row, column=min_col+1).value]) - 1
        ws_figure.merge_cells(start_row=start_row, end_row=end_row, start_column=min_col+1, end_column=min_col+1)
        start_row = end_row + 1
        end_row = start_row + len(self.dict_macro[ws_figure.cell(row=start_row, column=min_col+1).value]) - 1
        ws_figure.merge_cells(start_row=start_row, end_row=end_row, start_column=min_col+1, end_column=min_col+1)
        start_row = end_row + 1
        end_row = start_row + len(self.dict_macro[ws_figure.cell(row=start_row, column=min_col+1).value]) - 1
        ws_figure.merge_cells(start_row=start_row, end_row=end_row, start_column=min_col+1, end_column=min_col+1)
        # Footer
        max_row = max_row + 1
        for col in ws_figure.iter_rows(min_row=max_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[col[0].coordinate].value = 'TOTALE'
            ws_figure[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
            ws_figure[col[0].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[col[0].coordinate].font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
            ws_figure[col[3].coordinate].value = sum(dict_peso_micro.values())
            ws_figure[col[3].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[col[3].coordinate].font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
            ws_figure[col[3].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
            ws_figure[col[3].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws_figure[col[4].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
            ws_figure[col[5].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.merge_cells(start_row=max_row, end_row=max_row, start_column=min_col, end_column=min_col+2)
        assert sum(ws_figure.cell(i, max_col-2).value for i in range(min_row, max_row)) == ws_figure.cell(max_row, max_col-2).value

        # Grafico micro asset class
        chart = PieChart()
        chart.height = 4.77
        chart.width = 6.77
        labels = Reference(ws_figure, min_col=min_col+2, max_col=min_col+2, min_row=min_row, max_row=max_row-1)
        data = Reference(ws_figure, min_col=min_col+3, max_col=min_col+3, min_row=min_row, max_row=max_row-1)
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(labels)
        chart.dataLabels = DataLabelList(dLblPos='bestFit')
        chart.dataLabels.showVal = True
        chart.dataLabels.textProperties = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1100, b=True)), endParaRPr=CharacterProperties(sz=1100, b=True))])
        chart.legend = None
        chart_fonts = ['E4DFEC', 'CCC0DA', 'B1A0C7', '92CDDC', '00B0F0', '0033CC', '0070C0', '1F497D', '000080', 'F79646', 'FFCC66', 'DA5300', 'F62F00', 'EDF06A'] # cambia colori delle fette
        for _ in range(0,14):
            series = chart.series[0]
            pt = DataPoint(idx=_)
            pt.graphicalProperties.solidFill = chart_fonts[_]
            series.dPt.append(pt)
        chart.layout = Layout(manualLayout=ManualLayout(x=0.5, y=0.5, h=1, w=1)) # posizione e dimensione figura
        ws_figure.add_chart(chart, 'L17')
        # Grafico micro pie
        plt.subplots(figsize=(4,4))
        try:
            plt.pie([dict_peso_micro[self.micro_asset_class[_]] for _ in range(0, len(self.micro_asset_class))], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.05 else '' for key, value in dict_peso_micro.items()], radius=1.2, colors=['#E4DFEC', '#CCC0DA', '#B1A0C7', '#92CDDC', '#00B0F0', '#0033CC', '#0070C0', '#1F497D', '#000080', '#F79646', '#FFCC66', '#DA5300', '#F62F00', '#EDF06A'], pctdistance=0.2, labeldistance=0.5, textprops={'fontsize':14}, normalize=False)
        except ValueError:
            plt.pie([dict_peso_micro[self.micro_asset_class[_]] for _ in range(0, len(self.micro_asset_class))], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.05 else '' for key, value in dict_peso_micro.items()], radius=1.2, colors=['#E4DFEC', '#CCC0DA', '#B1A0C7', '#92CDDC', '#00B0F0', '#0033CC', '#0070C0', '#1F497D', '#000080', '#F79646', '#FFCC66', '#DA5300', '#F62F00', '#EDF06A'], pctdistance=0.2, labeldistance=0.5, textprops={'fontsize':14}, normalize=True)
        finally:
            plt.savefig('Media/micro_pie.png', bbox_inches='tight', pad_inches=0)
        # Grafico micro bar
        plt.subplots(figsize=(18.5,5))
        plt.bar(x=[_.replace('Altre Valute', 'Altro').replace('Obbligazionario', 'Obb').replace('Governativo', 'Gov').replace('All Maturities', '').replace('Aggregate', '').replace('North America', 'Nord america').replace('Pacific', 'Pacifico').replace('Emerging Markets', 'Emergenti') for _ in self.micro_asset_class], height=[dict_peso_micro[self.micro_asset_class[_]] for _ in range(0, len(self.micro_asset_class))], width=1, color=['#E4DFEC', '#CCC0DA', '#B1A0C7', '#92CDDC', '#00B0F0', '#0033CC', '#0070C0', '#1F497D', '#000080', '#F79646', '#FFCC66', '#DA5300', '#F62F00', '#EDF06A'])
        plt.xticks(rotation=25)
        plt.savefig('Media/micro_bar.png', bbox_inches='tight', pad_inches=0)

        # Strumenti #
        dict_peso_strumenti = self.peso_strumenti()['strumenti_figure']
        
        #---Tabella strumenti---#
        # Header
        header_strumenti = ['STRUMENTI', '', 'Peso', 'Warning']
        dim_strumenti = [3.4, 47, 9.5, 9.5]
        min_row, max_row, min_col, max_col = 18, 18, 1, 4
        for col in ws_figure.iter_cols(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[col[0].coordinate].value = header_strumenti[col[0].column-min_col]
            ws_figure[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws_figure[col[0].coordinate].font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
            ws_figure[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
            ws_figure[col[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure.column_dimensions[ws_figure[col[0].coordinate].column_letter].width = dim_strumenti[col[0].column-min_col]
        ws_figure.merge_cells(start_row=min_row, end_row=max_row, start_column=min_col, end_column=min_col+1)
        # Body
        fonts_strumenti = ['B1A0C7', '93DEFF', 'FFFF66', 'F79646', '00B0F0', '0066FF', 'FF3737', 'FB9FDA', 'FFC000', '92D050', 'BFBFBF']
        min_row = min_row + 1
        max_row = min_row + len(dict_peso_strumenti.keys()) - 1
        for row in ws_figure.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor=fonts_strumenti[row[0].row-min_row])
            ws_figure[row[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[1].coordinate].value = list(dict_peso_strumenti.keys())[row[0].row-min_row]
            ws_figure[row[1].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[2].coordinate].value = dict_peso_strumenti[ws_figure[row[1].coordinate].value]
            ws_figure[row[2].coordinate].number_format = '0.0%'
            ws_figure[row[2].coordinate].alignment = Alignment(horizontal='center')
            # warnings
            ws_figure[row[2].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            if ws_figure[row[1].coordinate].value == 'Obbligazioni strutturate / Certificates' and dict_peso_strumenti.get('Obbligazioni strutturate / Certificates', 0.00) > 0.10:
                ws_figure[row[3].coordinate].value = '!C'
                ws_figure[row[3].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[1].coordinate].value == 'Hedge funds' and dict_peso_strumenti.get('Hedge funds', 0.00) > 0.25:
                ws_figure[row[3].coordinate].value = '!C'
                ws_figure[row[3].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            ws_figure[row[3].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[3].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[row[3].coordinate].font = Font(color='000000', bold=True)
        # Footer
        max_row = max_row + 1
        ws_figure.cell(max_row, min_col, value='TOTALE')
        ws_figure.cell(max_row, min_col).alignment = Alignment(horizontal='center', vertical='center')
        ws_figure.cell(max_row, min_col).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row, min_col).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row, min_col).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_figure.merge_cells(start_row=max_row, end_row=max_row, start_column=min_col, end_column=min_col+1)
        ws_figure.cell(max_row, min_col+2, value=sum(dict_peso_strumenti.values()))
        assert sum(ws_figure.cell(i, min_col+2).value for i in range(min_row, max_row)) == ws_figure.cell(max_row, min_col+2).value
        ws_figure.cell(max_row, min_col+2).alignment = Alignment(horizontal='center', vertical='center')
        ws_figure.cell(max_row, min_col+2).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row, min_col+2).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row, min_col+2).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_figure.cell(max_row, min_col+2).number_format = FORMAT_PERCENTAGE_00
        ws_figure.cell(max_row, min_col+3).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row, min_col+3).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row, min_col+3).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        
        # Grafico strumenti
        chart = PieChart()
        chart.height = 4.77
        chart.width = 6.77
        labels = Reference(ws_figure, min_col=min_col+1, max_col=min_col+1, min_row=min_row, max_row=max_row-1)
        data = Reference(ws_figure, min_col=min_col+2, max_col=min_col+2, min_row=min_row, max_row=max_row-1)
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(labels)
        chart.dataLabels = DataLabelList(dLblPos='bestFit')
        chart.dataLabels.showVal = True
        chart.dataLabels.textProperties = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1100, b=True)), endParaRPr=CharacterProperties(sz=1100, b=True))])
        chart.legend = None
        # cambia colori delle fette
        chart_fonts = ['B1A0C7', '93DEFF', 'FFFF66', 'F79646', '00B0F0', '0066FF', 'FF3737', 'FB9FDA', 'FFC000', '92D050', 'BFBFBF']
        for _ in range(0,11):
            series = chart.series[0]
            pt = DataPoint(idx=_)
            pt.graphicalProperties.solidFill = chart_fonts[_]
            series.dPt.append(pt)
        # posizione e dimensione figura
        chart.layout = Layout(manualLayout=ManualLayout(x=0.5, y=0.5, h=1, w=1))
        ws_figure.add_chart(chart, 'E18')
        # Grafico strumenti
        plt.subplots(figsize=(4,4))
        try:
            plt.pie([value for value in dict_peso_strumenti.values()], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.05 else '' for value in dict_peso_strumenti.values()], radius=1.2, colors=[{'Conto corrente' : '#B1A0C7', 'Obbligazioni' : '#93DEFF', 'Obbligazioni strutturate / Certificates' : '#FFFF66', 'Azioni' : '#F79646', 'ETF/ETC' : '#00B0F0', 'Fondi comuni/Sicav' : '#0066FF', 'Real Estate' : '#FF3737', 'Hedge funds' : '#FB9FDA', 'Polizze' : '#FFC000', 'Gestioni patrimoniali' : '#92D050', 'Fondi pensione' : '#BFBFBF'}[key] for key, value in dict_peso_strumenti.items()], pctdistance=0.2, labeldistance=0.5, textprops={'fontsize':14}, normalize=False)
        except ValueError:
            plt.pie([value for value in dict_peso_strumenti.values()], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.05 else '' for value in dict_peso_strumenti.values()], radius=1.2, colors=[{'Conto corrente' : '#B1A0C7', 'Obbligazioni' : '#93DEFF', 'Obbligazioni strutturate / Certificates' : '#FFFF66', 'Azioni' : '#F79646', 'ETF/ETC' : '#00B0F0', 'Fondi comuni/Sicav' : '#0066FF', 'Real Estate' : '#FF3737', 'Hedge funds' : '#FB9FDA', 'Polizze' : '#FFC000', 'Gestioni patrimoniali' : '#92D050', 'Fondi pensione' : '#BFBFBF'}[key] for key, value in dict_peso_strumenti.items()], pctdistance=0.2, labeldistance=0.5, textprops={'fontsize':14}, normalize=True)
        finally:
            plt.savefig('Media/strumenti_pie.png', bbox_inches='tight', pad_inches=0)

        # Valute #
        dict_peso_valute = self.peso_valuta_ibrido()

        #---Tabella valute---#
        # Header
        header_valute = ['', 'VALUTE', 'Peso', 'Warning']
        dim_valute = [3.4, 9.5, 9.5, 9.5]
        min_row, max_row, min_col, max_col = 1, 1, 16, 19
        for col in ws_figure.iter_cols(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_figure[col[0].coordinate].value = header_valute[col[0].column-min_col]
            ws_figure[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center')
            ws_figure[col[0].coordinate].font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
            ws_figure[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='595959')
            ws_figure[col[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure.column_dimensions[ws_figure[col[0].coordinate].column_letter].width = dim_valute[col[0].column-min_col]
        # Body
        fonts_valute = ['3366FF', '339966', 'FF99CC', 'FF6600', 'B7DEE8', 'FF9900', 'FFFF66']
        min_row = min_row + 1
        max_row = min_row + len(dict_peso_valute) -1
        for row in ws_figure.iter_rows(min_row=2, max_row=8, min_col=16, max_col=19):
            ws_figure[row[0].coordinate].fill = PatternFill(fill_type='solid', fgColor=fonts_valute[row[0].row-min_row])
            ws_figure[row[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[1].coordinate].value = list(dict_peso_valute.keys())[row[0].row-min_row]
            ws_figure[row[1].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[row[1].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[2].coordinate].value = dict_peso_valute[ws_figure[row[1].coordinate].value]
            ws_figure[row[2].coordinate].number_format = '0.0%'
            ws_figure[row[2].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[row[2].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            if ws_figure[row[1].coordinate].value == 'EUR' and dict_peso_valute.get('EUR', 0.00) < 0.40:
                ws_figure[row[3].coordinate].value = '!!C'
                ws_figure[row[3].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[1].coordinate].value == 'EUR' and dict_peso_valute.get('EUR', 0.00) < 0.50:
                ws_figure[row[3].coordinate].value = '!!C'
                ws_figure[row[3].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif ws_figure[row[1].coordinate].value == 'EUR' and dict_peso_valute.get('EUR', 0.00) < 0.60:
                ws_figure[row[3].coordinate].value = '!C'
                ws_figure[row[3].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            ws_figure[row[3].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_figure[row[3].coordinate].alignment = Alignment(horizontal='center')
            ws_figure[row[3].coordinate].font = Font(color='000000', bold=True)
        # Footer
        max_row = max_row + 1
        ws_figure.cell(max_row, min_col, value='TOTALE')
        ws_figure.cell(max_row, min_col).alignment = Alignment(horizontal='center', vertical='center')
        ws_figure.cell(max_row, min_col).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row, min_col).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row, min_col).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_figure.merge_cells(start_row=max_row, end_row=max_row, start_column=min_col, end_column=min_col+1)
        ws_figure.cell(max_row, min_col+2).value = sum(dict_peso_valute.values())
        assert sum(ws_figure.cell(i, min_col+2).value for i in range(min_row, max_row)) == ws_figure.cell(max_row, min_col+2).value
        ws_figure.cell(max_row, min_col+2).alignment = Alignment(horizontal='center', vertical='center')
        ws_figure.cell(max_row, min_col+2).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row, min_col+2).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row, min_col+2).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_figure.cell(max_row, min_col+2).number_format = FORMAT_PERCENTAGE_00
        ws_figure.cell(max_row, min_col+3).font = Font(name='Arial', size=11, color='FFFFFF', bold=True)
        ws_figure.cell(max_row, min_col+3).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_figure.cell(max_row, min_col+3).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))

        # Grafico valute
        chart = PieChart()
        chart.height = 4.77
        chart.width = 6.77
        labels = Reference(ws_figure, min_col=min_col+1, max_col=min_col+1, min_row=min_row, max_row=max_row-1)
        data = Reference(ws_figure, min_col=min_col+2, max_col=min_col+2, min_row=min_row, max_row=max_row-1)
        chart.add_data(data, titles_from_data=False)
        chart.set_categories(labels)
        chart.dataLabels = DataLabelList(dLblPos='bestFit')
        chart.dataLabels.showVal = True
        chart.dataLabels.textProperties = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=CharacterProperties(sz=1100, b=True)), endParaRPr=CharacterProperties(sz=1100, b=True))])
        chart.legend = None
        chart_fonts = ['3366FF', '339966', 'FF99CC', 'FF6600', 'B7DEE8', 'FF9900', 'FFFF66'] # cambia colori delle fette
        for _ in range(0,7):
            series = chart.series[0]
            pt = DataPoint(idx=_)
            pt.graphicalProperties.solidFill = chart_fonts[_]
            series.dPt.append(pt)
        chart.layout = Layout(manualLayout=ManualLayout(x=0.5, y=0.5, h=1, w=1)) # posizione e dimensione figura
        ws_figure.add_chart(chart, 'Q11')
        # Grafico valute
        plt.subplots(figsize=(4,4))
        try:
            plt.pie([value for value in dict_peso_valute.values()], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.05 else '' for value in dict_peso_valute.values()], radius=1.2, colors=[{'EUR' : '#3366FF', 'USD' : '#339966', 'YEN' : '#FF99CC', 'CHF' : '#FF6600', 'GBP' : '#B7DEE8', 'AUD' : '#FF9900', 'ALTRO' : '#FFFF66'}[key] for key, value in dict_peso_valute.items()], pctdistance=0.2, labeldistance=0.5, textprops={'fontsize':14}, normalize=False)
        except ValueError:
            plt.pie([value for value in dict_peso_valute.values()], labels=[str(round((value*100),2)).replace('.',',')+'%' if value > 0.05 else '' for value in dict_peso_valute.values()], radius=1.2, colors=[{'EUR' : '#3366FF', 'USD' : '#339966', 'YEN' : '#FF99CC', 'CHF' : '#FF6600', 'GBP' : '#B7DEE8', 'AUD' : '#FF9900', 'ALTRO' : '#FFFF66'}[key] for key, value in dict_peso_valute.items()], pctdistance=0.2, labeldistance=0.5, textprops={'fontsize':14}, normalize=True)
        finally:
            plt.savefig('Media/valute_pie.png', bbox_inches='tight', pad_inches=0)

    def mappatura_fondi(self):
        """Crea la tabella e il grafico a barre della mappatura dei fondi."""
        # Carica il foglio fondi   
        fondi = self.wb['fondi']
        self.wb.active = fondi
        # Fondi
        df_portfolio = self.df_portfolio
        prodotti_gestiti = df_portfolio.loc[df_portfolio['strumento']=='fund']
        numero_prodotti_gestiti = prodotti_gestiti.nome.count()
        if numero_prodotti_gestiti > 0:
            # Mappatura dei fondi
            df_mappatura = self.df_mappatura
            df_mappatura_fondi = df_mappatura.loc[df_mappatura['ISIN'].isin(prodotti_gestiti['ISIN'])]
            # Header
            header = list(['ISIN', 'Nome']) + self.micro_asset_class
            dimensions = [23, 55, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23, 23]
            min_row = 1
            max_row = 1
            min_col = 11
            max_col = min_col + len(self.micro_asset_class) + 1
            for col in fondi.iter_cols(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                fondi[col[0].coordinate].value = header[col[0].column-min_col]
                fondi[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                fondi[col[0].coordinate].font = Font(name='Calibri', size=16, color='FFFFFF', bold=True)
                fondi[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
                fondi[col[0].coordinate].border = Border(top=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                fondi.column_dimensions[col[0].column_letter].width = dimensions[col[0].column-min_col]
            # Body
            min_row = min_row + 1
            max_row = max_row + numero_prodotti_gestiti
            for row in fondi.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                fondi[row[0].coordinate].value = df_mappatura_fondi['ISIN'].values[row[0].row-min_row]
                fondi[row[0].coordinate].font = Font(name='Calibri', size=18, color='000000')
                fondi[row[0].coordinate].border = Border(top=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                fondi[row[1].coordinate].value = df_mappatura_fondi.loc[df_mappatura_fondi['ISIN']==fondi[row[0].coordinate].value, 'nome'].values[0]
                fondi[row[1].coordinate].font = Font(name='Calibri', size=18, color='000000')
                fondi[row[1].coordinate].border = Border(top=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                for _ in range(2, 16):
                    fondi[row[_].coordinate].value = df_mappatura_fondi.loc[df_mappatura_fondi['ISIN']==fondi[row[0].coordinate].value, fondi.cell(row=min_row-1, column=row[_].column).value].values[0]
                    fondi[row[_].coordinate].font = Font(name='Calibri', size=18, color='000000')
                    fondi[row[_].coordinate].border = Border(top=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                    fondi[row[_].coordinate].alignment = Alignment(horizontal='center')
                    fondi[row[_].coordinate].number_format = FORMAT_PERCENTAGE_00
            # Footer
            min_row = max_row + 1
            max_row = max_row + 2
            footer = ['CONTROVALORE TOTALE', 'PESO PERCENTUALE']
            for row in fondi.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                fondi[row[1].coordinate].value = footer[row[0].row-min_row]
                fondi[row[1].coordinate].font = Font(name='Calibri', size=18, color='000000', bold=True)
                fondi[row[1].coordinate].border = Border(top=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                if row[0].row == min_row:
                    for _ in range(2, 16):
                        fondi[row[_].coordinate].value = np.nan_to_num(np.array(df_mappatura_fondi[fondi.cell(row=min_row-numero_prodotti_gestiti-1, column=row[_].column).value]), nan=0.0) @ np.array(prodotti_gestiti['controvalore_in_euro'])
                        fondi[row[_].coordinate].font = Font(name='Calibri', size=18, color='000000')
                        fondi[row[_].coordinate].border = Border(top=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                        fondi[row[_].coordinate].alignment = Alignment(horizontal='center')
                        fondi[row[_].coordinate].number_format = FORMAT_NUMBER_COMMA_SEPARATED1
                if row[0].row == max_row:
                    for _ in range(2, 16):
                        fondi[row[_].coordinate].value = fondi.cell(row=max_row-1, column=row[_].column).value / sum(prodotti_gestiti['controvalore_in_euro'])
                        fondi[row[_].coordinate].font = Font(name='Calibri', size=18, color='000000')
                        fondi[row[_].coordinate].border = Border(top=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
                        fondi[row[_].coordinate].alignment = Alignment(horizontal='center')
                        fondi[row[_].coordinate].number_format = FORMAT_PERCENTAGE_00
            # Grafico micro bar
            plt.subplots(figsize=(18.5,5))
            plt.bar(x=[_.replace('Altre Valute', 'Altro').replace('Obbligazionario', 'Obb').replace('Governativo', 'Gov').replace('All Maturities', '').replace('Aggregate', '').replace('North America', 'Nord america').replace('Pacific', 'Pacifico').replace('Emerging Markets', 'Emergenti') for _ in self.micro_asset_class], height=[fondi.cell(row=max_row, column=_).value for _ in range(min_col+2, max_col+1)], width=1, color=['#E4DFEC', '#CCC0DA', '#B1A0C7', '#92CDDC', '#00B0F0', '#0033CC', '#0070C0', '#1F497D', '#000080', '#F79646', '#FFCC66', '#DA5300', '#F62F00', '#EDF06A'])
            plt.xticks(rotation=25)
            plt.savefig('Media/map_fondi_bar.png', bbox_inches='tight', pad_inches=0)

    def sintesi(self):
        """Crea la tabella da piazzare in fondo alla presentazione."""
        self.wb.create_sheet('sintesi')
        ws_sintesi = self.wb['sintesi']
        self.wb.active = ws_sintesi
        df_p = self.df_portfolio
        df_m = self.df_mappatura.drop(['TOTALE'], axis=1)
        dict_peso_macro = self.peso_macro()
        # Header
        header = ['ISIN', 'Asset class', 'Prodotto', 'Valore di mercato in euro', 'Peso']
        min_row = 1
        max_row = 2
        min_col = 1
        max_col = len(header)
        for col in ws_sintesi.iter_cols(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            ws_sintesi[col[0].coordinate].value = header[col[0].column-min_col]
            ws_sintesi[col[0].coordinate].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ws_sintesi[col[0].coordinate].font = Font(name='Century Gothic', size=16, color='FFFFFF', bold=True)
            ws_sintesi[col[0].coordinate].fill = PatternFill(fill_type='solid', fgColor='808080')
            ws_sintesi[col[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi.row_dimensions[col[0].row].height = 30
            ws_sintesi.row_dimensions[col[1].row].height = 30
            ws_sintesi.merge_cells(start_row=min_row, end_row=max_row, start_column=col[0].column, end_column=col[0].column)
            if ws_sintesi[col[0].coordinate].value == 'ISIN':
                ws_sintesi.column_dimensions[col[0].column_letter].width = 25
            elif ws_sintesi[col[0].coordinate].value == 'Asset class':
                ws_sintesi.column_dimensions[col[0].column_letter].width = 56
            elif ws_sintesi[col[0].coordinate].value == 'Prodotto':
                ws_sintesi.column_dimensions[col[0].column_letter].width = max([len(nome) for nome in df_m['nome'].values])*1.7
            elif ws_sintesi[col[0].coordinate].value == 'Valore di mercato in euro':
                ws_sintesi.column_dimensions[col[0].column_letter].width = max(24.3, max([len(str(round(controvalore_in_euro,2))) for controvalore_in_euro in df_p['controvalore_in_euro'].values])*2.5)
            elif ws_sintesi[col[0].coordinate].value == 'Peso':
                ws_sintesi.column_dimensions[col[0].column_letter].width = 13
        ws_sintesi.cell(row=min_row, column=len(header)+1, value='Warning')
        ws_sintesi.cell(row=min_row, column=len(header)+1).alignment = Alignment(horizontal='center', vertical='center')
        ws_sintesi.cell(row=min_row, column=len(header)+1).font = Font(name='Century Gothic', size=16, color='FFFFFF', bold=True)
        ws_sintesi.cell(row=min_row, column=len(header)+1).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_sintesi.cell(row=min_row, column=len(header)+1).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_sintesi.merge_cells(start_row=min_row, end_row=min_row, start_column=len(header)+1, end_column=len(header)+3)
        ws_sintesi.cell(row=min_row+1, column=len(header)+1, value='C/R')
        ws_sintesi.cell(row=min_row+1, column=len(header)+1).alignment = Alignment(horizontal='center', vertical='center')
        ws_sintesi.cell(row=min_row+1, column=len(header)+1).font = Font(name='Century Gothic', size=16, color='FFFFFF', bold=True)
        ws_sintesi.cell(row=min_row+1, column=len(header)+1).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_sintesi.cell(row=min_row+1, column=len(header)+1).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_sintesi.cell(row=min_row+1, column=len(header)+2, value='L')
        ws_sintesi.cell(row=min_row+1, column=len(header)+2).alignment = Alignment(horizontal='center', vertical='center')
        ws_sintesi.cell(row=min_row+1, column=len(header)+2).font = Font(name='Century Gothic', size=16, color='FFFFFF', bold=True)
        ws_sintesi.cell(row=min_row+1, column=len(header)+2).fill = PatternFill(fill_type='solid', fgColor='595959')
        ws_sintesi.cell(row=min_row+1, column=len(header)+2).border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        ws_sintesi.cell(row=min_row+1, column=len(header)+3).fill = PatternFill(fill_type='solid', fgColor='595959')
        # Body
        df_m['asset_class'] = df_m.apply(lambda x : x[self.micro_asset_class].index[x[self.micro_asset_class] == 1.00].values[0] if any(x[self.micro_asset_class]==1.00) else 'Prodotto multi asset', axis=1)
        custom_dict = {value : num for num, value in enumerate(self.micro_asset_class)}
        custom_dict['Prodotto multi asset'] = 14
        # ordina il dataframe portfolio_valori per tipo di strumento così come è stato mappato (dalla liquidità al prodotto composto)
        df_m.sort_values(by=['asset_class'], inplace=True, key=lambda x : x.map(custom_dict))
        # print(df_m)
        min_row = min_row + 2
        max_row = min_row + len(df_m) - 1
        df_p.fillna('', inplace=True)
        df_m.fillna('', inplace=True)
        for row in ws_sintesi.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col+3):
            ws_sintesi[row[0].coordinate].value = df_m['ISIN'].values[row[0].row-min_row]
            ws_sintesi[row[0].coordinate].font = Font(name='Century Gothic', size=16, color='000000')
            ws_sintesi[row[0].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[1].coordinate].value = df_m['asset_class'].values[row[0].row-min_row].replace('Obbligazionario Euro Governativo All Maturities', 'Obbligazionario Euro Governativo')
            ws_sintesi[row[1].coordinate].font = Font(name='Century Gothic', size=16, color='000000')
            ws_sintesi[row[1].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[2].coordinate].value = df_m['nome'].values[row[0].row-min_row]
            ws_sintesi[row[2].coordinate].font = Font(name='Century Gothic', size=16, color='000000')
            ws_sintesi[row[2].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[3].coordinate].value = df_p.loc[(df_p['ISIN']==ws_sintesi[row[0].coordinate].value) & (df_p['nome']==ws_sintesi[row[2].coordinate].value), 'controvalore_in_euro'].values[0]
            ws_sintesi[row[3].coordinate].font = Font(name='Century Gothic', size=16, color='000000')
            ws_sintesi[row[3].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[3].coordinate].number_format = '€ #,0.00'
            ws_sintesi[row[4].coordinate].value = df_p.loc[(df_p['ISIN']==ws_sintesi[row[0].coordinate].value) & (df_p['nome']==ws_sintesi[row[2].coordinate].value), 'controvalore_in_euro'].values[0] / df_p['controvalore_in_euro'].sum()
            ws_sintesi[row[4].coordinate].font = Font(name='Century Gothic', size=16, color='000000')
            ws_sintesi[row[4].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[4].coordinate].number_format = FORMAT_PERCENTAGE_00
            if df_p.loc[(df_p['ISIN']==ws_sintesi[row[0].coordinate].value) & (df_p['nome']==ws_sintesi[row[2].coordinate].value), 'strumento'].values[0] == 'gov_bond' or df_p.loc[(df_p['ISIN']==ws_sintesi[row[0].coordinate].value) & (df_p['nome']==ws_sintesi[row[2].coordinate].value), 'strumento'].values[0] == 'corp_bond':
                if ws_sintesi[row[4].coordinate].value / dict_peso_macro['Obbligazionario'] > 0.30:
                    ws_sintesi[row[5].coordinate].value = '!!!C'
                    ws_sintesi[row[5].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif ws_sintesi[row[4].coordinate].value / dict_peso_macro['Obbligazionario'] > 0.20:
                    ws_sintesi[row[5].coordinate].value = '!!C'
                    ws_sintesi[row[5].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif ws_sintesi[row[4].coordinate].value / dict_peso_macro['Obbligazionario'] > 0.10:
                    ws_sintesi[row[5].coordinate].value = '!C'
                    ws_sintesi[row[5].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            elif df_p.loc[(df_p['ISIN']==ws_sintesi[row[0].coordinate].value) & (df_p['nome']==ws_sintesi[row[2].coordinate].value), 'strumento'].values[0] == 'equity':
                if ws_sintesi[row[4].coordinate].value / dict_peso_macro['Azionario'] > 0.30:
                    ws_sintesi[row[5].coordinate].value = '!!!C'
                    ws_sintesi[row[5].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif ws_sintesi[row[4].coordinate].value / dict_peso_macro['Azionario'] > 0.20:
                    ws_sintesi[row[5].coordinate].value = '!!C'
                    ws_sintesi[row[5].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
                elif ws_sintesi[row[4].coordinate].value / dict_peso_macro['Azionario'] > 0.10:
                    ws_sintesi[row[5].coordinate].value = '!C'
                    ws_sintesi[row[5].coordinate].fill = PatternFill(fill_type='solid', fgColor='FFD700')
            ws_sintesi[row[5].coordinate].font = Font(name='Century Gothic', size=16, color='000000', bold=True)
            ws_sintesi[row[5].coordinate].alignment = Alignment(horizontal='center')
            ws_sintesi[row[5].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[6].coordinate].font = Font(name='Century Gothic', size=16, color='FF0000', bold=True)
            ws_sintesi[row[6].coordinate].alignment = Alignment(horizontal='center')
            ws_sintesi[row[6].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[7].coordinate].border = Border(right=Side(border_style='thin', color='000000'), left=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
        # Footer
        max_row = max_row + 1
        for row in ws_sintesi.iter_rows(min_row=max_row, max_row=max_row, min_col=min_col+1, max_col=max_col+3):
            ws_sintesi[row[0].coordinate].value = 'TOTALE PORTAFOGLIO'
            ws_sintesi[row[0].coordinate].font = Font(name='Century Gothic', size=16, color='000000', bold=True)
            ws_sintesi[row[0].coordinate].alignment = Alignment(horizontal='center')
            ws_sintesi[row[0].coordinate].border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi.merge_cells(start_row=max_row, end_row=max_row, start_column=min_col+1, end_column=min_col+2)
            ws_sintesi[row[2].coordinate].value = df_p['controvalore_in_euro'].sum()
            ws_sintesi[row[2].coordinate].font = Font(name='Century Gothic', size=16, bold=True)
            ws_sintesi[row[2].coordinate].number_format = '€ #,0.00'
            ws_sintesi[row[2].coordinate].border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            np.testing.assert_almost_equal(actual=sum([ws_sintesi.cell(row=_, column=min_col+3).value for _ in range(min_row, max_row)]), desired=ws_sintesi[row[2].coordinate].value, decimal=2, err_msg='La somma dei valori dei singoli strumenti non è uguale al controvalore totale di portafoglio', verbose=True)
            ws_sintesi[row[3].coordinate].value = sum([ws_sintesi.cell(row=_, column=min_col+4).value for _ in range(min_row, max_row)])
            ws_sintesi[row[3].coordinate].font = Font(name='Century Gothic', size=16, bold=True)
            np.testing.assert_almost_equal(actual=ws_sintesi[row[3].coordinate].value, desired=1.00, decimal=1, err_msg='La sommma dei pesi non fa 100', verbose=True)
            ws_sintesi[row[3].coordinate].number_format = FORMAT_PERCENTAGE_00
            ws_sintesi[row[3].coordinate].border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi[row[4].coordinate].border = Border(left=Side(border_style='thin', color='000000'), right=Side(border_style='thin', color='000000'), top=Side(border_style='thin', color='000000'), bottom=Side(border_style='thin', color='000000'))
            ws_sintesi.merge_cells(start_row=max_row, end_row=max_row, start_column=max_col+1, end_column=max_col+3)

    def salva_file_portafoglio(self):
        """Salva il file excel."""
        self.wb.save(self.file_elaborato)

class Presentazione(Portfolio):
    """Tentativo di ricreare la presentazione."""

    def __init__(self, tipo_sap, file_elaborato, file_presentazione, **dimensioni):
        """
        Initialize the class. Inherits from SAP.

        Parameters:
        tipo_sap(str) = completo o light
        file_elaborato(str) = file excel elaborato
        file_presentazione(str) = file word da compilare
        **dimensioni(dict) = dimensioni delle pagine word
        """
        super().__init__(file_portafoglio=PTF, path=PATH)
        self.tipo_sap = tipo_sap
        if tipo_sap != 'completo' and tipo_sap != 'light':
            print('Il tipo di SAP può essere completo o light!')
            quit()
        self.document = Document()
        # Aggiorna proprietà documento
        self.document.core_properties.title = 'SAP'
        self.document.core_properties.subject = 'Analisi di portafoglio'
        self.document.core_properties.category = 'Financial analysis'
        self.document.core_properties.author = 'B&S'
        self.document.core_properties.comments = ''
        self.wb = load_workbook(file_elaborato)
        self.file_elaborato = file_elaborato
        self.file_presentazione = file_presentazione
        self.page_height = dimensioni['page_height']
        self.page_width = dimensioni['page_width']
        self.top_margin = dimensioni['top_margin']
        self.bottom_margin = dimensioni['bottom_margin']
        self.left_margin = dimensioni['left_margin']
        self.right_margin = dimensioni['right_margin']
        self.larghezza_pagina = self.page_width - self.left_margin - self.right_margin

    def copertina_1(self):
        """Copertina della presentazione."""
        # 1.Copertina
        section = self.document.sections[0]
        # Imposta dimensioni A4
        section.page_height = shared.Cm(self.page_height)
        section.page_width = shared.Cm(self.page_width)
        # Imposta margini
        left_margin = 0.60
        right_margin = 0.60
        top_margin = 0.45
        bottom_margin = 0.45
        section.left_margin = shared.Cm(left_margin)
        section.right_margin = shared.Cm(right_margin)
        section.top_margin = shared.Cm(top_margin)
        section.bottom_margin = shared.Cm(bottom_margin)
        section.header_distance = shared.Cm(0)
        section.footer_distance = shared.Cm(0)
        # Image
        paragraph = self.document.add_paragraph()
        paragraph.alignment = 1
        copertina = 'copertina_completo.jpg' if self.tipo_sap=='completo' else 'copertina_light.jpg' if self.tipo_sap=='light' else print('Il tipo di SAP può essere completo o light!')
        paragraph.add_run().add_picture(self.path+'\Media\default\\'+copertina, height=shared.Cm(self.page_height-top_margin-bottom_margin), width=shared.Cm(self.page_width-left_margin-right_margin))

    def indice_2(self):
        """Indice della presentazione."""
        # 2.Indice
        paragraph_format = self.document.styles['Normal'].paragraph_format
        paragraph_format.space_after = 0 # Annulla lo spazio dopo il testo per tutte le stringhe di tipo normale.
        self.document.add_section()
        section = self.document.sections[1]
        # Imposta margini
        section.top_margin = shared.Cm(self.top_margin)
        section.bottom_margin = shared.Cm(self.bottom_margin)
        section.right_margin = shared.Cm(self.right_margin)
        section.left_margin = shared.Cm(self.left_margin)
        # Header
        header = section.header
        header.is_linked_to_previous = False # Se True crea l'header anche per la pagina precedente
        paragraph = header.paragraphs[0]
        paragraph.add_run('\n\n').add_picture(self.path+'\Media\default\logo_azimut.bmp', height=shared.Cm(1.4), width=shared.Cm(3.72))
        # Title
        paragraph_0 = self.document.add_paragraph()
        run_0 = paragraph_0.add_run('\n')
        run_0 = paragraph_0.add_run(text='INDICE', style=None)
        run_0.bold = True
        run_0.font.name = 'Century Gothic'
        run_0.font.size = shared.Pt(24)
        run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
        run_0 = paragraph_0.add_run('\n')
        # Body - list numbers
        paragraph_1 = self.document.add_paragraph('PORTAFOGLIO ATTUALE', style='List Number')
        paragraph_1.style.font.name = 'Century Gothic'
        paragraph_1.style.font.size = shared.Pt(14)
        paragraph_1.style.font.color.rgb = shared.RGBColor(127, 127, 127)
        paragraph_1.add_run('\n')
        paragraph_2 = self.document.add_paragraph('ANALISI DEL PORTAFOGLIO', style='List Number')
        run_2 = paragraph_2.add_run('\n\tPer macro asset class\n\tPer micro asset class\n\tPer tipologia di prodotto\n\tPer valuta')
        run_2.font.name = 'Century Gothic'
        run_2.font.size = shared.Pt(12)
        run_2.font.color.rgb = shared.RGBColor(127, 127, 127)
        paragraph_2.add_run('\n')
        paragraph_3 = self.document.add_paragraph('ANALISI DEI SINGOLI STRUMENTI', style='List Number')
        paragraph_3.add_run('\n')
        paragraph_4 = self.document.add_paragraph('ANALISI DEL RISCHIO', style='List Number')
        paragraph_4.add_run('\n')
        self.document.add_paragraph('NOTE METODOLOGICHE', style='List Number')
        self.document.add_page_break()

    def portafoglio_attuale_3(self):
        """
        Portafoglio complessivo diviso per strumenti.
        Metodo 1 : stampa solo i primi 57 senza riportare come prima riga dopo l'intestazione l'etichetta del primo strumento a comparire.
        """
        df = self.df_portfolio
        if all(df['quantità'].isnull()):
            print("Mancano le quantità")
        if all(df['controvalore_iniziale'].isnull()):
            print("Mancano i controvalori iniziali")
        if all(df['prezzo_di_carico'].isnull()):
            print("Mancano i prezzi di carico")
        sheet = self.wb['agglomerato']
        self.wb.active = sheet
        # Nascondi colonne vuote
        hidden_columns = 0
        if all(df['quantità'].isnull()):
            sheet.column_dimensions['C'].hidden= True
            hidden_columns += 1
        if all(df['controvalore_iniziale'].isnull()):
            sheet.column_dimensions['D'].hidden= True
            hidden_columns += 1
        if all(df['prezzo_di_carico'].isnull()):
            sheet.column_dimensions['E'].hidden= True
            hidden_columns += 1
        self.wb.save(self.file_elaborato)
        c = Counter(list(df.loc[:, 'strumento']))
        strumenti_in_ptf = [strumento for strumento in self.strumenti if c[strumento] > 0]
        max_row = 1 + df['nome'].count() + len(strumenti_in_ptf) + 1
        # print(f'ci sono {max_row} righe')
        LIMITE= 63
        if max_row <= LIMITE:
            tabelle_agglomerato = 1
        else:
            if max_row % LIMITE != 0:
                tabelle_agglomerato = max_row // LIMITE + 1
            elif max_row % LIMITE == 0:
                tabelle_agglomerato = max_row // LIMITE
        for tabella in range(1, tabelle_agglomerato+1):
            if tabella != tabelle_agglomerato:
                excel2img.export_img(self.file_elaborato, self.path+'\Media\\agglomerato_'+str(tabella-1)+'.png', page='agglomerato', _range="A1:I"+str(LIMITE*tabella))
                sheet.row_dimensions.group(2,LIMITE*tabella,hidden=True)
                self.wb.save(self.file_elaborato)
            elif tabella == tabelle_agglomerato:
                excel2img.export_img(self.file_elaborato, self.path+'\Media\\agglomerato_'+str(tabella-1)+'.png', page='agglomerato', _range="A1:I"+str(max_row))
            print(f"sto aggiungendo l'agglomerato {tabella-1} alla presentazione.")
            self.document.add_section()
            paragraph_0 = self.document.add_paragraph(text='\n', style=None)
            paragraph_0.paragraph_format.space_after = shared.Pt(6)
            run_0 = paragraph_0.add_run('1. PORTAFOGLIO ATTUALE')
            run_0.bold = True
            run_0.font.name = 'Century Gothic'
            run_0.font.size = shared.Pt(14)
            run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
            paragraph_1 = self.document.add_paragraph(style=None)
            run_1 = paragraph_1.add_run()
            width = self.larghezza_pagina if hidden_columns==0 else self.larghezza_pagina - 1 if hidden_columns==1 else self.larghezza_pagina - 2 if hidden_columns==2 else self.larghezza_pagina - 3 if hidden_columns==3 else 18
            run_1.add_picture(self.path+'\Media\\agglomerato_'+ str(tabella-1) +'.png', width=shared.Cm(width))
        sheet.row_dimensions.group(2,LIMITE*(tabelle_agglomerato),hidden=False)
        
    def new_portafoglio_attuale_3(self):
        """Portafoglio complessivo diviso per strumenti.
        Metodo 2 : stampa i primi 57 riportando sempre come prima riga dopo l'intestazione l'etichetta del primo strumento a comparire."""
        df = pd.read_excel(self.file_elaborato, sheet_name='portfolio_valori')
        if all(df['quantità'].isnull()):
            print("Mancano le quantità")
        if all(df['controvalore_iniziale'].isnull()):
            print("Mancano i controvalori iniziali")
        if all(df['prezzo_di_carico'].isnull()):
            print("Mancano i prezzi di carico")
        sheet = self.wb['agglomerato']
        self.wb.active = sheet
        # Nascondi colonne vuote
        hidden_columns = 0
        if all(df['quantità'].isnull()):
            sheet.column_dimensions['C'].hidden= True
            hidden_columns += 1
        if all(df['controvalore_iniziale'].isnull()):
            sheet.column_dimensions['D'].hidden= True
            hidden_columns += 1
        if all(df['prezzo_di_carico'].isnull()):
            sheet.column_dimensions['E'].hidden= True
            hidden_columns += 1
        self.wb.save(self.file_elaborato)
        c = Counter(list(df.loc[:, 'strumento']))
        strumenti_in_ptf = [strumento for strumento in self.strumenti if c[strumento] > 0]
        max_row = 1 + df['nome'].count() + len(strumenti_in_ptf) + 1
        print("Ultima riga:", max_row)
        # limite = 57
        # # Prima tabella
        # excel2img.export_img(self.file_elaborato, 'C:\\Users\\Administrator\\Desktop\\Sbwkrq\\SAP\\Media\\'+'agglomerato_0.bmp', page='agglomerato', _range="A1:I"+str(limite))
        # while _ < max_row:
        riga = 1
        i = 0
        while riga <= 114:
            i += 1
            print(c)
            limite = 57
            # Foto prima tabella
            excel2img.export_img(self.file_elaborato, 'C:\\Users\\Administrator\\Desktop\\Sbwkrq\\SAP\\Media\\'+'agglomerato_'+str(i)+'.bmp', page='agglomerato', _range="A1:I"+str(limite))
            riga += 1
            # Nascondi la prima tabella lasciando l'etichetta dello strumento che ha più strumenti di quanti ce ne stanno nella prima pagina
            for strumento in strumenti_in_ptf:
                numerosità_strumento = c[strumento]
                if numerosità_strumento > 0:
                    print('numerosità',strumento, numerosità_strumento)
                    c_strumento = Counter({strumento : numerosità_strumento})
                    print(c_strumento)
                    if numerosità_strumento > limite - riga: # 57 - 2
                        sheet.row_dimensions.group(riga+1, limite, hidden=True)
                        riga += limite - 1
                        limite -= 1
                        # c_strumento = Counter({strumento : numerosità_strumento - })
                        c.subtract(c_strumento)
                        break
                    else:
                        sheet.row_dimensions.group(riga, riga + numerosità_strumento, hidden=True)
                        riga += numerosità_strumento + 1
                        c.subtract(c_strumento)
                        continue

     
        
        
        # # 1
        # numerosità_primo_strumento = c[strumenti_in_ptf[0]]
        # print('numerosità',strumenti_in_ptf[0], numerosità_primo_strumento)
        # if numerosità_primo_strumento > 55:
        #     sheet.row_dimensions.group(3,limite*1,hidden=True)
        #     if numerosità_primo_strumento - 55 > 55:
        #         sheet.row_dimensions.group(limite*1 + 1, limite*2, hidden=True)
        # elif numerosità_primo_strumento <= 55:
        #     sheet.row_dimensions.group(2,2+numerosità_primo_strumento,hidden=True)
        #     # 2
        #     numerosità_secondo_strumento = c[strumenti_in_ptf[1]]
        #     print('numerosità',strumenti_in_ptf[1], numerosità_secondo_strumento)
        #     if numerosità_secondo_strumento > 55 - numerosità_primo_strumento + 1:
        #         sheet.row_dimensions.group(2+(numerosità_primo_strumento+1),limite*1,hidden=True)
        #     elif numerosità_secondo_strumento <= 55 - (numerosità_primo_strumento + 1):
        #         sheet.row_dimensions.group(2+numerosità_primo_strumento+1,2+(numerosità_primo_strumento+1)+numerosità_secondo_strumento,hidden=True)
        #         # 3
        #         numerosità_terzo_strumento = c[strumenti_in_ptf[2]]
        #         print('numerosità',strumenti_in_ptf[2], numerosità_terzo_strumento)
        #         if numerosità_terzo_strumento > 55 - (numerosità_secondo_strumento + 1) - (numerosità_primo_strumento + 1):
        #             sheet.row_dimensions.group(2+(numerosità_secondo_strumento+1)+(numerosità_primo_strumento+1),limite*1,hidden=True)
        #         elif numerosità_terzo_strumento <= 55 - (numerosità_secondo_strumento + 1) - (numerosità_primo_strumento + 1):
        #             sheet.row_dimensions.group(2+(numerosità_secondo_strumento+1)+(numerosità_primo_strumento+1),2+(numerosità_secondo_strumento+1)+(numerosità_primo_strumento+1)+numerosità_terzo_strumento,hidden=True)





        # self.wb.save(self.file_elaborato)
        # paragraph = self.document.add_paragraph(text='', style=None)
        # run = paragraph.add_run('\n\nCaratteristiche finanziarie dei fondi comuni di investimento')
        # run.bold = True
        # run.font.name = 'Century Gothic'
        # run.font.size = shared.Pt(12)
        # run.font.color.rgb = shared.RGBColor(127, 127, 127)
        # paragraph = self.document.add_paragraph(text='', style=None)
        # run = paragraph.add_run()
        # run.add_picture(r'C:\Users\Administrator\Desktop\Sbwkrq\SAP\Media\barra.png', width=shared.Cm(18.5))
        # paragraph = self.document.add_paragraph(text='', style=None)
        # run = paragraph.add_run()
        # run.add_picture(r'C:\Users\Administrator\Desktop\Sbwkrq\SAP\Media\fondi_0.bmp', width=shared.Cm(18.5) if hidden_columns==0 else shared.Cm(13.5))
        # Seconda tabella
        # cosa c'era nella prima tabella?
        # intestazione nella prima riga
        # nascondi etichetta + relativi prodotti uno strumento alla volta, in modo che ci sia sempre un'etichetta sotto l'intestazione
        # conta quanti prodotti + etichetta ci sono del primo strumento; se > 56 nascondi la riga dalla 2 alla 55, altrimenti nascondi dalla riga 2 alla 2 + numero prodotti primo strumento

    def old_portafoglio_attuale_3(self):
        """Portafoglio complessivo diviso per strumenti."""
        # Carica dataframe del portfolio e controlla quali colonne sono interamente vuote
        df = pd.read_excel(self.file_elaborato, sheet_name='portfolio_valori')
        if all(df['quantità'].isnull()):
            print("Mancano le quantità")
        if all(df['controvalore_iniziale'].isnull()):
            print("Mancano i controvalori iniziali")
        if all(df['prezzo_di_carico'].isnull()):
            print("Mancano i prezzi di carico")
        # Ottieni immagini dei file agglomerati
        for ws in self.wb.sheetnames:
            if ws.startswith('agglomerato'):
                sheet = self.wb[ws]
                self.wb.active = sheet
                # Nascondi colonne vuote
                hidden_columns = 0
                if all(df['quantità'].isnull()):
                    sheet.column_dimensions['C'].hidden= True
                    hidden_columns += 1
                if all(df['controvalore_iniziale'].isnull()):
                    sheet.column_dimensions['D'].hidden= True
                    hidden_columns += 1
                if all(df['prezzo_di_carico'].isnull()):
                    sheet.column_dimensions['E'].hidden= True
                    hidden_columns += 1
                self.wb.save(self.file_elaborato)
                # Cancella immagini di file agglomerati salvati in precedenza (non è necessario, excel2img fa l'overwrite)
                # if os.path.exists('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\SAP\\Media\\'+ ws +'.bmp'):
                #     print(f'rimuovo il file {ws}...')
                    # os.remove('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\SAP\\Media\\'+ ws +'.bmp')
                # Importa agglomerato
                excel2img.export_img(self.file_elaborato, 'C:\\Users\\Administrator\\Desktop\\Sbwkrq\\SAP\\Media\\'+ ws +'.bmp', page=ws, _range="A1:I"+str(sheet.max_row))
                # Riempi la pagina con il titolo '1. PORTAFOGLIO ATTUALE' e con il corpo l'immagine appena salvata.
                print(f"\nsto aggiungendo l'{ws} alla presentazione.")
                self.document.add_section()
                # section = self.document.sections[2]
                # section.top_margin = shared.Cm(2.54)
                # section.right_margin = shared.Cm(1.5)
                # section.left_margin = shared.Cm(1.5)
                # section.bottom_margin = shared.Cm(2.54)
                paragraph_0 = self.document.add_paragraph(text='\n', style=None)
                paragraph_0.paragraph_format.space_after = shared.Pt(6)
                run_0 = paragraph_0.add_run('1. PORTAFOGLIO ATTUALE')
                run_0.bold = True
                run_0.font.name = 'Century Gothic'
                run_0.font.size = shared.Pt(14)
                run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph_1 = self.document.add_paragraph(style=None)
                run_1 = paragraph_1.add_run()
                width = 18 if hidden_columns==0 else 17 if hidden_columns==1 else 16 if hidden_columns==2 else 15 if hidden_columns==3 else 18
                run_1.add_picture('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\SAP\\Media\\'+ ws +'.bmp', width=shared.Cm(width))

    def commento_4(self):
        """Commento alla composizione del portafoglio."""
        self.document.add_section()
        paragraph_0 = self.document.add_paragraph(text='\n', style=None)
        run_0 = paragraph_0.add_run('1. PORTAFOGLIO ATTUALE')
        run_0.bold = True
        run_0.font.name = 'Century Gothic'
        run_0.font.size = shared.Pt(14)
        run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
        paragraph_1 = self.document.add_paragraph(text='\n', style=None)
        run_1 = paragraph_1.add_run('Commento generale sul portafoglio')
        run_1.bold = True
        run_1.font.name = 'Century Gothic'
        run_1.font.size = shared.Pt(14)
        run_1.font.color.rgb = shared.RGBColor(127, 127, 127)
        paragraph_2 = self.document.add_paragraph(style=None)
        paragraph_2.paragraph_format.space_after = shared.Pt(6)
        paragraph_2.paragraph_format.line_spacing_rule = 1

        # Carica il dataframe del portafoglio per estrane la composizione ed eventuali alert.
        df_portfolio = self.df_portfolio
        # Carica i dizionari delle macro e micro classi.
        dict_peso_macro = self.peso_macro()
        dict_peso_micro = self.peso_micro()
        # Crea il commento
        run_2 = paragraph_2.add_run(f'\nIl portafoglio attuale è investito ')
        run_2.font.name = 'Century Gothic'
        run_2.font.size = shared.Pt(10)
        dict_peso_strumenti_attivi = self.peso_strumenti()['strumenti_commento']
        for strumento, peso in dict_peso_strumenti_attivi.items():
            articolo = 'il ' if int(str(peso)[0]) in (2, 3, 4, 5, 6, 7, 9) else 'lo ' if int(str(peso)[0]) == 0 else "l'" if int(str(peso)[0]) == 8 else "l'" if int(str(peso)[0]) == 1 and peso < 2 else "il "
            if  strumento not in list(dict_peso_strumenti_attivi.keys())[-1]:
                run = paragraph_2.add_run(f"per {articolo}{str(round(peso,2)).replace('.',',')}% in {strumento}, ")
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(10)
            else:
                run = paragraph_2.add_run(f"e per {articolo}{str(round(peso,2)).replace('.',',')}% in {strumento}.")
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(10)
        
        # Alert mercati azionari
        paragraph_3 = self.document.add_paragraph()
        paragraph_3.paragraph_format.space_after = shared.Pt(6)
        paragraph_3.paragraph_format.line_spacing_rule = 1

        dict_peso_micro_azionarie = {item : dict_peso_micro[item]/dict_peso_macro[key] for key, value in self.dict_macro.items() for item in value if key=='Azionario'}
        pesi_limite_azionari = {'Azionario Europa' : 0.60, 'Azionario North America' : 0.60, 'Azionario Pacific' : 0.20, 'Azionario Emerging Markets' : 0.10}
        dict_alert_azionari = {k : True if v >= pesi_limite_azionari[k] else False for k, v in dict_peso_micro_azionarie.items()}
        nome_mercati_azionari = {'Azionario Europa' : 'europei', 'Azionario North America' : 'nordamericani', 'Azionario Pacific' : "nell'area del pacifico", 'Azionario Emerging Markets' : "emergenti"}
        
        if sum(dict_alert_azionari.values()) == 1:
            run_3 = paragraph_3.add_run(f"""Per la parte investita nel mercato azionario, si segnala l’eccessiva concentrazione verso i mercati {''.join([nome_mercati_azionari[k] for k, v in dict_alert_azionari.items() if v==True])}, che pesano {' '.join([str('il ' if str(dict_peso_micro_azionarie[k]*100)[0].startswith(('1','2','3','4','5','6','7','9')) and str(dict_peso_micro_azionarie[k]*100)[:2]!='11' else "l'")+str(round(dict_peso_micro_azionarie[k]*100, 2)).replace('.', ',')+str('%') for k, v in dict_alert_azionari.items() if v==True])} del comparto azionario, come indicato dal relativo warning. """)
            run_3.font.name = 'Century Gothic'
            run_3.font.size = shared.Pt(10)
        elif sum(dict_alert_azionari.values()) > 1:
            run_3 = paragraph_3.add_run(f"""Per la parte investita nel mercato azionario, si segnala l’eccessiva concentrazione verso i mercati {' e '.join([nome_mercati_azionari[k] for k, v in dict_alert_azionari.items() if v==True])} che pesano rispettivamente {', e '.join([str('il ' if str(dict_peso_micro_azionarie[k]*100)[0].startswith(('1','2','3','4','5','6','7','9')) and str(dict_peso_micro_azionarie[k]*100)[:2]!='11' else "l'")+str(round(dict_peso_micro_azionarie[k]*100, 2)).replace('.', ',')+str('%') for k, v in dict_alert_azionari.items() if v==True])} del comparto azionario, come indicato dai relativi warning. """)
            run_3.font.name = 'Century Gothic'
            run_3.font.size = shared.Pt(10)
        
        # Alert prodotti azionari
        quote_prodotti_azionari = df_portfolio.loc[df_portfolio['strumento']=='equity', ['nome', 'controvalore_in_euro']]
        quote_prodotti_azionari['peso_totale'] = quote_prodotti_azionari['controvalore_in_euro'] / df_portfolio['controvalore_in_euro'].sum()
        quote_prodotti_azionari['peso_su_azionario'] = quote_prodotti_azionari['peso_totale'] / dict_peso_macro['Azionario']
        if any(quote_prodotti_azionari['peso_su_azionario'] > 0.10):
            quote_prodotti_azionari_alert = dict(zip(quote_prodotti_azionari.loc[quote_prodotti_azionari['peso_su_azionario']>0.10, 'nome'].values, quote_prodotti_azionari.loc[quote_prodotti_azionari['peso_su_azionario']>0.10, 'peso_su_azionario']))
            run_3 = paragraph_3.add_run(f"""Riguardo agli strumenti azionari, si segnala il peso eccessivo di {' e '.join([k for k,v in quote_prodotti_azionari_alert.items()])} {'che pesa' if len(quote_prodotti_azionari_alert)==1 else 'che pesano rispettivamente'} {', e '.join([str('il ' if str(v*100)[0].startswith(('1','2','3','4','5','6','7','9')) and str(v*100)[:2]!='11' else "l'")+str(round(v*100, 2)).replace('.', ',')+str('%') for k,v in quote_prodotti_azionari_alert.items()])} dell’intero comparto azionario, come indicato {'dal relativo warning' if len(quote_prodotti_azionari_alert)==1 else 'dai relativi warning'} nella sezione di analisi del rischio dei singoli strumenti.""")
            run_3.font.name = 'Century Gothic'
            run_3.font.size = shared.Pt(10)
      
        # Alert mercati obbligazionari
        paragraph_4 = self.document.add_paragraph()
        paragraph_4.paragraph_format.space_after = shared.Pt(6)
        paragraph_4.paragraph_format.line_spacing_rule = 1

        dict_peso_micro_obbligazionarie = {item : dict_peso_micro[item]/dict_peso_macro[key] for key, value in self.dict_macro.items() for item in value if key=='Obbligazionario'}
        dict_peso_micro_obbligazionarie['Obbligazionario High Yield'] = dict_peso_micro_obbligazionarie['Obbligazionario Euro High Yield'] + dict_peso_micro_obbligazionarie['Obbligazionario Globale High Yield']
        dict_peso_micro_obbligazionarie.pop('Obbligazionario Euro High Yield', None)
        dict_peso_micro_obbligazionarie.pop('Obbligazionario Globale High Yield', None)
        pesi_limite_obbligazionari = {'Obbligazionario Euro Corporate' : 0.40, 'Obbligazionario Paesi Emergenti' : 0.20, 'Obbligazionario High Yield' : 0.20}
        dict_alert_obbligazionari = {k : True if v >= pesi_limite_obbligazionari[k] else False for k, v in dict_peso_micro_obbligazionarie.items() if k in pesi_limite_obbligazionari.keys()}
        nome_mercati_obbligazionari = {'Obbligazionario Euro Corporate' : 'corporate europeo', 'Obbligazionario Paesi Emergenti' : 'emergente', 'Obbligazionario High Yield' : 'high yield'}
        
        if sum(dict_alert_obbligazionari.values()) == 1:
            run_4 = paragraph_4.add_run(f"""Per la parte investita nel mercato obbligazionario, si segnala l’eccessiva concentrazione verso il comparto {''.join([nome_mercati_obbligazionari[k] for k, v in dict_alert_obbligazionari.items() if v==True])}, che pesa {' '.join([str('il ' if str(dict_peso_micro_obbligazionarie[k]*100)[0].startswith(('1','2','3','4','5','6','7','9')) and str(dict_peso_micro_obbligazionarie[k]*100)[:2]!='11' else "l'")+str(round(dict_peso_micro_obbligazionarie[k]*100, 2)).replace('.', ',')+str('%') for k, v in dict_alert_obbligazionari.items() if v==True])} del comparto obbligazionario, come indicato dal relativo warning. """)
            run_4.font.name = 'Century Gothic'
            run_4.font.size = shared.Pt(10)
        elif sum(dict_alert_obbligazionari.values()) > 1:
            run_4 = paragraph_4.add_run(f"""Per la parte investita nel mercato obbligazionario, si segnala l’eccessiva concentrazione verso i comparti {' e '.join([nome_mercati_obbligazionari[k] for k, v in dict_alert_obbligazionari.items() if v==True])} che pesano rispettivamente {', e '.join([str('il ' if str(dict_peso_micro_obbligazionarie[k]*100)[0].startswith(('1','2','3','4','5','6','7','9')) and str(dict_peso_micro_obbligazionarie[k]*100)[:2]!='11' else "l'")+str(round(dict_peso_micro_obbligazionarie[k]*100, 2)).replace('.', ',')+str('%') for k, v in dict_alert_obbligazionari.items() if v==True])} del comparto obbligazionario, come indicato dai relativi warning. """)
            run_4.font.name = 'Century Gothic'
            run_4.font.size = shared.Pt(10)
        
        # Alert prodotti obbligazionari
        quote_prodotti_obbligazionari = df_portfolio.loc[(df_portfolio['strumento']=='gov_bond') | (df_portfolio['strumento']=='corp_bond'), ['nome', 'controvalore_in_euro']]
        quote_prodotti_obbligazionari['peso_totale'] = quote_prodotti_obbligazionari['controvalore_in_euro'] / df_portfolio['controvalore_in_euro'].sum()
        quote_prodotti_obbligazionari['peso_su_obbligazionario'] = quote_prodotti_obbligazionari['peso_totale'] / dict_peso_macro['Obbligazionario']
        if any(quote_prodotti_obbligazionari['peso_su_obbligazionario'] > 0.10):
            quote_prodotti_obbligazionari_alert = dict(zip(quote_prodotti_obbligazionari.loc[quote_prodotti_obbligazionari['peso_su_obbligazionario']>0.10, 'nome'].values, quote_prodotti_obbligazionari.loc[quote_prodotti_obbligazionari['peso_su_obbligazionario']>0.10, 'peso_su_obbligazionario']))
            run_4 = paragraph_4.add_run(f"""Riguardo agli strumenti obbligazionari, si segnala il peso eccessivo di {' e '.join([k for k,v in quote_prodotti_obbligazionari_alert.items()])} {'che pesa' if len(quote_prodotti_obbligazionari_alert)==1 else 'che pesano rispettivamente'} {', e '.join([str('il ' if str(v*100)[0].startswith(('1','2','3','4','5','6','7','9')) and str(v*100)[:2]!='11' else "l'")+str(round(v*100, 2)).replace('.', ',')+str('%') for k,v in quote_prodotti_obbligazionari_alert.items()])} dell’intero comparto obbligazionario, come indicato {'dal relativo warning' if len(quote_prodotti_obbligazionari_alert)==1 else 'dai relativi warning'} nella sezione di analisi del rischio dei singoli strumenti.""")
            run_4.font.name = 'Century Gothic'
            run_4.font.size = shared.Pt(10)
        
    def analisi_di_portafoglio_5(self):
        """Incolla tabelle e grafici a torta."""
        self.document.add_section()
        paragraph_0 = self.document.add_paragraph(text='', style=None)
        run_0 = paragraph_0.add_run('\n')
        run_0 = paragraph_0.add_run('2. ANALISI DEL PORTAFOGLIO')
        run_0.bold = True
        run_0.font.name = 'Century Gothic'
        run_0.font.size = shared.Pt(14)
        run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
        table_0 = self.document.add_table(rows=9, cols=2)
        cell_1 = table_0.cell(0,0).merge(table_0.cell(0,1))
        paragraph_1 = cell_1.paragraphs[0]
        print('sto aggiungendo le macro categorie...')
        run_1 = paragraph_1.add_run('\nAnalisi per Macro Asset Class')
        run_1.bold = True
        run_1.font.name = 'Century Gothic'
        run_1.font.size = shared.Pt(12)
        run_1.font.color.rgb = shared.RGBColor(127, 127, 127)
        cell_2 = table_0.cell(1,0).merge(table_0.cell(1,1))
        paragraph_2 = cell_2.paragraphs[0]
        run_2 = paragraph_2.add_run()
        run_2.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(self.larghezza_pagina))
        cell_3 = table_0.cell(2,0)
        paragraph_3 = cell_3.paragraphs[0]
        run_3 = paragraph_3.add_run()
        excel2img.export_img(self.file_elaborato, self.path+r'\Media\macro.bmp', page='figure', _range="A1:C6")
        run_3.add_picture(self.path+r'\Media\macro.bmp', width=shared.Cm(9.5))
        cell_4 = table_0.cell(2,1)
        paragraph_4 = cell_4.paragraphs[0]
        paragraph_4.paragraph_format.alignment = 2
        run_4 = paragraph_4.add_run()
        run_4.add_picture(self.path+r'\Media\macro_pie.png', height=shared.Cm(4.2), width=shared.Cm(5.2))
        cell_5 = table_0.cell(3,0).merge(table_0.cell(3,1))
        paragraph_5 = cell_5.paragraphs[0]
        run_5 = paragraph_5.add_run()
        run_5.add_picture(self.path+r'\Media\default\macro_info.bmp', height=shared.Cm(1.64), width=shared.Cm(18))
        cell_6 = table_0.cell(4,0).merge(table_0.cell(4,1))
        paragraph_6 = cell_6.paragraphs[0]
        run_6 = paragraph_6.add_run('')
        run_6.font.size = shared.Pt(22)
        cell_7 = table_0.cell(5,0).merge(table_0.cell(5,1))
        paragraph_7 = cell_7.paragraphs[0]
        print('sto aggiungendo le micro categorie...')
        run_7 = paragraph_7.add_run('Analisi per Micro Asset Class')
        run_7.bold = True
        run_7.font.name = 'Century Gothic'
        run_7.font.size = shared.Pt(12)
        run_7.font.color.rgb = shared.RGBColor(127, 127, 127)
        cell_8 = table_0.cell(6,0).merge(table_0.cell(6,1))
        paragraph_8 = cell_8.paragraphs[0]
        run_8 = paragraph_8.add_run()
        run_8.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
        cell_9 = table_0.cell(7,0).merge(table_0.cell(7,1))
        paragraph_9 = cell_9.paragraphs[0]
        run_9 = paragraph_9.add_run()
        excel2img.export_img(self.file_elaborato, self.path+r'\Media\micro.bmp', page='figure', _range="I1:N16")
        run_9.add_picture(self.path+r'\Media\micro.bmp', height=shared.Cm(7), width=shared.Cm(18))
        cell_10 = table_0.cell(8,0).merge(table_0.cell(8,1))
        paragraph_10 = cell_10.paragraphs[0]
        run_10 = paragraph_10.add_run()
        run_10.add_picture(self.path+r'\Media\micro_bar.png', height=shared.Cm(5), width=shared.Cm(18))
        
    def analisi_di_portafoglio_6(self):
        """Incolla tabelle e grafici a torta"""
        self.document.add_section()
        paragraph_0 = self.document.add_paragraph(text='', style=None)
        run_0 = paragraph_0.add_run('\n')
        run_0 = paragraph_0.add_run('2. ANALISI DEL PORTAFOGLIO')
        run_0.bold = True
        run_0.font.name = 'Century Gothic'
        run_0.font.size = shared.Pt(14)
        run_0.font.color.rgb = shared.RGBColor(127, 127, 127)    
        table_0 = self.document.add_table(rows=9, cols=2)
        cell_1 = table_0.cell(0,0).merge(table_0.cell(0,1))
        paragraph_1 = cell_1.paragraphs[0]
        #paragraph_1.paragraph_format.space_after = 0
        print('sto aggiungendo gli strumenti...')
        run_1 = paragraph_1.add_run('\nAnalisi per Strumenti')
        run_1.bold = True
        run_1.font.name = 'Century Gothic'
        run_1.font.size = shared.Pt(12)
        run_1.font.color.rgb = shared.RGBColor(127, 127, 127)
        #row_1 = table_0.add_row()
        cell_2 = table_0.cell(1,0).merge(table_0.cell(1,1))
        paragraph_2 = cell_2.paragraphs[0]
        run_2 = paragraph_2.add_run()
        run_2.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
        cell_3 = table_0.cell(2,0)
        paragraph_3 = cell_3.paragraphs[0]
        run_3 = paragraph_3.add_run()
        excel2img.export_img(self.file_elaborato, self.path+r'\Media\strumenti.bmp', page='figure', _range="A18:D30")
        run_3.add_picture(self.path+r'\Media\strumenti.bmp', width=shared.Cm(10.5))
        cell_4 = table_0.cell(2,1)
        paragraph_4 = cell_4.paragraphs[0]
        paragraph_4.paragraph_format.alignment = 2
        run_4 = paragraph_4.add_run()
        run_4.add_picture(self.path+r'\Media\strumenti_pie.png', height=shared.Cm(4.2), width=shared.Cm(5.2))
        cell_5 = table_0.cell(5,0).merge(table_0.cell(5,1))
        paragraph_5 = cell_5.paragraphs[0]
        print('sto aggiungendo le valute...')
        run_5 = paragraph_5.add_run('\n\n\nAnalisi per Valute')
        run_5.bold = True
        run_5.font.name = 'Century Gothic'
        run_5.font.size = shared.Pt(12)
        run_5.font.color.rgb = shared.RGBColor(127, 127, 127)
        cell_6 = table_0.cell(6,0).merge(table_0.cell(6,1))
        paragraph_6 = cell_6.paragraphs[0]
        run_6 = paragraph_6.add_run()
        run_6.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
        cell_7 = table_0.cell(7,0)
        paragraph_7 = cell_7.paragraphs[0]
        run_7 = paragraph_7.add_run()
        excel2img.export_img(self.file_elaborato, self.path+r'\Media\valute.bmp', page='figure', _range="P1:S9")
        run_7.add_picture(self.path+r'\Media\valute.bmp', height=shared.Cm(3.7), width=shared.Cm(5))
        cell_8 = table_0.cell(7,1)
        paragraph_8 = cell_8.paragraphs[0]
        paragraph_8.paragraph_format.alignment = 2
        run_8 = paragraph_8.add_run()
        run_8.add_picture(self.path+r'\Media\valute_pie.png', height=shared.Cm(4.2), width=shared.Cm(5.2))
        cell_9 = table_0.cell(8,0).merge(table_0.cell(8,1))
        paragraph_9 = cell_9.paragraphs[0]
        run_9 = paragraph_9.add_run()
        run_9.add_picture(self.path+r'\Media\default\valute_info.bmp', width=shared.Cm(18))

    def analisi_strumenti_7(self):
        """Incolla tabelle di obbligazioni e azioni."""
        # Obbligazioni #
        df_portfolio = self.df_portfolio
        prodotti_obbligazionari = df_portfolio.loc[(df_portfolio['strumento']=='gov_bond') | (df_portfolio['strumento']=='corp_bond')]
        numero_prodotti_obbligazionari = prodotti_obbligazionari.nome.count()
        print('numero titoli obbligazionari:',numero_prodotti_obbligazionari)
        MAX_OBB_DES_PER_PAGINA = 54 # 54
        MAX_OBB_DATI_PER_PAGINA = 43 # 43
        MAX_AZIONI_PER_PAGINA = 48 # 48
        MAX_FONDI_PER_PAGINA = 42 # 42
        MAX_MAP_FONDI_PER_PAGINA = 75 #
        if numero_prodotti_obbligazionari > 0:
            # Carica il foglio obbligazioni
            obbligazioni = self.wb['obbligazioni']
            self.wb.active = obbligazioni
            # Nascondi le colonne del prezzo di carico e della variazione di prezzo
            hidden_columns = 0
            if all(df_portfolio['prezzo_di_carico'].isnull()):
                obbligazioni.column_dimensions['L'].hidden= True
                hidden_columns += 1
                obbligazioni.column_dimensions['O'].hidden= True
                hidden_columns += 1
            self.wb.save(self.file_elaborato)
            
            # Descrizione obbligazioni
            if numero_prodotti_obbligazionari > MAX_OBB_DES_PER_PAGINA and numero_prodotti_obbligazionari % MAX_OBB_DES_PER_PAGINA != 0:
                tabelle_des = int(numero_prodotti_obbligazionari // MAX_OBB_DES_PER_PAGINA + 1)
            elif numero_prodotti_obbligazionari > MAX_OBB_DES_PER_PAGINA and numero_prodotti_obbligazionari % MAX_OBB_DES_PER_PAGINA == 0:
                tabelle_des = int(numero_prodotti_obbligazionari // MAX_OBB_DES_PER_PAGINA)
            else:
                tabelle_des = 1
            print('tabelle_des:',tabelle_des)
            for tabella in range(1, tabelle_des+1):
                print(tabella)
                if tabella != tabelle_des:
                    excel2img.export_img(self.file_elaborato, self.path+r'\Media\obbligazioni_des_' + str(tabella) + '.bmp', page='obbligazioni', _range="B1:I"+str(MAX_OBB_DES_PER_PAGINA*tabella+1))
                    obbligazioni.row_dimensions.group(2+MAX_OBB_DES_PER_PAGINA*(tabella-1),MAX_OBB_DES_PER_PAGINA*tabella+1,hidden=True)
                    self.wb.save(self.file_elaborato)
                else:
                    excel2img.export_img(self.file_elaborato, self.path+r'\Media\obbligazioni_des_' + str(tabella) + '.bmp', page='obbligazioni', _range="B1:I"+str(prodotti_obbligazionari.nome.count()+1))
            
            obbligazioni.row_dimensions.group(1,MAX_OBB_DES_PER_PAGINA*tabelle_des,hidden=False)
            self.wb.save(self.file_elaborato)

            for tabella in range(1, tabelle_des+1):
                self.document.add_section()
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\n')
                run = paragraph.add_run('3. ANALISI DEI SINGOLI STRUMENTI')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(14)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\nCaratteristiche anagrafiche dei titoli obbligazionari')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(12)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\obbligazioni_des_'+str(tabella)+'.bmp', width=shared.Cm(18))

            # Dati obbligazioni
            # Calcolo numero titoli nell'ultima tabella
            if numero_prodotti_obbligazionari > MAX_OBB_DES_PER_PAGINA:
                if numero_prodotti_obbligazionari % MAX_OBB_DES_PER_PAGINA == 0:
                    num_obb_des_ultima_pagina = MAX_OBB_DES_PER_PAGINA
                elif numero_prodotti_obbligazionari % MAX_OBB_DES_PER_PAGINA != 0:
                    num_obb_des_ultima_pagina = numero_prodotti_obbligazionari % MAX_OBB_DES_PER_PAGINA
            elif numero_prodotti_obbligazionari <= MAX_OBB_DES_PER_PAGINA:
                num_obb_des_ultima_pagina = numero_prodotti_obbligazionari
            print("prodotti nell'ultima pagina:",num_obb_des_ultima_pagina)
            # Calcolo numero titoli nell'eventuale tabella sotto l'ultima
            if MAX_OBB_DATI_PER_PAGINA - num_obb_des_ultima_pagina - 7 > 0: # se rimane spazio sufficiente sotto l'ultima tabella precedente
                if (MAX_OBB_DATI_PER_PAGINA - num_obb_des_ultima_pagina - 7) < numero_prodotti_obbligazionari:
                    numerosita_tabella_obb_dati_sotto_la_precedente = MAX_OBB_DATI_PER_PAGINA - num_obb_des_ultima_pagina - 7
                else: # se tutte le obbligazioni ci stanno in quello spazio rimasto
                    numerosita_tabella_obb_dati_sotto_la_precedente = numero_prodotti_obbligazionari
            else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                numerosita_tabella_obb_dati_sotto_la_precedente = 0
            print("numerosità tabella obb dati sotto la precedente:",numerosita_tabella_obb_dati_sotto_la_precedente)  
            # Inserisci l'eventuale tabella sotto l'ultima
            if numerosita_tabella_obb_dati_sotto_la_precedente > 0:
                # Prima tabella dati obbligazioni
                excel2img.export_img(self.file_elaborato, self.path+r'\Media\obbligazioni_dati_0.bmp', page='obbligazioni', _range="K1:Q"+str(numerosita_tabella_obb_dati_sotto_la_precedente+1))
                obbligazioni.row_dimensions.group(2,numerosita_tabella_obb_dati_sotto_la_precedente+1,hidden=True)
                self.wb.save(self.file_elaborato)
                print(0)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\n\nCaratteristiche finanziarie dei titoli obbligazionari')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(12)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\obbligazioni_dati_0.bmp', width=shared.Cm(18) if hidden_columns==0 else shared.Cm(14.5))
            # Inserisci le tabelle rimanenti
            if numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente == 0: # tutti i titoli sono contenuti nella tabella sotto l'ultima
                tabelle_dati = 1
                print('tabelle_dati:',tabelle_dati)
            else:
                if numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente <= MAX_OBB_DATI_PER_PAGINA:
                    tabelle_dati = 1
                    # numerosita_tabella_obb_dati_sotto_la_precedente = numero_prodotti_obbligazionari
                elif numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente > MAX_OBB_DATI_PER_PAGINA:
                    if (numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) % MAX_OBB_DATI_PER_PAGINA != 0:
                        tabelle_dati = int((numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) // MAX_OBB_DATI_PER_PAGINA) + 1
                    else:
                        tabelle_dati = int((numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) // MAX_OBB_DATI_PER_PAGINA)
                    # numerosita_tabella_obb_dati_sotto_la_precedente = MAX_OBB_DATI_PER_PAGINA
                print('tabelle_dati:',tabelle_dati if numerosita_tabella_obb_dati_sotto_la_precedente == 0 else tabelle_dati+1)
                for tabella in range(1, tabelle_dati+1):
                    print(tabella)
                    if tabella != tabelle_dati:
                        excel2img.export_img(self.file_elaborato, self.path+'\Media\obbligazioni_dati_' + str(tabella) + '.bmp', page='obbligazioni', _range="K1:Q"+str(numerosita_tabella_obb_dati_sotto_la_precedente+MAX_OBB_DATI_PER_PAGINA*tabella+1))
                        obbligazioni.row_dimensions.group(2+MAX_OBB_DATI_PER_PAGINA*(tabella-1),numerosita_tabella_obb_dati_sotto_la_precedente+MAX_OBB_DATI_PER_PAGINA*tabella+1,hidden=True)
                        self.wb.save(self.file_elaborato)
                    else:
                        excel2img.export_img(self.file_elaborato, self.path+'\Media\obbligazioni_dati_' + str(tabella) + '.bmp', page='obbligazioni', _range="K1:Q"+str(numero_prodotti_obbligazionari+1))
                
                obbligazioni.row_dimensions.group(1,MAX_OBB_DATI_PER_PAGINA*(tabelle_dati+1),hidden=False)
                self.wb.save(self.file_elaborato)
                
                for tabella in range(1, tabelle_dati+1):
                    self.document.add_section()
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run('\n')
                    run = paragraph.add_run('3. ANALISI DEI SINGOLI STRUMENTI')
                    run.bold = True
                    run.font.name = 'Century Gothic'
                    run.font.size = shared.Pt(14)
                    run.font.color.rgb = shared.RGBColor(127, 127, 127)
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run('\nCaratteristiche finanziarie dei titoli obbligazionari')
                    run.bold = True
                    run.font.name = 'Century Gothic'
                    run.font.size = shared.Pt(12)
                    run.font.color.rgb = shared.RGBColor(127, 127, 127)
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run()
                    run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run()
                    run.add_picture(self.path+r'\Media\obbligazioni_dati_'+str(tabella)+'.bmp', width=shared.Cm(18) if hidden_columns==0 else shared.Cm(14.5))

        elif numero_prodotti_obbligazionari == 0:
            tabelle_dati = 0
            numerosita_tabella_obb_dati_sotto_la_precedente = 0


        # Azioni
        prodotti_azionari = df_portfolio.loc[df_portfolio['strumento']=='equity']
        numero_prodotti_azionari = prodotti_azionari.nome.count()
        print('numero titoli azionari:',numero_prodotti_azionari)
        if numero_prodotti_azionari > 0:
            # Carica il foglio azioni
            azioni = self.wb['azioni']
            self.wb.active = azioni
            # Nascondi le colonne del prezzo di carico e della variazione di prezzo
            hidden_columns = 0
            if all(df_portfolio['prezzo_di_carico'].isnull()):
                azioni.column_dimensions['F'].hidden= True
                hidden_columns += 1
                azioni.column_dimensions['H'].hidden= True
                hidden_columns += 1
                azioni.column_dimensions['I'].hidden= True
                hidden_columns += 1
            self.wb.save(self.file_elaborato)

            # Calcola il numero di prodotti nell'ultima pagina
            if numero_prodotti_obbligazionari > MAX_OBB_DES_PER_PAGINA:
                num_prodotti_ultima_pagina = (numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) % MAX_OBB_DATI_PER_PAGINA
            elif numero_prodotti_obbligazionari <= MAX_OBB_DES_PER_PAGINA:
                if numerosita_tabella_obb_dati_sotto_la_precedente == 0:
                    if numero_prodotti_obbligazionari <= MAX_OBB_DATI_PER_PAGINA:
                        num_prodotti_ultima_pagina = numero_prodotti_obbligazionari
                    elif numero_prodotti_obbligazionari > MAX_OBB_DATI_PER_PAGINA:
                        num_prodotti_ultima_pagina = (numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) % MAX_OBB_DATI_PER_PAGINA
                elif numerosita_tabella_obb_dati_sotto_la_precedente > 0:
                    if numerosita_tabella_obb_dati_sotto_la_precedente == numero_prodotti_obbligazionari:
                        num_prodotti_ultima_pagina = numero_prodotti_obbligazionari + numero_prodotti_obbligazionari
                    elif numerosita_tabella_obb_dati_sotto_la_precedente < numero_prodotti_obbligazionari:
                        num_prodotti_ultima_pagina = (numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) % MAX_OBB_DATI_PER_PAGINA
            elif numero_prodotti_obbligazionari == 0:
                num_prodotti_ultima_pagina = 0
            print("prodotti nell'ultima pagina:",num_prodotti_ultima_pagina)

            # Calcolo il numero titoli nell'eventuale tabella sotto l'ultima
            if num_prodotti_ultima_pagina == 0: # se non ci sono obbligazioni
                numerosita_tabella_azioni_sotto_la_precedente = 0
            elif num_prodotti_ultima_pagina == (numero_prodotti_obbligazionari * 2): # se le tabelle des e dati sono sulla stessa pagina
                if MAX_AZIONI_PER_PAGINA - num_prodotti_ultima_pagina - 22 > 0: # se rimane spazio sufficiente sotto le due tabelle precedenti
                    if (MAX_AZIONI_PER_PAGINA - num_prodotti_ultima_pagina - 22) < numero_prodotti_azionari: # ma non ce nè abbastanza per tutte le azioni
                        numerosita_tabella_azioni_sotto_la_precedente = MAX_AZIONI_PER_PAGINA - num_prodotti_ultima_pagina - 22
                    else: # e tutte le azioni ci stanno in quello spazio rimasto
                        numerosita_tabella_azioni_sotto_la_precedente = numero_prodotti_azionari
                else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                    numerosita_tabella_azioni_sotto_la_precedente = 0
            else: # ci sono obbligazioni ma le tabelle des e dati non sono sulla stessa pagina
                if MAX_AZIONI_PER_PAGINA - num_prodotti_ultima_pagina - 16 > 0: # se rimane spazio sufficiente sotto le due tabelle precedenti
                    if (MAX_AZIONI_PER_PAGINA - num_prodotti_ultima_pagina - 16) < numero_prodotti_azionari: # ma non ce nè abbastanza per tutte le azioni
                        numerosita_tabella_azioni_sotto_la_precedente = MAX_AZIONI_PER_PAGINA - num_prodotti_ultima_pagina - 16
                    else: # e tutte le azioni ci stanno in quello spazio rimasto
                        numerosita_tabella_azioni_sotto_la_precedente = numero_prodotti_azionari
                else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                    numerosita_tabella_azioni_sotto_la_precedente = 0
            print("numerosità tabella azioni sotto la precedente:",numerosita_tabella_azioni_sotto_la_precedente)

            # Inserisci l'eventuale tabella sotto l'ultima
            if numerosita_tabella_azioni_sotto_la_precedente > 0:
                # Prima tabella dati azioni
                excel2img.export_img(self.file_elaborato, self.path+r'\Media\azioni_0.bmp', page='azioni', _range="B1:K"+str(numerosita_tabella_azioni_sotto_la_precedente+1))
                azioni.row_dimensions.group(2,numerosita_tabella_azioni_sotto_la_precedente+1,hidden=True)
                self.wb.save(self.file_elaborato)
                print(0)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\n\nCaratteristiche dei titoli azionari')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(12)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\azioni_0.bmp', width=shared.Cm(18.5) if hidden_columns==0 else shared.Cm(14.5))
            # Inserisci le tabelle rimanenti
            if numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente == 0: # tutti i titoli sono contenuti nella tabella sotto l'ultima
                tabelle_azioni = 1
                print('tabelle_azioni:',tabelle_azioni)
            else:
                if numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente <= MAX_AZIONI_PER_PAGINA:
                    tabelle_azioni = 1
                    # numerosita_tabella_azioni_sotto_la_precedente = numero_prodotti_azionari
                elif numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente > MAX_AZIONI_PER_PAGINA:
                    if (numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente) % MAX_AZIONI_PER_PAGINA != 0:
                        tabelle_azioni = int((numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente) // MAX_AZIONI_PER_PAGINA) + 1
                    else:
                        tabelle_azioni = int((numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente) // MAX_AZIONI_PER_PAGINA)
                    # numerosita_tabella_azioni_sotto_la_precedente = MAX_AZIONI_PER_PAGINA
                print('tabelle_azioni:',tabelle_azioni if numerosita_tabella_azioni_sotto_la_precedente == 0 else tabelle_azioni+1)
                for tabella in range(1, tabelle_azioni+1):
                    print(tabella)
                    if tabella != tabelle_azioni:
                        excel2img.export_img(self.file_elaborato, self.path+r'\Media\azioni_' + str(tabella) + '.bmp', page='azioni', _range="B1:K"+str(numerosita_tabella_azioni_sotto_la_precedente+MAX_AZIONI_PER_PAGINA*tabella+1))
                        azioni.row_dimensions.group(2+MAX_AZIONI_PER_PAGINA*(tabella-1),numerosita_tabella_azioni_sotto_la_precedente+MAX_AZIONI_PER_PAGINA*tabella+1,hidden=True)
                        self.wb.save(self.file_elaborato)
                    else:
                        excel2img.export_img(self.file_elaborato, self.path+r'\Media\azioni_' + str(tabella) + '.bmp', page='azioni', _range="B1:K"+str(numero_prodotti_azionari+1))
                for tabella in range(1, tabelle_azioni+1):
                    self.document.add_section()
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run('\n')
                    run = paragraph.add_run('3. ANALISI DEI SINGOLI STRUMENTI')
                    run.bold = True
                    run.font.name = 'Century Gothic'
                    run.font.size = shared.Pt(14)
                    run.font.color.rgb = shared.RGBColor(127, 127, 127)
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run('\nCaratteristiche dei titoli azionari')
                    run.bold = True
                    run.font.name = 'Century Gothic'
                    run.font.size = shared.Pt(12)
                    run.font.color.rgb = shared.RGBColor(127, 127, 127)
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run()
                    run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run()
                    run.add_picture(self.path+r'\Media\azioni_'+str(tabella)+'.bmp', width=shared.Cm(18.5) if hidden_columns==0 else shared.Cm(14.5))

            azioni.row_dimensions.group(1,MAX_AZIONI_PER_PAGINA*(tabelle_dati+1),hidden=False)
            self.wb.save(self.file_elaborato)


        # Fondi
        prodotti_gestiti = df_portfolio.loc[df_portfolio['strumento']=='fund']
        numero_prodotti_gestiti = prodotti_gestiti.nome.count()
        print('numero fondi:',numero_prodotti_gestiti)
        if numero_prodotti_gestiti > 0:
            # Carica il foglio fondi
            fondi = self.wb['fondi']
            self.wb.active = fondi
            # Nascondi le colonne del prezzo di carico e della variazione di prezzo
            hidden_columns = 0
            if all(df_portfolio['prezzo_di_carico'].isnull()):
                fondi.column_dimensions['F'].hidden= True
                hidden_columns += 1
                fondi.column_dimensions['H'].hidden= True
                hidden_columns += 1
                fondi.column_dimensions['I'].hidden= True
                hidden_columns += 1
            self.wb.save(self.file_elaborato)

            # Calcola il numero di prodotti nell'ultima pagina
            if numero_prodotti_obbligazionari == 0 and numero_prodotti_azionari == 0: # non ci sono obbligazioni nè azioni
                num_prodotti_ultima_pagina = 0
            elif numero_prodotti_obbligazionari == 0 and numero_prodotti_azionari > 0: # ci sono azioni ma non obbligazioni
                if numero_prodotti_azionari % MAX_AZIONI_PER_PAGINA == 0:
                    num_prodotti_ultima_pagina = MAX_AZIONI_PER_PAGINA
                elif numero_prodotti_azionari % MAX_AZIONI_PER_PAGINA != 0:
                    num_prodotti_ultima_pagina = numero_prodotti_azionari % MAX_AZIONI_PER_PAGINA
            elif numero_prodotti_obbligazionari > 0 and numero_prodotti_azionari == 0: # ci sono obbligazioni ma non azioni
                if numero_prodotti_obbligazionari > MAX_OBB_DES_PER_PAGINA:
                    num_prodotti_ultima_pagina = (numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) % MAX_OBB_DATI_PER_PAGINA
                elif numero_prodotti_obbligazionari <= MAX_OBB_DES_PER_PAGINA:
                    if numerosita_tabella_obb_dati_sotto_la_precedente == 0:
                        if numero_prodotti_obbligazionari <= MAX_OBB_DATI_PER_PAGINA:
                            num_prodotti_ultima_pagina = numero_prodotti_obbligazionari
                        elif numero_prodotti_obbligazionari > MAX_OBB_DATI_PER_PAGINA:
                            num_prodotti_ultima_pagina = (numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) % MAX_OBB_DATI_PER_PAGINA
                    elif numerosita_tabella_obb_dati_sotto_la_precedente > 0:
                        if numerosita_tabella_obb_dati_sotto_la_precedente == numero_prodotti_obbligazionari:
                            num_prodotti_ultima_pagina = numero_prodotti_obbligazionari + numero_prodotti_obbligazionari
                        elif numerosita_tabella_obb_dati_sotto_la_precedente < numero_prodotti_obbligazionari:
                            num_prodotti_ultima_pagina = (numero_prodotti_obbligazionari - numerosita_tabella_obb_dati_sotto_la_precedente) % MAX_OBB_DATI_PER_PAGINA
            elif numero_prodotti_obbligazionari > 0 and numero_prodotti_azionari > 0: # ci sono obbligazioni e azioni
                if numerosita_tabella_azioni_sotto_la_precedente == numero_prodotti_azionari: # le azioni stanno tutte sotto l'ultima tabella precedente
                    num_prodotti_ultima_pagina = num_prodotti_ultima_pagina + numerosita_tabella_azioni_sotto_la_precedente
                elif numerosita_tabella_azioni_sotto_la_precedente == 0: # le tabelle delle azioni cominciano da pagina nuova
                    num_prodotti_ultima_pagina = numero_prodotti_azionari % MAX_AZIONI_PER_PAGINA
                elif numerosita_tabella_azioni_sotto_la_precedente < numero_prodotti_azionari: # le azioni non stanno tutte sotto l'ultima tabella precedente
                    if (numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente) % MAX_AZIONI_PER_PAGINA == 0: # l'ultima pagina contiene il numero massimo di azioni
                        num_prodotti_ultima_pagina = MAX_AZIONI_PER_PAGINA
                    else: # l'ultima pagina contiene un numero di azioni inferiore al numero massimo di azioni per pagina
                        num_prodotti_ultima_pagina = (numero_prodotti_azionari - numerosita_tabella_azioni_sotto_la_precedente) % MAX_AZIONI_PER_PAGINA
            print("prodotti nell'ultima pagina:",num_prodotti_ultima_pagina)

            # Calcolo il numero dei fondi da inserire nell'eventuale tabella sotto l'ultima
            if num_prodotti_ultima_pagina == 0: # se non ci sono obbligazioni nè azioni
                numerosita_tabella_fondi_sotto_la_precedente = 0
            elif numero_prodotti_obbligazionari == 0: # non ci sono obbligazioni
                if MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 7 > 0: # se rimane spazio sufficiente sotto la tabella precedente
                    if (MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 7) <= numero_prodotti_gestiti: # ma non ce nè abbastanza per tutti i fondi
                        numerosita_tabella_fondi_sotto_la_precedente = MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 7
                    else: # e tutti i fondi ci stanno in quello spazio rimasto
                        numerosita_tabella_fondi_sotto_la_precedente = numero_prodotti_gestiti
                else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                    numerosita_tabella_fondi_sotto_la_precedente = 0
            elif numero_prodotti_azionari == 0: # non ci sono azioni
                if num_prodotti_ultima_pagina == (numero_prodotti_obbligazionari * 2): # se le tabelle des e dati sono sulla stessa pagina
                    if MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 15 > 0: # se rimane spazio sufficiente sotto le due tabelle precedenti
                        if (MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 15) < numero_prodotti_gestiti: # ma non ce nè abbastanza per tutti i fondi
                            numerosita_tabella_fondi_sotto_la_precedente = MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 15
                        else: # e tutti i fondi ci stanno in quello spazio rimasto
                            numerosita_tabella_fondi_sotto_la_precedente = numero_prodotti_gestiti
                    else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                        numerosita_tabella_fondi_sotto_la_precedente = 0
                else: # ci sono obbligazioni ma le tabelle des e dati non sono sulla stessa pagina
                    if MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 12 > 0: # se rimane spazio sufficiente sotto le due tabelle precedenti
                        if (MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 12) < numero_prodotti_gestiti: # ma non ce nè abbastanza per tutti i fondi
                            numerosita_tabella_fondi_sotto_la_precedente = MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 12
                        else: # e tutti i fondi ci stanno in quello spazio rimasto
                            numerosita_tabella_fondi_sotto_la_precedente = numero_prodotti_gestiti
                    else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                        numerosita_tabella_fondi_sotto_la_precedente = 0
            elif num_prodotti_ultima_pagina == (numero_prodotti_obbligazionari * 2) + numero_prodotti_azionari: # se le tabelle delle obbligazioni e delle azioni sono sulla stessa pagina
                if MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 24 > 0: # se rimane spazio sufficiente sotto le tre tabelle precedenti
                    if (MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 24) < numero_prodotti_gestiti: # ma non ce nè abbastanza per tutti i fondi
                        numerosita_tabella_fondi_sotto_la_precedente = MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 24
                    elif MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 24 >= numero_prodotti_gestiti: # e tutti i fondi ci stanno in quello spazio rimasto
                        numerosita_tabella_fondi_sotto_la_precedente = numero_prodotti_gestiti
                else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                    numerosita_tabella_fondi_sotto_la_precedente = 0
            else: # ci sono obbligazioni e/o azioni, ma le tabelle non sono sulla stessa pagina
                if numero_prodotti_azionari <= numerosita_tabella_azioni_sotto_la_precedente: # l'ultima pagina ha la tabella dati e la tabella azioni
                    if MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 18 > 0: # se rimane spazio sufficiente sotto le tabelle precedenti
                        if (MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 18) < numero_prodotti_gestiti: # ma non ce nè abbastanza per tutti i fondi
                            numerosita_tabella_fondi_sotto_la_precedente = MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 18
                        elif MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 18 >= numero_prodotti_gestiti: # e tutti i fondi ci stanno in quello spazio rimasto
                            numerosita_tabella_fondi_sotto_la_precedente = numero_prodotti_gestiti
                    else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                        numerosita_tabella_fondi_sotto_la_precedente = 0
                elif numero_prodotti_azionari > numerosita_tabella_azioni_sotto_la_precedente: # l'ultima pagina ha una sola tabella di azioni
                    if MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 6 > 0: # se rimane spazio sufficiente sotto la tabella precedente
                        if (MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 6) < numero_prodotti_gestiti: # ma non ce nè abbastanza per tutti i fondi
                            numerosita_tabella_fondi_sotto_la_precedente = MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 6
                        elif MAX_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 6 >= numero_prodotti_gestiti: # e tutti i fondi ci stanno in quello spazio rimasto
                            numerosita_tabella_fondi_sotto_la_precedente = numero_prodotti_gestiti
                    else: # se non rimane spazio a sufficienza per una tabella sotto la precedente
                        numerosita_tabella_fondi_sotto_la_precedente = 0
            print("numerosità tabella fondi sotto la precedente:",numerosita_tabella_fondi_sotto_la_precedente)

            # Inserisci l'eventuale tabella sotto l'ultima
            if numerosita_tabella_fondi_sotto_la_precedente > 0:
                # Prima tabella dati fondi
                excel2img.export_img(self.file_elaborato, self.path+r'\Media\fondi_0.bmp', page='fondi', _range="B1:I"+str(numerosita_tabella_fondi_sotto_la_precedente+1))
                fondi.row_dimensions.group(2,numerosita_tabella_fondi_sotto_la_precedente+1,hidden=True)
                self.wb.save(self.file_elaborato)
                print(0)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\n\nCaratteristiche finanziarie dei fondi comuni di investimento')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(12)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\fondi_0.bmp', width=shared.Cm(18.5) if hidden_columns==0 else shared.Cm(13.5))
            # Inserisci le tabelle rimanenti
            if numero_prodotti_gestiti - numerosita_tabella_fondi_sotto_la_precedente == 0: # tutti i titoli sono contenuti nella tabella sotto l'ultima
                tabelle_fondi = 1
                print('tabelle_fondi:',tabelle_fondi)
            else:
                if numero_prodotti_gestiti - numerosita_tabella_fondi_sotto_la_precedente <= MAX_FONDI_PER_PAGINA:
                    tabelle_fondi = 1
                    # numerosita_tabella_fondi_sotto_la_precedente = numero_prodotti_gestiti
                elif numero_prodotti_gestiti - numerosita_tabella_fondi_sotto_la_precedente > MAX_FONDI_PER_PAGINA:
                    if (numero_prodotti_gestiti - numerosita_tabella_fondi_sotto_la_precedente) % MAX_FONDI_PER_PAGINA != 0:
                        tabelle_fondi = int((numero_prodotti_gestiti - numerosita_tabella_fondi_sotto_la_precedente) // MAX_FONDI_PER_PAGINA) + 1
                    else:
                        tabelle_fondi = int((numero_prodotti_gestiti - numerosita_tabella_fondi_sotto_la_precedente) // MAX_FONDI_PER_PAGINA)
                    # numerosita_tabella_fondi_sotto_la_precedente = MAX_FONDI_PER_PAGINA
                print('tabelle_fondi:',tabelle_fondi if numerosita_tabella_fondi_sotto_la_precedente == 0 else tabelle_fondi+1)
                for tabella in range(1, tabelle_fondi+1):
                    print(tabella)
                    if tabella != tabelle_fondi:
                        excel2img.export_img(self.file_elaborato, self.path+r'\Media\fondi_' + str(tabella) + '.bmp', page='fondi', _range="B1:I"+str(numerosita_tabella_fondi_sotto_la_precedente+MAX_FONDI_PER_PAGINA*tabella+1))
                        fondi.row_dimensions.group(2+MAX_FONDI_PER_PAGINA*(tabella-1),numerosita_tabella_fondi_sotto_la_precedente+MAX_FONDI_PER_PAGINA*tabella+1,hidden=True)
                        self.wb.save(self.file_elaborato)
                    else:
                        excel2img.export_img(self.file_elaborato, self.path+r'\Media\fondi_' + str(tabella) + '.bmp', page='fondi', _range="B1:I"+str(numero_prodotti_gestiti+1))
                for tabella in range(1, tabelle_fondi+1):
                    self.document.add_section()
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run('\n')
                    run = paragraph.add_run('3. ANALISI DEI SINGOLI STRUMENTI')
                    run.bold = True
                    run.font.name = 'Century Gothic'
                    run.font.size = shared.Pt(14)
                    run.font.color.rgb = shared.RGBColor(127, 127, 127)
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run('\nCaratteristiche finanziarie dei fondi comuni di investimento')
                    run.bold = True
                    run.font.name = 'Century Gothic'
                    run.font.size = shared.Pt(12)
                    run.font.color.rgb = shared.RGBColor(127, 127, 127)
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run()
                    run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                    paragraph = self.document.add_paragraph(text='', style=None)
                    run = paragraph.add_run()
                    run.add_picture(self.path+r'\Media\fondi_'+str(tabella)+'.bmp', width=shared.Cm(18.5) if hidden_columns==0 else shared.Cm(13.5))

            fondi.row_dimensions.group(1,MAX_FONDI_PER_PAGINA*(tabelle_fondi+1),hidden=False)
            self.wb.save(self.file_elaborato)

            # Mappatura fondi #
            numero_prodotti_gestiti_map = numero_prodotti_gestiti + 2
            if numero_prodotti_gestiti_map > MAX_MAP_FONDI_PER_PAGINA and numero_prodotti_gestiti_map % MAX_MAP_FONDI_PER_PAGINA != 0:
                tabelle_map_fondi = int(numero_prodotti_gestiti_map // MAX_MAP_FONDI_PER_PAGINA + 1)
            elif numero_prodotti_gestiti_map > MAX_MAP_FONDI_PER_PAGINA and numero_prodotti_gestiti_map % MAX_MAP_FONDI_PER_PAGINA == 0:
                tabelle_map_fondi = int(numero_prodotti_gestiti_map // MAX_MAP_FONDI_PER_PAGINA)
            else:
                tabelle_map_fondi = 1
            print('tabelle_map_fondi:',tabelle_map_fondi)
            for tabella in range(1, tabelle_map_fondi+1):
                print(tabella)
                if tabella != tabelle_map_fondi:
                    excel2img.export_img(self.file_elaborato, self.path+r'\Media\map_fondi_' + str(tabella) + '.bmp', page='fondi', _range="L1:Z"+str(MAX_MAP_FONDI_PER_PAGINA*tabella+1))
                    fondi.row_dimensions.group(2+MAX_MAP_FONDI_PER_PAGINA*(tabella-1),MAX_MAP_FONDI_PER_PAGINA*tabella+1,hidden=True)
                    self.wb.save(self.file_elaborato)
                else:
                    excel2img.export_img(self.file_elaborato, self.path+r'\Media\map_fondi_' + str(tabella) + '.bmp', page='fondi', _range="L1:Z"+str(numero_prodotti_gestiti_map+1))
            
            fondi.row_dimensions.group(1,MAX_MAP_FONDI_PER_PAGINA*tabelle_map_fondi,hidden=False)
            self.wb.save(self.file_elaborato)

            for tabella in range(1, tabelle_map_fondi+1):
                self.document.add_section()
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\n')
                run = paragraph.add_run('3. ANALISI DEI SINGOLI STRUMENTI')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(14)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\nMappatura dei fondi comuni di investimento')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(12)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\map_fondi_info.bmp', width=shared.Cm(18.5))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\map_fondi_'+str(tabella)+'.bmp', width=shared.Cm(18.5))

            fondi.row_dimensions.group(1,MAX_MAP_FONDI_PER_PAGINA*(tabelle_map_fondi+1),hidden=False)
            self.wb.save(self.file_elaborato)

            # Calcola numero fondi mappati nell'ultima pagina
            if numero_prodotti_gestiti_map <= MAX_MAP_FONDI_PER_PAGINA:
                num_prodotti_ultima_pagina = numero_prodotti_gestiti_map
            elif numero_prodotti_gestiti_map > MAX_MAP_FONDI_PER_PAGINA:
                if numero_prodotti_gestiti_map % MAX_MAP_FONDI_PER_PAGINA != 0:
                    num_prodotti_ultima_pagina = numero_prodotti_gestiti_map % MAX_MAP_FONDI_PER_PAGINA
                elif numero_prodotti_gestiti_map % MAX_MAP_FONDI_PER_PAGINA == 0:
                    num_prodotti_ultima_pagina = MAX_MAP_FONDI_PER_PAGINA
            print("numerosità ultima tabella mappatura fondi:",num_prodotti_ultima_pagina)

            if MAX_MAP_FONDI_PER_PAGINA - num_prodotti_ultima_pagina - 24 > 0: # c'è spazio per inserire il grafico a barre
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\map_fondi_bar.png', width=shared.Cm(18.5))
            else: # non c'è spazio per inserire il grafico a barre
                self.document.add_section()
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\n')
                run = paragraph.add_run('3. ANALISI DEI SINGOLI STRUMENTI')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(14)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run('\nMappatura dei fondi comuni di investimento')
                run.bold = True
                run.font.name = 'Century Gothic'
                run.font.size = shared.Pt(12)
                run.font.color.rgb = shared.RGBColor(127, 127, 127)
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\barra.png', width=shared.Cm(18.1))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\map_fondi_info.bmp', width=shared.Cm(18.5))
                paragraph = self.document.add_paragraph(text='', style=None)
                run = paragraph.add_run()
                run.add_picture(self.path+r'\Media\default\map_fondi_bar.png', width=shared.Cm(18.5))

    def rischio_8(self):
        """Inserisci la parte di rischio"""
        self.document.add_section()

    def note_metodologiche_9(self):
        """Inserisci le note metodologiche e le avvertenze più la pagina di chiusura."""
        # Note metodologiche 1
        self.document.add_section()
        paragraph_0 = self.document.add_paragraph(text='\n', style=None)
        run_0 = paragraph_0.add_run('5. NOTE METODOLOGICHE')
        run_0.bold = True
        run_0.font.name = 'Century Gothic'
        run_0.font.size = shared.Pt(14)
        run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
        paragraph_1 = self.document.add_paragraph(text='\n', style=None)
        paragraph_1.paragraph_format.alignment = 3
        paragraph_1.paragraph_format.space_after = shared.Pt(6)
        run_1 = paragraph_1.add_run('Nello svolgimento di questa analisi ci siamo avvalsi della documentazione fornitaci da Azimut Wealth Management. Tali informazioni saranno assunte come attendibili da Benchmark&Style. Sono inoltre stati analizzati i dati di mercato tratti da MorningStar e Bloomberg.')
        run_1.bold = True
        run_1.font.name = 'Century Gothic'
        run_1.font.size = shared.Pt(10)
        paragraph_2 = self.document.add_paragraph(text='', style=None)
        paragraph_2.paragraph_format.alignment = 3
        paragraph_2.paragraph_format.space_after = shared.Pt(6)
        run_2_1 = paragraph_2.add_run('Sezione A): ')
        run_2_1.bold = True
        run_2_1.font.name = 'Century Gothic'
        run_2_1.font.size = shared.Pt(10)
        run_2_2 = paragraph_2.add_run('nella sezione A viene riportata la composizione del portafoglio. Le quantità dei titoli/prodotti in portafoglio fanno riferimento ai dati di reportistica cliente ricevuta, mentre il controvalore fa riferimento al valore degli stessi alla data di analisi del portafoglio.')
        run_2_2.font.name = 'Century Gothic'
        run_2_2.font.size = shared.Pt(10)
        paragraph_3 = self.document.add_paragraph(text='', style=None)
        paragraph_3.paragraph_format.alignment = 3
        paragraph_3.paragraph_format.space_after = shared.Pt(6)
        run_3_1 = paragraph_3.add_run('Sezione B): ')
        run_3_1.bold = True
        run_3_1.font.name = 'Century Gothic'
        run_3_1.font.size = shared.Pt(10)
        run_3_2 = paragraph_3.add_run('le segnalazioni riportate nelle pagine dedicate all’analisi del portafoglio evidenziano eventuali concen-trazioni del portafoglio stesso su specifiche tipologie di prodotto/strumento, su particolari asset class e sulle valute; per le concentrazioni in valuta si distingue tra investimenti in Euro e in valute diverse dall’Euro.')
        run_3_2.font.name = 'Century Gothic'
        run_3_2.font.size = shared.Pt(10)
        paragraph_4 = self.document.add_paragraph(text='', style=None)
        paragraph_4.paragraph_format.alignment = 0
        paragraph_4.paragraph_format.space_after = shared.Pt(6)
        run_4_1 = paragraph_4.add_run('Sezione C): ')
        run_4_1.bold = True
        run_4_1.font.name = 'Century Gothic'
        run_4_1.font.size = shared.Pt(10)
        run_4_2 = paragraph_4.add_run('per la descrizione sintetica dei titoli obbligazionari sono stati adottati i seguenti criteri:')
        run_4_2.font.name = 'Century Gothic'
        run_4_2.font.size = shared.Pt(10)
        paragraph_5 = self.document.add_paragraph(text='', style=None)
        paragraph_5.paragraph_format.alignment = 0
        paragraph_5.paragraph_format.space_after = shared.Pt(6)
        run_5 = paragraph_5.add_run('- i rating indicano il giudizio espresso dalle principali società di rating (Moodys, Standard and Poor’s e Fitch);')
        run_5.font.name = 'Century Gothic'
        run_5.font.size = shared.Pt(10)
        paragraph_6 = self.document.add_paragraph(text='', style=None)
        paragraph_6.paragraph_format.alignment = 0
        paragraph_6.paragraph_format.space_after = shared.Pt(6)
        run_6 = paragraph_6.add_run('- con la tipologia FIXED facciamo riferimento ad obbligazioni a tasso fisso, mentre VARIABLE indica obbligazioni a tasso variabile;')
        run_6.font.name = 'Century Gothic'
        run_6.font.size = shared.Pt(10)
        paragraph_7 = self.document.add_paragraph(text='', style=None)
        paragraph_7.paragraph_format.alignment = 0
        paragraph_7.paragraph_format.space_after = shared.Pt(6)
        run_7 = paragraph_7.add_run('- la duration ponderata fa riferimento alla sola componente obbligazionaria del portafoglio.')
        run_7.font.name = 'Century Gothic'
        run_7.font.size = shared.Pt(10)
        paragraph_8 = self.document.add_paragraph(text='', style=None)
        paragraph_8.paragraph_format.alignment = 3
        paragraph_8.paragraph_format.space_after = shared.Pt(6)
        run_8 = paragraph_8.add_run('La duration di un portafoglio titoli (o di un singolo titolo) è la vita finanziaria media ponderata dei titoli presenti nel portafoglio (o di un singolo titolo); si tratta di un valore espresso in anni ed indica l’arco temporale entro il quale l’investitore rientrerà in possesso del capitale inizialmente investito, tenendo conto sia delle cedole, sia del rimborso finale.')
        run_8.font.name = 'Century Gothic'
        run_8.font.size = shared.Pt(10)
        paragraph_9 = self.document.add_paragraph(text='', style=None)
        paragraph_9.paragraph_format.alignment = 3
        paragraph_9.paragraph_format.space_after = shared.Pt(6)
        run_9_1 = paragraph_9.add_run('Sezione D): ')
        run_9_1.bold = True
        run_9_1.font.name = 'Century Gothic'
        run_9_1.font.size = shared.Pt(10)
        run_9_2 = paragraph_9.add_run('per il calcolo del portafoglio di rischio si è proceduto come segue: tutti gli strumenti e i prodotti presenti in portafoglio sono stati ricondotti a delle specifiche asset class. Conoscendo il grado di rischio delle singole asset class e il grado di correlazione che le lega fra loro, si procede al calcolo del grado di rischio complessivo del portafoglio.')
        run_9_2.font.name = 'Century Gothic'
        run_9_2.font.size = shared.Pt(10)
        paragraph_10 = self.document.add_paragraph(text='', style=None)
        paragraph_10.paragraph_format.alignment = 3
        paragraph_10.paragraph_format.space_after = shared.Pt(6)
        run_10_1 = paragraph_10.add_run('Sezione E): ')
        run_10_1.bold = True
        run_10_1.font.name = 'Century Gothic'
        run_10_1.font.size = shared.Pt(10)
        run_10_2 = paragraph_10.add_run('nell’analisi di dettaglio si riportano le segnalazioni a livello di singolo prodotto/strumento distinguendo tra segnalazioni di concentrazione e di liquidabilità. Per quanto riguarda la concentrazione, si evidenzia sia l’eventuale eccesso di concentrazione in termini di peso sulla rispettiva asset class, sia quella sul singolo emittente.')
        run_10_2.font.name = 'Century Gothic'
        run_10_2.font.size = shared.Pt(10)
        paragraph_11 = self.document.add_paragraph(text='', style=None)
        paragraph_11.paragraph_format.alignment = 3
        paragraph_11.paragraph_format.space_after = shared.Pt(6)
        run_11 = paragraph_11.add_run('Con riferimento alla liquidabilità degli strumenti, si evidenzia l’eventuale basso grado di liquidabilità dell’investimento, dovuto ai tempi lunghi sia a costi elevati di smobilizzo. In questa sezione, inoltre, viene fornito un giudizio a livello di singolo prodotto/strumento:')
        run_11.font.name = 'Century Gothic'
        run_11.font.size = shared.Pt(10)
        paragraph_12 = self.document.add_paragraph(text='', style=None)
        paragraph_12.paragraph_format.alignment = 3
        paragraph_12.paragraph_format.space_after = shared.Pt(6)
        run_12_1 = paragraph_12.add_run('- “poco efficiente” ')
        run_12_1.bold = True
        run_12_1.font.name = 'Century Gothic'
        run_12_1.font.size = shared.Pt(10)
        run_12_2 = paragraph_12.add_run('(in termini di rapporto rendimento/rischio);')
        run_12_2.font.name = 'Century Gothic'
        run_12_2.font.size = shared.Pt(10)
        paragraph_13 = self.document.add_paragraph(text='', style=None)
        paragraph_13.paragraph_format.alignment = 3
        paragraph_13.paragraph_format.space_after = shared.Pt(6)
        run_13_1 = paragraph_13.add_run('- “neutro” ')
        run_13_1.bold = True
        run_13_1.font.name = 'Century Gothic'
        run_13_1.font.size = shared.Pt(10)
        run_13_2 = paragraph_13.add_run('(in termini di rapporto rendimento/rischio);')
        run_13_2.font.name = 'Century Gothic'
        run_13_2.font.size = shared.Pt(10)
        paragraph_14 = self.document.add_paragraph(text='', style=None)
        paragraph_14.paragraph_format.alignment = 3
        paragraph_14.paragraph_format.space_after = shared.Pt(6)
        run_14 = paragraph_14.add_run('- “default”;')
        run_14.bold = True
        run_14.font.name = 'Century Gothic'
        run_14.font.size = shared.Pt(10)
        paragraph_15 = self.document.add_paragraph(text='', style=None)
        paragraph_15.paragraph_format.alignment = 3
        paragraph_15.paragraph_format.space_after = shared.Pt(6)
        run_15 = paragraph_15.add_run('- “scaduto”.')
        run_15.bold = True
        run_15.font.name = 'Century Gothic'
        run_15.font.size = shared.Pt(10)
        # Note metodologiche 2
        self.document.add_section()
        paragraph_0 = self.document.add_paragraph(text='\n', style=None)
        run_0 = paragraph_0.add_run('5. NOTE METODOLOGICHE')
        run_0.bold = True
        run_0.font.name = 'Century Gothic'
        run_0.font.size = shared.Pt(14)
        run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
        paragraph_1 = self.document.add_paragraph(text='\n', style=None)
        paragraph_1.paragraph_format.alignment = 0
        paragraph_1.paragraph_format.line_spacing_rule = 1
        paragraph_1.paragraph_format.space_after = shared.Pt(6)
        run_1_1 = paragraph_1.add_run('In particolare, per quanto riguarda le sezioni B) ed E), di seguito viene riportata una tabella riassuntiva con gli «alert» di concentrazione analizzati:\n\n')
        run_1_1.font.name = 'Century Gothic'
        run_1_1.font.size = shared.Pt(10)
        run_1_2 = paragraph_1.add_run()
        run_1_2.add_picture(self.path+r'\Media\default\note_metodologiche.bmp', width=shared.Cm(18.5))
        paragraph_2 = self.document.add_paragraph(text='\n', style=None)
        paragraph_2.paragraph_format.alignment = 3
        paragraph_2.paragraph_format.line_spacing_rule = 1
        paragraph_2.paragraph_format.space_after = shared.Pt(6)
        run_2 = paragraph_2.add_run('Nel processo di analisi del portafoglio, quindi, vengono dapprima calcolate le esposizioni delle singole micro asset class rispetto alla macro asset class di riferimento. Per esempio, per quanto riguarda il compar-to azionario, se la componente europea pesa più del 60% dell’intera esposizione azionaria, scatterà il warning «!C», se pesa più del 70%, l’alert «!!C», altrimenti se rappresenta più dell’80%, il warning «!!!C».')
        run_2.font.name = 'Century Gothic'
        run_2.font.size = shared.Pt(10)
        paragraph_3 = self.document.add_paragraph(text='', style=None)
        paragraph_3.paragraph_format.alignment = 3
        paragraph_3.paragraph_format.line_spacing_rule = 1
        paragraph_3.paragraph_format.space_after = shared.Pt(6)
        run_3 = paragraph_3.add_run('Per quanto concerne l’esposizione dei singoli strumenti (azioni ed obbligazioni) sulla macro asset class di riferimento, ad esempio se un titolo azionario pesa più del 10%, del 20% o del 30% dell’intero comparto equity, allora scatteranno rispettivamente uno dei seguenti warning: «!C», «!!C», «!!!C». Per le obbligazioni tale controllo non viene svolto nel caso di bond governativi emessi dai paesi sviluppati.')
        run_3.font.name = 'Century Gothic'
        run_3.font.size = shared.Pt(10)
        paragraph_4 = self.document.add_paragraph(text='', style=None)
        paragraph_4.paragraph_format.alignment = 3
        paragraph_4.paragraph_format.line_spacing_rule = 1
        paragraph_4.paragraph_format.space_after = shared.Pt(6)
        run_4 = paragraph_4.add_run('Per le classi di strumenti più complessi (obbligazioni strutturate, Hedge Fund, Altri) si confronta, da un lato, l’esposizione rispetto all’intero portafoglio, dall’altro i livelli percentuali critici indicati nella terza tabella.')
        run_4.font.name = 'Century Gothic'
        run_4.font.size = shared.Pt(10)
        paragraph_5 = self.document.add_paragraph(text='', style=None)
        paragraph_5.paragraph_format.alignment = 3
        paragraph_5.paragraph_format.line_spacing_rule = 1
        paragraph_5.paragraph_format.space_after = shared.Pt(6)
        run_5 = paragraph_5.add_run('Per quanto riguarda l’esposizione valutaria, infine, scatterà un warning quando il peso delle valute diverse dall’Euro è maggiore del 40%, 50% o 60% dell’intero portafoglio.')
        run_5.font.name = 'Century Gothic'
        run_5.font.size = shared.Pt(10)
        # Note metodologiche 3
        self.document.add_section()
        paragraph_0 = self.document.add_paragraph(text='\n', style=None)
        run_0 = paragraph_0.add_run('AVVERTENZE')
        run_0.bold = True
        run_0.font.name = 'Century Gothic'
        run_0.font.size = shared.Pt(14)
        run_0.font.color.rgb = shared.RGBColor(127, 127, 127)
        paragraph_1 = self.document.add_paragraph(text='\n', style=None)
        paragraph_1.paragraph_format.alignment = 3
        paragraph_1.paragraph_format.line_spacing_rule = 1
        paragraph_1.paragraph_format.space_after = shared.Pt(6)
        run_1 = paragraph_1.add_run('Questo documento è stato prodotto a solo scopo informativo. Di conseguenza non è fornita alcuna garanzia circa la completezza, l’accuratezza, l’affidabilità delle informazioni in esso contenute.')
        run_1.font.name = 'Century Gothic'
        run_1.font.size = shared.Pt(10)
        paragraph_2 = self.document.add_paragraph(text='', style=None)
        paragraph_2.paragraph_format.alignment = 3
        paragraph_2.paragraph_format.line_spacing_rule = 1
        paragraph_2.paragraph_format.space_after = shared.Pt(6)
        run_2 = paragraph_2.add_run('Di conseguenza nessuna garanzia, esplicita o implicita è fornita da parte o per conto della Società o di alcuno dei suoi membri, dirigenti, funzionari o impiegati o altre persone. Né la Società né alcuno dei suoi membri, dirigenti, funzionari o impiegati o altre persone che agiscano per conto della Società accetta alcuna responsabilità per qualsiasi perdita potesse derivare dall’uso di questa presentazione o dei suoi contenuti o altrimenti connesso con la presentazione e i suoi contenuti.')
        run_2.font.name = 'Century Gothic'
        run_2.font.size = shared.Pt(10)
        paragraph_3 = self.document.add_paragraph(text='', style=None)
        paragraph_3.paragraph_format.alignment = 3
        paragraph_3.paragraph_format.line_spacing_rule = 1
        paragraph_3.paragraph_format.space_after = shared.Pt(6)
        run_3 = paragraph_3.add_run('Le informazioni e opinioni contenute in questa presentazione sono aggiornate alla data indicata sulla presentazione e possono essere cambiate senza preavviso.')
        run_3.font.name = 'Century Gothic'
        run_3.font.size = shared.Pt(10)
        paragraph_4 = self.document.add_paragraph(text='', style=None)
        paragraph_4.paragraph_format.alignment = 3
        paragraph_4.paragraph_format.line_spacing_rule = 1
        paragraph_4.paragraph_format.space_after = shared.Pt(6)
        run_4 = paragraph_4.add_run('Questo documento non costituisce una sollecitazione o un’offerta e nessuna parte di esso può costituire la base o il riferimento per qualsivoglia contratto o impegno.')
        run_4.font.name = 'Century Gothic'
        run_4.font.size = shared.Pt(10)
        paragraph_5 = self.document.add_paragraph(text='', style=None)
        paragraph_5.paragraph_format.alignment = 3
        paragraph_5.paragraph_format.line_spacing_rule = 1
        paragraph_5.paragraph_format.space_after = shared.Pt(6)
        run_5 = paragraph_5.add_run('All’investimento descritto è associato il rischio di andamento dei tassi di interesse nominali e reali, dell’inflazione, dei cambi e dei mercati azionari e il rischio legato al possibile deterioramento del merito di credito degli emittenti.')
        run_5.font.name = 'Century Gothic'
        run_5.font.size = shared.Pt(10)
        paragraph_6 = self.document.add_paragraph(text='', style=None)
        paragraph_6.paragraph_format.alignment = 3
        paragraph_6.paragraph_format.line_spacing_rule = 1
        paragraph_6.paragraph_format.space_after = shared.Pt(6)
        run_6 = paragraph_6.add_run('Relativamente all’investimento in AZ Fund e Azimut Fondi si rimanda ai prospetti informativi dei relativi fondi che raccomandiamo di leggere prima della sottoscrizione.')
        run_6.font.name = 'Century Gothic'
        run_6.font.size = shared.Pt(10)
        paragraph_7 = self.document.add_paragraph(text='', style=None)
        paragraph_7.paragraph_format.alignment = 3
        paragraph_7.paragraph_format.line_spacing_rule = 1
        paragraph_7.paragraph_format.space_after = shared.Pt(6)
        run_7 = paragraph_7.add_run('L’investimento descritto non assicura il mantenimento del capitale, né offre garanzie di rendimento.')
        run_7.font.name = 'Century Gothic'
        run_7.font.size = shared.Pt(10)
        # Pagina di chiusura
        self.document.add_section()
        header = self.document.sections[-1].header
        header.is_linked_to_previous = False
        section = self.document.sections[-1]
        left_margin = 0.60
        right_margin = 0.60
        top_margin = 0.45
        bottom_margin = 0.45
        section.left_margin = shared.Cm(left_margin)
        section.right_margin = shared.Cm(right_margin)
        section.top_margin = shared.Cm(top_margin)
        section.bottom_margin = shared.Cm(bottom_margin)
        section.header_distance = shared.Cm(0)
        section.footer_distance = shared.Cm(0)
        paragraph_0 = self.document.add_paragraph(text='', style=None)
        paragraph_0.alignment = 1
        paragraph_0.add_run().add_picture(self.path+'\Media\default\pagina_di_chiusura.jpg', height=shared.Cm(28.8), width=shared.Cm(20.14))
    
    def salva_file_portafoglio(self):
        """Salva il file excel."""
        self.wb.save(self.file_elaborato)

    def salva_file_presentazione(self):
        """Salva il file della presentazione con nome."""
        try:
            self.document.save(self.file_presentazione)
        except PermissionError:
            print(f'\nChiudi il file {self.file_presentazione}')



if __name__ == "__main__":
    start = time.time()
    # separa le tre classi in tre file diversi
    PTF = 'ptf_20.xlsx'
    PTF_ELABORATO = PTF[:-5] + '_elaborato.xlsx'
    PATH = r'C:\Users\Administrator\Desktop\Sbwkrq\SAP'

    __ = Elaborazione(file_elaborato=PTF_ELABORATO)
    __.new_agglomerato()
    # # __.old_agglomerato()
    __.figure()
    __.mappatura_fondi()
    __.sintesi()
    __.salva_file_portafoglio()
    
    ___ = Presentazione(tipo_sap='completo', file_elaborato=PTF_ELABORATO, file_presentazione='ahah.docx', page_height = 29.7, page_width = 21, top_margin = 2.5, bottom_margin = 2.5, left_margin = 1.5, right_margin = 1.5)
    ___.copertina_1()
    ___.indice_2()
    ___.portafoglio_attuale_3()
    # # ___.new_portafoglio_attuale_3()
    # # ___.old_portafoglio_attuale_3()
    # ___.commento_4()
    ___.analisi_di_portafoglio_5()
    ___.analisi_di_portafoglio_6()
    ___.analisi_strumenti_7()
    # ___.rischio_8()
    # ___.note_metodologiche_9()

    ___.salva_file_portafoglio()
    ___.salva_file_presentazione()
    end = time.time()
    print("Elapsed time: ", end - start, 'seconds')