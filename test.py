import os
from datetime import date, datetime, timedelta
import numpy as np
import pandas as pd
with os.add_dll_directory('C:\\Users\\Administrator\\Desktop\\Sbwkrq\\_blpapi'):
    import blpapi
from xbbg import blp
import matplotlib.pyplot as plt
import seaborn as sns
# fondi = ['IT0005334062', 'LU0717749021', 'LU1097689522', 'LU1439460020', 'LU0115139569', 'LU0432616901', 'LU1388496355', 
#     'LU0937587227', 'LU0814405493', 'LU0937587144', 'LU0937586252', 'LU1005158651', 'LU0349157924', 'LU1713432752', 
#     'IT0005186108', 'IT0005185985',
#     ] # testa per fondi pubblici che hanno tutta la storia
# fondi = ['IT0005334062', 'LU0717749021', 'LU2368226135'] # testa per fondi che non hanno tutta la storia
# fondi = ['IT0005334062', 'LU0717749021', 'IT0005408718'] # testa per fondi che su bloomberg non esistono (privati)
fondi = ['IT0005334062', 'LU0717749021', 'LU1097689522', 'LU1439460020', 'LU0115139569', 'LU0432616901', 'LU1388496355', 
    'LU0937587227', 'LU0814405493', 'LU0937587144', 'LU2368226135', 'LU0937586252', 'LU1005158651', 'IT0005408718',
    'LU0349157924', 'LU1713432752', 'IT0005186108', 'IT0005185985', 'GB00B15KXV33', 'JE00B1VS3770', 'LU0539144625',
    'LU0335987003',
    ]
first_day_of_current_month = date.today().replace(day=1)
last_day_of_previous_month = first_day_of_current_month - timedelta(days=1)
last_day_of_previous_month_of_previous_year = last_day_of_previous_month.replace(year=last_day_of_previous_month.year - 1)

day = datetime(2022, 10, 31).replace(day=1)
day_of_previous_month = day - timedelta(days=1)
day_of_previous_month_of_previous_year = day_of_previous_month.replace(year=day_of_previous_month.year - 1)


serie_storica = blp.bdh(['/isin/' + fondo for fondo in fondi], flds="DAY_TO_DAY_TOT_RETURN_GROSS_DVDS",
    start_date=last_day_of_previous_month_of_previous_year, end_date=last_day_of_previous_month, Days="A", Period="W")
serie_storica.columns = [column[0][6:] for column in serie_storica.columns] # rinomina solo i fondi che esistono
# serie_storica.to_excel('ahah.xlsx')
print(serie_storica)
# print(len(serie_storica.index)) # = 52
corr_matrix = serie_storica.corr(min_periods=len(serie_storica.index))
print(corr_matrix)
upper_triangle_corr_matrix = np.triu(corr_matrix)
lower_triangle_corr_matrix = np.tril(corr_matrix)
sns.heatmap(data=corr_matrix, vmin=-1, vmax=+1, annot=True, cmap="turbo")#, mask=upper_triangle_corr_matrix)
plt.xticks(rotation=75, fontsize=9)
plt.show()