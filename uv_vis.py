'''
This program merges data from .csv files with the read values of a
spectrophotometer Genesys 10S UV-Vis. Furthermore, the program applies
Saviztky-Golay filter in order to smooth data and calculate the derivative
spectroscopy in Savizty-Golay values.
'''

# Imports
import pandas as pd
import glob
import xlsxwriter
import os
from scipy.signal import savgol_filter
import string
import datetime


# Functions
def acquire_files():
    '''
    Collect all .csv files and assign them to a variable using 'glob' library.
    It's important to define a path in which the .csv files will be obtained.
    Returns a list with these files.
    '''
    path = r'C:/Users/UFC/Desktop/Espectrofotometro/A analisar'
    files = glob.glob(path + '/*.csv')
    files.sort(key=os.path.getmtime)
    return files


def date_today():
    '''
    Get the today's date and returns it. The date will be used in the .xlsx
    filename.
    '''
    date = datetime.date.today()
    date_today = (date.day, date.month, date.year)
    return date_today


def create_dicts(files, window_length=11, polyorder=2):
    '''
    Creates three dicts with absorbances values and a list with wavelength
    values. The first dict (files_dict) has the read values.
    Second dict (savgol_dict) has the Savizty-Golay values. The third
    dict (deriv_dict) has the derivative values of Savitzky-Golay filtered
    values. Here 'pandas' library is used to filter the data of original files
    in order to get the wavelength and absorbance values.
    Returns a tuple with these data: wavelength_list, files_dict, savgol_dict
    and deriv_dict.
    '''

    files_dict = dict()
    savgol_dict = dict()
    deriv_dict = dict()

    for file in files:
        # Below it's used [:-4] notation to delete '.csv' of the filename
        file_name = os.path.basename(file)[:-4]
        df = pd.read_csv(file, sep=';', encoding='utf-8')
        df = df.dropna(axis='columns')
        df = df.iloc[1:]
        df.columns = ['nm', 'A']
        df['nm'] = df['nm'].astype(int)
        df['A'] = df['A'].str.replace(',', '.').astype(float)

        wavelength_list = list(set(df['nm']))
        files_dict[file_name] = list(df['A'].values)
        savgol_dict[file_name] = savgol_filter(list(df['A'].values), 11, 2)

        abs_list_sav = savgol_filter(df['A'].values, 11, 2).tolist()
        partial_results = dict(zip(wavelength_list, abs_list_sav))
        deriv_dict[file_name] = list(
                                     create_derivatives(
                                                        **{
                                                           str(k): v
                                                           for k, v
                                                           in partial_results.items()}
                                                        )
                                    )
    return (wavelength_list, files_dict, savgol_dict, deriv_dict)


def create_derivatives(deriv_delta=3, **partial_results):
    '''
    Creates derivative values of Saviztky-Golay absorbance values using
    the principles of derivative spectroscopy. The deriv_delta variable
    is delta lambda of derivative calculus (Default = 3).
    '''
    partial_results = {int(k): v for k, v in partial_results.items()}
    absorbances = list(partial_results.values())
    derivative_dict = dict()
    wv_list = list(partial_results.keys())
    for index, wavelength in enumerate(wv_list):
        if index < deriv_delta:
            derivative_dict[wavelength] = ((absorbances[index + deriv_delta] -
                                            absorbances[index]) /
                                           (deriv_delta)) * 1/2
        elif index + deriv_delta > (wv_list[-1] - wv_list[0])/3:
            derivative_dict[wavelength] = ((absorbances[index] -
                                            absorbances[index - deriv_delta]) /
                                           (deriv_delta)) * 1/2
        else:
            derivative_dict[wavelength] = ((absorbances[index + 1] -
                                            absorbances[index - 1]) /
                                           (deriv_delta)) * 1/2
    return derivative_dict.values()


def create_workbook(workbook_name):
    '''
    Creates the workbook which will receive the data.
    '''
    date = date_today()
    workbook = xlsxwriter.Workbook('C:/Users/UFC/Desktop/'
                                   'Espectrofotometro/Analisados/'
                                   f'{workbook_name}_'
                                   f'{date[0]:02d}{date[1]:02d}'
                                   f'{date[2]:04d}.xlsx')
    return workbook


def create_worksheet(workbook, workbook_name, cat_calc, chart_subtype,
                     deriv=False, remainder=None, init_deriv=None,
                     delta=None, *wavelength_list, **files_dict):
    '''
    The function in charge to write the worksheets inside the workbook.
    The parameters are:
    # workbook         -> the created workbook
    # workbook_name    -> the name the user typed when using the program
    # cat_calc         -> the category of the data. If the worksheet is the
    worksheet with data from Savitzky-Golay values, the argument is '_sav_gol'.
    If the worksheet is the derivative worksheet, this argument is '_deriv'
    # chart_subtype    -> the subtype value of chart. The options are
    'straight' or 'smooth'
    # deriv            -> a boolean parameter to define if will be made the
    derivative calculus
    # remainder        -> like 'init_deriv' parameter, it's a variable to
    control the chart representation of derivative spectroscopy. It's used
    to create a new wavelength list with filtered wavelength values.
    # init_deriv       -> like 'remainder', it's a variable to control the
    chart representation of derivative spectroscopy. Defines the initial
    value of "abs_list", a list with filtered absorbances of derivative
    spectroscopy calculus
    # delta             -> delta lambda value
    # wavelength_list   -> list with read wavelength values
    # files_dict        -> a dict with data that will be analyzed. The key
    value of this dict is the filename and the values are the absorbance
    read values.
    The first for loop is responsible to write the filenames to identify the
    absorbance values.
    The first 'if' statement writes the wavelength values in first column and
    the absorbance values in other columns.
    The second 'if' statement writes is in charge to process data in order to
    write derivative spectroscopy values.
    The for loops in string.ascii values are responsible to create new columns
    labels. This will be important if a lot of files will be read, to avoid
    errors on the creation of charts.
    The final part of this function creates a chart for each worksheet.
    '''
    worksheet = workbook.add_worksheet(
                                       f'{(workbook_name + cat_calc).replace(" ", "")}'
                                       )

    worksheet.write('A1', 'nm')
    row = 0
    col = 1
    for k in files_dict.keys():
        worksheet.write(row, col, k)
        col += 1
    col = 0
    if not deriv:
        row += 1
        worksheet.write_column(1, 0, wavelength_list)
        for v in files_dict.values():
            worksheet.write_column(row, col + 1, v)
            col += 1
    row = 1
    col = 0
    if deriv:
        abs_list = []
        wavelength_list = [i for i in wavelength_list if i % 3 == remainder]
        for values in files_dict.values():
            abs_list.append(list(values)[init_deriv::3])
        worksheet.write_column(row, col, wavelength_list)
        col += 1
        for ab in abs_list:
            worksheet.write_column(row, col, ab)
            col += 1

    chart = workbook.add_chart({'type': 'scatter', 'subtype': chart_subtype})
    letters = list(string.ascii_uppercase)
    for i in list(string.ascii_uppercase):
        letters.append('A' + i)
    for i in list(string.ascii_uppercase):
        letters.append('B' + i)
    for i in range(len(files_dict)):
        chart.add_series({
            'categories': f'{(workbook_name + cat_calc).replace(" ", "")}'
                          f'!$A$2:$A${len(wavelength_list) + 1}',
            'values': f'{(workbook_name + cat_calc).replace(" ", "")}'
                      f'!${letters[i+1]}$2:'
                      f'${letters[i+1]}${len(wavelength_list) + 1}',
        })
    chart.set_title({'name': f'{workbook_name + cat_calc}'})
    chart.set_x_axis({
        'name': 'nm',
        'min': wavelength_list[0],
        'max': wavelength_list[len(wavelength_list) - 1],
    })
    chart.set_y_axis({'name': 'A', })
    chart.set_size({'width': 1300, 'height': 600})
    worksheet.insert_chart('E4', chart)


def workbook_launcher(workbook):
    '''
    Launches the workbook in Windows OS computers.
    '''
    date = date_today()
    os.startfile('C:/Users/UFC/Desktop/'
                 'Espectrofotometro/Analisados/'
                 f'{workbook_name}_'
                 f'{date[0]:02d}{date[1]:02d}'
                 f'{date[2]:04d}.xlsx')


# Init
print('#'*20, 'Compilador de .csv do GENESYS 10S UV-Vis', '#'*20)
workbook_name = str(input('Digite um nome para a planilha (.xlsx): '))
files = acquire_files()
print('Você quer mudar os parâmetros de Saviztky-Golay?')
print('Por default, os parâmetros são: ')
print('Window length: 11\nPolyorder: 2\nDerivada: 1\nDelta: 3.0')
op_savgol = str(input('Digite "S" (sem aspas) se sim: ')).strip().upper()
if op_savgol == 'S':
    window_length = int(input('Window length (Inteiro positivo ímpar): '))
    polyorder = int(input('Polyorder (2, 3 ou 4): '))
    deriv = int(input('Derivada (Ordem da derivação): '))
    delta = float(input('Delta (Intervalo de comprimento de onda): '))
    data = create_dicts(files, window_length, polyorder, deriv, delta)
else:
    data = create_dicts(files)

print('Aguarde sua planilha ser preparada...')

workbook = create_workbook(workbook_name)

for i in range(6):
    if i == 0:
        create_worksheet(
                         workbook, workbook_name,
                         '', 'straight', False, None, None, None,
                         *data[0], **data[1]
                         )
    elif i == 1:
        create_worksheet(
                         workbook, workbook_name,
                         '_sav_gol', 'smooth', False, None, None, None,
                         *data[0], **data[2]
                         )
    elif i == 2:
        create_worksheet(
                         workbook, workbook_name,
                         '_deriv', 'smooth', False, None, None, None,
                         *data[0], **data[3]
                         )
    elif i == 3:
        create_worksheet(
                         workbook, workbook_name,
                         f'_deriv{i-2}', 'smooth', True, 2, 0, 3,
                         *data[0], **data[3]
                         )
    elif i == 4:
        create_worksheet(
                         workbook, workbook_name,
                         f'_deriv{i-2}', 'smooth', True, 0, 1, 3,
                         *data[0], **data[3]
                         )
    elif i == 5:
        create_worksheet(
                         workbook, workbook_name,
                         f'_deriv{i-2}', 'smooth', True, 1, 2, 3,
                         *data[0], **data[3]
                         )
workbook.close()
workbook_launcher(workbook)
