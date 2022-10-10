########################################################################################################################
#
#                                       INCOMPLETE - PENDING OPTIMIZATION
#
########################################################################################################################
import datetime

import os
import datetime

import pandas as pd
import win32com
import xlwings as xw

from os.path import exists

from PIL import ImageGrab

from emailSend import email

global file_path
global mod_file_paths


########################################################################################################################
#
#                                               EXCEL FILE GENERATION
#
########################################################################################################################

# get the image of the
def getImage(sht, xlrange):
    win32c = win32com.client.constants
    sht.range(xlrange).api.CopyPicture(Format=win32c.xlBitmap)  # 2 copies as when it is printed

    img = ImageGrab.grabclipboard()
    img.save(r'report.png')


# Format the inventory files by removing blank/unnecessary lines
def formatExcel(hj_path, sap_path):
    global mod_file_paths
    mod_file_paths = []
    #
    # Remove unwanted lines from the Excel file
    #
    with xw.App(visible=False) as app:
        wb = app.books.open(hj_path)

        # check if this value exists on the first line of the file
        # if yes then remove that line
        # if no then do nothing
        check = '<div'
        if check in wb.sheets['Sheet1'].range('A1').value:
            wb.sheets['Sheet1'].range('1:1').delete()

            wb.save(hj_path)


def toExcel(df_sap, df_hj):
    #
    # Export the SAP and HJ data to Excel file
    #
    global file_path
    date = datetime.datetime.now()
    desktop_path = "D:\Sending Confirmation"
    file_path = desktop_path + "\\Sending Confirmation Report " + date.strftime("%d %b %I %M %p") + ".xlsx"
    print(file_path)

    if exists(file_path):
        wb = xw.Book(file_path)
    else:
        wb = xw.Book()

    # define the Excel sheets
    sht_sap = wb.sheets['Sheet1']
    sht_sap.name = 'MB51'

    sht_hj = wb.sheets.add()
    sht_hj.name = 'HJ GDN'

    df_sap = df_sap.applymap(lambda x: str(x) if isinstance(x, datetime.time) else x)
    # paste dataframes in table form
    sht_sap.range("A1").options(index=False).value = df_sap
    sht_hj.range("A1").options(index=False).value = df_hj

    wb.save(file_path)


def toExcelVariance(df_variance, df_variance_2):
    #
    # Export Variance Identified to Excel
    #
    global file_path

    if exists(file_path):
        wb = xw.Book(file_path)
    else:
        wb = xw.Book()

    # define the Excel sheets
    sht_var = wb.sheets.add()
    sht_var.name = 'Variance'

    # paste dataframes in table form
    sht_var.range("A2").value = 'Quantity Summary Comparison in YDS'
    sht_var.range("A2").expand("right").api.Font.Bold = True

    sht_var.range("A4").options(index=False).value = df_variance

    table1 = sht_var.tables.add(source=sht_var.range('A4').expand())

    # -----------------------------------------------------------------

    last_table = sht_var.tables[-1]
    last_row = last_table.range.last_cell.row

    sht_var.range(last_row + 2, 1).value = 'Quantity Summary Comparison in Rolls'
    sht_var.range(last_row + 2, 1).expand("right").api.Font.Bold = True

    sht_var.range(last_row + 4, 1).options(index=False).value = df_variance_2

    table2 = sht_var.tables.add(source=sht_var.range(last_row + 4, 1).expand(), table_style_name='TableStyleMedium5')

    # -----------------------------------------------------------------

    last_table = sht_var.tables[-1]
    last_row = last_table.range.last_cell.row

    sht_var.range(last_row + 2, 1).value = 'Sending Confirmation Summary'
    sht_var.range(last_row + 2, 1).expand("right").api.Font.Bold = True

    # create summary dataframe and convert it to a table in excel
    df_summary = pd.DataFrame()
    df_summary['Category'] = ['YDS', 'Rolls']

    df_summary['SAP'] = [df_variance['SAP Pick Quantity (YDS)'].sum(),
                         df_variance_2['SAP Pick Quantity (Roll)'].sum()]

    df_summary['HJ'] = [df_variance['HJ GDN Quantity (YDS)'].sum(),
                        df_variance_2['HJ GDN Quantity (Roll)'].sum()]

    df_summary['Variance'] = [df_variance['SAP Pick Quantity (YDS)'].sum() - df_variance['HJ GDN Quantity (YDS)'].sum(),
                              df_variance_2['SAP Pick Quantity (Roll)'].sum() - df_variance_2[
                                  'HJ GDN Quantity (Roll)'].sum()]

    # identifying the larger quantity
    qty = []
    for var in df_summary['Variance']:
        if var > 0:
            qty.append("HJ+")
        elif var < 0:
            qty.append("SAP+")
        else:
            qty.append("Quantity Tally")

    df_summary['Remarks'] = qty

    sht_var.range(last_row + 4, 1).options(index=False).value = df_summary

    table3 = sht_var.tables.add(source=sht_var.range(last_row + 4, 1).expand(), table_style_name='TableStyleMedium7')

    expand_cells = sht_var.range("A1").expand("right")
    expand_cells.row_height = 15
    expand_cells.column_width = 30

    sht_var.range("B:B").number_format = "#,##0.000"
    sht_var.range("C:C").number_format = "#,##0.000"

    image_range = sht_var.range((1, 1),
                                (table3.range.last_cell.row, table3.range.last_cell.column))

    getImage(sht_var, image_range)
    wb.save(file_path)


########################################################################################################################
#
#                                                 DATA EXTRACTION
#
########################################################################################################################

# get the SAP and HJ Reports

def sendingConfirmation(file_paths, status, progress_bar):
    global mod_file_paths

    try:
        formatExcel(file_paths[1], file_paths[0])

        # get the SAP and HJ Reports
        df_sap = pd.read_excel(file_paths[0])
        df_hj = pd.read_excel(file_paths[1])

        # Creating Material and trimmed load ID code
        material = []
        load_id = []
        references = []

        for mat in df_hj['Item Number']:
            temp = mat.split('-')
            print(temp)
            material.append(temp[1])

        for load in df_hj['Load Id']:
            temp_load = load.split('-')
            load_id.append(temp_load[0])

        for ref in df_sap['Reference']:
            temp_load = str(ref).split('-')
            references.append(temp_load[0])

        df_hj['Material'] = material
        df_hj['Load ID Trimmed'] = load_id
        df_sap['Reference'] = references

        # get the dataframe for the sap report and re-arrange the ID column
        df_sap["Unique"] = df_sap['Material'].astype(str) + "-" + df_sap["Reference"].astype(str)
        sap_ids = df_sap['Unique'].to_list()

        # remove the decimal, nan values & random '-' IDs
        new_sap_ids = [i.replace('.0', '') for i in sap_ids]
        newer_sap_ids = [i.replace('nan', '') for i in new_sap_ids]
        df_sap.drop('Unique', axis=1, inplace=True)
        df_sap['Unique'] = newer_sap_ids

        first_column_sap = df_sap.pop('Unique')
        df_sap.insert(0, 'Unique', first_column_sap)
        df_sap = df_sap[df_sap['Unique'] != '-']

        # get the data frame for the HJ report and re-arrange the ID column
        df_hj["Unique"] = df_hj['Material'].astype(str) + "-" + df_hj["Load ID Trimmed"].astype(str)
        first_column_hj = df_hj.pop('Unique')
        df_hj.insert(0, 'Unique', first_column_hj)

        toExcel(df_sap, df_hj)

        df_sap_cpy = df_sap
        df_hj_cpy = df_hj

        #
        # Quantity Summary Comparison in YDS
        #
        df_sap = df_sap[['Unique', 'Qty in Un. of Entry']]
        df_sap = df_sap.groupby('Unique').sum()
        df_sap.reset_index(inplace=True)

        df_hj = df_hj[['Unique', 'Quantity']]
        df_hj = df_hj.groupby('Unique').sum()
        df_hj.reset_index(inplace=True)

        # combine IDs and Quantities
        df_variance = pd.merge(df_sap, df_hj, on="Unique", how="outer")
        df_variance.fillna(0, inplace=True)
        df_variance.columns = ['Unique', 'SAP Pick Quantity (YDS)', 'HJ GDN Quantity (YDS)']
        df_variance['VARIANCE'] = round(
            df_variance['HJ GDN Quantity (YDS)'].astype(float) - df_variance['SAP Pick Quantity (YDS)'].astype(float), 3)

        # identifying the larger quantity
        qty = []
        for var in df_variance['VARIANCE']:
            if var > 0:
                qty.append("HJ+")
            elif var < 0:
                qty.append("SAP+")
            else:
                qty.append("Quantity Tally")

        # adding status and comments columns
        df_variance['SAP+/HJ+ in QTY Level'] = qty

        #
        # Quantity Summary Comparison in Rolls
        #
        df_sap_2 = df_sap_cpy[['Unique', 'Qty in Un. of Entry']]
        df_sap_2 = df_sap_2.groupby('Unique').count()
        df_sap_2.reset_index(inplace=True)

        df_hj_2 = df_hj_cpy[['Unique', 'Quantity']]
        df_hj_2 = df_hj_2.groupby('Unique').count()
        df_hj_2.reset_index(inplace=True)

        # combine IDs and Quantities
        df_variancej_2 = pd.merge(df_sap_2, df_hj_2, on="Unique", how="outer")
        df_variancej_2.fillna(0, inplace=True)
        df_variancej_2.columns = ['Unique', 'SAP Pick Quantity (Roll)', 'HJ GDN Quantity (Roll)']
        df_variancej_2['VARIANCE'] = round(
            df_variancej_2['SAP Pick Quantity (Roll)'].astype(float) - df_variancej_2['HJ GDN Quantity (Roll)'].astype(
                float), 3)

        # identifying the larger quantity
        qty = []
        for var in df_variancej_2['VARIANCE']:
            if var > 0:
                qty.append("HJ+")
            elif var < 0:
                qty.append("SAP+")
            else:
                qty.append("Quantity Tally")

        # adding status and comments columns
        df_variancej_2['SAP+/HJ+ in QTY Level'] = qty

        toExcelVariance(df_variance, df_variancej_2)

        email(file_paths)
        status.configure(text="Completed")
        progress_bar.stop()

    except Exception as e:
        status.configure(text=e)
        progress_bar.stop()
