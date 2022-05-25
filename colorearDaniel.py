# -*- coding: utf-8 -*-
"""
Created on Mon Mar 28 13:17:58 2022

@author: DANIEL AGUDELO-MARTINEZ
"""

import sys, os
import re
import win32com.client
import tkinter as tk
from tkinter import filedialog
import pandas as pd

days = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']
try:
    # ========================================================================
    # Seleccionar RecEle
    root = tk.Tk()
    root.withdraw()
    address = filedialog.askopenfilename(
                               filetypes=[("Excel files", ".xlsm .xlsx")],
                               title = "Seleccione un RecEle"
                               )
    # address=r"d:\ambiente\escritorio\pytests\Copia de RecEle13_Ant_TEST.xlsm"
    print("Archivo seleccionado: "+address)
    print("Ingresando causales del día: ")
    # ========================================================================
    # Para cada día
    for day in days:
        # ========================================================================    
        xls = pd.ExcelFile(address)
        df = pd.read_excel(xls,day)
        sheet1 = df.values.tolist()
        csgs = [row[2] for row in sheet1]
        # print(csgs)
        # ========================================================================   
        pattern_csg = r"^\(C\d+\).+"
        pattern_voltage = r"^\(C\d+\)\s.+\s(.+)\s[kK][vV]$"
        idx_csgs = [int(idx_csg) for idx_csg,csg in enumerate(csgs) \
                    if re.search(pattern_csg, str(csg))
                    ]
        # print (idx_csgs)
        if any(idx_csgs):
            idx_voltages = [int(idx_volt) for idx_volt,csg in enumerate(csgs) \
                    if re.search(pattern_voltage, str(csg))
                    ]
            voltage_csgs_aux = [re.findall(pattern_voltage, str(csg)) for csg in csgs]
            voltage_csgs_aux2 = []
            for elem in voltage_csgs_aux:
                voltage_csgs_aux2.extend(elem)
            
            voltage_csgs_aux3 = []
            pattern_trx = r"^(\d+)[/.+]?"
            voltage_csgs_aux3 = [re.findall(pattern_trx,str(voltage)) for voltage in voltage_csgs_aux2]
            voltage_csgs = []
            for elem in voltage_csgs_aux3:
                voltage_csgs.extend(elem)
            # print (voltage_csgs)
            
            causal_csgs = ["STN" if int(voltage.replace(".",","))>=220 else "NI" for voltage in voltage_csgs]
            # print (causal_csgs)
            # ========================================================================
            xlApp = win32com.client.Dispatch('Excel.Application')
            xlApp.DisplayAlerts = False
            xlApp.EnableEvents = False
            wb = xlApp.Workbooks.Open(
                address,
                ReadOnly = False)
            ws = wb.Worksheets(day)
            for idx,aux in enumerate(idx_voltages):
                # ws.Range(ws.Cells(idx_csgs[idx]+2,24),ws.Cells(idx_csgs[idx]+2,24)).Value = 'DANIEL'
                ws.Cells(idx_voltages[idx]+2,24).Value=causal_csgs[idx]
            wb.Save()
            wb.Close(True)
            # del(wb)
            xlApp.EnableEvents = True
            xlApp.DisplayAlerts = True
            xlApp.Application.Quit()
            # xlApp.Quit()
            # del(xlApp)
            # ========================================================================
            print("\t"+str(day)+": "+str(len(causal_csgs)))
    # ========================================================================
    print("\nCausales actualizadas correctamente")
except Exception as e:
    print("\nOcurrió un error en el proceso:"+str(e))
# ========================================================================
input("Presione una tecla para salir...")
# ========================================================================

# https://www.mrexcel.com/board/threads/use-vba-to-open-workbook-without-running-macros.368961/
# https://pythonexcels.com/python/2009/09/29/basic-excel-driving-with-python
