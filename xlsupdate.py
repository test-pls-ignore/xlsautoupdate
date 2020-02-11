# -*- coding: utf-8 -*-
"""
Created on Tue Feb 11 07:15:56 2020

@author: Christian.Reiners
"""

import win32com.client

# Start an instance of Excel
xlapp = win32com.client.DispatchEx("Excel.Application")

# Open the workbook in said instance of Excel
wb = xlapp.workbooks.open("U:\PMO-Übersicht.xlsm")

# create backup
wb.SaveCopyAs("U:\PMO-Übersicht_backup.xlsm")

# Optional to see the instance
# xlapp.Visible = True

# Refresh all data connections.
wb.RefreshAll()
wb.Save()

# Quit
xlapp.Quit()
