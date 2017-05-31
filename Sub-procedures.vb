Attribute VB_Name = "Module1"
'Subprocedure module

Public Sub SOLVER()
    SolverSolve userFinish:=True
End Sub
Public Sub HideSheets()
    Application.ScreenUpdating = False
    Sheets("Locate Lactate Minimum").Visible = 2
    Sheets("Lactate Normalized Intensities").Visible = 2
    Sheets("Lactate Norm Int Smoothed").Visible = 2
    Sheets("Lactate Analysis").Visible = 2
    Sheets("PO2 Analysis").Visible = 2
    Sheets("PCO2 Analysis").Visible = 2
    Sheets("pH Analysis").Visible = 2
    Sheets("B-Lac Barcode & Results").Visible = 2
    Sheets("Statistical Analysis").Visible = 2
    Sheets("Analyte Targets").Visible = 2
    Sheets("Instrument 1 Data Base").Visible = 2
    Sheets("Instrument 2 Data Base").Visible = 2
    Sheets("Instrument 3 Data Base").Visible = 2
    Sheets("Instrument 4 Data Base").Visible = 2
End Sub
Public Sub UnhideSheets()
    Application.ScreenUpdating = False
    Sheets("Locate Lactate Minimum").Visible = -1
    Sheets("Lactate Normalized Intensities").Visible = -1
    Sheets("Lactate Norm Int Smoothed").Visible = -1
    Sheets("Lactate Analysis").Visible = -1
    Sheets("PO2 Analysis").Visible = -1
    Sheets("PCO2 Analysis").Visible = -1
    Sheets("pH Analysis").Visible = -1
    Sheets("B-Lac Barcode & Results").Visible = -1
    Sheets("Statistical Analysis").Visible = -1
    Sheets("Analyte Targets").Visible = -1
    Sheets("Instrument 1 Data Base").Visible = -1
    Sheets("Instrument 2 Data Base").Visible = -1
    Sheets("Instrument 3 Data Base").Visible = -1
    Sheets("Instrument 4 Data Base").Visible = -1
End Sub
Public Sub LockDownWorkbook()
    Const strPassword As String = "12345"
    Application.ScreenUpdating = False
    Sheets("Locate Lactate Minimum").Protect Password:=strPassword
    Sheets("Lactate Normalized Intensities").Protect Password:=strPassword
    Sheets("Lactate Norm Int Smoothed").Protect Password:=strPassword
    Sheets("Lactate Analysis").Protect Password:=strPassword
    Sheets("PO2 Analysis").Protect Password:=strPassword
    Sheets("PCO2 Analysis").Protect Password:=strPassword
    Sheets("pH Analysis").Protect Password:=strPassword
    Sheets("B-Lac Barcode & Results").Protect Password:=strPassword
    Sheets("Statistical Analysis").Protect Password:=strPassword
    Sheets("Analyte Targets").Protect Password:=strPassword
    Sheets("Instrument 1 Data Base").Protect Password:=strPassword
    Sheets("Instrument 2 Data Base").Protect Password:=strPassword
    Sheets("Instrument 3 Data Base").Protect Password:=strPassword
    Sheets("Instrument 4 Data Base").Protect Password:=strPassword
    Sheets("Controls").Protect Password:=strPassword
    ActiveWorkbook.Protect Password:=strPassword
End Sub
Public Sub UnlockWorkbook()
    Const strPassword As String = "12345"
    Application.ScreenUpdating = False
    Sheets("Locate Lactate Minimum").Unprotect Password:=strPassword
    Sheets("Lactate Normalized Intensities").Unprotect Password:=strPassword
    Sheets("Lactate Norm Int Smoothed").Unprotect Password:=strPassword
    Sheets("Lactate Analysis").Unprotect Password:=strPassword
    Sheets("PO2 Analysis").Unprotect Password:=strPassword
    Sheets("PCO2 Analysis").Unprotect Password:=strPassword
    Sheets("pH Analysis").Unprotect Password:=strPassword
    Sheets("B-Lac Barcode & Results").Unprotect Password:=strPassword
    Sheets("Statistical Analysis").Unprotect Password:=strPassword
    Sheets("Analyte Targets").Unprotect Password:=strPassword
    Sheets("Instrument 1 Data Base").Unprotect Password:=strPassword
    Sheets("Instrument 2 Data Base").Unprotect Password:=strPassword
    Sheets("Instrument 3 Data Base").Unprotect Password:=strPassword
    Sheets("Instrument 4 Data Base").Unprotect Password:=strPassword
    Sheets("Controls").Unprotect Password:=strPassword
    ActiveWorkbook.Unprotect Password:=strPassword
End Sub

