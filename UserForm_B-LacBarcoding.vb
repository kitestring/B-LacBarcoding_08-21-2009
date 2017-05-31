VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSampleNumbers 
   Caption         =   "B-Lac Barcode User Inputs"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   OleObjectBlob   =   "frmSampleNumbers.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSampleNumbers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const strPassword As String = "12345"
Dim CCA As Byte 'Network Address & CCA S/N Cycling Variable
Dim SN As Byte 'Sample Number Cycling Variable
Dim wkbDataBase As Workbook
Dim wkbBarcode As Workbook
Dim intRawDataStartRow As Integer
Dim intRawDataEndColumn As Integer
Dim SampleNumberFound As Boolean
Dim ValidEntry As Boolean
Dim Row As Double
Dim Column As Double
Dim ErrorCode As Byte
Dim ErrorMessage As String
Dim Continue As Boolean
Dim FirstRun As Boolean
Dim n As Byte

Private Sub cmdAnalyze_Click()

Call VBACheck
If Continue = False Then
    Unload frmSampleNumbers
    Exit Sub
End If

frmSampleNumbers.Hide
Application.ScreenUpdating = False
Call Lock_Unlock_Workbook("unlock")

Dim intCCA_SampleNumbers(3, 17) As String
Dim varDataBaseAddresses(3) As String
Dim strCCNSN(3) As String
Dim RawDataSheet(3) As Worksheet

SearchFailure = False
LocateSampleNumberFailure = False
FirstRun = FirstRunCheck
ErrorCode = 0

Set wkbBarcode = ActiveWorkbook

        intCCA_SampleNumbers(0, 0) = txtOCL1_1_CCA1.Value
        intCCA_SampleNumbers(0, 1) = txtOCL1_2_CCA1.Value
        intCCA_SampleNumbers(0, 2) = txtOCL2_1_CCA1.Value
        intCCA_SampleNumbers(0, 3) = txtOCL2_2_CCA1.Value
        intCCA_SampleNumbers(0, 4) = txtOCL3_1_CCA1.Value
        intCCA_SampleNumbers(0, 5) = txtOCL3_2_CCA1.Value
        intCCA_SampleNumbers(0, 6) = txtRNA5_1_CCA1.Value
        intCCA_SampleNumbers(0, 7) = txtRNA5_2_CCA1.Value
        intCCA_SampleNumbers(0, 8) = txtLowBlood_1_CCA1.Value
        intCCA_SampleNumbers(0, 9) = txtLowBlood_2_CCA1.Value
        intCCA_SampleNumbers(0, 10) = txtMidBlood_1_CCA1.Value
        intCCA_SampleNumbers(0, 11) = txtMidBlood_2_CCA1.Value
        intCCA_SampleNumbers(0, 12) = txtHighBlood_1_CCA1.Value
        intCCA_SampleNumbers(0, 13) = txtHighBlood_2_CCA1.Value
        intCCA_SampleNumbers(0, 14) = txtRawBlood_1_CCA1.Value
        intCCA_SampleNumbers(0, 15) = txtRawBlood_2_CCA1.Value
        intCCA_SampleNumbers(0, 16) = txtRBLowtHb_1_CCA1.Value
        intCCA_SampleNumbers(0, 17) = txtRBLowtHb_2_CCA1.Value
        
        intCCA_SampleNumbers(1, 0) = txtOCL1_1_CCA2.Value
        intCCA_SampleNumbers(1, 1) = txtOCL1_2_CCA2.Value
        intCCA_SampleNumbers(1, 2) = txtOCL2_1_CCA2.Value
        intCCA_SampleNumbers(1, 3) = txtOCL2_2_CCA2.Value
        intCCA_SampleNumbers(1, 4) = txtOCL3_1_CCA2.Value
        intCCA_SampleNumbers(1, 5) = txtOCL3_2_CCA2.Value
        intCCA_SampleNumbers(1, 6) = txtRNA5_1_CCA2.Value
        intCCA_SampleNumbers(1, 7) = txtRNA5_2_CCA2.Value
        intCCA_SampleNumbers(1, 8) = txtLowBlood_1_CCA2.Value
        intCCA_SampleNumbers(1, 9) = txtLowBlood_2_CCA2.Value
        intCCA_SampleNumbers(1, 10) = txtMidBlood_1_CCA2.Value
        intCCA_SampleNumbers(1, 11) = txtMidBlood_2_CCA2.Value
        intCCA_SampleNumbers(1, 12) = txtHighBlood_1_CCA2.Value
        intCCA_SampleNumbers(1, 13) = txtHighBlood_2_CCA2.Value
        intCCA_SampleNumbers(1, 14) = txtRawBlood_1_CCA2.Value
        intCCA_SampleNumbers(1, 15) = txtRawBlood_2_CCA2.Value
        intCCA_SampleNumbers(1, 16) = txtRBLowtHb_1_CCA2.Value
        intCCA_SampleNumbers(1, 17) = txtRBLowtHb_2_CCA2.Value
        
        intCCA_SampleNumbers(2, 0) = txtOCL1_1_CCA3.Value
        intCCA_SampleNumbers(2, 1) = txtOCL1_2_CCA3.Value
        intCCA_SampleNumbers(2, 2) = txtOCL2_1_CCA3.Value
        intCCA_SampleNumbers(2, 3) = txtOCL2_2_CCA3.Value
        intCCA_SampleNumbers(2, 4) = txtOCL3_1_CCA3.Value
        intCCA_SampleNumbers(2, 5) = txtOCL3_2_CCA3.Value
        intCCA_SampleNumbers(2, 6) = txtRNA5_1_CCA3.Value
        intCCA_SampleNumbers(2, 7) = txtRNA5_2_CCA3.Value
        intCCA_SampleNumbers(2, 8) = txtLowBlood_1_CCA3.Value
        intCCA_SampleNumbers(2, 9) = txtLowBlood_2_CCA3.Value
        intCCA_SampleNumbers(2, 10) = txtMidBlood_1_CCA3.Value
        intCCA_SampleNumbers(2, 11) = txtMidBlood_2_CCA3.Value
        intCCA_SampleNumbers(2, 12) = txtHighBlood_1_CCA3.Value
        intCCA_SampleNumbers(2, 13) = txtHighBlood_2_CCA3.Value
        intCCA_SampleNumbers(2, 14) = txtRawBlood_1_CCA3.Value
        intCCA_SampleNumbers(2, 15) = txtRawBlood_2_CCA3.Value
        intCCA_SampleNumbers(2, 16) = txtRBLowtHb_1_CCA3.Value
        intCCA_SampleNumbers(2, 17) = txtRBLowtHb_2_CCA3.Value
        
        intCCA_SampleNumbers(3, 0) = txtOCL1_1_CCA4.Value
        intCCA_SampleNumbers(3, 1) = txtOCL1_2_CCA4.Value
        intCCA_SampleNumbers(3, 2) = txtOCL2_1_CCA4.Value
        intCCA_SampleNumbers(3, 3) = txtOCL2_2_CCA4.Value
        intCCA_SampleNumbers(3, 4) = txtOCL3_1_CCA4.Value
        intCCA_SampleNumbers(3, 5) = txtOCL3_2_CCA4.Value
        intCCA_SampleNumbers(3, 6) = txtRNA5_1_CCA4.Value
        intCCA_SampleNumbers(3, 7) = txtRNA5_2_CCA4.Value
        intCCA_SampleNumbers(3, 8) = txtLowBlood_1_CCA4.Value
        intCCA_SampleNumbers(3, 9) = txtLowBlood_2_CCA4.Value
        intCCA_SampleNumbers(3, 10) = txtMidBlood_1_CCA4.Value
        intCCA_SampleNumbers(3, 11) = txtMidBlood_2_CCA4.Value
        intCCA_SampleNumbers(3, 12) = txtHighBlood_1_CCA4.Value
        intCCA_SampleNumbers(3, 13) = txtHighBlood_2_CCA4.Value
        intCCA_SampleNumbers(3, 14) = txtRawBlood_1_CCA4.Value
        intCCA_SampleNumbers(3, 15) = txtRawBlood_2_CCA4.Value
        intCCA_SampleNumbers(3, 16) = txtRBLowtHb_1_CCA4.Value
        intCCA_SampleNumbers(3, 17) = txtRBLowtHb_2_CCA4.Value
        
        varDataBaseAddresses(0) = txtDBNetworkAddress1.Value
        varDataBaseAddresses(1) = txtDBNetworkAddress2.Value
        varDataBaseAddresses(2) = txtDBNetworkAddress3.Value
        varDataBaseAddresses(3) = txtDBNetworkAddress4.Value
        
        strCCNSN(0) = "CCA SN " & txtCCA1SN.Value & " Data Base"
        strCCNSN(1) = "CCA SN " & txtCCA2SN.Value & " Data Base"
        strCCNSN(2) = "CCA SN " & txtCCA3SN.Value & " Data Base"
        strCCNSN(3) = "CCA SN " & txtCCA4SN.Value & " Data Base"
        
            
        '1)Drop user inputs onto Sheets("Targets & Limits")
            Sheets("Targets & Limits").Select
            Call DropGeneralInfo
            Call Drop_pHReferenceAnalyzer
            Call Drop_LacReferenceAnalyzer
            Call DropCCASN
            
        '2)Define raw data worksheets
            For n = 0 To 3
                Set RawDataSheet(n) = Worksheets(n + 11)
            Next n
            
        '3)Drop sample numbers into respective raw data sheets
            For CCA = 0 To 3
                RawDataSheet(CCA).Select
                For SN = 0 To 17
                    Call Drop_SampleNumber(intCCA_SampleNumbers(CCA, SN), SN)
                    Call DropCSVFileLocation(varDataBaseAddresses(CCA))
                Next
            Next CCA
            
        '4)Check for required inputs if true then _
            rename raw data worksheets and continue with analysis _
            otherwise alert user and terminate macro.
            Sheets("Targets & Limits").Select
            If Range("RequiredInputsCheck").Value = False Then
                ErrorCode = 3
                GoTo ErrorHandler
            End If
            For n = 0 To 3
                RawDataSheet(n).Name = strCCNSN(n)
            Next n
            
        '5)Extract raw data from *.csv file & _
        drop it onto the respective CCA SN raw data sheet.
            For CCA = 0 To 3
                RawDataSheet(CCA).Select
                Call OpenCSVFile(varDataBaseAddresses(CCA))
                For SN = 0 To 17
                    Call TransferDataRow(intCCA_SampleNumbers(CCA, SN), SN)
                    If ErrorCode <> 0 Then GoTo ErrorHandler
                Next SN
                wkbDataBase.Close False
                wkbBarcode.Activate
            Next CCA
            
        'X)Finalize macro: save file as, disable full screen view.
            Call Lock_Unlock_Workbook("lock")
            Worksheets(10).Select
            Call UserSaveAs(txtLotNo.Value)
            Unload frmSampleNumbers
            
Exit Sub
ErrorHandler:
    Select Case ErrorCode
        Case 1
            ErrorMessage = "Data base is corrupt or incompatable." & _
                "Check that you have selected the correct data base file" & _
                "or redownload the data base." & Chr(10) & Chr(10) & varDataBaseAddresses(CCA)
            wkbDataBase.Close False
        Case 2
            ErrorMessage = SampleNumberNotFound(CCA, SN, intCCA_SampleNumbers(CCA, SN))
            wkbDataBase.Close False
        Case 3
            ErrorMessage = "One or more required user inputs have be omitted." & Chr(10) & _
                "Please check your inputs and reanalyze."
    End Select
    
    MsgBox ErrorMessage, vbCritical, "Error Has Occured"
    
    Call Lock_Unlock_Workbook("lock")
    Call UserSaveAs(txtLotNo.Value)
    Unload frmSampleNumbers
    
End Sub

Private Sub cmdSampleNumbersHelp_Click()
    'MsgBox "Please enter the CCA S/N in text box on the upper left corner of the user form." _
            & Chr(10) & " Do this for all 4 of the CCA's used during barcoding on each of the CCA tabs above." _
            , vbInformation, "CCA S/N Entry"
    'MsgBox "If you ran your samples in the following order with no gaps in the sequence," _
            & Chr(10) & "then be sure the check box at the upper right corner of the user form is checked." _
            , vbInformation, "Sample Numbers Sequential?"
    'MsgBox "2 Repeats OPTICheck Level 1" & Chr(10) & "2 Repeats OPTICheck Level 2" _
            & Chr(10) & "2 Repeats OPTICheck Level 3" & Chr(10) & "2 Repeats RNA Level 0" _
            & Chr(10) & "2 Repeats RNA Level 5" & Chr(10) & "2 Repeats Low Blood" _
            & Chr(10) & "2 Repeats Mid Blood" & Chr(10) & "2 Repeats High Blood" _
            & Chr(10) & "2 Repeats Raw Blood", vbInformation, "Sample Numbers Sequential?"
    'MsgBox "Please enter a sample number in each of the highlighted cells." _
            & " If you ran your samples in the order listed above you need only enter the first number in the sequence." _
            & " Other wise click the Yes button corresponding to that instrument which will change the Yes to No and enter each sample number individually." _
            , vbInformation, "Sample Numbers"
        
        txtOCL1_1_CCA1.Value = 36
        txtOCL1_2_CCA1.Value = 37
        txtOCL2_1_CCA1.Value = 38
        txtOCL2_2_CCA1.Value = 54
        txtOCL3_1_CCA1.Value = 40
        txtOCL3_2_CCA1.Value = 41
        txtRNA5_1_CCA1.Value = 42
        txtRNA5_2_CCA1.Value = 43
        txtLowBlood_1_CCA1.Value = 44
        txtLowBlood_2_CCA1.Value = 45
        txtMidBlood_1_CCA1.Value = 46
        txtMidBlood_2_CCA1.Value = 47
        txtHighBlood_1_CCA1.Value = 48
        txtHighBlood_2_CCA1.Value = 49
        txtRawBlood_1_CCA1.Value = 50
        txtRawBlood_2_CCA1.Value = 51
        txtRBLowtHb_1_CCA1.Value = 52
        txtRBLowtHb_2_CCA1.Value = 53
        
        txtOCL1_1_CCA2.Value = 13
        txtOCL1_2_CCA2.Value = 14
        txtOCL2_1_CCA2.Value = 15
        txtOCL2_2_CCA2.Value = 31
        txtOCL3_1_CCA2.Value = 17
        txtOCL3_2_CCA2.Value = 18
        txtRNA5_1_CCA2.Value = 19
        txtRNA5_2_CCA2.Value = 20
        txtLowBlood_1_CCA2.Value = 21
        txtLowBlood_2_CCA2.Value = 22
        txtMidBlood_1_CCA2.Value = 23
        txtMidBlood_2_CCA2.Value = 24
        txtHighBlood_1_CCA2.Value = 25
        txtHighBlood_2_CCA2.Value = 26
        txtRawBlood_1_CCA2.Value = 27
        txtRawBlood_2_CCA2.Value = 28
        txtRBLowtHb_1_CCA2.Value = 29
        txtRBLowtHb_2_CCA2.Value = 30
        
        txtOCL1_1_CCA3.Value = 13
        txtOCL1_2_CCA3.Value = 14
        txtOCL2_1_CCA3.Value = 15
        txtOCL2_2_CCA3.Value = 31
        txtOCL3_1_CCA3.Value = 17
        txtOCL3_2_CCA3.Value = 18
        txtRNA5_1_CCA3.Value = 19
        txtRNA5_2_CCA3.Value = 20
        txtLowBlood_1_CCA3.Value = 21
        txtLowBlood_2_CCA3.Value = 22
        txtMidBlood_1_CCA3.Value = 23
        txtMidBlood_2_CCA3.Value = 24
        txtHighBlood_1_CCA3.Value = 25
        txtHighBlood_2_CCA3.Value = 26
        txtRawBlood_1_CCA3.Value = 27
        txtRawBlood_2_CCA3.Value = 28
        txtRBLowtHb_1_CCA3.Value = 29
        txtRBLowtHb_2_CCA3.Value = 30
        
        txtOCL1_1_CCA4.Value = 13
        txtOCL1_2_CCA4.Value = 14
        txtOCL2_1_CCA4.Value = 15
        txtOCL2_2_CCA4.Value = 31
        txtOCL3_1_CCA4.Value = 17
        txtOCL3_2_CCA4.Value = 18
        txtRNA5_1_CCA4.Value = 19
        txtRNA5_2_CCA4.Value = 20
        txtLowBlood_1_CCA4.Value = 21
        txtLowBlood_2_CCA4.Value = 22
        txtMidBlood_1_CCA4.Value = 23
        txtMidBlood_2_CCA4.Value = 24
        txtHighBlood_1_CCA4.Value = 25
        txtHighBlood_2_CCA4.Value = 26
        txtRawBlood_1_CCA4.Value = 27
        txtRawBlood_2_CCA4.Value = 28
        txtRBLowtHb_1_CCA4.Value = 29
        txtRBLowtHb_2_CCA4.Value = 30
        
        txtDBNetworkAddress1.Value = "N:\RD\2009 Experiments\Lactate\09-024 Lactate Barcoding Macro (Development)\LogFiles & DataBases\8-21-09 (693399)\DB1716.821"
        txtDBNetworkAddress2.Value = "N:\RD\2009 Experiments\Lactate\09-024 Lactate Barcoding Macro (Development)\LogFiles & DataBases\8-21-09 (693399)\DB1718.821"
        txtDBNetworkAddress3.Value = "N:\RD\2009 Experiments\Lactate\09-024 Lactate Barcoding Macro (Development)\LogFiles & DataBases\8-21-09 (693399)\DB2557.821"
        txtDBNetworkAddress4.Value = "N:\RD\2009 Experiments\Lactate\09-024 Lactate Barcoding Macro (Development)\LogFiles & DataBases\8-21-09 (693399)\DB9223.821"
        
        txtCCA1SN.Value = "1716"
        txtCCA2SN.Value = "1718"
        txtCCA3SN.Value = "2557"
        txtCCA4SN.Value = "9223"
End Sub

Private Sub cmdDBBrowseCCA1_Click()
Application.ScreenUpdating = False
Dim NetworkAddress As Variant
Dim CCA_SN As String
    NetworkAddress = UserDefineDataBaseAddress
    If NetworkAddress = False Then Exit Sub
    CCA_SN = ExtractSN(NetworkAddress)
    If CCA_SN = "Invalid DB" Then
        txtDBNetworkAddress1.Value = NetworkAddress
        lblDBCCA1.Caption = "Invalid File:"
        MsgBox "Cannot locate CCA SN within data base." & Chr(10) & _
            "Suggested corrective actions:" & Chr(10) & "1) Check that you have selected the correct file." _
            & Chr(10) & "2) Re-download the data base." _
            , vbCritical, "Invalid Data Base File"
    ElseIf CCA_SN = txtCCA2SN.Value Or CCA_SN = txtCCA3SN.Value Or CCA_SN = txtCCA4SN.Value Then
        MsgBox "This instrument has already been loaded.", vbCritical, "Duplicate Instrument"
    Else
        txtDBNetworkAddress1.Value = NetworkAddress
        txtCCA1SN.Value = CCA_SN
        fraSampleNumbersCCA1.Caption = CCA_SN & " Sample Numbers"
        lblDBCCA1.Caption = CCA_SN & " Data Base:"
    End If
Application.ScreenUpdating = True
End Sub

Private Sub cmdDBBrowseCCA2_Click()
Application.ScreenUpdating = False
Dim NetworkAddress As Variant
Dim CCA_SN As String
    NetworkAddress = UserDefineDataBaseAddress
    If NetworkAddress = False Then Exit Sub
    CCA_SN = ExtractSN(NetworkAddress)
    If CCA_SN = "Invalid DB" Then
        txtDBNetworkAddress2.Value = NetworkAddress
        lblDBCCA2.Caption = "Invalid File:"
        MsgBox "Cannot locate CCA SN within data base." & Chr(10) & _
            "Suggested corrective actions:" & Chr(10) & "1) Check that you have selected the correct file." _
            & Chr(10) & "2) Re-download the data base." _
            , vbCritical, "Invalid Data Base File"
    ElseIf CCA_SN = txtCCA1SN.Value Or CCA_SN = txtCCA3SN.Value Or CCA_SN = txtCCA4SN.Value Then
        MsgBox "This instrument has already been loaded.", vbCritical, "Duplicate Instrument"
    Else
        txtDBNetworkAddress2.Value = NetworkAddress
        txtCCA2SN.Value = CCA_SN
        fraSampleNumbersCCA2.Caption = CCA_SN & " Sample Numbers"
        lblDBCCA2.Caption = CCA_SN & " Data Base:"
    End If
Application.ScreenUpdating = True
End Sub

Private Sub cmdDBBrowseCCA3_Click()
Application.ScreenUpdating = False
Dim NetworkAddress As Variant
Dim CCA_SN As String
    NetworkAddress = UserDefineDataBaseAddress
    If NetworkAddress = False Then Exit Sub
    CCA_SN = ExtractSN(NetworkAddress)
    If CCA_SN = "Invalid DB" Then
        txtDBNetworkAddress3.Value = NetworkAddress
        lblDBCCA3.Caption = "Invalid File:"
        MsgBox "Cannot locate CCA SN within data base." & Chr(10) & _
            "Suggested corrective actions:" & Chr(10) & "1) Check that you have selected the correct file." _
            & Chr(10) & "2) Re-download the data base." _
            , vbCritical, "Invalid Data Base File"
    ElseIf CCA_SN = txtCCA1SN.Value Or CCA_SN = txtCCA2SN.Value Or CCA_SN = txtCCA4SN.Value Then
        MsgBox "This instrument has already been loaded.", vbCritical, "Duplicate Instrument"
    Else
        txtDBNetworkAddress3.Value = NetworkAddress
        txtCCA3SN.Value = CCA_SN
        fraSampleNumbersCCA3.Caption = CCA_SN & " Sample Numbers"
        lblDBCCA3.Caption = CCA_SN & " Data Base:"
    End If
Application.ScreenUpdating = True
End Sub

Private Sub cmdDBBrowseCCA4_Click()
Application.ScreenUpdating = False
Dim NetworkAddress As Variant
Dim CCA_SN As String
    NetworkAddress = UserDefineDataBaseAddress
    If NetworkAddress = False Then Exit Sub
    CCA_SN = ExtractSN(NetworkAddress)
    If CCA_SN = "Invalid DB" Then
        txtDBNetworkAddress4.Value = NetworkAddress
        lblDBCCA4.Caption = "Invalid File:"
        MsgBox "Cannot locate CCA SN within data base." & Chr(10) & _
            "Suggested corrective actions:" & Chr(10) & "1) Check that you have selected the correct file." _
            & Chr(10) & "2) Re-download the data base." _
            , vbCritical, "Invalid Data Base File"
    ElseIf CCA_SN = txtCCA1SN.Value Or CCA_SN = txtCCA2SN.Value Or CCA_SN = txtCCA3SN.Value Then
        MsgBox "This instrument has already been loaded.", vbCritical, "Duplicate Instrument"
    Else
        txtDBNetworkAddress4.Value = NetworkAddress
        txtCCA4SN.Value = CCA_SN
        fraSampleNumbersCCA4.Caption = CCA_SN & " Sample Numbers"
        lblDBCCA4.Caption = CCA_SN & " Data Base:"
    End If
Application.ScreenUpdating = True
End Sub

Private Function UserDefineDataBaseAddress() As String
    UserDefineDataBaseAddress = Application.GetOpenFilename(FileFilter:="All Files Types(*.*),*.*", _
        Title:="Select a data base", MultiSelect:=False)
End Function

Private Function ExtractSN(ByVal DBAddress As String) As String
On Error GoTo ErrorHandler
Dim TrackError As Byte
    TrackError = 0

    Workbooks.OpenText FileName:=DBAddress, _
        Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
        , Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
    Set wkbDataBase = ActiveWorkbook

    TrackError = 1
    Cells.Find(What:="dbg > set sn", After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate

    ExtractSN = Right(ActiveCell.Value, 4)
    wkbDataBase.Close
Exit Function
ErrorHandler:
    ExtractSN = "Invalid DB"
    If TrackError = 1 Then wkbDataBase.Close
End Function

Private Sub OpenDelimited(DBAddress As String)
    Workbooks.OpenText FileName:=DBAddress, _
        Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
        , Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
        TrailingMinusNumbers:=True
End Sub

Private Function QuickDBCheck(ByVal CCA_SN As String) As Boolean
    
    If CCA_SN = "Invalid DB" Then
        MsgBox "Cannot locate CCA SN within data base." & Chr(10) & _
            "Suggested corrective actions:" & Chr(10) & "1) Check that you have selected the correct file." _
            & Chr(10) & "2) Re-download the data base." _
            , vbCritical, "Invalid Data Base File"
    End If
    
End Function

'OPTI Check Auto Complete
Private Sub txtOCL1LotNo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo QuickExit
    If txtOCL1LotNo.Value = "" Then Exit Sub
    txtOCL2LotNo.Value = txtOCL1LotNo.Value + 100
    txtOCL3LotNo.Value = txtOCL1LotNo.Value + 200
QuickExit:
End Sub

'OPTI Check Auto Complete
Private Sub txtOCL1ExpDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
On Error GoTo QuickExit
    If txtOCL1ExpDate.Value = "" Then Exit Sub
    txtOCL2ExpDate.Value = txtOCL1ExpDate.Value
    txtOCL3ExpDate.Value = txtOCL1ExpDate.Value
QuickExit:
End Sub

Private Function SampleNumberInputValidation(ByVal SampleNumberInput As Variant) As Boolean
Dim strErrorMessage As String
    If SampleNumberInput = "" Then
        SampleNumberInputValidation = False
        Exit Function
    ElseIf Not (IsNumeric(SampleNumberInput)) Then
        SampleNumberInputValidation = False
        strErrorMessage = "Sample number cannot contain" & Chr(10) & _
            "symbolic or alphabetic characters."
    ElseIf Int(SampleNumberInput) - SampleNumberInput <> 0 Then
        SampleNumberInputValidation = False
        strErrorMessage = "Sample number cannot be fractional" & Chr(10) & _
            "values or contain decimals."
    ElseIf SampleNumberInput < 1 Then
        SampleNumberInputValidation = False
        strErrorMessage = "Sample number cannot be less than 1."
    Else
        SampleNumberInputValidation = True
        Exit Function
    End If
    MsgBox strErrorMessage, vbCritical, "Invalid Data Entry"
End Function

'CCA1 Data Validation OPTICheck Level 1 (1)
'Auto fill sample numbers if sequential is checked
Private Sub txtOCL1_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_1_CCA1.Value)
    If chkSequential_CCA1.Value = True And ValidEntry = True Then
        txtOCL1_2_CCA1.Value = txtOCL1_1_CCA1.Value + 1
        txtOCL2_1_CCA1.Value = txtOCL1_1_CCA1.Value + 2
        txtOCL2_2_CCA1.Value = txtOCL1_1_CCA1.Value + 3
        txtOCL3_1_CCA1.Value = txtOCL1_1_CCA1.Value + 4
        txtOCL3_2_CCA1.Value = txtOCL1_1_CCA1.Value + 5
        txtRNA5_1_CCA1.Value = txtOCL1_1_CCA1.Value + 6
        txtRNA5_2_CCA1.Value = txtOCL1_1_CCA1.Value + 7
        txtLowBlood_1_CCA1.Value = txtOCL1_1_CCA1.Value + 8
        txtLowBlood_2_CCA1.Value = txtOCL1_1_CCA1.Value + 9
        txtMidBlood_1_CCA1.Value = txtOCL1_1_CCA1.Value + 10
        txtMidBlood_2_CCA1.Value = txtOCL1_1_CCA1.Value + 11
        txtHighBlood_1_CCA1.Value = txtOCL1_1_CCA1.Value + 12
        txtHighBlood_2_CCA1.Value = txtOCL1_1_CCA1.Value + 13
        txtRawBlood_1_CCA1.Value = txtOCL1_1_CCA1.Value + 14
        txtRawBlood_2_CCA1.Value = txtOCL1_1_CCA1.Value + 15
        txtRBLowtHb_1_CCA1.Value = txtOCL1_1_CCA1.Value + 16
        txtRBLowtHb_2_CCA1.Value = txtOCL1_1_CCA1.Value + 17
    ElseIf ValidEntry = False Then
        txtOCL1_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation OPTICheck Level 1 (2)
Private Sub txtOCL1_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_2_CCA1.Value)
    If ValidEntry = False Then
        txtOCL1_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation OPTICheck Level 2 (1)
Private Sub txtOCL2_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_1_CCA1.Value)
    If ValidEntry = False Then
        txtOCL2_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation OPTICheck Level 2 (2)
Private Sub txtOCL2_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_2_CCA1.Value)
    If ValidEntry = False Then
        txtOCL2_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation OPTICheck Level 3 (1)
Private Sub txtOCL3_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_1_CCA1.Value)
    If ValidEntry = False Then
        txtOCL3_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation OPTICheck Level 3 (2)
Private Sub txtOCL3_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_2_CCA1.Value)
    If ValidEntry = False Then
        txtOCL3_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Raw Blood Low tHb (1)
Private Sub txtRBLowtHb_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_1_CCA1.Value)
    If ValidEntry = False Then
        txtRBLowtHb_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Raw Blood Low tHb (2)
Private Sub txtRBLowtHb_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_2_CCA1.Value)
    If ValidEntry = False Then
        txtRBLowtHb_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation RNA Level 5 (1)
Private Sub txtRNA5_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_1_CCA1.Value)
    If ValidEntry = False Then
        txtRNA5_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation RNA Level 5 (2)
Private Sub txtRNA5_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_2_CCA1.Value)
    If ValidEntry = False Then
        txtRNA5_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Low Blood (1)
Private Sub txtLowBlood_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_1_CCA1.Value)
    If ValidEntry = False Then
        txtLowBlood_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Low Blood (2)
Private Sub txtLowBlood_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_2_CCA1.Value)
    If ValidEntry = False Then
        txtLowBlood_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Mid Blood (1)
Private Sub txtMidBlood_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_1_CCA1.Value)
    If ValidEntry = False Then
        txtMidBlood_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Mid Blood (2)
Private Sub txtMidBlood_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_2_CCA1.Value)
    If ValidEntry = False Then
        txtMidBlood_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation High Blood (1)
Private Sub txtHighBlood_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_1_CCA1.Value)
    If ValidEntry = False Then
        txtHighBlood_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation High Blood (2)
Private Sub txtHighBlood_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_2_CCA1.Value)
    If ValidEntry = False Then
        txtHighBlood_2_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Raw Blood (1)
Private Sub txtRawBlood_1_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_1_CCA1.Value)
    If ValidEntry = False Then
        txtRawBlood_1_CCA1.Value = Null
    End If
End Sub

'CCA1 Data Validation Raw Blood (2)
Private Sub txtRawBlood_2_CCA1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_2_CCA1.Value)
    If ValidEntry = False Then
        txtRawBlood_2_CCA1.Value = Null
    End If
End Sub
'
'CCA2 Data Validation OPTICheck Level 1 (1)
'Auto fill sample numbers if sequential is checked
Private Sub txtOCL1_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_1_CCA2.Value)
    If chkSequential_CCA2.Value = True And ValidEntry = True Then
        txtOCL1_2_CCA2.Value = txtOCL1_1_CCA2.Value + 1
        txtOCL2_1_CCA2.Value = txtOCL1_1_CCA2.Value + 2
        txtOCL2_2_CCA2.Value = txtOCL1_1_CCA2.Value + 3
        txtOCL3_1_CCA2.Value = txtOCL1_1_CCA2.Value + 4
        txtOCL3_2_CCA2.Value = txtOCL1_1_CCA2.Value + 5
        txtRNA5_1_CCA2.Value = txtOCL1_1_CCA2.Value + 6
        txtRNA5_2_CCA2.Value = txtOCL1_1_CCA2.Value + 7
        txtLowBlood_1_CCA2.Value = txtOCL1_1_CCA2.Value + 8
        txtLowBlood_2_CCA2.Value = txtOCL1_1_CCA2.Value + 9
        txtMidBlood_1_CCA2.Value = txtOCL1_1_CCA2.Value + 10
        txtMidBlood_2_CCA2.Value = txtOCL1_1_CCA2.Value + 11
        txtHighBlood_1_CCA2.Value = txtOCL1_1_CCA2.Value + 12
        txtHighBlood_2_CCA2.Value = txtOCL1_1_CCA2.Value + 13
        txtRawBlood_1_CCA2.Value = txtOCL1_1_CCA2.Value + 14
        txtRawBlood_2_CCA2.Value = txtOCL1_1_CCA2.Value + 15
        txtRBLowtHb_1_CCA2.Value = txtOCL1_1_CCA2.Value + 16
        txtRBLowtHb_2_CCA2.Value = txtOCL1_1_CCA2.Value + 17
    ElseIf ValidEntry = False Then
        txtOCL1_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation OPTICheck Level 1 (2)
Private Sub txtOCL1_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_2_CCA2.Value)
    If ValidEntry = False Then
        txtOCL1_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation OPTICheck Level 2 (1)
Private Sub txtOCL2_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_1_CCA2.Value)
    If ValidEntry = False Then
        txtOCL2_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation OPTICheck Level 2 (2)
Private Sub txtOCL2_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_2_CCA2.Value)
    If ValidEntry = False Then
        txtOCL2_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation OPTICheck Level 3 (1)
Private Sub txtOCL3_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_1_CCA2.Value)
    If ValidEntry = False Then
        txtOCL3_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation OPTICheck Level 3 (2)
Private Sub txtOCL3_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_2_CCA2.Value)
    If ValidEntry = False Then
        txtOCL3_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Raw Blood Low tHb (1)
Private Sub txtRBLowtHb_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_1_CCA2.Value)
    If ValidEntry = False Then
        txtRBLowtHb_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Raw Blood Low tHb (2)
Private Sub txtRBLowtHb_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_2_CCA2.Value)
    If ValidEntry = False Then
        txtRBLowtHb_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation RNA Level 5 (1)
Private Sub txtRNA5_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_1_CCA2.Value)
    If ValidEntry = False Then
        txtRNA5_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation RNA Level 5 (2)
Private Sub txtRNA5_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_2_CCA2.Value)
    If ValidEntry = False Then
        txtRNA5_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Low Blood (1)
Private Sub txtLowBlood_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_1_CCA2.Value)
    If ValidEntry = False Then
        txtLowBlood_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Low Blood (2)
Private Sub txtLowBlood_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_2_CCA2.Value)
    If ValidEntry = False Then
        txtLowBlood_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Mid Blood (1)
Private Sub txtMidBlood_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_1_CCA2.Value)
    If ValidEntry = False Then
        txtMidBlood_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Mid Blood (2)
Private Sub txtMidBlood_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_2_CCA2.Value)
    If ValidEntry = False Then
        txtMidBlood_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation High Blood (1)
Private Sub txtHighBlood_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_1_CCA2.Value)
    If ValidEntry = False Then
        txtHighBlood_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation High Blood (2)
Private Sub txtHighBlood_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_2_CCA2.Value)
    If ValidEntry = False Then
        txtHighBlood_2_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Raw Blood (1)
Private Sub txtRawBlood_1_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_1_CCA2.Value)
    If ValidEntry = False Then
        txtRawBlood_1_CCA2.Value = Null
    End If
End Sub

'CCA2 Data Validation Raw Blood (2)
Private Sub txtRawBlood_2_CCA2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_2_CCA2.Value)
    If ValidEntry = False Then
        txtRawBlood_2_CCA2.Value = Null
    End If
End Sub

'CCA3 Data Validation OPTICheck Level 1 (1)
'Auto fill sample numbers if sequential is checked
Private Sub txtOCL1_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_1_CCA3.Value)
    If chkSequential_CCA3.Value = True And ValidEntry = True Then
        txtOCL1_2_CCA3.Value = txtOCL1_1_CCA3.Value + 1
        txtOCL2_1_CCA3.Value = txtOCL1_1_CCA3.Value + 2
        txtOCL2_2_CCA3.Value = txtOCL1_1_CCA3.Value + 3
        txtOCL3_1_CCA3.Value = txtOCL1_1_CCA3.Value + 4
        txtOCL3_2_CCA3.Value = txtOCL1_1_CCA3.Value + 5
        txtRNA5_1_CCA3.Value = txtOCL1_1_CCA3.Value + 6
        txtRNA5_2_CCA3.Value = txtOCL1_1_CCA3.Value + 7
        txtLowBlood_1_CCA3.Value = txtOCL1_1_CCA3.Value + 8
        txtLowBlood_2_CCA3.Value = txtOCL1_1_CCA3.Value + 9
        txtMidBlood_1_CCA3.Value = txtOCL1_1_CCA3.Value + 10
        txtMidBlood_2_CCA3.Value = txtOCL1_1_CCA3.Value + 11
        txtHighBlood_1_CCA3.Value = txtOCL1_1_CCA3.Value + 12
        txtHighBlood_2_CCA3.Value = txtOCL1_1_CCA3.Value + 13
        txtRawBlood_1_CCA3.Value = txtOCL1_1_CCA3.Value + 14
        txtRawBlood_2_CCA3.Value = txtOCL1_1_CCA3.Value + 15
        txtRBLowtHb_1_CCA3.Value = txtOCL1_1_CCA3.Value + 16
        txtRBLowtHb_2_CCA3.Value = txtOCL1_1_CCA3.Value + 17
    ElseIf ValidEntry = False Then
        txtOCL1_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation OPTICheck Level 1 (2)
Private Sub txtOCL1_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_2_CCA3.Value)
    If ValidEntry = False Then
        txtOCL1_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation OPTICheck Level 2 (1)
Private Sub txtOCL2_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_1_CCA3.Value)
    If ValidEntry = False Then
        txtOCL2_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation OPTICheck Level 2 (2)
Private Sub txtOCL2_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_2_CCA3.Value)
    If ValidEntry = False Then
        txtOCL2_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation OPTICheck Level 3 (1)
Private Sub txtOCL3_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_1_CCA3.Value)
    If ValidEntry = False Then
        txtOCL3_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation OPTICheck Level 3 (2)
Private Sub txtOCL3_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_2_CCA3.Value)
    If ValidEntry = False Then
        txtOCL3_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Raw Blood Low tHb (1)
Private Sub txtRBLowtHb_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_1_CCA3.Value)
    If ValidEntry = False Then
        txtRBLowtHb_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Raw Blood Low tHb (2)
Private Sub txtRBLowtHb_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_2_CCA3.Value)
    If ValidEntry = False Then
        txtRBLowtHb_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation RNA Level 5 (1)
Private Sub txtRNA5_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_1_CCA3.Value)
    If ValidEntry = False Then
        txtRNA5_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation RNA Level 5 (2)
Private Sub txtRNA5_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_2_CCA3.Value)
    If ValidEntry = False Then
        txtRNA5_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Low Blood (1)
Private Sub txtLowBlood_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_1_CCA3.Value)
    If ValidEntry = False Then
        txtLowBlood_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Low Blood (2)
Private Sub txtLowBlood_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_2_CCA3.Value)
    If ValidEntry = False Then
        txtLowBlood_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Mid Blood (1)
Private Sub txtMidBlood_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_1_CCA3.Value)
    If ValidEntry = False Then
        txtMidBlood_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Mid Blood (2)
Private Sub txtMidBlood_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_2_CCA3.Value)
    If ValidEntry = False Then
        txtMidBlood_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation High Blood (1)
Private Sub txtHighBlood_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_1_CCA3.Value)
    If ValidEntry = False Then
        txtHighBlood_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation High Blood (2)
Private Sub txtHighBlood_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_2_CCA3.Value)
    If ValidEntry = False Then
        txtHighBlood_2_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Raw Blood (1)
Private Sub txtRawBlood_1_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_1_CCA3.Value)
    If ValidEntry = False Then
        txtRawBlood_1_CCA3.Value = Null
    End If
End Sub

'CCA3 Data Validation Raw Blood (2)
Private Sub txtRawBlood_2_CCA3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_2_CCA3.Value)
    If ValidEntry = False Then
        txtRawBlood_2_CCA3.Value = Null
    End If
End Sub

'CCA4 Data Validation OPTICheck Level 1 (1)
'Auto fill sample numbers if sequential is checked
Private Sub txtOCL1_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_1_CCA4.Value)
    If chkSequential_CCA4.Value = True And ValidEntry = True Then
        txtOCL1_2_CCA4.Value = txtOCL1_1_CCA4.Value + 1
        txtOCL2_1_CCA4.Value = txtOCL1_1_CCA4.Value + 2
        txtOCL2_2_CCA4.Value = txtOCL1_1_CCA4.Value + 3
        txtOCL3_1_CCA4.Value = txtOCL1_1_CCA4.Value + 4
        txtOCL3_2_CCA4.Value = txtOCL1_1_CCA4.Value + 5
        txtRNA5_1_CCA4.Value = txtOCL1_1_CCA4.Value + 6
        txtRNA5_2_CCA4.Value = txtOCL1_1_CCA4.Value + 7
        txtLowBlood_1_CCA4.Value = txtOCL1_1_CCA4.Value + 8
        txtLowBlood_2_CCA4.Value = txtOCL1_1_CCA4.Value + 9
        txtMidBlood_1_CCA4.Value = txtOCL1_1_CCA4.Value + 10
        txtMidBlood_2_CCA4.Value = txtOCL1_1_CCA4.Value + 11
        txtHighBlood_1_CCA4.Value = txtOCL1_1_CCA4.Value + 12
        txtHighBlood_2_CCA4.Value = txtOCL1_1_CCA4.Value + 13
        txtRawBlood_1_CCA4.Value = txtOCL1_1_CCA4.Value + 14
        txtRawBlood_2_CCA4.Value = txtOCL1_1_CCA4.Value + 15
        txtRBLowtHb_1_CCA4.Value = txtOCL1_1_CCA4.Value + 16
        txtRBLowtHb_2_CCA4.Value = txtOCL1_1_CCA4.Value + 17
    ElseIf ValidEntry = False Then
        txtOCL1_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation OPTICheck Level 1 (2)
Private Sub txtOCL1_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL1_2_CCA4.Value)
    If ValidEntry = False Then
        txtOCL1_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation OPTICheck Level 2 (1)
Private Sub txtOCL2_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_1_CCA4.Value)
    If ValidEntry = False Then
        txtOCL2_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation OPTICheck Level 2 (2)
Private Sub txtOCL2_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL2_2_CCA4.Value)
    If ValidEntry = False Then
        txtOCL2_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation OPTICheck Level 3 (1)
Private Sub txtOCL3_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_1_CCA4.Value)
    If ValidEntry = False Then
        txtOCL3_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation OPTICheck Level 3 (2)
Private Sub txtOCL3_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtOCL3_2_CCA4.Value)
    If ValidEntry = False Then
        txtOCL3_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Raw Blood Low tHb (1)
Private Sub txtRBLowtHb_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_1_CCA4.Value)
    If ValidEntry = False Then
        txtRBLowtHb_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Raw Blood Low tHb (2)
Private Sub txtRBLowtHb_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRBLowtHb_2_CCA4.Value)
    If ValidEntry = False Then
        txtRBLowtHb_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation RNA Level 5 (1)
Private Sub txtRNA5_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_1_CCA4.Value)
    If ValidEntry = False Then
        txtRNA5_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation RNA Level 5 (2)
Private Sub txtRNA5_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRNA5_2_CCA4.Value)
    If ValidEntry = False Then
        txtRNA5_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Low Blood (1)
Private Sub txtLowBlood_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_1_CCA4.Value)
    If ValidEntry = False Then
        txtLowBlood_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Low Blood (2)
Private Sub txtLowBlood_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtLowBlood_2_CCA4.Value)
    If ValidEntry = False Then
        txtLowBlood_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Mid Blood (1)
Private Sub txtMidBlood_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_1_CCA4.Value)
    If ValidEntry = False Then
        txtMidBlood_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Mid Blood (2)
Private Sub txtMidBlood_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtMidBlood_2_CCA4.Value)
    If ValidEntry = False Then
        txtMidBlood_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation High Blood (1)
Private Sub txtHighBlood_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_1_CCA4.Value)
    If ValidEntry = False Then
        txtHighBlood_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation High Blood (2)
Private Sub txtHighBlood_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtHighBlood_2_CCA4.Value)
    If ValidEntry = False Then
        txtHighBlood_2_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Raw Blood (1)
Private Sub txtRawBlood_1_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_1_CCA4.Value)
    If ValidEntry = False Then
        txtRawBlood_1_CCA4.Value = Null
    End If
End Sub

'CCA4 Data Validation Raw Blood (2)
Private Sub txtRawBlood_2_CCA4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    ValidEntry = SampleNumberInputValidation(txtRawBlood_2_CCA4.Value)
    If ValidEntry = False Then
        txtRawBlood_2_CCA4.Value = Null
    End If
End Sub

Private Sub VBACheck()
    If txtOperator.Value = "unlock" Then
        Call Lock_Unlock_Workbook("unlock")
        MsgBox strPassword
        Continue = False
    ElseIf txtOperator.Value = "lock" Then
        Call Lock_Unlock_Workbook("lock")
        Continue = False
    Else
        Continue = True
    End If
End Sub

Private Sub UserForm_Activate()

    cmb_pHRefInst1.AddItem "Bayer"
    cmb_pHRefInst1.AddItem "OPTI CCA"
    cmb_pHRefInst1.AddItem "OPTI LION"
    cmb_pHRefInst1.AddItem "Radiometer"
    cmb_pHRefInst1.AddItem "I-Stat"
    
    cmb_pHRefInst2.AddItem "Bayer"
    cmb_pHRefInst2.AddItem "OPTI CCA"
    cmb_pHRefInst2.AddItem "OPTI LION"
    cmb_pHRefInst2.AddItem "Radiometer"
    cmb_pHRefInst2.AddItem "I-Stat"
    cmb_pHRefInst2.AddItem "NA"
    
    cmb_LacRefInst1.AddItem "Bayer"
    cmb_LacRefInst1.AddItem "I-Stat"
    
    cmb_LacRefInst2.AddItem "Bayer"
    cmb_LacRefInst2.AddItem "I-Stat"
    cmb_LacRefInst2.AddItem "NA"
    
    Application.ScreenUpdating = False
    
        Sheets("Targets & Limits").Select
        txtOperator.Value = Range("Operator").Value
        If Range("TestDate").Value = "" Then
            txtTestDate.Value = Date
        Else
            txtTestDate.Value = Range("TestDate").Value
        End If
        txtLotNo.Value = Range("Lot_No").Value
        txtBLacExpirationDate.Value = Range("LotExpDate").Value
        txtOCL1LotNo.Value = Range("OCL1_Lot").Value
        txtOCL1ExpDate.Value = Range("OCL1_ExpDate").Value
        txtOCL2LotNo.Value = Range("OCL2_Lot").Value
        txtOCL2ExpDate.Value = Range("OCL2_ExpDate").Value
        txtOCL3LotNo.Value = Range("OCL3_Lot").Value
        txtOCL3ExpDate.Value = Range("OCL3_ExpDate").Value
        txtRNA5LotNo.Value = Range("RNA5_Lot").Value
        txtRNA5ExpDate.Value = Range("RNA5_ExpDate").Value
        
        txtCCA1SN.Value = Range("CCA_1_S_N").Value
        txtCCA2SN.Value = Range("CCA_2_S_N").Value
        txtCCA3SN.Value = Range("CCA_3_S_N").Value
        txtCCA4SN.Value = Range("CCA_4_S_N").Value
        
        cmb_pHRefInst1.Value = Range("pHRef1_AnalyzerType").Value
        txt_pHRefInst1SN.Value = Range("pHRef1_AnalyzerSN").Value
        txt_pHRefInst1CassetteType.Value = Range("pHRef1_CassetteStyle").Value
        txt_pHRefInst1CassettLot.Value = Range("pHRef1_CassetteLot").Value
        txt_pHRefInst1_Low1.Value = Range("pHRef1_LowGas1").Value
        txt_pHRefInst1_Low2.Value = Range("pHRef1_LowGas2").Value
        txt_pHRefInst1_Mid1.Value = Range("pHRef1_MidGas1").Value
        txt_pHRefInst1_Mid2.Value = Range("pHRef1_MidGas2").Value
        txt_pHRefInst1_High1.Value = Range("pHRef1_HighGas1").Value
        txt_pHRefInst1_High2.Value = Range("pHRef1_HighGas2").Value
    
        cmb_pHRefInst2.Value = Range("phRef2_AnalyzerType").Value
        txt_pHRefInst2SN.Value = Range("phRef2_AnalyzerSN").Value
        txt_pHRefInst2CassetteType.Value = Range("phRef2_CassetteStyle").Value
        txt_pHRefInst2CassettLot.Value = Range("phRef2_CassetteLot").Value
        txt_pHRefInst2_Low1.Value = Range("phRef2_LowGas1").Value
        txt_pHRefInst2_Low2.Value = Range("phRef2_LowGas2").Value
        txt_pHRefInst2_Mid1.Value = Range("phRef2_MidGas1").Value
        txt_pHRefInst2_Mid2.Value = Range("phRef2_MidGas2").Value
        txt_pHRefInst2_High1.Value = Range("phRef2_HighGas1").Value
        txt_pHRefInst2_High2.Value = Range("phRef2_HighGas2").Value
        
        cmb_LacRefInst1.Value = Range("LacRef1_AnalyzerType").Value
        txt_LacRefInst1SN.Value = Range("LacRef1_AnalyzerSN").Value
        txt_LacRefInst1CassetteType.Value = Range("LacRef1_CassetteStyle").Value
        txt_LacRefInst1CassettLot.Value = Range("LacRef1_CassetteLot").Value
        txt_LacRefInst1RawLowtHb1.Value = Range("LacRef1_RBLowtHb1").Value
        txt_LacRefInst1RawLowtHb2.Value = Range("LacRef1_RBLowtHb2").Value
        txt_LacRefInst1_Raw1.Value = Range("LacRef1_Raw1").Value
        txt_LacRefInst1_Raw2.Value = Range("LacRef1_Raw2").Value
    
        cmb_LacRefInst2.Value = Range("LacRef2_AnalyzerType").Value
        txt_LacRefInst2SN.Value = Range("LacRef2_AnalyzerSN").Value
        txt_LacRefInst2CassetteType.Value = Range("LacRef2_CassetteStyle").Value
        txt_LacRefInst2CassettLot.Value = Range("LacRef2_CassetteLot").Value
        txt_LacRefInst2RawLowtHb1.Value = Range("LacRef2_RBLowtHb1").Value
        txt_LacRefInst2RawLowtHb2.Value = Range("LacRef2_RBLowtHb2").Value
        txt_LacRefInst2_Raw1.Value = Range("LacRef2_Raw1").Value
        txt_LacRefInst2_Raw2.Value = Range("LacRef2_Raw2").Value
    
        Worksheets(11).Select
        txtOCL1_1_CCA1.Value = Range("C44").Value
        txtOCL1_2_CCA1.Value = Range("C45").Value
        txtOCL2_1_CCA1.Value = Range("C46").Value
        txtOCL2_2_CCA1.Value = Range("C47").Value
        txtOCL3_1_CCA1.Value = Range("C48").Value
        txtOCL3_2_CCA1.Value = Range("C49").Value
        txtRNA5_1_CCA1.Value = Range("C50").Value
        txtRNA5_2_CCA1.Value = Range("C51").Value
        txtLowBlood_1_CCA1.Value = Range("C52").Value
        txtLowBlood_2_CCA1.Value = Range("C53").Value
        txtMidBlood_1_CCA1.Value = Range("C54").Value
        txtMidBlood_2_CCA1.Value = Range("C55").Value
        txtHighBlood_1_CCA1.Value = Range("C56").Value
        txtHighBlood_2_CCA1.Value = Range("C57").Value
        txtRawBlood_1_CCA1.Value = Range("C58").Value
        txtRawBlood_2_CCA1.Value = Range("C59").Value
        txtRBLowtHb_1_CCA1.Value = Range("C60").Value
        txtRBLowtHb_2_CCA1.Value = Range("C61").Value
        txtDBNetworkAddress1.Value = Range("D20").Value
        
        Worksheets(12).Select
        txtOCL1_1_CCA2.Value = Range("C44").Value
        txtOCL1_2_CCA2.Value = Range("C45").Value
        txtOCL2_1_CCA2.Value = Range("C46").Value
        txtOCL2_2_CCA2.Value = Range("C47").Value
        txtOCL3_1_CCA2.Value = Range("C48").Value
        txtOCL3_2_CCA2.Value = Range("C49").Value
        txtRNA5_1_CCA2.Value = Range("C50").Value
        txtRNA5_2_CCA2.Value = Range("C51").Value
        txtLowBlood_1_CCA2.Value = Range("C52").Value
        txtLowBlood_2_CCA2.Value = Range("C53").Value
        txtMidBlood_1_CCA2.Value = Range("C54").Value
        txtMidBlood_2_CCA2.Value = Range("C55").Value
        txtHighBlood_1_CCA2.Value = Range("C56").Value
        txtHighBlood_2_CCA2.Value = Range("C57").Value
        txtRawBlood_1_CCA2.Value = Range("C58").Value
        txtRawBlood_2_CCA2.Value = Range("C59").Value
        txtRBLowtHb_1_CCA2.Value = Range("C60").Value
        txtRBLowtHb_2_CCA2.Value = Range("C61").Value
        txtDBNetworkAddress2.Value = Range("D20").Value
        
        Worksheets(13).Select
        txtOCL1_1_CCA3.Value = Range("C44").Value
        txtOCL1_2_CCA3.Value = Range("C45").Value
        txtOCL2_1_CCA3.Value = Range("C46").Value
        txtOCL2_2_CCA3.Value = Range("C47").Value
        txtOCL3_1_CCA3.Value = Range("C48").Value
        txtOCL3_2_CCA3.Value = Range("C49").Value
        txtRNA5_1_CCA3.Value = Range("C50").Value
        txtRNA5_2_CCA3.Value = Range("C51").Value
        txtLowBlood_1_CCA3.Value = Range("C52").Value
        txtLowBlood_2_CCA3.Value = Range("C53").Value
        txtMidBlood_1_CCA3.Value = Range("C54").Value
        txtMidBlood_2_CCA3.Value = Range("C55").Value
        txtHighBlood_1_CCA3.Value = Range("C56").Value
        txtHighBlood_2_CCA3.Value = Range("C57").Value
        txtRawBlood_1_CCA3.Value = Range("C58").Value
        txtRawBlood_2_CCA3.Value = Range("C59").Value
        txtRBLowtHb_1_CCA3.Value = Range("C60").Value
        txtRBLowtHb_2_CCA3.Value = Range("C61").Value
        txtDBNetworkAddress3.Value = Range("D20").Value
        
        Worksheets(14).Select
        txtOCL1_1_CCA4.Value = Range("C44").Value
        txtOCL1_2_CCA4.Value = Range("C45").Value
        txtOCL2_1_CCA4.Value = Range("C46").Value
        txtOCL2_2_CCA4.Value = Range("C47").Value
        txtOCL3_1_CCA4.Value = Range("C48").Value
        txtOCL3_2_CCA4.Value = Range("C49").Value
        txtRNA5_1_CCA4.Value = Range("C50").Value
        txtRNA5_2_CCA4.Value = Range("C51").Value
        txtLowBlood_1_CCA4.Value = Range("C52").Value
        txtLowBlood_2_CCA4.Value = Range("C53").Value
        txtMidBlood_1_CCA4.Value = Range("C54").Value
        txtMidBlood_2_CCA4.Value = Range("C55").Value
        txtHighBlood_1_CCA4.Value = Range("C56").Value
        txtHighBlood_2_CCA4.Value = Range("C57").Value
        txtRawBlood_1_CCA4.Value = Range("C58").Value
        txtRawBlood_2_CCA4.Value = Range("C59").Value
        txtRBLowtHb_1_CCA4.Value = Range("C60").Value
        txtRBLowtHb_2_CCA4.Value = Range("C61").Value
        txtDBNetworkAddress4.Value = Range("D20").Value
        
    If FirstRunCheck = False Then
        chkSequential_CCA1.Value = False
        chkSequential_CCA2.Value = False
        chkSequential_CCA3.Value = False
        chkSequential_CCA4.Value = False
    End If
    
    Sheets("Targets & Limits").Select
    Application.ScreenUpdating = True
    
End Sub

Private Sub cmb_pHRefInst1_Change()
    If cmb_pHRefInst1.Value = "" Then Exit Sub
    
    Select Case cmb_pHRefInst1
        Case "Bayer"
            txt_pHRefInst1SN.Value = "6245"
            txt_pHRefInst1CassetteType.Value = "NA"
            txt_pHRefInst1CassettLot.Value = "NA"
        Case "Radiometer"
            txt_pHRefInst1SN.Value = "126R0159N02"
            txt_pHRefInst1CassetteType.Value = "NA"
            txt_pHRefInst1CassettLot.Value = "NA"
        Case "I-Stat"
            txt_pHRefInst1SN.Value = "318221"
            txt_pHRefInst1CassetteType.Value = "CG4+"
            txt_pHRefInst1CassettLot.Value = ""
        Case "OPTI CCA"
            txt_pHRefInst1SN.Value = ""
            txt_pHRefInst1CassetteType.Value = ""
            txt_pHRefInst1CassettLot.Value = ""
        Case "OPTI LION"
            txt_pHRefInst1SN.Value = ""
            txt_pHRefInst1CassetteType.Value = ""
            txt_pHRefInst1CassettLot.Value = ""
    End Select
    
End Sub

Private Sub cmb_pHRefInst2_Change()
    If cmb_pHRefInst2.Value = "" Then Exit Sub
    
    Select Case cmb_pHRefInst2
        Case "Bayer"
            txt_pHRefInst2SN.Value = "6245"
            txt_pHRefInst2CassetteType.Value = "NA"
            txt_pHRefInst2CassettLot.Value = "NA"
        Case "Radiometer"
            txt_pHRefInst2SN.Value = "126R0159N02"
            txt_pHRefInst2CassetteType.Value = "NA"
            txt_pHRefInst2CassettLot.Value = "NA"
        Case "I-Stat"
            txt_pHRefInst2SN.Value = "318221"
            txt_pHRefInst2CassetteType.Value = "CG4+"
            txt_pHRefInst2CassettLot.Value = ""
        Case "OPTI CCA"
            txt_pHRefInst2SN.Value = ""
            txt_pHRefInst2CassetteType.Value = ""
            txt_pHRefInst2CassettLot.Value = ""
        Case "OPTI LION"
            txt_pHRefInst2SN.Value = ""
            txt_pHRefInst2CassetteType.Value = ""
            txt_pHRefInst2CassettLot.Value = ""
        Case "NA"
            txt_pHRefInst2SN.Value = "NA"
            txt_pHRefInst2CassetteType.Value = "NA"
            txt_pHRefInst2CassettLot.Value = "NA"
            txt_pHRefInst2_Low1.Value = "NA"
            txt_pHRefInst2_Low2.Value = "NA"
            txt_pHRefInst2_Mid1.Value = "NA"
            txt_pHRefInst2_Mid2.Value = "NA"
            txt_pHRefInst2_High1.Value = "NA"
            txt_pHRefInst2_High2.Value = "NA"
    End Select

End Sub

Private Sub cmb_LacRefInst1_Change()
    If cmb_LacRefInst1.Value = "" Then Exit Sub
    
    Select Case cmb_LacRefInst1
        Case "Bayer"
            txt_LacRefInst1SN.Value = "6245"
            txt_LacRefInst1CassetteType.Value = "NA"
            txt_LacRefInst1CassettLot.Value = "NA"
        Case "I-Stat"
            txt_LacRefInst1SN.Value = "318221"
            txt_LacRefInst1CassetteType.Value = "CG4+"
            txt_LacRefInst1CassettLot.Value = ""
    End Select
End Sub

Private Sub cmb_LacRefInst2_Change()
    If cmb_LacRefInst2.Value = "" Then Exit Sub
    
    Select Case cmb_LacRefInst2
        Case "Bayer"
            txt_LacRefInst2SN.Value = "6245"
            txt_LacRefInst2CassetteType.Value = "NA"
            txt_LacRefInst2CassettLot.Value = "NA"
        Case "I-Stat"
            txt_LacRefInst2SN.Value = "318221"
            txt_LacRefInst2CassetteType.Value = "CG4+"
            txt_LacRefInst2CassettLot.Value = ""
        Case "NA"
            txt_LacRefInst2SN.Value = "NA"
            txt_LacRefInst2CassetteType.Value = "NA"
            txt_LacRefInst2CassettLot.Value = "NA"
            txt_LacRefInst2RawLowtHb1.Value = "NA"
            txt_LacRefInst2RawLowtHb2.Value = "NA"
            txt_LacRefInst2_Raw1.Value = "NA"
            txt_LacRefInst2_Raw2.Value = "NA"
    End Select
End Sub

Private Sub txtLotNo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim InvalidInput As Boolean
Dim strMessage As String
    InvalidInput = False
    
    If txtLotNo.Value = "" Then Exit Sub
    If Not (IsNumeric(txtLotNo.Value)) Then
        InvalidInput = True
        strMessage = "Lot number must be numeric."
    ElseIf Int(txtLotNo.Value) - txtLotNo.Value <> 0 Then
        InvalidInput = True
        strMessage = "Lot number must be an integer."
    ElseIf txtLotNo.Value < 100000 Or txtLotNo.Value > 999999 Then
        InvalidInput = True
        strMessage = "Lot number must be a postive number of 6 digits."
    ElseIf StyleNumberCheck(CDbl(txtLotNo.Value)) <> 65 Then
        InvalidInput = True
        strMessage = "The style number (digits 4 & 5) must equal 65."
    End If
    
    If InvalidInput = True Then
        MsgBox strMessage, vbCritical, "Invalid lot number input"
        txtLotNo.Value = ""
    End If
End Sub

Private Sub DropGeneralInfo()
    Range("Operator").Value = txtOperator.Value
    Range("TestDate").Value = txtTestDate.Value
    Range("Lot_No").Value = txtLotNo.Value
    Range("LotExpDate").Value = txtBLacExpirationDate.Value
    Range("AnalysisDate").Value = Date
    Range("OCL1_Lot").Value = txtOCL1LotNo.Value
    Range("OCL1_ExpDate").Value = txtOCL1ExpDate.Value
    Range("OCL2_Lot").Value = txtOCL2LotNo.Value
    Range("OCL2_ExpDate").Value = txtOCL2ExpDate.Value
    Range("OCL3_Lot").Value = txtOCL3LotNo.Value
    Range("OCL3_ExpDate").Value = txtOCL3ExpDate.Value
    Range("RNA5_Lot").Value = txtRNA5LotNo.Value
    Range("RNA5_ExpDate").Value = txtRNA5ExpDate.Value
End Sub

Private Sub DropCCASN()
    Range("CCA_1_S_N").Value = txtCCA1SN.Value
    Range("CCA_2_S_N").Value = txtCCA2SN.Value
    Range("CCA_3_S_N").Value = txtCCA3SN.Value
    Range("CCA_4_S_N").Value = txtCCA4SN.Value
End Sub

Private Sub Drop_pHReferenceAnalyzer()
    Range("pHRef1_AnalyzerType").Value = cmb_pHRefInst1.Value
    Range("pHRef1_AnalyzerSN").Value = txt_pHRefInst1SN.Value
    Range("pHRef1_CassetteStyle").Value = txt_pHRefInst1CassetteType.Value
    Range("pHRef1_CassetteLot").Value = txt_pHRefInst1CassettLot.Value
    Range("pHRef1_LowGas1").Value = txt_pHRefInst1_Low1.Value
    Range("pHRef1_LowGas2").Value = txt_pHRefInst1_Low2.Value
    Range("pHRef1_MidGas1").Value = txt_pHRefInst1_Mid1.Value
    Range("pHRef1_MidGas2").Value = txt_pHRefInst1_Mid2.Value
    Range("pHRef1_HighGas1").Value = txt_pHRefInst1_High1.Value
    Range("pHRef1_HighGas2").Value = txt_pHRefInst1_High2.Value
    
    Range("phRef2_AnalyzerType").Value = cmb_pHRefInst2.Value
    Range("phRef2_AnalyzerSN").Value = txt_pHRefInst2SN.Value
    Range("phRef2_CassetteStyle").Value = txt_pHRefInst2CassetteType.Value
    Range("phRef2_CassetteLot").Value = txt_pHRefInst2CassettLot.Value
    Range("phRef2_LowGas1").Value = txt_pHRefInst2_Low1.Value
    Range("phRef2_LowGas2").Value = txt_pHRefInst2_Low2.Value
    Range("phRef2_MidGas1").Value = txt_pHRefInst2_Mid1.Value
    Range("phRef2_MidGas2").Value = txt_pHRefInst2_Mid2.Value
    Range("phRef2_HighGas1").Value = txt_pHRefInst2_High1.Value
    Range("phRef2_HighGas2").Value = txt_pHRefInst2_High2.Value
End Sub

Private Sub Drop_LacReferenceAnalyzer()
    Range("LacRef1_AnalyzerType").Value = cmb_LacRefInst1.Value
    Range("LacRef1_AnalyzerSN").Value = txt_LacRefInst1SN.Value
    Range("LacRef1_CassetteStyle").Value = txt_LacRefInst1CassetteType.Value
    Range("LacRef1_CassetteLot").Value = txt_LacRefInst1CassettLot.Value
    Range("LacRef1_RBLowtHb1").Value = txt_LacRefInst1RawLowtHb1.Value
    Range("LacRef1_RBLowtHb2").Value = txt_LacRefInst1RawLowtHb2.Value
    Range("LacRef1_Raw1").Value = txt_LacRefInst1_Raw1.Value
    Range("LacRef1_Raw2").Value = txt_LacRefInst1_Raw2.Value
    
    Range("LacRef2_AnalyzerType").Value = cmb_LacRefInst2.Value
    Range("LacRef2_AnalyzerSN").Value = txt_LacRefInst2SN.Value
    Range("LacRef2_CassetteStyle").Value = txt_LacRefInst2CassetteType.Value
    Range("LacRef2_CassetteLot").Value = txt_LacRefInst2CassettLot.Value
    Range("LacRef2_RBLowtHb1").Value = txt_LacRefInst2RawLowtHb1.Value
    Range("LacRef2_RBLowtHb2").Value = txt_LacRefInst2RawLowtHb2.Value
    Range("LacRef2_Raw1").Value = txt_LacRefInst2_Raw1.Value
    Range("LacRef2_Raw2").Value = txt_LacRefInst2_Raw2.Value
End Sub

Private Sub Drop_SampleNumber(ByVal SampleNumberValue As String, ByVal SampleNumberSequence As Byte)
    
    Range("C" & 44 + SampleNumberSequence).Value = SampleNumberValue
    
End Sub

Private Sub UserSaveAs(ByVal CassetteLotNumber As String)
On Error GoTo ErrorHandler
Dim varBarcodeFile As Variant
Sheets("Results Summary").Select
Reattemptsave:
    varBarcodeFile = Application.GetSaveAsFilename("B-Lac_" & CassetteLotNumber & "_Barcode.xls", _
        "Excel Files(*.xls),*.xls", , "Save B-Lac Lot " & CassetteLotNumber & " Barcoding File")
    If varBarcodeFile = False Then
        MsgBox "Cannot abort file save dialogue.", vbCritical, "MUST SAVE FILE"
        GoTo Reattemptsave
    End If
    ActiveWorkbook.SaveAs varBarcodeFile
Exit Sub
ErrorHandler:
GoTo Reattemptsave
End Sub

Private Sub OpenCSVFile(ByVal CSVNetworkAddress As String)

    Workbooks.OpenText FileName:=CSVNetworkAddress _
    , Origin:=437, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
    xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, Semicolon:=False _
    , Comma:=True, Space:=False, Other:=False, FieldInfo:=Array(1, 1), _
    TrailingMinusNumbers:=True
    Set wkbDataBase = ActiveWorkbook
    
End Sub

Private Sub TransferDataRow(ByVal SampleNumberFromPrintOut As String, _
ByVal SampleSequenceNumber As Byte)

Dim Search(1) As String
Dim Col(1) As Double
Dim NumberSearch As String
Dim RowOffset As Byte
Dim CopyRow As Double
    
    NumberSearch = SampleNumberFromPrintOut
    
    Col(1) = 4
    Search(1) = "Raw Data"
    RowOffset = 23
    
    If SampleSequenceNumber = 0 Or SampleSequenceNumber = 1 Then
        Col(0) = 6
        Search(0) = "QC Level 1 Data"
    ElseIf SampleSequenceNumber = 2 Or SampleSequenceNumber = 3 Then
        Col(0) = 6
        Search(0) = "QC Level 2 Data"
    ElseIf SampleSequenceNumber = 4 Or SampleSequenceNumber = 5 Then
        Col(0) = 6
        Search(0) = "QC Level 3 Data"
    ElseIf SampleSequenceNumber = 6 Or SampleSequenceNumber = 7 Then
        Col(0) = 35
        Search(0) = "Patient Data"
    ElseIf SampleSequenceNumber > 7 Then
        Col(0) = 35
        Search(0) = "Patient Data"
        RowOffset = 24
    End If
    
    For n = 0 To 1
        Continue = True
        Call ExcelSearch(Search(n))
        If ErrorCode <> 0 Then Exit Sub
        Row = ActiveCell.Row + 3
        Do While Continue = True
            If IsEmpty(Cells(Row, Col(n))) Then
                ErrorCode = 2
                Exit Sub
            End If
            If Cells(Row, Col(n)).Value = NumberSearch Then
                Continue = False
                NumberSearch = Cells(Row, 4).Value
                CopyRow = Row
                If n = 0 Then
                    Call CopyPasteDataRow(CopyRow, SampleSequenceNumber + RowOffset)
                End If
            End If
            Row = Row + 1
        Loop
    Next n
    
    Call CopyPasteDataRow(CopyRow, SampleSequenceNumber)
    
End Sub

Private Sub ExcelSearch(ByVal SearchValue As String)
On Error GoTo SearchFail
    Range("A1").Select
    Cells.Find(What:=SearchValue, After:=ActiveCell, LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
Exit Sub
SearchFail:
    ErrorCode = 1
End Sub

Private Sub CopyPasteDataRow(ByVal DataLineRow As Double, ByVal SampleSequenceNumber As Byte)
    Range(Cells(DataLineRow, 1), Cells(DataLineRow, 234)).Copy
    wkbBarcode.Activate
    Range("A" & 1 + SampleSequenceNumber).Select
    Call PasteSpecialAsValues
    wkbDataBase.Activate
End Sub

Private Sub PasteSpecialAsValues()
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
Application.CutCopyMode = False
End Sub

Private Sub DropCSVFileLocation(ByVal CSVNetworkAddress As String)
    Range("D20").Value = CSVNetworkAddress
End Sub

Private Function FirstRunCheck() As Boolean
    Sheets("Targets & Limits").Select
    FirstRunCheck = Range("First_Run_Check").Value
End Function

Private Function SampleNumberNotFound(ByVal CCA_Number As Byte, ByVal SampleReferenceNo As Byte _
, ByVal UserEnteredNumber As String) As String

Dim SampleTitle(17) As String
Dim CCASN(3) As String
        
        SampleTitle(0) = "OPTI Check Level 1 (1)"
        SampleTitle(1) = "OPTI Check Level 1 (2)"
        SampleTitle(2) = "OPTI Check Level 2 (1)"
        SampleTitle(3) = "OPTI Check Level 2 (2)"
        SampleTitle(4) = "OPTI Check Level 3 (1)"
        SampleTitle(5) = "OPTI Check Level 3 (2)"
        SampleTitle(6) = "RNA QC 823 Level 5 (1)"
        SampleTitle(7) = "RNA QC 823 Level 5 (2)"
        SampleTitle(8) = "Low-Gas Tonometry Blood (1)"
        SampleTitle(9) = "Low-Gas Tonometry Blood (2)"
        SampleTitle(10) = "Mid-Gas Tonometry Blood (1)"
        SampleTitle(11) = "Mid-Gas Tonometry Blood (2)"
        SampleTitle(12) = "High-Gas Tonometry Blood (1)"
        SampleTitle(13) = "High-Gas Tonometry Blood (2)"
        SampleTitle(14) = "Raw Blood (1)"
        SampleTitle(15) = "Raw Blood (2)"
        SampleTitle(16) = "Raw Blood Low tHb (1)"
        SampleTitle(17) = "Raw Blood Low tHb (2)"
        
        CCASN(0) = txtCCA1SN.Value
        CCASN(1) = txtCCA2SN.Value
        CCASN(2) = txtCCA3SN.Value
        CCASN(3) = txtCCA4SN.Value
        
        SampleNumberNotFound = "The following user entered sample number could not be found:" & Chr(10) _
            & Chr(10) & "CCA SN: " & CCASN(CCA_Number) & Chr(10) & SampleTitle(SampleReferenceNo) _
            & Chr(10) & "Sample number entered: " & UserEnteredNumber & Chr(10) & Chr(10) & _
            "Please check the number and reattempt analysis."
        
End Function

Private Sub Lock_Unlock_Workbook(ByVal Choice As String)
    If Choice = "lock" Then
        For i = 1 To Worksheets.Count
            Worksheets(i).Protect Password:=strPassword
        Next i
        ActiveWorkbook.Protect Password:=strPassword
    ElseIf Choice = "unlock" Then
        For i = 1 To Worksheets.Count
            Worksheets(i).Unprotect Password:=strPassword
        Next i
        ActiveWorkbook.Unprotect Password:=strPassword
    End If
End Sub

Private Function StyleNumberCheck(ByVal LotNumberEntered As Double) As Double
Dim x As Double
Dim y As Double
    x = LotNumberEntered
    y = (Int(x / 1000) * 1000)
    StyleNumberCheck = Int((x - y) / 10)
End Function
