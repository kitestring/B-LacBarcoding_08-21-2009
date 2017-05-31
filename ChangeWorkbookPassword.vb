Attribute VB_Name = "Module1"
Option Explicit

Sub ChangeWorkbookPassword()
On Error GoTo BadPassword
Dim OldPassword As String
Dim NewPassword1 As String
Dim NewPassword2 As String
Dim i As Byte

OldPassword = InputBox("Enter old password", "Workbook password change")
NewPassword1 = InputBox("Enter new password", "Workbook password change")
NewPassword2 = InputBox("Reenter new password", "Workbook password change")

    If NewPassword1 <> NewPassword2 Then
        MsgBox "New password inconsistency, please reenter."
        Exit Sub
    End If
    
    For i = 1 To Worksheets.Count
        Worksheets(i).Unprotect Password:=OldPassword
    Next i
    ActiveWorkbook.Unprotect Password:=OldPassword
    
    For i = 1 To Worksheets.Count
        Worksheets(i).Protect Password:=NewPassword1
    Next i
    ActiveWorkbook.Protect Password:=NewPassword1
Exit Sub

BadPassword:
    MsgBox "Old password is incorrect"
End Sub
