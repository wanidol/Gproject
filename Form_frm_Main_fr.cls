VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Main_fr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Db As Database
Dim rs As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Sql As String
Dim S1 As String

Private Sub cmdConsult_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_consultation_en", acNormal
DoCmd.Close acForm, "frm_Main_fr"

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub cmdEng_Click()
On Error GoTo ErrHandler

    DoCmd.OpenForm "frm_Main_en", acNormal
    DoCmd.Close acForm, "frm_Main_fr"
    
Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandler
Call QUITFORM

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdWorkRequest_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_workRequest_en", acNormal
DoCmd.Close acForm, "frm_Main_fr"

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub




