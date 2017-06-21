VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_Main"
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

Private Sub UserForm_Initialize()
On Error GoTo ErrHandler

    'default Eng cmdFr open cmdEng Close
    If GB_LANG = "EN" Then
        'Me.cmdExit.SetFocus 'Move focus before hiding button
        txtHelper.Locked = True
        Me.txtHelper.SetFocus
        cmdEng.Visible = False
        cmdEng.Enabled = False
        cmdFr.Visible = True
        cmdFr.Enabled = True
        lbl_MenuConsult.Caption = "Consultation/Research Documents"
        lbl_MenuWorkRequest.Caption = "Work Request"
        'cmdConsult.SetFocus
    Else
        'Me.cmdExit.SetFocus 'Move focus before hiding button
        txtHelper.Locked = True
        Me.txtHelper.SetFocus
        cmdEng.Visible = True
        cmdEng.Enabled = True
        cmdFr.Visible = False
        cmdFr.Enabled = False
        lbl_MenuConsult.Caption = "Consutation/Recherche des documents liés"
        lbl_MenuWorkRequest.Caption = "Demande de travaux"
        'cmdConsult.SetFocus
    End If
    
    

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub


Private Sub cmdConsult_Click()
On Error GoTo ErrHandler

    DoCmd.OpenForm "frm_consultation", acNormal
    DoCmd.Close acForm, "frm_Main"

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub cmdEng_Click()


    If GB_LANG = "EN" Then
        GB_LANG = "FR"
    Else
        GB_LANG = "EN"
    End If
    
   Call UserForm_Initialize
   

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

Private Sub cmdFr_Click()


    If GB_LANG = "EN" Then
        GB_LANG = "FR"
    Else
        GB_LANG = "EN"
    End If
    
   Call UserForm_Initialize
   

End Sub



Private Sub cmdWorkRequest_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_workRequest", acNormal
DoCmd.Close acForm, "frm_Main"

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
    
    'DoCmd.Maximize
    
 
    If GB_LANG = "" Then
        Call InitGlobal
    End If
    
    Call UserForm_Initialize
    
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
