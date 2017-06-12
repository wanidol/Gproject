VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_consultation_en"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database



Private Sub cmd_A_Click()
On Error GoTo ErrHandler
DoCmd.OpenForm "frm_listDocByID_en", acNormal
DoCmd.Close acForm, "frm_consultation_en"



Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmd_B_Click()
On Error GoTo ErrHandler
DoCmd.OpenForm "frm_listDocByMark_en", acNormal
DoCmd.Close acForm, "frm_consultation_en"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub Cmd_C_Click()
On Error GoTo ErrHandler
DoCmd.OpenForm "frm_listDocByCritere_en", acNormal
DoCmd.Close acForm, "frm_consultation_en"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub


Private Sub cmdExit_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_Main_en", acNormal
DoCmd.Close acForm, "frm_consultation_en"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
