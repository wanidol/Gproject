VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_consultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : frm_consultation
' Author    : Wanidol Lievaux
' Website   : https://www.linkedin.com/in/wanidol-lievaux-0a7bb7124/
' Purpose   :
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged
'
' Revision History:
' Rev   App.version    Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         1.0        2017-jul-25
' 2         2.0        05/07/2017 - Changed font family from copper black to Arial black because
'                                   runtime doesn't support
'                                 - changed caption in v.english from assignement to assignment
'---------------------------------------------------------------------------------------Option Compare Database

Private Sub cmd_A_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_listDocByID", acNormal
DoCmd.Close acForm, "frm_consultation"



Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmd_B_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_listDocByMark", acNormal
DoCmd.Close acForm, "frm_consultation"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub Cmd_C_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_listDocByCritere", acNormal
DoCmd.Close acForm, "frm_consultation"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub


Private Sub cmdExit_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_Main", acNormal
DoCmd.Close acForm, "frm_consultation"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler

    If GB_LANG = "EN" Then
        lblHeader.Caption = "Consultation/Research Documents"
        lblDocument.Caption = "Document Follow-up "
        lblMark.Caption = "Document list related to a part stamp "
        lblCritere.Caption = "Document list related to an access criteria"
    Else 'FR
        lblHeader.Caption = "Consutation/Recherche des documents liés"
        lblDocument.Caption = "Consultation Suivi Document"
        lblMark.Caption = "Liste des documents liés à un marquage de pièce "
        lblCritere.Caption = "Liste des documents liés à un critère d'accès"
    End If
        
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

