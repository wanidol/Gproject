VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_workRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : frm_workRequest
' Author    : Wanidol Lievaux
' Website   : https://www.linkedin.com/in/wanidol-lievaux-0a7bb7124/
' Purpose   :
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2017-jul-25
'---------------------------------------------------------------------------------------

Option Compare Database

Option Explicit





Private Sub cmdAffection_Click()
On Error GoTo ErrHandler

    DoCmd.OpenForm "frm_DocByAffectation", acNormal
    DoCmd.Close acForm, "frm_workRequest"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub CmdCritere_Click()
On Error GoTo ErrHandler

    DoCmd.OpenForm "frm_DocByCritere", acNormal
    DoCmd.Close acForm, "frm_workRequest"


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
    DoCmd.Close acForm, "frm_workRequest"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdMark_Click()
On Error GoTo ErrHandler

    DoCmd.OpenForm "frm_DocByMark", acNormal
    DoCmd.Close acForm, "frm_workRequest"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdNature_Click()
On Error GoTo ErrHandler

    DoCmd.OpenForm "frm_DocByNature", acNormal
    DoCmd.Close acForm, "frm_workRequest"


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
        lblHeader.Caption = "Work Request"
        lblNature.Caption = "Document List by Nature"
        lblMark.Caption = "Document List by Stamp"
        lblCritere.Caption = "Document List by Criteria"
        lblAffectation.Caption = "Document List by Assignement"
    Else 'FR
        lblHeader.Caption = "Demandes de travaux"
        lblNature.Caption = "Liste des documents d'une nature"
        lblMark.Caption = "Liste des documents d'un marquage de pièce"
        lblCritere.Caption = "Liste des documents d'une critère"
        lblAffectation.Caption = "Liste des documents d'une affectation"
    End If
        
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

