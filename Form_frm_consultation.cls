VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_consultation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

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
        lblHeader.Caption = "Consultation - Document Search"
        lblDocument.Caption = "Consultation Follow-up Document"
        lblMark.Caption = "List of documents related to a marking piece "
        lblCritere.Caption = "List of documents related to an access criterion"
    Else 'FR
        lblHeader.Caption = "Consultation - Document Search"
        lblDocument.Caption = "Consultation Suivi Document"
        lblMark.Caption = "Liste des documents liés à une pièce de marquage "
        lblCritere.Caption = "Liste des documents liés à un critère d'accès"
    End If
        
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub
