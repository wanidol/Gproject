VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_workRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Option Explicit
Dim Db As Database
Dim rst As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String


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
        lblNature.Caption = "List Of Nature Documents"
        lblMark.Caption = "List of Marking Documents"
        lblCritere.Caption = "List of Criteria Documents"
        lblAffectation.Caption = "List of Affectation Documents"
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
