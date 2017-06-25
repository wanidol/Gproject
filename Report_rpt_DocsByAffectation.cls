VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_DocsByAffectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Load()
On Error GoTo ErrHandler

    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

'Private Sub PrintReport()
'
'    Dim oDoc As clsDocuments
'
'    For Each oDoc In GB_cDocs
'
'         lblDdocNo.Caption = oDoc.DocID
'         lblDNature.Caption = oDoc.NatureID
'         lblDIndex.Caption = oDoc.Indice
'         lblDLibelle.Caption = oDoc.Libelle
'         lblDComment.Caption = oDoc.Commentaire
'         lblDEtat.Caption = oDoc.Etat
'         lblDOrigin.Caption = oDoc.Origin
'
'    Next
'
'End Sub
Private Sub Report_NoData(Cancel As Integer)
On Error GoTo ErrHandler

    MsgBox "No data found! Closing report."
    'DoCmd.Close acReport, "rpt_docsByAffectation", acSaveNo

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub Report_Open(Cancel As Integer)
On Error GoTo ErrHandler
    Dim strOpenArgs() As String
    Dim Str As String
    
    If GB_LANG = "EN" Then
    
        lblRHAffect.Caption = "Affectation Code : "
        lblRHEnd.Caption = "end"
        lblRHComment.Caption = "Comment : "
        lblRHClientLinked.Caption = "Related to customers : "
        lblRHeader.Caption = "List of Affectation Documents"
        lblRHEtat.Caption = "Status start : "
        lblRHNature.Caption = "Nature Code : "
        
        lblHComment.Caption = "Comment"
        lblHDesc.Caption = "Description"
        lblHDocNo.Caption = "Document No."
        lblHEtat.Caption = "Status"
        lblHIndex.Caption = "Index"
        lblHnature.Caption = "Nature Code"
        lblHOrigin.Caption = "Date Origin"
        lblHAffectation.Caption = "Affectation"
        
        
    Else 'FR
    
        lblRHAffect.Caption = "Affectation : "
        lblRHEnd.Caption = "fin"
        lblRHComment.Caption = "Commentaires : "
        lblRHClientLinked.Caption = "Liés à un client : "
        lblRHeader.Caption = "Liste des documents d'une affectation"
        lblRHEtat.Caption = "Etat début: "
        lblRHNature.Caption = "Nature : "
        
        lblHComment.Caption = "Commentaires"
        lblHDesc.Caption = "Libellé du document"
        lblHDocNo.Caption = "N° Document"
        lblHEtat.Caption = "Etat"
        lblHIndex.Caption = "Indice"
        lblHnature.Caption = "Nature"
        lblHOrigin.Caption = "Date"
        lblHAffectation.Caption = "Affectation"
    
    
    End If
    
    
    
    
    If Not IsNull(GB_RPT_TITLE) Then
             
        strOpenArgs = Split(GB_RPT_TITLE, ";")
        Me.lblAffectID.Caption = strOpenArgs(0)
        Me.lblNatureID.Caption = strOpenArgs(1)
        Me.lblEtatStart.Caption = strOpenArgs(2)
        Me.lblEtatStop.Caption = strOpenArgs(3)
        Me.lblCritereLinked.Caption = strOpenArgs(4)
        Me.lblComment.Caption = strOpenArgs(5)
    
        
    End If
    
    'Call PrintReport

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

    
End Sub
Private Sub Report_Close()
On Error GoTo ErrHandler


'DoCmd.Close acReport, "rpt_DocsByAffectation", acSaveNo
DoCmd.ShowToolbar "Ribbon", acToolbarNo
DoCmd.OpenForm "frm_DocByAffectation", acNormal

 

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

