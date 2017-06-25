VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_DocsByCritere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Load()

    DoCmd.ShowToolbar "Ribbon", acToolbarYes

End Sub

Private Sub Report_NoData(Cancel As Integer)
On Error GoTo ErrHandler

    MsgBox "No data found! Closing report."
    'DoCmd.Close acReport, "rpt_docsByCritere", acSaveNo

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
    
        lblRHType.Caption = "Type :"
        lblRHCritere.Caption = "N° criteria:"
        lblRHNature.Caption = "Nature : "
        lblRHEnd1.Caption = "end"
        lblRHEnd1.Caption = "end"
        lblRHComment.Caption = "Comment : "
        lblRHCritereLinked.Caption = "Criteria Linked : "
        lblRHeader.Caption = "List of Criteria documents"
        lblRHEtat.Caption = "Status start: "
        lblRHOrigin.Caption = "Date start: "
        
               
'        lblHCritere.Caption = "N° criteria"
'        lblHType.Caption = "Type"
        lblHComment.Caption = "Comment"
        lblHDesc.Caption = "Description"
        lblHDocNo.Caption = "Document No."
        lblHEtat.Caption = "Status"
        lblHIndex.Caption = "Index"
        lblHnature.Caption = "Nature Code"
        lblHOrigin.Caption = "Date Origin"
        
    Else 'FR
    
        lblRHType.Caption = "Type"
        lblRHCritere.Caption = "N° critère :"
        lblRHNature.Caption = "Nature : "
        lblRHEnd1.Caption = "fin "
        lblRHEnd1.Caption = "fin "
        lblRHComment.Caption = "Commentaires : "
        lblRHCritereLinked.Caption = "Critères liés : "
        lblRHeader.Caption = "Liste des documents d'un critère"
        lblRHEtat.Caption = "Status début : "
        lblRHOrigin.Caption = "Date début : "
        
'
'        lblHCritere.Caption = "N° critère"
''        lblHType.Caption = "Type"
        lblHComment.Caption = "Commentaires"
        lblHDesc.Caption = "Libellé du document"
        lblHDocNo.Caption = "N° Document"
        lblHEtat.Caption = "Etat"
        lblHIndex.Caption = "Indice"
        lblHnature.Caption = "Nature"
        lblHOrigin.Caption = "Origine"
    
    End If
    
    
    If Not IsNull(GB_RPT_TITLE) Then
      strOpenArgs = Split(GB_RPT_TITLE, ";")
      Me.lblTypeID.Caption = strOpenArgs(0)
      Me.lblCritereID.Caption = strOpenArgs(1)
      Me.lblNatureID.Caption = strOpenArgs(2)
      Me.lblEtatStart.Caption = strOpenArgs(3)
      Me.lblEtatStop.Caption = strOpenArgs(4)
      Me.lblDateStart.Caption = strOpenArgs(5)
      Me.lblDateStop.Caption = strOpenArgs(6)
      Me.lblCritereLinked.Caption = strOpenArgs(7)
      Me.lblComment.Caption = strOpenArgs(8)
      
    End If

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

    
End Sub


Private Sub Report_Close()
On Error GoTo ErrHandler


'DoCmd.Close acReport, "rpt_DocByCritere", acSaveNo
DoCmd.ShowToolbar "Ribbon", acToolbarNo
DoCmd.OpenForm "frm_DocByCritere", acNormal



Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

