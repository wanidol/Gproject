VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_DocsByMark"
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

Private Sub Report_NoData(Cancel As Integer)
On Error GoTo ErrHandler

    MsgBox "No data found! Closing report."
    'DoCmd.Close acReport, "rpt_DocsByMark", acSaveNo

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub
Private Sub Report_Close()
On Error GoTo ErrHandler


'DoCmd.Close acReport, "rpt_DocsByMark", acSaveNo
DoCmd.ShowToolbar "Ribbon", acToolbarNo
DoCmd.OpenForm "frm_DocByMark", acNormal

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
        
        lblRHComment.Caption = "Comment : "
        lblRHCritere.Caption = "Criteria :"
        lblRHeader.Caption = "List of Marking Documents"
        lblRHEtat.Caption = "Etat début "
        lblRHMarkID.Caption = "Mark Code : "
        lblRHEnd1.Caption = "end "
        lblRHEnd1.Caption = "end "
        lblRHNatureID.Caption = "Nature documents : "
        lblRHOrigin.Caption = "Date start "
        
       
'        lblHCritere.Caption = "Criteria"
        lblHmark.Caption = "Mark Code"
        lblHComment.Caption = "Comment"
        lblHDesc.Caption = "Description"
        lblHDocNo.Caption = "Document No."
        lblHEtat.Caption = "Status"
        lblHIndex.Caption = "Index"
        lblHnature.Caption = "Nature Code"
        lblHOrigin.Caption = "Date Origin"
        
    Else 'FR
    
        lblRHComment.Caption = "Commentaires : "
        lblRHCritere.Caption = "Critères :"
        lblRHeader.Caption = "liste des documents d'un marquage"
        lblRHEtat.Caption = "Etat début "
        lblRHMarkID.Caption = "Marquage de pièce : "
        lblRHEnd1.Caption = "fin "
        lblRHEnd1.Caption = "fin "
        lblRHNatureID.Caption = "Nature documents : "
        lblRHOrigin.Caption = "Date début "
        
   
'
'        lblHCritere.Caption = "Critères"
        lblHmark.Caption = "Marquage"
        lblHComment.Caption = "Commentaire"
        lblHDesc.Caption = "Libellé du document"
        lblHDocNo.Caption = "N° Document"
        lblHEtat.Caption = "Etat"
        lblHIndex.Caption = "Indice"
        lblHnature.Caption = "Nature"
        lblHOrigin.Caption = "Origine Date"
    
    End If
    
    
    If Not IsNull(GB_RPT_TITLE) Then
      strOpenArgs = Split(GB_RPT_TITLE, ";")
      Me.lblMarkID.Caption = strOpenArgs(0)
      Me.lblNatureID.Caption = strOpenArgs(1)
      Me.lblEtatStart.Caption = strOpenArgs(2)
      Me.lblEtatStop.Caption = strOpenArgs(3)
      Me.lblDateStart.Caption = strOpenArgs(4)
      Me.lblDateStop.Caption = strOpenArgs(5)
      Me.lblComment.Caption = strOpenArgs(6)
      Me.lblCritere.Caption = strOpenArgs(7)
    End If

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

    
End Sub


