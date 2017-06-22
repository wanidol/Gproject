VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_DocsByNature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Sub Report_NoData(Cancel As Integer)
On Error GoTo ErrHandler

    MsgBox "No data found! Closing report."
    'DoCmd.Close acReport, "rpt_DocsByNature", acSaveNo

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
    
        lblRHComment.Caption = "Comment :"
        lblRHeader.Caption = "List of Nature Documents"
        lblRHEnd1.Caption = "end "
        lblRHEnd2.Caption = "end "
        lblRHEtat.Caption = "Status start :"
        lblRHNature.Caption = "Nature :"
        lblRHOrigin.Caption = "Date start"
        
        
        lblHComment.Caption = "Comment"
        lblHDesc.Caption = "Description"
        lblHDocNo.Caption = "Document No."
        lblHEtat.Caption = "Status"
        lblHIndex.Caption = "Index"
        lblHnature.Caption = "Nature Code"
        lblHOrigin.Caption = "Date Origin"
        
    Else 'FR
    
       lblRHComment.Caption = "Commentaires :"
        lblRHeader.Caption = "Liste des documents d'une nature"
        lblRHEnd1.Caption = "end "
        lblRHEnd2.Caption = "end "
        lblRHEtat.Caption = "Etat début :"
        lblRHNature.Caption = "Nature des documents : "
        lblRHOrigin.Caption = "Date début"
        
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
      Me.lblNatureID.Caption = strOpenArgs(0)
      Me.lblEtatStart.Caption = strOpenArgs(1)
      Me.lblEtatStop.Caption = strOpenArgs(2)
      Me.lblDateStart.Caption = strOpenArgs(3)
      Me.lblDateStop.Caption = strOpenArgs(4)
      Me.lblComment.Caption = strOpenArgs(5)
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


'DoCmd.Close acReport, "rpt_DocsByNature", acSaveNo
DoCmd.OpenForm "frm_DocByNature", acNormal

'Dim frm As Form
'
'    For Each frm In Application.Forms
'        If frm.Name = PreviousForm Then
'            frm.Visible = True
'            Exit For
'        End If
'    Next frm


    'DoCmd.Close acReport, "rpt_DocsByMark", acSaveNo
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

