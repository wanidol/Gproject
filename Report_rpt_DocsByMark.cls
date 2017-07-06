VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_DocsByMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : rpt_DocsByMark
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
'---------------------------------------------------------------------------------------

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


Private Sub Report_Close()
On Error GoTo ErrHandler


'show toolbar when report open / hide toolbar when report close
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
        lblRHeader.Caption = "Document List by stamp"
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
        lblHDocNo.Caption = "Document Ref"
        lblHEtat.Caption = "Status"
        lblHIndex.Caption = "Index"
        lblHnature.Caption = "Nature Code"
        lblHOrigin.Caption = "Date Origin"
        
        lblNoRecord.Caption = "Number Of Record"
        
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
        lblHDocNo.Caption = "N° de Document"
        lblHEtat.Caption = "Etat"
        lblHIndex.Caption = "Indice"
        lblHnature.Caption = "Nature"
        lblHOrigin.Caption = "Origine Date"
        
        lblNoRecord.Caption = "Nombre d'enregistrement"
    
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


