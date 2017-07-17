VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_DocsByAffectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : rpt_DocsByAffectation
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



Private Sub Report_Open(Cancel As Integer)
On Error GoTo ErrHandler
    Dim strOpenArgs() As String
    Dim Str As String
    
    If Not IsNull(GB_LANG) And Not (GB_LANG = "") Then
        If GB_LANG = "EN" Then
        
            lblRHAffect.Caption = "Assignment Code : "
            lblRHEnd.Caption = "end"
            lblRHComment.Caption = "Comment : "
            lblRHClientLinked.Caption = "Related to customers : "
            lblRHeader.Caption = "Document List by Assignment"
            lblRHEtat.Caption = "Status start : "
            lblRHNature.Caption = "Nature Code : "
            
            lblHComment.Caption = "Comment"
            lblHDesc.Caption = "Description"
            lblHDocNo.Caption = "Document Ref"
            lblHEtat.Caption = "Status"
            lblHIndex.Caption = "Index"
            lblHnature.Caption = "Nature"
            lblHOrigin.Caption = "Date Origin"
'            lblHAffectation.Caption = "Assignment"
            
            lblNoRecord.Caption = "Number Of Record"
            
            
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
            lblHDocNo.Caption = "N° de Document"
            lblHEtat.Caption = "Etat"
            lblHIndex.Caption = "Indice"
            lblHnature.Caption = "Nature"
            lblHOrigin.Caption = "Origine"
'            lblHAffectation.Caption = "Affectation"
            
            lblNoRecord.Caption = "Nombre d'enregistrement"
        
        
        End If
    End If
        
    
    
    If Not IsNull(GB_RPT_TITLE) And Not (GB_RPT_TITLE = "") Then
             
        strOpenArgs = Split(GB_RPT_TITLE, ";")
        Me.lblAffectID.Caption = strOpenArgs(0)
        Me.lblNatureID.Caption = strOpenArgs(1)
        Me.lblEtatStart.Caption = strOpenArgs(2)
        Me.lblEtatStop.Caption = strOpenArgs(3)
        Me.lblCritereLinked.Caption = strOpenArgs(4)
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


'show toolbar when report open / hide toolbar when report close


DoCmd.ShowToolbar "Ribbon", acToolbarNo
DoCmd.OpenForm "frm_DocByAffectation", acNormal

 

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

