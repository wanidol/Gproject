VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_TablesDeSidoq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure : frm_TablesDeSidoq
' Author    : Wanidol Lievaux
' Website   : https://www.linkedin.com/in/wanidol-lievaux-0a7bb7124/
' Purpose   :
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged
'
' Revision History:
' Rev   App.version    Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         1.0        2017-jul-25 - Show all data from each table
' 2         2.0        05/07/2017 - Changed font family from copper black to Arial black because
'                                   runtime doesn't support
'                                 - changed caption in v.english from assignement to assignment
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Dim Db As Database


Dim cn As New ADODB.Connection
Dim Status As String

Dim sql As String

Private Sub cmdExit_Click()
On Error GoTo ErrHandler

DoCmd.OpenForm "frm_Main", acNormal
DoCmd.Close acForm, "frm_TablesDeSidoq"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Private Sub Form_Load()
    On Error GoTo ErrHandler

    Set Db = CurrentDb
    
    DoCmd.Maximize
    'DoCmd.Minimize
    
    
    If GB_LANG = "EN" Then
        lblHeader.Caption = "All Tables Of Sidoq"
    Else
    'Fr
        lblHeader.Caption = "Tables De Sidoq"
    End If
    
    lstAffect.ColumnCount = 5
    
   sql = "SELECT * FROM  AFFECTATION ORDER BY CODE_AFFECTATION "
          
    Call fReturnRecordset(sql, lstAffect)
    'Call UserForm_Initialize

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub TabCtl_Change()
 On Error GoTo ErrHandler
    
        Select Case TabCtl.Value
        
            Case 0
               ' Affectation
                lstAffect.ColumnCount = 5
                sql = "SELECT * FROM  AFFECTATION ORDER BY CODE_AFFECTATION "
                Call fReturnRecordset(sql, lstAffect)
            
            Case 1
             '   EtatNature
                lstEtatNature.ColumnCount = 5
                sql = "SELECT * FROM  ETATS_DOCUMENTS_NATURE ORDER BY CODE_NATURE_DOCUMENT "
                Call fReturnRecordset(sql, lstEtatNature)
            
            Case 2
             '   Liens
                lstLiens.ColumnCount = 7
                sql = "SELECT * FROM  LIENS ORDER BY  CODE_TYPE_CRITERE"
                Call fReturnRecordset(sql, lstLiens)
            
            Case 3
             '   Critere
                lstCritere.ColumnCount = 5
                sql = "SELECT * FROM  MOTS_CLE_CRITERE ORDER BY CODE_NATURE_DOCUMENT "
                Call fReturnRecordset(sql, lstCritere)
            
            Case 4
              ' Nature
                lstNature.ColumnCount = 11
                lstNature.ColumnWidths = "4.2cm;8.4cm"
                sql = "SELECT * FROM  NATURES_DOCUMENT ORDER BY CODE_NATURE_DOCUMENT "
                Call fReturnRecordset(sql, lstNature)
            
            Case 5
               ' Marquage
                lstmark.ColumnCount = 4
                sql = "SELECT * FROM  REPERES_MARQUAGE ORDER BY CODE_REPERE_MARQUAGE "
                Call fReturnRecordset(sql, lstmark)
            
            Case 6
                'TypeCritere
                lstType.ColumnCount = 7
                lstType.ColumnWidths = "6.6cm;6.6cm;6.6cm;12.12cm"
                sql = "SELECT * FROM  TYPES_CRITERES_NATURE ORDER BY CODE_NATURE_DOCUMENT"
                Call fReturnRecordset(sql, lstType)
            
            Case 7
               ' Unites
                lstUnites.ColumnCount = 5
                sql = "SELECT * FROM  UNITES ORDER BY CODE_UNITE "
                Call fReturnRecordset(sql, lstUnites)
            
            Case 8
                'Documents
                lstDocument.ColumnCount = 11
                lstDocument.ColumnWidths = "4.2cm;4.2cm;2.1cm;4.2cm;2.1cm;8.4cm;4.2cm;4.2cm;8.4cm"
                
                 sql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,CODE_ETAT_DOCUMENT,CODE_AFFECTATION," & _
                        "INDICE_MODIFICATION,LIBELLE_DOCUMENT,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT," & _
                        "COMMENTAIRES_SUIVI,CODE_REPERE_MARQUAGE,CONTRACTUEL " & _
                        "FROM  DOCUMENTS ORDER BY NUMERO_DOCUMENT "
                Call fReturnRecordset(sql, lstDocument)
                
        End Select
Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub


