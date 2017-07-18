VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : frm_DocsByMark
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
Dim Db As Database
Dim rst As New ADODB.Recordset
Dim rstMark As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
Dim rstEtatdate As New ADODB.Recordset
Dim rstDate As New ADODB.Recordset
Dim rstComment As New ADODB.Recordset
Dim rstCritere As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim NatureID As String
Dim critereID As String
Dim publicSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler

    moreSql = ""
        
    If IsNull(cmbMarkID.Value) Or (cmbMarkID.Value = "") Or (cmbMarkID.ListIndex = -1) Then
        
        
        If GB_LANG = "EN" Then
            MsgBox "Invalid Mark Code", vbOKOnly
        Else
            MsgBox "N° de marquage Invalide ", vbOKOnly
        End If
        
        cmbMarkID.SetFocus
        CheckValid = False
        
        Exit Function
        
    Else 'else main if
    
        If Not (cmbNatureID.ListIndex = -1) Then
            
             NatureID = cmbNatureID.Value
        
        Else 'case invalid data
            If Not IsNull(cmbNatureID.Value) Then
                If GB_LANG = "Fr" Then
                   MsgBox " Nature Invalide", vbOKOnly
                Else
                    MsgBox "Invalid Nature Code", vbOKOnly
                End If
                
                cmbNatureID.SetFocus
                CheckValid = False
                Exit Function
            End If
       
        End If
        
        If Not (cmbEtatStart.ListIndex = -1) Then
            
            'start ok
                If Not (cmbEtatEnd.ListIndex = -1) Then
                 'end ok
                    moreSql = moreSql & " And DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatEnd.Value & "'"
                    
                Else 'valid data listindex = -1 and not null end not ok
                    If GB_LANG = "Fr" Then
                        MsgBox "Date de fin Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid end date", vbOKOnly
                    End If
                    
                    cmbEtatEnd.SetFocus
                    CheckValid = False
                    Exit Function
                 
                End If
        Else 'case invalid data
            If Not IsNull(cmbEtatStart.Value) Then
                If GB_LANG = "Fr" Then
                    MsgBox "Date de début Invalide ", vbOKOnly
                Else
                    MsgBox "Invalid start date", vbOKOnly
                End If
                
                cmbEtatStart.SetFocus
                CheckValid = False
                Exit Function
            End If
        End If
        
        
        
        If Not (cmbStart.ListIndex = -1) Then
            
            'Start ok
            If Not (cmbEnd.ListIndex = -1) Then
            'End ok
              moreSql = moreSql & " And DATE_ORIGINE_DOCUMENT BETWEEN '" & cmbStart.Value & "' AND '" & cmbEnd.Value & "'"
            Else
            'end not Ok
                 If GB_LANG = "Fr" Then
                        MsgBox "Date de fin Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid end date", vbOKOnly
                    End If
                
                cmbEnd.SetFocus
                CheckValid = False
                Exit Function
           End If
            
        Else 'case invalid data listindex = -1
            If Not IsNull(cmbStart.Value) Then
                
                If GB_LANG = "Fr" Then
                    MsgBox "Date de début Invalide ", vbOKOnly
                Else
                    MsgBox "Invalid start date", vbOKOnly
                End If
                
                cmbStart.SetFocus
                CheckValid = False
                Exit Function
            End If
       
        End If
              
        
        If Not (cmbComment.ListIndex = -1) Then
            moreSql = moreSql & " AND COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
        Else 'case invalid data listindex = -1
            If Not IsNull(cmbComment.Value) Then
                If GB_LANG = "Fr" Then
                    MsgBox "Commentaire Invalide ", vbOKOnly
                Else
                    MsgBox "Invalid Comment", vbOKOnly
                End If
                cmbComment.SetFocus
                CheckValid = False
                Exit Function
            End If
        End If
        
        If Not (cmbCritere.ListIndex = -1) Then
           
           critereID = cmbCritere.Value
             
        Else 'case invalid data listindex = -1
            If Not IsNull(cmbCritere.Value) Then
                If GB_LANG = "Fr" Then
                    MsgBox "Critère Invalide ", vbOKOnly
                Else
                    MsgBox "Invalid Criteria Code", vbOKOnly
                End If
                cmbCritere.SetFocus
                CheckValid = False
                Exit Function
            End If
        End If
        
        CheckValid = True
    End If 'end main if
  
Exit_Sub:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Function

Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
    Dim sql As String
    Dim typeID As String
    Dim critId As String
    Dim docNum As String

    'Clear Variable
    GB_RPT_TITLE = ""
    moreSql = ""
    NatureID = ""
    critereID = ""
    
    'Clear all Controls
    Call ClearAll(Me)
    
   If GB_LANG = "EN" Then
        lblHeader.Caption = "Document List by stamp"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblMark.Caption = "Part Stamp :"
        lblComment.Caption = "Comment : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "start : "
        lblEStop.Caption = "end : "
        lblEtat.Caption = "Status"
        lblNature.Caption = "Nature Code : "
        lblStart.Caption = "start : "
        lblStop.Caption = "end : "
        lblCritere.Caption = "Criteria : "
        
    Else 'FR
        lblHeader.Caption = "Documents d'un marquage de pièce"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblMark.Caption = "Marquage de pièce: "
        lblComment.Caption = "Commentaires : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "début : "
        lblEStop.Caption = "fin : "
        lblEtat.Caption = "Etat"
        lblNature.Caption = "Nature Code : "
        lblStart.Caption = "début : "
        lblStop.Caption = "fin : "
        lblCritere.Caption = "Critères : "
    End If
    
    'Initial values to combobox
    If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        
        If rstMark.State = adStateOpen Then rstMark.Close
        If rstNature.State = adStateOpen Then rstNature.Close
        If rstEtatdate.State = adStateOpen Then rstEtatdate.Close
        If rstDate.State = adStateOpen Then rstDate.Close
        If rstComment.State = adStateOpen Then rstComment.Close
        If rstCritere.State = adStateOpen Then rstCritere.Close
        
        'Mark
        With rstMark
            sql = "SELECT CODE_REPERE_MARQUAGE FROM REPERES_MARQUAGE order by CODE_REPERE_MARQUAGE  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        Set cmbMarkID.Recordset = rstMark
        Set rstMark = Nothing
        'Nature
        With rstNature
            sql = "SELECT CODE_NATURE_DOCUMENT,LIBELLE_NATURE_DOCUMENT FROM Natures_document ORDER BY CODE_NATURE_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        Set cmbNatureID.Recordset = rstNature
        Set rstNature = Nothing
        
        'Etat Date
        With rstEtatdate
            sql = "SELECT DISTINCT DATE_ETAT_DOCUMENT FROM Documents WHERE DATE_ETAT_DOCUMENT IS NOT NULL  ORDER BY DATE_ETAT_DOCUMENT "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbEtatStart.Recordset = rstEtatdate
        Set cmbEtatEnd.Recordset = rstEtatdate
        Set rstEtatdate = Nothing
        
        cmbEtatEnd.Enabled = False
        
        
        'Origine Date
        With rstDate
            sql = "SELECT DISTINCT DATE_ORIGINE_DOCUMENT FROM Documents ORDER BY DATE_ORIGINE_DOCUMENT "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbStart.Recordset = rstDate
        Set cmbEnd.Recordset = rstDate
        Set rstDate = Nothing
        cmbEnd.Enabled = False
        
        
        'comment
        
         With rstComment
            sql = "SELECT COMMENTAIRES_SUIVI FROM Documents ORDER BY NUMERO_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbComment.Recordset = rstComment
        Set rstComment = Nothing
        
        
       'critere
       With rstCritere
        sql = "SELECT DISTINCT CRITERE_MOT_CLE FROM LIENS order by CRITERE_MOT_CLE "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbCritere.Recordset = rstCritere
        Set rstCritere = Nothing
        
       
       
       
       
       With Me.cmbMarkID
            .Value = .ItemData(.ListIndex + 1)
        End With
        cmbMarkID.SetFocus

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub cmbEtatStart_Change()
On Error GoTo ErrHandler

    If Not (cmbEtatStart.ListIndex = -1) Then
        cmbEtatEnd.Value = ""
        cmbEtatEnd.Enabled = True
    Else
        cmbEtatEnd.Value = ""
        cmbEtatEnd.Enabled = False
    End If
Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmbStart_Change()
On Error GoTo ErrHandler

    If Not (cmbStart.ListIndex = -1) Then
        cmbEnd.Value = ""
        cmbEnd.Enabled = True
    Else
        cmbEnd.Value = ""
        cmbEnd.Enabled = False
    End If
Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
    
End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandler

    Set rst = Nothing
    DoCmd.OpenForm "frm_workRequest", acNormal
    DoCmd.Close acForm, "frm_DocByMark"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Private Sub BuildSql()
On Error GoTo ErrHandler
Dim MarkId As String

    publicSql = ""
    MarkId = cmbMarkID.Value
    
   
    If (NatureID = "") And (critereID = "") Then
                    
        publicSql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,INDICE_MODIFICATION," & _
            "LIBELLE_DOCUMENT,COMMENTAIRES_SUIVI,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT,CODE_REPERE_MARQUAGE " & _
            "FROM DOCUMENTS WHERE CODE_REPERE_MARQUAGE ='" & MarkId & "'" & moreSql & _
            " ORDER BY NUMERO_DOCUMENT"
    Else
        'case inner join both
        If (NatureID <> "") And (critereID <> "") Then
            publicSql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                    "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                    "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT,d.CODE_REPERE_MARQUAGE " & _
                    "FROM DOCUMENTS d " & _
                    "INNER JOIN (Select NUMERO_DOCUMENT from Liens " & _
                    "WHERE CODE_NATURE_DOCUMENT='" & NatureID & "'" & _
                    "And CRITERE_MOT_CLE ='" & critereID & "') doc_num " & _
                    "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                    "WHERE d.CODE_REPERE_MARQUAGE ='" & MarkId & "'" & moreSql & _
                    " ORDER BY d.NUMERO_DOCUMENT"
        Else
        'case inner join by critere
            If (NatureID = "") And (critereID <> "") Then
                publicSql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                        "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                        "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT,d.CODE_REPERE_MARQUAGE " & _
                        "FROM DOCUMENTS d " & _
                        "INNER JOIN (Select NUMERO_DOCUMENT from Liens " & _
                        "WHERE CRITERE_MOT_CLE ='" & critereID & "') doc_num " & _
                        "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                        "WHERE d.CODE_REPERE_MARQUAGE ='" & MarkId & "'" & moreSql & _
                        " ORDER BY d.NUMERO_DOCUMENT"
            Else
            'case inner join by nature
                publicSql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,INDICE_MODIFICATION," & _
                    "LIBELLE_DOCUMENT,COMMENTAIRES_SUIVI,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT,CODE_REPERE_MARQUAGE " & _
                    "FROM DOCUMENTS WHERE CODE_REPERE_MARQUAGE ='" & MarkId & "'" & _
                    " AND CODE_NATURE_DOCUMENT = '" & NatureID & "' " & _
                    "ORDER BY NUMERO_DOCUMENT"
            End If
        End If
    End If


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrHandler
   'Dim MarkId As String
   Dim rptSql As String
   
    Dim rep As Report
      
    If CheckValid() Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rst.State = adStateOpen Then rst.Close

        With rst
        
            Call BuildSql
            
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open publicSql, cn, , , adCmdText
            End With

            'Record Found
            If Not (rst.BOF And rst.EOF) Then
            
                CurrentDb.QueryDefs("qry_DocByMark").sql = publicSql
                
                GB_RPT_TITLE = Nz(cmbMarkID.Value, "-") & ";" & _
                    Nz(cmbNatureID.Value, "-") & ";" & _
                    Nz(cmbEtatStart.Value, "-") & ";" & _
                    Nz(cmbEtatEnd.Value, "-") & ";" & _
                    Nz(cmbStart.Value, "-") & ";" & _
                    Nz(cmbEnd.Value, "-") & ";" & _
                    Nz(cmbComment.Value, "-") & ";" & _
                    Nz(cmbCritere.Value, "-") & ";"
               
                DoCmd.OpenReport "rpt_DocsByMark", acViewPreview
                                
                PreviousForm = ""
                PreviousForm = Me.Name
                
                DoCmd.Close acForm, "frm_DocByMark"
'                Debug.Print ("1")
                
            
            Else
                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
                cmbMarkID.SetFocus
            End If
               
        'Close/Open Form?
        'DoCmd.Close acForm, "frm_DocByMark"
    End If


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
    Call UserForm_Initialize

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub




