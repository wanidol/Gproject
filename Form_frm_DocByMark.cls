VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
Dim CritereId As String
Dim publicSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler

    moreSql = ""
        
    If IsNull(cmbMarkID.Value) Or (cmbMarkID.Value = "") Or (cmbMarkID.ListIndex = -1) Then
        CheckValid = False
        
        If GB_LANG = "EN" Then
            MsgBox "Invalid Mark Code", vbOKOnly
        Else
            MsgBox "Invalide N° de marquage", vbOKOnly
        End If
        
        cmbMarkID.SetFocus
        
        Exit Function
    Else
        If Not IsNull(cmbNatureID.Value) Or Not (cmbNatureID.Value = "") Or Not (cmbNatureID.ListIndex = -1) Then
            'moreSql = " And CODE_NATURE_DOCUMENT = '" & cmbNatureID.Value & "'"
             NatureID = cmbNatureID.Value
       
        End If
        
        If Not IsNull(cmbEtatStart.Value) Or Not (cmbEtatStart.Value = "") Or Not (cmbEtatStart.ListIndex = -1) Then
            moreSql = moreSql & " And DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatStop.Value & "'"
        End If
        
        If Not IsNull(cmbStart.Value) Or Not (cmbStart.Value = "") Or Not (cmbStart.ListIndex = -1) Then
            moreSql = moreSql & " And DATE_ORIGINE_DOCUMENT BETWEEN '" & cmbStart.Value & "' AND '" & cmbStop.Value & "'"
        End If
        
        If Not IsNull(cmbComment.Value) Or Not (cmbComment.Value = "") Or Not (cmbComment.ListIndex = -1) Then
            moreSql = moreSql & " AND COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
        End If
        
        If Not IsNull(cmbCritere.Value) Or Not (cmbCritere.Value = "") Or Not (cmbCritere.ListIndex = -1) Then
            'moreSql = moreSql & " And CRITERE_MOT_CLE = '" & cmbCritere.Value & "'"
             CritereId = cmbCritere.Value
        End If
        
        CheckValid = True
    End If
  
Exit_Sub:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Function

Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
    Dim Sql As String
    Dim typeId As String
    Dim critId As String
    Dim docNum As String

    'Clear Variable
    GB_RPT_TITLE = ""
    moreSql = ""
    NatureID = ""
    CritereId = ""
    
    'Clear all Controls
    Call ClearAll(Me)
    
   If GB_LANG = "EN" Then
        lblHeader.Caption = "List of documents of a part marking"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblComment.Caption = "Comment : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "start : "
        lblEStop.Caption = "end : "
        lblEtat.Caption = "Status : "
        lblNature.Caption = "Nature Code : "
        lblStart.Caption = "start : "
        lblStop.Caption = "end : "
        lblCritere.Caption = "Criteria : "
        
    Else 'FR
        lblHeader.Caption = "Liste des documents d'un marquage de pièce"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblComment.Caption = "Commentaires : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "début : "
        lblEStop.Caption = "fin : "
        lblEtat.Caption = "Etat : "
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
            Sql = "SELECT CODE_REPERE_MARQUAGE FROM REPERES_MARQUAGE order by CODE_REPERE_MARQUAGE  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        Set cmbMarkID.Recordset = rstMark
        
        'Nature
        With rstNature
            Sql = "SELECT CODE_NATURE_DOCUMENT,LIBELLE_NATURE_DOCUMENT FROM Natures_document ORDER BY CODE_NATURE_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        Set cmbNatureID.Recordset = rstNature
        
        'Etat Date
        With rstEtatdate
            Sql = "SELECT DISTINCT DATE_ETAT_DOCUMENT FROM Documents WHERE DATE_ETAT_DOCUMENT IS NOT NULL  ORDER BY DATE_ETAT_DOCUMENT "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbEtatStart.Recordset = rstEtatdate
        Set cmbEtatStop.Recordset = rstEtatdate
        'cmbEtatStop.Enabled = False
        
        
        'Origine Date
        With rstDate
            Sql = "SELECT DISTINCT DATE_ORIGINE_DOCUMENT FROM Documents ORDER BY DATE_ORIGINE_DOCUMENT "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbStart.Recordset = rstDate
        Set cmbStop.Recordset = rstDate
        'cmbStop.Enabled = False
        
        
        'comment
        
         With rstComment
            Sql = "SELECT COMMENTAIRES_SUIVI FROM Documents ORDER BY NUMERO_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbComment.Recordset = rstComment
        
        
       'critere
       With rstCritere
        Sql = "SELECT DISTINCT CRITERE_MOT_CLE FROM LIENS order by CRITERE_MOT_CLE "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbCritere.Recordset = rstCritere
        
       
       
       
       
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

Private Sub cmdExit_Click()
On Error GoTo ErrHandler

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
   
    If (NatureID = "") And (CritereId = "") Then
                    
        publicSql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,INDICE_MODIFICATION," & _
            "LIBELLE_DOCUMENT,COMMENTAIRES_SUIVI,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT,CODE_REPERE_MARQUAGE " & _
            "FROM DOCUMENTS WHERE CODE_REPERE_MARQUAGE ='" & MarkId & "' ORDER BY NUMERO_DOCUMENT"
    Else
        'case inner join both
        If (NatureID <> "") And (CritereId <> "") Then
            publicSql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                    "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                    "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT,d.CODE_REPERE_MARQUAGE " & _
                    "FROM DOCUMENTS d " & _
                    "INNER JOIN (Select NUMERO_DOCUMENT from Liens " & _
                    "WHERE CODE_NATURE_DOCUMENT='" & NatureID & "'" & _
                    "And CRITERE_MOT_CLE ='" & CritereId & "') doc_num " & _
                    "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                    "WHERE d.CODE_REPERE_MARQUAGE ='" & MarkId & "'" & moreSql & _
                    " ORDER BY d.NUMERO_DOCUMENT"
        Else
        'case inner join by critere
            If (NatureID = "") And (CritereId <> "") Then
                publicSql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                        "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                        "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT,d.CODE_REPERE_MARQUAGE " & _
                        "FROM DOCUMENTS d " & _
                        "INNER JOIN (Select NUMERO_DOCUMENT from Liens " & _
                        "WHERE CRITERE_MOT_CLE ='" & CritereId & "') doc_num " & _
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
            
                CurrentDb.QueryDefs("qry_DocByMark").Sql = publicSql
                
                GB_RPT_TITLE = Nz(cmbMarkID.Value, "-") & ";" & _
                    Nz(cmbNatureID.Value, "-") & ";" & _
                    Nz(cmbEtatStart.Value, "-") & ";" & _
                    Nz(cmbEtatStop.Value, "-") & ";" & _
                    Nz(cmbStart.Value, "-") & ";" & _
                    Nz(cmbStop.Value, "-") & ";" & _
                    Nz(cmbComment.Value, "-") & ";" & _
                    Nz(cmbCritere.Value, "-") & ";"
               
                DoCmd.OpenReport "rpt_DocsByMark", acViewPreview
            
            Else
                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
                'Call ClearAll(Me)
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


