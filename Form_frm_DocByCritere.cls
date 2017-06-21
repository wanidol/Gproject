VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByCritere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Db As Database
Dim rst As New ADODB.Recordset
Dim rstType As New ADODB.Recordset
Dim rstCritere As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
Dim rstEtatdate As New ADODB.Recordset
Dim rstDate As New ADODB.Recordset
Dim rstComment As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim publicSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler

    moreSql = ""
    
    If IsNull(cmbTypeID.Value) Or (cmbTypeID.Value = "") Or (cmbTypeID.ListIndex = -1) Then
        CheckValid = False
        
        If GB_LANG = "EN" Then
            MsgBox "Invalid Type Code", vbOKOnly
        Else
            MsgBox "Type Invalide", vbOKOnly
        End If
        
        'MsgBox "Invalid Type Code", vbOKOnly
        Exit Function
    Else
        If IsNull(cmbCritereId.Value) Or (cmbCritereId.Value = "") Or (cmbCritereId.ListIndex = -1) Then
            CheckValid = False
            
            If GB_LANG = "EN" Then
                MsgBox "Invalid Criteria Code", vbOKOnly
            Else
                MsgBox "Invalide Critère", vbOKOnly
            End If
           
            Exit Function
        Else
            If Not IsNull(cmbNatureID.Value) Or Not (cmbNatureID.Value = "") Or Not (cmbNatureID.ListIndex = -1) Then
                moreSql = "WHERE d.CODE_NATURE_DOCUMENT = '" & cmbNatureID.Value & "'"
            Else
                '
            End If
            'nature yes
            If Not IsNull(cmbEtatStart.Value) Or Not (cmbEtatStart.Value = "") Or Not (cmbEtatStart.ListIndex = -1) Then
                If (cmbNatureID.ListIndex = -1) Then
                    moreSql = "WHERE d.DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatStop.Value & "'"
                Else
                    moreSql = moreSql & " And d.DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatStop.Value & "'"
                End If
            End If
            
            If Not IsNull(cmbStart.Value) Or Not (cmbStart.Value = "") Or Not (cmbStart.ListIndex = -1) Then
                If (cmbNatureID.ListIndex = -1) And (cmbEtatStart.ListIndex = -1) Then
                    moreSql = "WHERE  d.DATE_ORIGINE_DOCUMENT BETWEEN '" & cmbStart.Value & "' AND '" & cmbStop.Value & "'"
                Else
                    moreSql = moreSql & " And d.DATE_ORIGINE_DOCUMENT BETWEEN '" & cmbStart.Value & "' AND '" & cmbStop.Value & "'"
                End If
            End If
            
            If Not IsNull(cmbCritereLinked.Value) Or Not (cmbCritereLinked.Value = "") Or Not (cmbCritereLinked.ListIndex = -1) Then
                If (cmbNatureID.ListIndex = -1) And (cmbEtatStart.ListIndex = -1) And (cmbStart.ListIndex = -1) Then
                    moreSql = "WHERE d.COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
                Else
                    moreSql = moreSql & " AND d.COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
                End If
            End If
            
            If Not IsNull(cmbComment.Value) Or Not (cmbComment.Value = "") Or Not (cmbComment.ListIndex = -1) Then
                If (cmbNatureID.ListIndex = -1) And (cmbEtatStart.ListIndex = -1) And (cmbStart.ListIndex = -1) And (cmbCritereLinked.ListIndex = -1) Then
                    moreSql = "WHERE d.COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
                Else
                    moreSql = moreSql & " AND d.COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
                End If
            End If
            CheckValid = True
        End If
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
    Dim i As Integer

    'Clear Variable
    GB_RPT_TITLE = ""
    'Clear all Controls
    Call ClearAll(Me)
    
    If GB_LANG = "EN" Then
        lblHeader.Caption = "Documents of a criterion"
        lblBox1.Caption = "Required information :"
        lblbox2.Caption = "Optional information :"
        lblComment.Caption = "Comment : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "start : "
        lblEStop.Caption = "end : "
        lblEtat.Caption = "Status : "
        lblNature.Caption = "Nature Code : "
        lblStart.Caption = "start : "
        lblStop.Caption = "end : "
        lblCritere.Caption = "Criterion No."
        lblCritereLinked.Caption = "Related Criteria"
        lblTypeCode.Caption = "Type Code"
        
        
    Else 'FR
        lblHeader.Caption = "Documents d'un critère"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblComment.Caption = "Commentaires : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "début : "
        lblEStop.Caption = "fin : "
        lblEtat.Caption = "Etat : "
        lblNature.Caption = "Nature : "
        lblStart.Caption = "début : "
        lblStop.Caption = "fin : "
        lblCritere.Caption = "N° critère :"
        lblCritereLinked.Caption = "Critères liés : "
        lblTypeCode.Caption = "Type : "
        
    End If
    
    
    
    'Initial values to combobox
        If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rstType.State = adStateOpen Then rstNature.Close
        If rstCritere.State = adStateOpen Then rstNature.Close
        If rstNature.State = adStateOpen Then rstNature.Close
        If rstEtatdate.State = adStateOpen Then rstEtatdate.Close
        If rstDate.State = adStateOpen Then rstDate.Close
        If rstComment.State = adStateOpen Then rstComment.Close
        
        'Type
        With rstType
            Sql = "SELECT DISTINCT  CODE_TYPE_CRITERE FROM liens "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbTypeID.Recordset = rstType
        
        'Critere
        
        With rstCritere
           Sql = "SELECT DISTINCT CRITERE_MOT_CLE FROM LIENS order by CRITERE_MOT_CLE "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbCritereId.Recordset = rstCritere
        
        'NatureID
        With rstNature
            Sql = "SELECT CODE_NATURE_DOCUMENT,LIBELLE_NATURE_DOCUMENT FROM Natures_document ORDER BY CODE_NATURE_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbNatureID.Recordset = rstNature
'        i = 0
'        Do While Not rstNature.EOF
'            With rstNature
'                cmbNatureID.AddItem (.Fields("CODE_NATURE_DOCUMENT").Value)
''                cmbNatureID.List(i, 0) = .Fields("CODE_NATURE_DOCUMENT").Value
''                cmbNatureID.Column(i, 1) = .Fields("LIBELLE_NATURE_DOCUMENT").Value
'               ' cmbNatureID.AddItem (.Fields("CODE_NATURE_DOCUMENT").Value & " : " & .Fields("LIBELLE_NATURE_DOCUMENT").Value)
'                i = i + 1
'                rstNature.MoveNext
'            End With
'
'        Loop
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
        
        
        'set focus on Nature Code default first listindex
        With Me.cmbTypeID
            .Value = .ItemData(.ListIndex + 1)
        End With
        cmbTypeID.SetFocus
            
       

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub
Private Sub BuildSql()
On Error GoTo ErrHandler
    Dim typeId As String
    Dim CritereId As String
    
    publicSql = ""
   
    typeId = cmbTypeID.Value
    CritereId = cmbCritereId.Value
    
    publicSql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                            "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                            "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT,d.CODE_REPERE_MARQUAGE " & _
                            "FROM DOCUMENTS d " & _
                            "INNER JOIN (Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='" & typeId & "' And CRITERE_MOT_CLE ='" & CritereId & "') doc_num " & _
                            "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT ORDER BY d.NUMERO_DOCUMENT "
                            


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Private Sub cmdPrint_Click()
On Error GoTo ErrHandler
   
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
                CurrentDb.QueryDefs("qry_DocCritere").Sql = publicSql
                GB_RPT_TITLE = Nz(cmbTypeID.Value, "-") & ";" & _
                    Nz(cmbCritereId.Value, "-") & ";" & _
                    Nz(cmbNatureID.Value, "-") & ";" & _
                    Nz(cmbEtatStart.Value, "-") & ";" & _
                    Nz(cmbEtatStop.Value, "-") & ";" & _
                    Nz(cmbStart.Value, "-") & ";" & _
                    Nz(cmbStop.Value, "-") & ";" & _
                    Nz(cmbCritereLinked.Value, "-") & ";" & _
                    Nz(cmbComment.Value, "-") & ";"
    
                DoCmd.OpenReport "rpt_DocsByCritere", acViewPreview
    
            Else

                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
               'Call ClearAll(Me)
               cmbTypeID.SetFocus

            End If

        'Close/Open Form?
        'DoCmd.Close acForm, "frm_DocByNature"
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

    DoCmd.OpenForm "frm_workRequest", acNormal
    DoCmd.Close acForm, "frm_DocByCritere"


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

