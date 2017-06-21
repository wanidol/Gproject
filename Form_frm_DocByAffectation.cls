VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByAffectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim Db As Database
Dim rstAffect As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
Dim rstEtatdate As New ADODB.Recordset
Dim rstClient As New ADODB.Recordset
Dim rstComment As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim publicSql As String

Private Function CreateReport() As Boolean

Dim coll As Collection


' Declare the Variable
    Dim oDocs As clsDocuments
    
' Create new Document
    Set oDocs = New clsDocuments
    Set GB_cDocs = New Collection
    
    CreateReport = False

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
    
    If Not (rst.BOF And rst.EOF) Then
        CreateReport = True
        
        Do While Not (rst.EOF)
             With rst.Fields
              oDocs.DocID = Nz(.Item("NUMERO_DOCUMENT").Value, "")
              oDocs.NatureID = Nz(.Item("CODE_NATURE_DOCUMENT").Value, "")
              oDocs.Indice = Nz(.Item("INDICE_MODIFICATION").Value, "")
              oDocs.Libelle = Nz(.Item("LIBELLE_DOCUMENT").Value, "")
              oDocs.Etat = Nz(.Item("DATE_ETAT_DOCUMENT").Value, "")
              oDocs.Origin = Nz(.Item("DATE_ORIGINE_DOCUMENT").Value, "")
              oDocs.Commentaire = Nz(.Item("COMMENTAIRES_SUIVI").Value, "")
              
              GB_cDocs.Add oDocs
            End With
            rst.MoveNext
        Loop
        
        GB_RPT_TITLE = Nz(cmbAffectID.Value) & ";" & _
                            Nz(cmbNatureID.Value, "-") & ";" & _
                            Nz(cmbEtatStart.Value, "-") & ";" & _
                            Nz(cmbEtatStop.Value, "-") & ";" & _
                            Nz(cmbClientLinked.Value, "-") & ";" & _
                            Nz(cmbComment.Value, "-") & ";"
'
'    Else
'        CreateReport = False
        
    End If


End Function

Function CheckValid() As Boolean
On Error GoTo ErrHandler

    moreSql = ""
    
    If IsNull(cmbAffectID.Value) Or (cmbAffectID.Value = "") Or (cmbAffectID.ListIndex = -1) Then
        CheckValid = False
        
        If GB_LANG = "EN" Then
            MsgBox "Invalid Affectation Code", vbOKOnly
        Else
            MsgBox "Invalide Affectation", vbOKOnly
        End If
        
        'MsgBox "Invalid Affectation Code", vbOKOnly
        Exit Function
    Else
        If Not IsNull(cmbNatureID.Value) Or Not (cmbNatureID.Value = "") Or Not (cmbNatureID.ListIndex = -1) Then
            moreSql = moreSql & " And CODE_NATURE_DOCUMENT = '" & cmbNatureID.Value & "'"
        End If
        If Not IsNull(cmbEtatStart.Value) Or Not (cmbEtatStart.Value = "") Or Not (cmbEtatStart.ListIndex = -1) Then
            moreSql = moreSql & " And DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatStop.Value & "'"
        End If
        
        If Not IsNull(cmbClientLinked.Value) Or Not (cmbClientLinked.Value = "") Or Not (cmbClientLinked.ListIndex = -1) Then
             moreSql = moreSql & " AND CONTRACTUEL = '" & cmbClientLinked.Value & "'"
        End If
        
        If Not IsNull(cmbComment.Value) Or Not (cmbComment.Value = "") Or Not (cmbComment.ListIndex = -1) Then
            moreSql = moreSql & " AND COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
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
    Dim i As Integer

    'Clear Variable
    GB_RPT_TITLE = ""
    'Clear all Controls
    Call ClearAll(Me)
    
    If GB_LANG = "EN" Then
        lblHeader.Caption = "Documents of an assignment"
        lblBox1.Caption = "Required information :"
        lblbox2.Caption = "Optional information :"
        lblComment.Caption = "Comment : "
        lblEStart.Caption = "start : "
        lblEStop.Caption = "end : "
        lblEtat.Caption = "Status : "
        lblNature.Caption = "Nature Code : "
        
        lblAffectation.Caption = "Assignment : "
        lblClientLinked.Caption = "Related to a client : "
        
        
        
    Else 'FR
        lblHeader.Caption = "Documents d'une affectation"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblComment.Caption = "Commentaires : "
       
        lblEStart.Caption = "début : "
        lblEStop.Caption = "fin : "
        lblEtat.Caption = "Etat : "
        lblNature.Caption = "Nature : "
        
        lblAffectation.Caption = "Affectation : "
        lblClientLinked.Caption = "Liés à un client : "
        
    End If
    
    'Initial values to combobox
        If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rstAffect.State = adStateOpen Then rstAffect.Close
        If rstNature.State = adStateOpen Then rstNature.Close
        If rstEtatdate.State = adStateOpen Then rstEtatdate.Close
        If rstClient.State = adStateOpen Then rstClient.Close
        If rstComment.State = adStateOpen Then rstComment.Close
        
        
        'Affectation
        With rstAffect
             Sql = "SELECT CODE_AFFECTATION FROM AFFECTATION ORDER BY CODE_AFFECTATION  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbAffectID.Recordset = rstAffect
        
        'NatureID
        With rstNature
            Sql = "SELECT CODE_NATURE_DOCUMENT FROM Natures_document ORDER BY CODE_NATURE_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
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
        
        'Linked to a client
        With rstClient
         Sql = "SELECT DISTINCT CONTRACTUEL FROM DOCUMENTS"
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbClientLinked.Recordset = rstClient
        
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
        With Me.cmbAffectID
            .Value = .ItemData(.ListIndex + 1)
        End With
        cmbAffectID.SetFocus
            
       

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
    DoCmd.Close acForm, "frm_DocByAffectation"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Private Sub BuildSql()
On Error GoTo ErrHandler
    Dim AffactID As String
    
    publicSql = ""
   
    AffactID = cmbAffectID.Value
    
    publicSql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,INDICE_MODIFICATION," & _
                            "LIBELLE_DOCUMENT,COMMENTAIRES_SUIVI,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT,CODE_AFFECTATION " & _
                            "FROM DOCUMENTS WHERE CODE_AFFECTATION ='" & AffactID & "'" & moreSql & _
                            "ORDER BY NUMERO_DOCUMENT"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrHandler

    Dim oDocument As New clsDocuments
      
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
                CurrentDb.QueryDefs("qry_DocByAffectation").Sql = publicSql
                GB_RPT_TITLE = Nz(cmbAffectID.Value) & ";" & _
                    Nz(cmbNatureID.Value, "-") & ";" & _
                    Nz(cmbEtatStart.Value, "-") & ";" & _
                    Nz(cmbEtatStop.Value, "-") & ";" & _
                    Nz(cmbClientLinked.Value, "-") & ";" & _
                    Nz(cmbComment.Value, "-") & ";"
        
                DoCmd.OpenReport "rpt_DocsByAffectation", acViewPreview
            Else
                'DoCmd.Close acReport, "rpt_DocsByAffectation", acSaveYes
                
                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
                'Call ClearAll(Me)
                cmbAffectID.SetFocus

            End If

        'Close/Open Form?
        'DoCmd.Close acForm, "frm_DocByAffectation"
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

