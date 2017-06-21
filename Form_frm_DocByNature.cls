VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByNature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Db As Database
Dim rstNature As New ADODB.Recordset
Dim rstEtatdate As New ADODB.Recordset
Dim rstDate As New ADODB.Recordset
Dim rstComment As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim publicSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler

    moreSql = ""
    
    If IsNull(cmbNatureID.Value) Or (cmbNatureID.Value = "") Or (cmbNatureID.ListIndex = -1) Then
        CheckValid = False
        
        If GB_LANG = "EN" Then
             MsgBox "Invalid Nature Code", vbOKOnly
        Else
            MsgBox "Invalide Nature", vbOKOnly
        End If
        
        cmbNatureID.SetFocus
        Exit Function
    Else
    If Not IsNull(cmbEtatStart.Value) Or Not (cmbEtatStart.Value = "") Or Not (cmbEtatStart.ListIndex = -1) Then
        'If Not (cmbEtatStart.Value = "") And Not (cmbEtatStop = "") Then
            moreSql = " And DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatStop.Value & "'"
        End If
        If Not IsNull(cmbStart.Value) Or Not (cmbStart.Value = "") Or Not (cmbStart.ListIndex = -1) Then
        'If Not (cmbStart.Value = "") And Not (cmbStop.Value = "") Then
            moreSql = moreSql & " And DATE_ORIGINE_DOCUMENT BETWEEN '" & cmbStart.Value & "' AND '" & cmbStop.Value & "'"
        End If
        If Not IsNull(cmbComment.Value) Or Not (cmbComment.Value = "") Or Not (cmbComment.ListIndex = -1) Then
        'If Not (cmbComment.Value = "") Then
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
        lblHeader.Caption = "Documents of a Nature"
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
        
        
    Else 'FR
        lblHeader.Caption = "Documents d'une nature"
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
    End If
    
    
    
    'Initial values to combobox
        If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rstNature.State = adStateOpen Then rstNature.Close
        If rstEtatdate.State = adStateOpen Then rstEtatdate.Close
        If rstDate.State = adStateOpen Then rstDate.Close
        If rstComment.State = adStateOpen Then rstComment.Close
        
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
        With Me.cmbNatureID
            .Value = .ItemData(.ListIndex + 1)
        End With
        cmbNatureID.SetFocus
            
       

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
    DoCmd.Close acForm, "frm_DocByNature"

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Private Sub BuildSql()
On Error GoTo ErrHandler

    Dim NatureID As String

    NatureID = cmbNatureID.Value
    publicSql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,INDICE_MODIFICATION," & _
                            "LIBELLE_DOCUMENT,COMMENTAIRES_SUIVI,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT " & _
                            "FROM DOCUMENTS WHERE CODE_NATURE_DOCUMENT ='" & NatureID & "'" & moreSql & _
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
    
    Dim rep As Report
     Dim NatureID As String
    
      
    If CheckValid() Then
        NatureID = cmbNatureID.Value
       
        
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
            
                CurrentDb.QueryDefs("qry_DocByNature").Sql = publicSql
                GB_RPT_TITLE = Nz(cmbNatureID.Value, "-") & ";" & _
                    Nz(cmbEtatStart.Value, "-") & ";" & _
                    Nz(cmbEtatStop.Value, "-") & ";" & _
                    Nz(cmbStart.Value, "-") & ";" & _
                    Nz(cmbStop.Value, "-") & ";" & _
                    Nz(cmbComment.Value, "-") & ";"
    
                DoCmd.OpenReport "rpt_DocsByNature", acViewPreview
    
            Else

                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly

                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
                'Call ClearAll(Me)
                cmbNatureID.SetFocus
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
