VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByMark_en"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Db As Database
Dim rstMark As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
Dim rstEtatdate As New ADODB.Recordset
Dim rstDate As New ADODB.Recordset
Dim rstComment As New ADODB.Recordset
Dim rstCritere As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim NatureId As String
Dim CritereId As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler

        
    If IsNull(cmbMarkID.Value) Or (cmbMarkID.Value = "") Or (cmbMarkID.ListIndex = -1) Then
        CheckValid = False
        MsgBox "Please Choose Mark Code", vbOKOnly
        Exit Function
    Else
        If Not IsNull(cmbNatureID.Value) Or Not (cmbNatureID.Value = "") Or Not (cmbNatureID.ListIndex = -1) Then
            'moreSql = " And CODE_NATURE_DOCUMENT = '" & cmbNatureID.Value & "'"
             NatureId = cmbNatureID.Value
       
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
    RPT_TITLE = ""
    moreSql = ""
    NatureId = ""
    CritereId = ""
    
    'Clear all Controls
    Call ClearAll(Me)
    
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
        Sql = "SELECT CRITERE_MOT_CLE FROM MOTS_CLE_CRITERE order by CRITERE_MOT_CLE "
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

    DoCmd.OpenForm "frm_workRequest_en", acNormal
    DoCmd.Close acForm, "frm_DocByMark_en"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Private Sub cmdPrint_Click()
On Error GoTo ErrHandler
    Dim MarkId As String
    Dim Sql As String
    Dim rep As Report
      
    If CheckValid() Then
        Sql = ""
        MarkId = cmbMarkID.Value
   
        DoCmd.OpenReport "rpt_DocsByMark_en", acViewDesign
        
        Set rep = Reports!rpt_DocsByMark_en
        
        'query by only mark code
        If (NatureId = "") And (CritereId = "") Then
            Sql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,INDICE_MODIFICATION," & _
                        "LIBELLE_DOCUMENT,COMMENTAIRES_SUIVI,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT " & _
                        "FROM DOCUMENTS WHERE CODE_REPERE_MARQUAGE ='" & MarkId & "' ORDER BY NUMERO_DOCUMENT"
        Else
        'case inner join both
            If (NatureId <> "") And (CritereId <> "") Then
                Sql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                            "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                            "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT " & _
                            "FROM DOCUMENTS d " & _
                            "INNER JOIN (Select NUMERO_DOCUMENT from Liens" & _
                            "WHERE CODE_NATURE_DOCUMENT='" & NatureId & "'" & _
                            "And CRITERE_MOT_CLE ='" & CritereId & "') doc_num " & _
                            "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                            "WHERE CODE_REPERE_MARQUAGE ='" & MarkId & "'" & moreSql & _
                            " ORDER BY NUMERO_DOCUMENT"
            Else
            'case inner join by Nature
                If (NatureId = "") And (CritereId <> "") Then
                    Sql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                            "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                            "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT " & _
                            "FROM DOCUMENTS d " & _
                            "INNER JOIN (Select NUMERO_DOCUMENT from Liens " & _
                            "WHERE CRITERE_MOT_CLE ='" & CritereId & "') doc_num " & _
                            "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                            "WHERE CODE_REPERE_MARQUAGE ='" & MarkId & "'" & moreSql & _
                            " ORDER BY NUMERO_DOCUMENT"
                Else
                'case inner join by Critere
                    Sql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                            "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                            "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT " & _
                            "FROM DOCUMENTS d " & _
                            "INNER JOIN (Select NUMERO_DOCUMENT from Liens where CODE_NATURE_DOCUMENT='" & NatureId & "') doc_num " & _
                            "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                            "WHERE CODE_REPERE_MARQUAGE ='" & MarkId & "'" & moreSql & _
                            " ORDER BY NUMERO_DOCUMENT"
                End If
            End If
            
        End If
        
        rep.RecordSource = Sql
                            
        RPT_TITLE = Nz(cmbMarkID.Value, "-") & ";" & Nz(cmbNatureID.Value, "-") & ";" & Nz(cmbEtatStart.Value, "-") & ";" & Nz(cmbEtatStop.Value, "-") & ";" & _
        Nz(cmbStart.Value, "-") & ";" & Nz(cmbStop.Value, "-") & ";" & Nz(cmbComment.Value, "-") & ";" & Nz(cmbCritere.Value, "-") & ";"
    
        DoCmd.OpenReport "rpt_DocsByMark_en", acViewPreview
        
        'Close/Open Form?
        'DoCmd.Close acForm, "frm_DocByNature_en"
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


