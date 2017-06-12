VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByNature_en"
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

Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler
    
    If IsNull(cmbNatureID.Value) Or (cmbNatureID.Value = "") Or (cmbNatureID.ListIndex = -1) Then
        CheckValid = False
        MsgBox "Please Choose Nature Code", vbOKOnly
        Exit Function
    Else
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

    'Clear all Controls
    Call ClearAll(Me)
    'Initial values to combobox
        If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rstNature.State = adStateOpen Then rstNature.Close
        If rstEtatdate.State = adStateOpen Then rstEtatdate.Close
        If rstDate.State = adStateOpen Then rstDate.Close
        If rstComment.State = adStateOpen Then rstComment.Close
        
        'NatureID
        With rstNature
            Sql = "SELECT CODE_NATURE_DOCUMENT,LIBELLE_NATURE_DOCUMENT FROM Natures_document "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Do While Not rstNature.EOF
            With rstNature
                cmbNatureID.AddItem (.Fields("CODE_NATURE_DOCUMENT").Value & " : " & .Fields("LIBELLE_NATURE_DOCUMENT").Value)
                rstNature.MoveNext
            End With
            
        Loop
       'Etat Date
        With rstEtatdate
            Sql = "SELECT DATE_ETAT_DOCUMENT FROM Documents "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbEtatStart.Recordset = rstEtatdate
        Set cmbEtatStop.Recordset = rstEtatdate
        
        'Origine Date
        With rstDate
            Sql = "SELECT DATE_ORIGINE_DOCUMENT FROM Documents "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbStart.Recordset = rstDate
        Set cmbStop.Recordset = rstDate
        
        'comment
        
         With rstComment
            Sql = "SELECT COMMENTAIRES_SUIVI FROM Documents "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Set cmbComment.Recordset = rstComment
        
        
        
        cmbNatureID.Selected(0) = True
         
       
            
       

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
    DoCmd.Close acForm, "frm_DocByNature_en"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrHandler
    Dim Sql As String
    Dim markId As String
    Dim docNum As String
    Dim natureId As String
    Dim etatId As String
    Dim i As Integer

    If CheckValid() Then
        natureId = cmbNatureID.Value
        'check date start stop


    DoCmd.OpenReport "rptReportName", acViewPreview, , , acWindowNormal, Me.RecordSource
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
