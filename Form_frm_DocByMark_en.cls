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
Dim rst As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler
    
    If IsNull(cmbMarkID.Value) Or (cmbMarkID.Value = "") Or (cmbMarkID.ListIndex = -1) Then
        CheckValid = False
        MsgBox "Please Choose Mark Code", vbOKOnly
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
        If rst.State = adStateOpen Then rst.Close
        If rs1.State = adStateOpen Then rs1.Close
        With rst
            Sql = "SELECT CODE_REPERE_MARQUAGE FROM REPERES_MARQUAGE order by CODE_REPERE_MARQUAGE  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        Set cmbMarkID.Recordset = rst
        
        With rs1
            Sql = "SELECT CODE_NATURE_DOCUMENT,LIBELLE_NATURE_DOCUMENT FROM Natures_document "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        
        Do While Not rs1.EOF
            With rs1
                cmbNatureID.AddItem (.Fields("CODE_NATURE_DOCUMENT").Value & " : " & .Fields("LIBELLE_NATURE_DOCUMENT").Value)
                rs1.MoveNext
            End With

        Loop
        cmbMarkID.Selected(0) = True
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
    DoCmd.Close acForm, "frm_DocByMark_en"


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


