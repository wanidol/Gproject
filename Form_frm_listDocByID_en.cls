VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_listDocByID_en"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Dim Db As Database
Dim rst As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler

'data test
'1. nature or etat
    '1.1 nature no etat ok
    '1.2 nature ok etat no
'2. nature ok and etat ok
'3. no nature no etat
    
    
    If IsNull(cmbDocId.Value) Or (cmbDocId.Value = "") Then
        CheckValid = False
        MsgBox "Please Choose A Document Number", vbOKOnly
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
DocumentKey = ""
lstDocument.Enabled = False




If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rs1.State = adStateOpen Then rs1.Close
        
        With rs1
            Sql = "SELECT NUMERO_DOCUMENT FROM DOCUMENTS "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        Set cmbDocId.Recordset = rs1
        cmbDocId.Selected(0) = True
        

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandler
DoCmd.OpenForm "frm_consultation_en", acNormal
DoCmd.Close acForm, "frm_listDocByID_en"

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdFind_Click()
On Error GoTo ErrHandler
    Dim Sql As String
    Dim docNum As String
    Dim i As Integer

    If CheckValid() Then
        docNum = cmbDocId.Value
'        natureId = txtCodeNature.Value
'        etatId = txtEtatCode.Value
    
      
        If cn.State = adStateOpen Then cn.Close
            Set cn = CurrentProject.AccessConnection
            If rst.State = adStateOpen Then rst.Close
            With rst
            
            'data test 3 rows => select * from documents where code_repere_marquage='ABQQ' AND code_nature_document='PQP' and code_etat_document = 'FM';
            'only nature=>SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,CODE_ETAT_DOCUMENT,INDICE_MODIFICATION,DATE_ETAT_DOCUMENT,LIBELLE_DOCUMENT,DATE_ORIGINE_DOCUMENT FROM DOCUMENTS Where code_repere_marquage ='ABQQ' And CODE_NATURE_DOCUMENT = 'pqp'
            'SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,CODE_ETAT_DOCUMENT,INDICE_MODIFICATION,DATE_ETAT_DOCUMENT,LIBELLE_DOCUMENT,DATE_ORIGINE_DOCUMENT FROM DOCUMENTS Where code_repere_marquage ='ABQQ' And CODE_NATURE_DOCUMENT = 'PQP' ORDER BY NUMERO_DOCUMENT
            'SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,CODE_ETAT_DOCUMENT,INDICE_MODIFICATION,DATE_ETAT_DOCUMENT,LIBELLE_DOCUMENT,DATE_ORIGINE_DOCUMENT FROM DOCUMENTS Where code_repere_marquage ='ABQQ' And CODE_NATURE_DOCUMENT = 'PQP' And CODE_ETAT_DOCUMENT = 'FM' ORDER BY NUMERO_DOCUMENT
             Sql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT," & _
                    "CODE_ETAT_DOCUMENT,INDICE_MODIFICATION," & _
                    "DATE_ETAT_DOCUMENT,LIBELLE_DOCUMENT,DATE_ORIGINE_DOCUMENT " & _
                    "FROM DOCUMENTS WHERE NUMERO_DOCUMENT ='" & docNum & "' ORDER BY NUMERO_DOCUMENT"
                .CursorType = adOpenDynamic
                .CursorLocation = adUseClient
                .LockType = adLockOptimistic
                .Open Sql, cn, , , adCmdText
            End With
            
            'Record Found
            If Not (rst.BOF And rst.EOF) Then
                Set lstDocument.Recordset = rst
                lstDocument.Selected(0) = True
                lstDocument.Enabled = True
    
            Else
                MsgBox "Record Not Found!", vbOKOnly
                lstDocument.Enabled = False
 
            End If
      
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


Private Sub lstDocument_DblClick(Cancel As Integer)
On Error GoTo ErrHandler
Dim passData As String

'passing document number for retriving data in frm_DocsManagement
passData = lstDocument.Column(0)


'for back to previous page

PreviousForm = ""

PreviousForm = Me.Name

DoCmd.OpenForm "frm_DocsManagement_en", acNormal, , , acFormReadOnly, , passData

'Hiding form
Me.Visible = False


Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub


