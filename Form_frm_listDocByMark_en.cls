VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_listDocByMark_en"
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
    
    
    If IsNull(cmbMarkID.Value) Or (cmbMarkID.Value = "") Then
        CheckValid = False
        MsgBox "Please Choose A Mark Code", vbOKOnly
        Exit Function
    Else
        
        If IsNull(txtCodeNature.Value) Or (txtCodeNature.Value = "") Then
            If IsNull(txtEtatCode.Value) Or (txtEtatCode.Value = "") Then
            'case 3. no nature no etat
               CheckValid = False
                MsgBox "Please Input Nature Code And/Or Etat Code", vbOKOnly
                Exit Function
            Else
                'case 1.1 nature no etat ok
                moreSql = " And CODE_ETAT_DOCUMENT = '" & Trim(txtEtatCode.Value) & "'"
                Debug.Print moreSql
                CheckValid = True
            End If
        Else
            'case 1.2 nature ok etat no
            If IsNull(txtEtatCode.Value) Or (txtEtatCode.Value = "") Then
                moreSql = " And CODE_NATURE_DOCUMENT = '" & Trim(txtCodeNature.Value) & "'"
                Debug.Print moreSql
            Else
            'case 2. nature ok etat ok
                moreSql = " And CODE_NATURE_DOCUMENT = '" & Trim(txtCodeNature.Value) & "'" & _
                            " And CODE_ETAT_DOCUMENT = '" & Trim(txtEtatCode.Value) & "'"
                Debug.Print moreSql
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

'Clear all Controls

    Call ClearAll(Me)
    
    DocumentKey = ""
    lstDocument.Enabled = False
    
    
    
    
    If cn.State = adStateOpen Then cn.Close
    Set cn = CurrentProject.AccessConnection
    If rs1.State = adStateOpen Then rs1.Close
    
    With rs1
        Sql = "SELECT CODE_REPERE_MARQUAGE FROM REPERES_MARQUAGE "
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    Set cmbMarkID.Recordset = rs1
    cmbMarkID.Selected(0) = True
        

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
DoCmd.Close acForm, "frm_listDocByMark_en"

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
    Dim markId As String
    Dim docNum As String
    Dim natureId As String
    Dim etatId As String
    Dim i As Integer

    If CheckValid() Then
        markId = cmbMarkID.Value
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
                    "FROM DOCUMENTS Where code_repere_marquage ='" & markId & "'" & moreSql & " ORDER BY NUMERO_DOCUMENT"
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

PreviousForm = ""
'for back to previous page

'Me.Visible = False
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


