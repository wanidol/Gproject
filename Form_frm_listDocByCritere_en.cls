VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_listDocByCritere_en"
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
    
    
    If IsNull(cmbTypeID.Value) Or (cmbTypeID.Value = "") Then
        CheckValid = False
        MsgBox "Please Choose A Type Code", vbOKOnly
        Exit Function
    Else
        If IsNull(cmbCritereId.Value) Or (cmbCritereId.Value = "") Then
            CheckValid = False
            MsgBox "Please Choose A Criteria Code", vbOKOnly
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
                    moreSql = " d.CODE_ETAT_DOCUMENT = '" & Trim(txtEtatCode.Value) & "'"
'                    Debug.Print moreSql
                    CheckValid = True
                End If
            Else
                'case 1.2 nature ok etat no
                If IsNull(txtEtatCode.Value) Or (txtEtatCode.Value = "") Then
                    moreSql = " d.CODE_NATURE_DOCUMENT = '" & Trim(txtCodeNature.Value) & "'"
'                    Debug.Print moreSql
                Else
                'case 2. nature ok etat ok
                    moreSql = " d.CODE_NATURE_DOCUMENT = '" & Trim(txtCodeNature.Value) & "'" & _
                                " And d.CODE_ETAT_DOCUMENT = '" & Trim(txtEtatCode.Value) & "'"
'                    Debug.Print moreSql
                End If
                CheckValid = True
            End If
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

'Initial values to combobox
If cn.State = adStateOpen Then cn.Close
    Set cn = CurrentProject.AccessConnection
    If rstCritere.State = adStateOpen Then rstCritere.Close
    If rstType.State = adStateOpen Then rstType.Close
    
    With rstType
        Sql = "SELECT DISTINCT  CODE_TYPE_CRITERE FROM liens "
    
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    
    With rstCritere
        'Sql = "SELECT DISTINCT CRITERE_MOT_CLE FROM liens "
        Sql = "SELECT FROM MOTS_CLE_CRITERE order by CRITERE_MOT_CLE "

        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With

        Set cmbTypeID.Recordset = rstType
        Set cmbCritereId.Recordset = rstCritere
    
        
        With Me.cmbTypeID
            .Value = .ItemData(.ListIndex + 1)
        End With
        'cmbTypeID.Selected(0) = True
        With Me.cmbCritereId
            .Value = .ItemData(.ListIndex + 1)
        End With
        'cmbCritereId.Selected(0) = True

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
DoCmd.Close acForm, "frm_listDocByCritere_en"

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
    Dim typeId As String
    Dim critId As String
    Dim docNum As String
    Dim NatureId As String
    Dim etatId As String
    Dim i As Integer


    If CheckValid() Then
        typeId = cmbTypeID.Value
        critId = cmbCritereId.Value
'        natureId = txtCodeNature.Value
'        etatId = txtEtatCode.Value
    
      
        If cn.State = adStateOpen Then cn.Close
            Set cn = CurrentProject.AccessConnection
            If rst.State = adStateOpen Then rst.Close
            With rst
            'select * from documents inner join (Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='MIC'and CRITERE_MOT_CLE ='17010') doc_num  On documents.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT ;
            'code_type_critere='MIC' and critere_mot_cle='21437';
            'SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,CODE_ETAT_DOCUMENT,INDICE_MODIFICATION,DATE_ETAT_DOCUMENT,LIBELLE_DOCUMENT,DATE_ORIGINE_DOCUMENT FROM DOCUMENTS WHERE NUMERO_DOCUMENT IN (Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='PQP'and CRITERE_MOT_CLE ='PQPF19972330') And CODE_NATURE_DOCUMENT = 'RLM' ORDER BY NUMERO_DOCUMENT
        
            'select d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT,d.CODE_ETAT_DOCUMENT,d.INDICE_MODIFICATION,d.DATE_ETAT_DOCUMENT,d.LIBELLE_DOCUMENT,d.DATE_ORIGINE_DOCUMENT
            'from documents d  Inner join (Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='PQP'and CRITERE_MOT_CLE ='PQRF19972330') doc_num  On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT ;
            Sql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                    "d.CODE_ETAT_DOCUMENT,d.INDICE_MODIFICATION," & _
                    "d.DATE_ETAT_DOCUMENT,d.LIBELLE_DOCUMENT,d.DATE_ORIGINE_DOCUMENT " & _
                    "FROM DOCUMENTS d Inner join (Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='" & typeId & "' And CRITERE_MOT_CLE ='" & critId & "') doc_num " & _
                    "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                    "WHERE " & moreSql & ""
'                Sql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT," & _
'                    "CODE_ETAT_DOCUMENT,INDICE_MODIFICATION," & _
'                    "DATE_ETAT_DOCUMENT,LIBELLE_DOCUMENT,DATE_ORIGINE_DOCUMENT " & _
'                    "FROM DOCUMENTS WHERE NUMERO_DOCUMENT IN " & _
'                    "(Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='" & typeId & "'" & _
'                     "and CRITERE_MOT_CLE ='" & critId & "')" & moreSql & " ORDER BY NUMERO_DOCUMENT"
'                    "and CRITERE_MOT_CLE ='" & critId & "') And CODE_NATURE_DOCUMENT ='PQP' and CODE_ETAT_DOCUMENT='FM' ORDER BY NUMERO_DOCUMENT"

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
