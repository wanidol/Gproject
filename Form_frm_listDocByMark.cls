VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_listDocByMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Compare Database
Option Explicit
Dim Db As Database
Dim rst As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
Dim rstEtat As New ADODB.Recordset
Dim rstMark As New ADODB.Recordset
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
    
    
    If IsNull(cmbMarkID.Value) Or (cmbMarkID.Value = "") Or (cmbMarkID.ListIndex = -1) Then
        CheckValid = False
        
        If GB_LANG = "FR" Then
            MsgBox "N° de Marguage Invalide", vbOKOnly
        Else
            MsgBox "Invalid Mark Code", vbOKOnly
        End If
        
        cmbMarkID.SetFocus
        Exit Function
    Else
        
        If IsNull(cmbNatureID.Value) Or (cmbNatureID.Value = "") Or (cmbNatureID.ListIndex = -1) Then
            If IsNull(cmbEtatID.Value) Or (cmbEtatID.Value = "") Or (cmbEtatID.ListIndex = -1) Then
            'case 3. no nature no etat
                CheckValid = False
               
                If GB_LANG = "FR" Then
                    MsgBox "Saisissez Le Code De La Nature Et/Ou Le Code De L'Etat", vbOKOnly
                Else
                    MsgBox "Please Input Nature Code And/Or Etat Code", vbOKOnly
                End If
                
                If (cmbNatureID.ListIndex = -1) Then
                    cmbNatureID.SetFocus
                Else
                    cmbEtatID.SetFocus
                End If
                
                Exit Function
            Else
                'case 1.1 nature no etat ok
                moreSql = " And CODE_ETAT_DOCUMENT = '" & Trim(cmbEtatID.Value) & "'"
'                Debug.Print moreSql
                CheckValid = True
            End If
        Else
            'case 1.2 nature ok etat no
            If IsNull(cmbEtatID.Value) Or (cmbEtatID.Value = "") Or (cmbEtatID.ListIndex = -1) Then
                moreSql = " And CODE_NATURE_DOCUMENT = '" & Trim(cmbNatureID.Value) & "'"
                'Debug.Print moreSql
            Else
            'case 2. nature ok etat ok
                moreSql = " And CODE_NATURE_DOCUMENT = '" & Trim(cmbNatureID.Value) & "'" & _
                            " And CODE_ETAT_DOCUMENT = '" & Trim(cmbEtatID.Value) & "'"
                'Debug.Print moreSql
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
    
    If GB_LANG = "EN" Then
        lblHeader.Caption = "List of Marking Documents"
        lblNature.Caption = "Nature Code : "
        lblMark.Caption = "Mark Number : "
        lblEtat.Caption = "Status :"
        'list
        lblListDoc.Caption = "Document Number"
        lblListNature.Caption = "Nature"
        lblListEtat.Caption = "Status"
        lblListInd.Caption = "Index"
        lblListEtatDate.Caption = "Status Date"
        lblListDesc.Caption = "Description"
        lblListOriginDate.Caption = "Original Date"
        
    Else 'FR
        lblHeader.Caption = "Liste des documents liés à une pièce de marquage"
        lblNature.Caption = "Nature : "
        lblMark.Caption = "N° de marquage : "
        lblEtat.Caption = "Etat :"
        
        'list
        lblListDoc.Caption = "N° Document"
        lblListNature.Caption = "Nature"
        lblListEtat.Caption = "Etat"
        lblListInd.Caption = "Ind"
        lblListEtatDate.Caption = "Date état"
        lblListDesc.Caption = "Libellé du document"
        lblListOriginDate.Caption = "Origine"
    End If
    
    DocumentKey = ""
    lstDocument.Enabled = False
    
    
    
    
    If cn.State = adStateOpen Then cn.Close
    Set cn = CurrentProject.AccessConnection
    If rstMark.State = adStateOpen Then rstMark.Close
    If rstNature.State = adStateOpen Then rstNature.Close
    If rstEtat.State = adStateOpen Then rstEtat.Close
    
    With rstMark
        Sql = "SELECT CODE_REPERE_MARQUAGE FROM REPERES_MARQUAGE ORDER BY CODE_REPERE_MARQUAGE "
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    Set cmbMarkID.Recordset = rstMark
        
    With rstNature
        Sql = "SELECT CODE_NATURE_DOCUMENT FROM Natures_document ORDER BY CODE_NATURE_DOCUMENT  "
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    Set cmbNatureID.Recordset = rstNature
    
    With rstEtat
        Sql = "SELECT DISTINCT CODE_ETAT_DOCUMENT FROM DOCUMENTS ORDER BY CODE_ETAT_DOCUMENT  "
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    Set cmbEtatID.Recordset = rstEtat
    
    
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

    DoCmd.OpenForm "frm_consultation", acNormal
    DoCmd.Close acForm, "frm_listDocByMark"

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
    Dim MarkId As String
    Dim docNum As String
    Dim NatureID As String
    Dim etatId As String
    Dim i As Integer

    If CheckValid() Then
        MarkId = cmbMarkID.Value
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
                    "FROM DOCUMENTS Where code_repere_marquage ='" & MarkId & "'" & moreSql & " ORDER BY NUMERO_DOCUMENT"
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
                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
            
                'MsgBox "Record Not Found!", vbOKOnly
                'Call ClearAll(Me)
                lstDocument.Enabled = False
                cmbMarkID.SetFocus
 
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
    
    DoCmd.OpenForm "frm_DocsManagement", acNormal, , , acFormReadOnly, , passData
    
    'Hiding form
    Me.Visible = False


Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub


