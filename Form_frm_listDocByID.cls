VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_listDocByID"
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
    
    
    If IsNull(cmbDocId.Value) Or (cmbDocId.Value = "") Or (cmbDocId.ListIndex = -1) Then
        CheckValid = False
        If GB_LANG = "EN" Then
            MsgBox "Invalid Document Number ", vbOKOnly
        Else
            MsgBox "Invalide Numéro De Document  ", vbOKOnly
        End If
        
        cmbDocId.SetFocus
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

    If GB_LANG = "EN" Then
        lblHeader.Caption = "Follow-up Document"
        lblDocID.Caption = "Document Number : "
        
        'list
        lblListDoc.Caption = "Document Number"
        lblListNature.Caption = "Nature"
        lblListEtat.Caption = "Status"
        lblListInd.Caption = "Index"
        lblListEtatDate.Caption = "Status Date"
        lblListDesc.Caption = "Description"
        lblListOriginDate.Caption = "Original Date"
        
    Else 'FR
        lblHeader.Caption = "Consultation Suivi Document"
        lblDocID.Caption = "N° de Document : "
    
        
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
        If rs1.State = adStateOpen Then rs1.Close
        
        With rs1
            Sql = "SELECT NUMERO_DOCUMENT FROM DOCUMENTS ORDER BY NUMERO_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql, cn, , , adCmdText
        End With
        Set cmbDocId.Recordset = rs1
        'cmbDocId.Selected(0) = True
        
        With Me.cmbDocId
            .Value = .ItemData(.ListIndex + 1)
        End With
        cmbDocId.SetFocus

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

'Private Sub cmbDocID_Change()
'Call ClearAll
'End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandler
DoCmd.OpenForm "frm_consultation", acNormal
DoCmd.Close acForm, "frm_listDocByID"

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
'        natureId = cmbNatureID.Value
'        etatId = cmbEtatID.Value
    
      
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
            
                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                                    
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
                'MsgBox "Record Not Found!", vbOKOnly
                'Call ClearAll(Me)
                lstDocument.Enabled = False
                cmbDocId.SetFocus
 
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


