VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_listDocByCritere"
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
Dim rstEtat As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
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
    
    
    If IsNull(cmbTypeID.Value) Or (cmbTypeID.Value = "") Or (cmbTypeID.ListIndex = -1) Then
        CheckValid = False
        
        If GB_LANG = "FR" Then
            MsgBox "Type Invalide", vbOKOnly
        Else
            MsgBox "Invalid Type Code", vbOKOnly
        End If
        
        cmbTypeID.SetFocus
        Exit Function
    Else
        If IsNull(cmbCritereId.Value) Or (cmbCritereId.Value = "") Or (cmbCritereId.ListIndex = -1) Then
            CheckValid = False
            
            If GB_LANG = "FR" Then
                MsgBox " N° Citère Invalide", vbOKOnly
            Else
                MsgBox "Invalid Criteria Code", vbOKOnly
            End If
            
            cmbCritereId.SetFocus
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
                    moreSql = " d.CODE_ETAT_DOCUMENT = '" & cmbEtatID.Value & "'"
'                    Debug.Print moreSql
                    CheckValid = True
                End If
            Else
                'case 1.2 nature ok etat no
                If IsNull(cmbEtatID.Value) Or (cmbEtatID.Value = "") Then
                    moreSql = " d.CODE_NATURE_DOCUMENT = '" & cmbNatureID.Value & "'"
'                    Debug.Print moreSql
                Else
                'case 2. nature ok etat ok
                    moreSql = " d.CODE_NATURE_DOCUMENT = '" & cmbNatureID.Value & "'" & _
                                " And d.CODE_ETAT_DOCUMENT = '" & cmbEtatID.Value & "'"
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
    If GB_LANG = "EN" Then
        lblHeader.Caption = "Liste des documents liés à un critère d'accès"
        lblTypeCode.Caption = "Type :"
        lblCritCode.Caption = "Criteria Code :"
        lblNature.Caption = "Nature :"
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
        lblHeader.Caption = "Liste des documents liés à un critère d'accès"
        lblTypeCode.Caption = "Type :"
        lblCritCode.Caption = "N° critère :"
        lblNature.Caption = "Nature :"
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
    
        Sql = "SELECT DISTINCT CRITERE_MOT_CLE FROM LIENS order by CRITERE_MOT_CLE "
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    
     With rstNature
        Sql = "SELECT CODE_NATURE_DOCUMENT FROM Natures_document ORDER BY CODE_NATURE_DOCUMENT  "
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    
    
    With rstEtat
        Sql = "SELECT DISTINCT CODE_ETAT_DOCUMENT FROM DOCUMENTS ORDER BY CODE_ETAT_DOCUMENT  "
        .CursorType = adOpenDynamic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .Open Sql, cn, , , adCmdText
    End With
    
    Set cmbEtatID.Recordset = rstEtat
    Set cmbTypeID.Recordset = rstType
    Set cmbCritereId.Recordset = rstCritere
    Set cmbNatureID.Recordset = rstNature
    
        
        With Me.cmbTypeID
            .Value = .ItemData(.ListIndex + 1)
        End With
        With Me.cmbCritereId
            .Value = .ItemData(.ListIndex + 1)
        End With
        
        cmbTypeID.SetFocus

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
    DoCmd.Close acForm, "frm_listDocByCritere"

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
    Dim NatureID As String
    Dim etatId As String
    Dim i As Integer


    If CheckValid() Then
        typeId = cmbTypeID.Value
        critId = cmbCritereId.Value
     
        If cn.State = adStateOpen Then cn.Close
            Set cn = CurrentProject.AccessConnection
            If rst.State = adStateOpen Then rst.Close
            With rst
            
            Sql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                    "d.CODE_ETAT_DOCUMENT,d.INDICE_MODIFICATION," & _
                    "d.DATE_ETAT_DOCUMENT,d.LIBELLE_DOCUMENT,d.DATE_ORIGINE_DOCUMENT " & _
                    "FROM DOCUMENTS d Inner join (Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='" & typeId & "' And CRITERE_MOT_CLE ='" & critId & "') doc_num " & _
                    "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT " & _
                    "WHERE " & moreSql & " ORDER BY d.NUMERO_DOCUMENT"

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
                cmbTypeID.SetFocus
 
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
