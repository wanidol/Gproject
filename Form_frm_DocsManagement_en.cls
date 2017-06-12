VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocsManagement_en"
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
Dim ParentFormName As String
Dim DocNumber As String


Private Sub UserForm_Initialize()
On Error GoTo ErrHandler


    Dim Sql1 As String
    Dim Sql2 As String
    Dim typeId As String
    Dim critId As String
    Dim docNum As String

    lstCritere.Enabled = False

    'Initial values to combobox
    If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rs1.State = adStateOpen Then rs1.Close
        If rs2.State = adStateOpen Then rs2.Close
        With rs1
            Sql1 = "SELECT * FROM documents where NUMERO_DOCUMENT ='" & DocNumber & "'"
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql1, cn, , , adCmdText
        End With
        With rs2
            Sql2 = "SELECT CODE_TYPE_CRITERE,CRITERE_MOT_CLE FROM liens where NUMERO_DOCUMENT ='" & DocNumber & "'"

            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql2, cn, , , adCmdText
        End With
       'Record found for Document data
        If Not (rs1.BOF And rs1.EOF) Then
            With rs1.Fields
                txtComment = .Item("COMMENTAIRES_SUIVI").Value
                txtCustAgree = .Item("DATE_ACCORD_CLIENT").Value
                txtCustReview = .Item("DATE_AVIS_CLIENT").Value
                TxtDocNo = .Item("NUMERO_DOCUMENT").Value
                txtEtatCode = .Item("CODE_ETAT_DOCUMENT").Value
                txtetatDate = .Item("DATE_ETAT_DOCUMENT").Value
                txtIndice = .Item("INDICE_MODIFICATION").Value
                txtNatureCode = .Item("CODE_NATURE_DOCUMENT").Value
                txtOA = .Item("NUMERO_OA").Value
                txtOF = .Item("NUMERO_OF").Value
                txtOriginDate = .Item("DATE_ORIGINE_DOCUMENT").Value
                txtPropo = .Item("DATE_PROPOSITION_CLIENT").Value
                txtRecepTime = .Item("DELAI_RECEPTION_OA").Value
                txtWording = .Item("LIBELLE_DOCUMENT").Value
                cmbAffectation = .Item("CODE_AFFECTATION").Value
                cmbContractual = .Item("CONTRACTUEL").Value
                cmbDiffuse = .Item("NOM_DIFFUSE").Value
                cmbLocal = .Item("TRAITEMENT_LOCAL_FNC").Value
                cmbMark = .Item("CODE_REPERE_MARQUAGE").Value
                cmbTransmis = .Item("SERVICE_ACTION").Value
                cmbTreatLevel = .Item("NIVEAU_TRAITEMENT_FNC").Value
                cmbUC = .Item("CODE_UNITE").Value
            End With
        Else
            MsgBox "Record Not Found!", vbOKOnly
        End If

        If Not (rs2.BOF And rs2.EOF) Then
            Set lstCritere.Recordset = rs2
            lstCritere.Selected(0) = True
            lstCritere.Enabled = True
        Else
            MsgBox "Record Not Found!", vbOKOnly
            lstCritere.Enabled = False
        End If
        
       
       

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandler

'DoCmd.OpenForm "frm_consultation_en", acNormal
DoCmd.Close acForm, "frm_DocsManagement_en"

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub Form_Close()
On Error GoTo ErrHandler

Dim frm As Form

    For Each frm In Application.Forms
        If frm.Name = PreviousForm Then
            frm.Visible = True
            Exit For
        End If
    Next frm

    
Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim DocNumber As String


    Set Db = CurrentDb
    Call UserForm_Initialize

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub Form_Open(Cancel As Integer)
On Error GoTo ErrHandler

    DocNumber = Forms!frm_DocsManagement_en.OpenArgs
    Call ClearAll(Me)

'    If Len(DocNumber) > 0 Then
'       TxtDocNo.Value = DocNumber
'    End If

  

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub
