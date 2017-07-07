VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocsManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : frm_DocsManagement
' Author    : Wanidol Lievaux
' Website   : https://www.linkedin.com/in/wanidol-lievaux-0a7bb7124/
' Purpose   :
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged
'
' Revision History:
' Rev   App.version    Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         1.0        2017-jul-25
' 2         2.0        05/07/2017 - Changed font family from copper black to Arial black because
'                                   runtime doesn't support
'                                 - changed caption in v.english from assignement to assignment
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit
Dim Db As Database
Dim rst As New ADODB.Recordset
Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim ParentFormName As String
Dim DocNumber As String


Private Sub UserForm_Initialize()
On Error GoTo ErrHandler


    Dim Sql1 As String
    Dim Sql2 As String
    Dim typeID As String
    Dim critId As String
    Dim docNum As String
    
    If GB_LANG = "EN" Then
        lblHeader.Caption = "Quality Document Management: "
        lblAccord.Caption = "Customer Agreement: "
        lblAffection.Caption = "Assignment: "
        lblAvis.Caption = "Customer reviews: "
        lblComment.Caption = "Comment: "
        lblContractuel.Caption = "Contractual: "
        lblCritere.Caption = "CRITERIA (TYPE+REF): "
        lblDelai.Caption = "Delivery time: "
        lblDiffuse.Caption = "Diffuse: "
        lblDocNo.Caption = "Document Ref: "
        lblEtat.Caption = "Status: "
        lblEtatDate.Caption = "Status Date: "
        lblIndice.Caption = "Index: "
        lblLibelle.Caption = "Wording: "
        lblLocal.Caption = "Local(O-N): "
        lblMark.Caption = "Stamp Part: "
        lblNature.Caption = "Nature: "
        lblNiveauT.Caption = "Treatment Level: "
        lblOA.Caption = "OA Ref: "
        lblOF.Caption = "OF Ref: "
        lblOriginDate.Caption = "Date Origin: "
        lblProposition.Caption = "Proposal: "
        lblTransmitted.Caption = "Transmitted for action: "
        lblUC.Caption = "UC: "
      
        'list
        lblListCritere.Caption = "Criteria"
        lblListType.Caption = "Type"
        
    Else 'FR
        lblHeader.Caption = "Gestion d'un document qualité"
        lblAccord.Caption = "Accord Client: "
        lblAffection.Caption = "Affectation: "
        lblAvis.Caption = "Avis Client: "
        lblComment.Caption = "Commentaire: "
        lblContractuel.Caption = "Contractuel: "
        lblCritere.Caption = "CRITERES (Type+Numéro): "
        lblDelai.Caption = "Délai réception: "
        lblDiffuse.Caption = "Diffusé: "
        lblDocNo.Caption = "N°Document: "
        lblEtat.Caption = "Etat: "
        lblEtatDate.Caption = "Date état: "
        lblIndice.Caption = "Indice: "
        lblLibelle.Caption = "Libellé: "
        lblLocal.Caption = "Local(O-N): "
        lblMark.Caption = "Marquage Pièce: "
        lblNature.Caption = "Nature: "
        lblNiveauT.Caption = "Niveau Traitement: "
        lblOA.Caption = "Numéro OA : "
        lblOF.Caption = "Numéro OF : "
        lblOriginDate.Caption = "Date Origine: "
        lblProposition.Caption = "Proposition: "
        lblTransmitted.Caption = "Transmitted pour action: "
        lblUC.Caption = "UC: "
      
        'list
        lblListCritere.Caption = "Critères"
        lblListType.Caption = "Type"
    End If
    
    lstCritere.Enabled = False

    'Initial values to combobox
    If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rst1.State = adStateOpen Then rst1.Close
        If rst2.State = adStateOpen Then rst2.Close
        With rst1
            Sql1 = "SELECT * FROM documents where NUMERO_DOCUMENT ='" & DocNumber & "'"
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql1, cn, , , adCmdText
        End With
        With rst2
            Sql2 = "SELECT CODE_TYPE_CRITERE,CRITERE_MOT_CLE FROM liens where NUMERO_DOCUMENT ='" & DocNumber & "'"

            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open Sql2, cn, , , adCmdText
        End With
       'Record found for Document data
        If Not (rst1.BOF And rst1.EOF) Then
            With rst1.Fields
                txtComment = .Item("COMMENTAIRES_SUIVI").Value
                txtCustAgree = .Item("DATE_ACCORD_CLIENT").Value
                txtCustReview = .Item("DATE_AVIS_CLIENT").Value
                TxtDocNo = .Item("NUMERO_DOCUMENT").Value
                txtEtatCode = .Item("CODE_ETAT_DOCUMENT").Value
                txtEtatDate = .Item("DATE_ETAT_DOCUMENT").Value
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
            If GB_LANG = "EN" Then
                MsgBox "Record Not Found!", vbOKOnly
            Else
                MsgBox "Enregistrement non trouvé", vbOKOnly
            End If
        End If

        If Not (rst2.BOF And rst2.EOF) Then
            Set lstCritere.Recordset = rst2
            lstCritere.Selected(0) = True
            lstCritere.Enabled = True
        Else
            'MsgBox "Record Not Found!", vbOKOnly
            lstCritere.Enabled = False
        End If
        
        Call LockedCtl(Me)
        Set rst1 = Nothing
        Set rst2 = Nothing
       
       

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub cmdExit_Click()
On Error GoTo ErrHandler

Set rst1 = Nothing
Set rst2 = Nothing
'DoCmd.OpenForm "frm_consultation", acNormal
DoCmd.Close acForm, "frm_DocsManagement"

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
            'Call ClearAll(frm)
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

    If Not IsNull(Forms!frm_DocsManagement.OpenArgs) Then
        DocNumber = Forms!frm_DocsManagement.OpenArgs
        Call ClearAll(Me)
    End If

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub
