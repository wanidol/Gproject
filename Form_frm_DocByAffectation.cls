VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByAffectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'---------------------------------------------------------------------------------------
' Procedure : frm_DocsByAffectation
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
Dim rstAffect As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
Dim rstEtatdate As New ADODB.Recordset
Dim rstClient As New ADODB.Recordset
Dim rstComment As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim publicSql As String


Function CheckValid() As Boolean
On Error GoTo ErrHandler

    moreSql = ""
    
    If IsNull(cmbAffectID.Value) Or (cmbAffectID.Value = "") Or (cmbAffectID.ListIndex = -1) Then
        
        If GB_LANG = "EN" Then
            MsgBox "Invalid Assignment Code", vbOKOnly
        Else
            MsgBox "Affectation Invalide ", vbOKOnly
        End If
        
        CheckValid = False
        cmbAffectID.SetFocus
        
        Exit Function
        
    Else 'else main if
        If Not (cmbNatureID.ListIndex = -1) Then
            moreSql = moreSql & " And CODE_NATURE_DOCUMENT = '" & cmbNatureID.Column(0) & "'"
        Else
            If Not IsNull(cmbNatureID.Value) Then
                If GB_LANG = "Fr" Then
                    MsgBox "Code de Nature Invalide ", vbOKOnly
                Else
                    MsgBox "Invalid Nature code", vbOKOnly
                End If
                cmbNatureID.SetFocus
                CheckValid = False
                Exit Function
            End If
        End If
        
        If Not (cmbEtatStart.ListIndex = -1) Then
        
            If Not (cmbEtatEnd.ListIndex = -1) Then
                moreSql = moreSql & " And DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatEnd.Value & "'"
            Else
'                If Not IsNull(cmbEtatEnd.Value) Then
                    If GB_LANG = "Fr" Then
                        MsgBox "Date de fin Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid end date", vbOKOnly
                    End If
                    cmbEtatEnd.SetFocus
                    CheckValid = False
                    Exit Function
'               End If
            End If
        Else
            If Not IsNull(cmbEtatStart.Value) Then
                If GB_LANG = "Fr" Then
                    MsgBox "Date de début Invalide ", vbOKOnly
                Else
                    MsgBox "Invalid start date", vbOKOnly
                End If
                
                cmbEtatStart.SetFocus
                CheckValid = False
                Exit Function
            End If
        
        End If
        
        If Not (cmbClientLinked.ListIndex = -1) Then
             moreSql = moreSql & " AND CONTRACTUEL = '" & cmbClientLinked.Value & "'"
        Else
            If Not IsNull(cmbClientLinked.Value) Then
                If GB_LANG = "Fr" Then
                    MsgBox "Lié client Invalide ", vbOKOnly
                Else
                    MsgBox "Invalid client linked", vbOKOnly
                End If
                
                cmbClientLinked.SetFocus
                CheckValid = False
                Exit Function
            End If
        End If
        
        If Not (cmbComment.ListIndex = -1) Then
            moreSql = moreSql & " AND COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
        Else
            If Not IsNull(cmbComment.Value) Then
                If GB_LANG = "Fr" Then
                    MsgBox "Invalide Commentaire", vbOKOnly
                Else
                    MsgBox "Invalid Comment Data", vbOKOnly
                End If
                
                cmbComment.SetFocus
                CheckValid = False
                Exit Function
            End If
        End If
        
        CheckValid = True
    End If ' end main if
  
Exit_Sub:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Function

Private Sub UserForm_Initialize()
On Error GoTo ErrHandler
    Dim sql As String
    Dim typeID As String
    Dim critId As String
    Dim docNum As String
    Dim i As Integer

    'Clear Variable
    GB_RPT_TITLE = ""
    'Clear all Controls
    Call ClearAll(Me)
    
    If GB_LANG = "EN" Then
        lblHeader.Caption = "Documents by assignment"
        lblBox1.Caption = "Required information :"
        lblbox2.Caption = "Optional information :"
        lblComment.Caption = "Comment : "
        lblEStart.Caption = "start : "
        lblEStop.Caption = "end : "
        lblEtat.Caption = "Status"
        lblNature.Caption = "Nature Code : "
        
        lblAffectation.Caption = "Assignment : "
        lblClientLinked.Caption = "Related to a client : "
        
        
        
    Else 'FR
        lblHeader.Caption = "Documents d'une affectation"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblComment.Caption = "Commentaires : "
       
        lblEStart.Caption = "début : "
        lblEStop.Caption = "fin : "
        lblEtat.Caption = "Etat"
        lblNature.Caption = "Nature : "
        
        lblAffectation.Caption = "Affectation : "
        lblClientLinked.Caption = "Liés à un client : "
        
    End If
    
    'Initial values to combobox
        If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rstAffect.State = adStateOpen Then rstAffect.Close
        If rstNature.State = adStateOpen Then rstNature.Close
        If rstEtatdate.State = adStateOpen Then rstEtatdate.Close
        If rstClient.State = adStateOpen Then rstClient.Close
        If rstComment.State = adStateOpen Then rstComment.Close
        
        
        'Affectation
        With rstAffect
             sql = "SELECT CODE_AFFECTATION,LIBELLE_AFFECTATION FROM AFFECTATION ORDER BY CODE_AFFECTATION  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbAffectID.Recordset = rstAffect
        Set rstAffect = Nothing
        
        'NatureID
        With rstNature
            sql = "SELECT CODE_NATURE_DOCUMENT,LIBELLE_NATURE_DOCUMENT FROM Natures_document ORDER BY CODE_NATURE_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        

        Set cmbNatureID.Recordset = rstNature
        Set rstNature = Nothing
        
       'Etat Date
        With rstEtatdate
            sql = "SELECT DISTINCT DATE_ETAT_DOCUMENT FROM Documents WHERE DATE_ETAT_DOCUMENT IS NOT NULL  ORDER BY DATE_ETAT_DOCUMENT "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbEtatStart.Recordset = rstEtatdate
        Set cmbEtatEnd.Recordset = rstEtatdate
        Set rstEtatdate = Nothing
        
        
        cmbEtatEnd.Enabled = False
        
        'Linked to a client
        With rstClient
         sql = "SELECT DISTINCT CONTRACTUEL FROM DOCUMENTS"
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbClientLinked.Recordset = rstClient
        Set rstClient = Nothing
        
        'comment
        
         With rstComment
            sql = "SELECT COMMENTAIRES_SUIVI FROM Documents ORDER BY NUMERO_DOCUMENT  "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbComment.Recordset = rstComment
        Set rstComment = Nothing
        
        
        'set focus on Nature Code default first listindex
        With Me.cmbAffectID
            .Value = .ItemData(.ListIndex + 1)
            '.Selected(1) = True
             '   .Value = .ItemData(0)
            '.Value = .ListIndex(0)
'            .ListIndex = 1
        End With
        cmbAffectID.SetFocus
            
       

Exit_Sub:
    Exit Sub
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

End Sub

Private Sub cmbEtatStart_Change()
On Error GoTo ErrHandler

    If Not (cmbEtatStart.ListIndex = -1) Then
        cmbEtatEnd.Value = ""
        cmbEtatEnd.Enabled = True
    Else
        cmbEtatEnd.Value = ""
        cmbEtatEnd.Enabled = False
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

    Set rst = Nothing
    DoCmd.OpenForm "frm_workRequest", acNormal
    DoCmd.Close acForm, "frm_DocByAffectation"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Private Sub BuildSql()
On Error GoTo ErrHandler
    Dim AffactID As String
    
    publicSql = ""
   
    AffactID = cmbAffectID.Value
 
    publicSql = "SELECT NUMERO_DOCUMENT,CODE_NATURE_DOCUMENT,INDICE_MODIFICATION," & _
                            "LIBELLE_DOCUMENT,COMMENTAIRES_SUIVI,DATE_ETAT_DOCUMENT,DATE_ORIGINE_DOCUMENT,CODE_AFFECTATION " & _
                            "FROM DOCUMENTS WHERE CODE_AFFECTATION ='" & AffactID & "'" & moreSql & _
                            "ORDER BY NUMERO_DOCUMENT"


Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrHandler

    Dim oDocument As New clsDocuments
      
    If CheckValid() Then
            
        If cn.State = adStateOpen Then cn.Close
            Set cn = CurrentProject.AccessConnection
            If rst.State = adStateOpen Then rst.Close

            With rst
                Call BuildSql

                .CursorType = adOpenDynamic
                .CursorLocation = adUseClient
                .LockType = adLockOptimistic
                .Open publicSql, cn, , , adCmdText
            End With

            'Record Found
            If Not (rst.BOF And rst.EOF) Then
                CurrentDb.QueryDefs("qry_DocByAffectation").sql = publicSql
                GB_RPT_TITLE = Nz(cmbAffectID.Value) & ";" & _
                    Nz(cmbNatureID.Value, "-") & ";" & _
                    Nz(cmbEtatStart.Value, "-") & ";" & _
                    Nz(cmbEtatEnd.Value, "-") & ";" & _
                    Nz(cmbClientLinked.Value, "-") & ";" & _
                    Nz(cmbComment.Value, "-") & ";"
        
                DoCmd.OpenReport "rpt_DocsByAffectation", acViewPreview
                PreviousForm = ""
                PreviousForm = Me.Name
                
                DoCmd.Close acForm, "frm_DocByAffectation"
            Else
                               
                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
                'Call ClearAll(Me)
                cmbAffectID.SetFocus

            End If

        'Close/Open Form?
        'DoCmd.Close acForm, "frm_DocByAffectation"
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

