VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frm_DocByCritere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Procedure : frm_DocsByCritere
' Author    : Wanidol Lievaux
' Website   : https://www.linkedin.com/in/wanidol-lievaux-0a7bb7124/
' Purpose   :
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2017-jul-25
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
Dim Db As Database
Dim rst As New ADODB.Recordset
Dim rstType As New ADODB.Recordset
Dim rstCritere As New ADODB.Recordset
Dim rstNature As New ADODB.Recordset
Dim rstEtatdate As New ADODB.Recordset
Dim rstDate As New ADODB.Recordset
Dim rstComment As New ADODB.Recordset
Dim rstDesc As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim Status As String

Dim moreSql As String
Dim publicSql As String

Function CheckValid() As Boolean
On Error GoTo ErrHandler

    moreSql = ""
    
    If IsNull(cmbTypeID.Value) Or (cmbTypeID.Value = "") Or (cmbTypeID.ListIndex = -1) Then
        
        
        If GB_LANG = "EN" Then
            MsgBox "Invalid Type Code", vbOKOnly
        Else
            MsgBox "Type Invalide", vbOKOnly
        End If
        
        'MsgBox "Invalid Type Code", vbOKOnly
        
        CheckValid = False
        cmbTypeID.SetFocus
        Exit Function
    
    Else 'else main if
        
        If IsNull(cmbCritereId.Value) Or (cmbCritereId.Value = "") Or (cmbCritereId.ListIndex = -1) Then
        
            
             
            If GB_LANG = "EN" Then
                MsgBox "Invalid Criteria Code", vbOKOnly
            Else
                MsgBox "Critère Invalide ", vbOKOnly
            End If
            
            CheckValid = False
            cmbCritereId.SetFocus
            Exit Function
        Else
        
            If Not (cmbNatureID.ListIndex = -1) Then
                moreSql = "WHERE d.CODE_NATURE_DOCUMENT = '" & cmbNatureID.Value & "'"
            Else
                'case invalid data
                If Not IsNull(cmbNatureID.Value) Then
                    If GB_LANG = "Fr" Then
                       MsgBox "Nature Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid Nature Code", vbOKOnly
                    End If
                    
                    cmbNatureID.SetFocus
                    CheckValid = False
                    Exit Function
                End If
            End If
           
            'nature yes
            If Not (cmbEtatStart.ListIndex = -1) Then
                'start ok
                If Not (cmbEtatEnd.ListIndex = -1) Then
                    'end ok
                    If (cmbNatureID.ListIndex = -1) Then
                        moreSql = "WHERE d.DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatEnd.Value & "'"
                    Else
                        moreSql = moreSql & " And d.DATE_ETAT_DOCUMENT BETWEEN '" & cmbEtatStart.Value & "' AND '" & cmbEtatEnd.Value & "'"
                    End If
                Else
                    'end invalid
                    If GB_LANG = "Fr" Then
                       MsgBox "Date de fin Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid end date", vbOKOnly
                    End If
                    
                    cmbEtatEnd.SetFocus
                    CheckValid = False
                    Exit Function
                
                End If
           
            Else
                'Case Invalid start
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
            
            If Not (cmbStart.ListIndex = -1) Then
                'start ok
                If Not (cmbEnd.ListIndex = -1) Then
                    'end ok
                    If (cmbNatureID.ListIndex = -1) And (cmbEtatStart.ListIndex = -1) Then
                        moreSql = "WHERE  d.DATE_ORIGINE_DOCUMENT BETWEEN '" & cmbStart.Value & "' AND '" & cmbEnd.Value & "'"
                    Else
                        moreSql = moreSql & " And d.DATE_ORIGINE_DOCUMENT BETWEEN '" & cmbStart.Value & "' AND '" & cmbEnd.Value & "'"
                    End If
                Else
                    'invalid end
                   If GB_LANG = "Fr" Then
                       MsgBox "Date de fin Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid end date", vbOKOnly
                    End If
                    
                    cmbEnd.SetFocus
                    CheckValid = False
                    Exit Function
                    
                End If
            Else
                'case Invalid start
                If Not IsNull(cmbStart.Value) Then
                    If GB_LANG = "Fr" Then
                       MsgBox "Date de début Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid start date", vbOKOnly
                    End If
                    
                    cmbStart.SetFocus
                    CheckValid = False
                    Exit Function
                End If
                
            End If
            
         
            
            If Not (cmbComment.ListIndex = -1) Then
         
                If (cmbNatureID.ListIndex = -1) And (cmbEtatStart.ListIndex = -1) And (cmbStart.ListIndex = -1) Then
                    moreSql = "WHERE d.COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
                Else
                    moreSql = moreSql & " AND d.COMMENTAIRES_SUIVI = '" & cmbComment.Value & "'"
                End If
            Else
                'Case Invalid Comment
                 If Not IsNull(cmbComment.Value) Then
                    If GB_LANG = "Fr" Then
                       MsgBox "Commentaire Invalide ", vbOKOnly
                    Else
                        MsgBox "Invalid Comment", vbOKOnly
                    End If
                    
                    cmbComment.SetFocus
                    CheckValid = False
                    Exit Function
                End If
            End If
            CheckValid = True
        End If
    End If 'end main if
  
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
        lblHeader.Caption = "Documents by Criteria"
        lblBox1.Caption = "Required information :"
        lblbox2.Caption = "Optional information :"
        lblComment.Caption = "Comment : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "start : "
        lblEStop.Caption = "end : "
        lblEtat.Caption = "Status"
        lblNature.Caption = "Nature Code : "
        lblStart.Caption = "start : "
        lblStop.Caption = "end : "
        lblCritere.Caption = "Criteria N°"
        lblTypeCode.Caption = "Type Code"
        
        
    Else 'FR
        lblHeader.Caption = "Documents d'un critère"
        lblBox1.Caption = "Informations obligatoires :"
        lblbox2.Caption = "Informations facultatives :"
        lblComment.Caption = "Commentaires : "
        lblDate.Caption = "Date"
        lblEStart.Caption = "début : "
        lblEStop.Caption = "fin : "
        lblEtat.Caption = "Etat"
        lblNature.Caption = "Nature : "
        lblStart.Caption = "début : "
        lblStop.Caption = "fin : "
        lblCritere.Caption = "N° critère :"
        lblTypeCode.Caption = "Type : "
        
    End If
    
    
    
    'Initial values to combobox
        If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rstType.State = adStateOpen Then rstNature.Close
        If rstCritere.State = adStateOpen Then rstNature.Close
        If rstNature.State = adStateOpen Then rstNature.Close
        If rstEtatdate.State = adStateOpen Then rstEtatdate.Close
        If rstDate.State = adStateOpen Then rstDate.Close
        If rstComment.State = adStateOpen Then rstComment.Close
        
        'Type
        With rstType
            sql = "SELECT DISTINCT  CODE_TYPE_CRITERE FROM liens "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbTypeID.Recordset = rstType
        Set rstType = Nothing
        
        'Critere
        
        With rstCritere
           sql = "SELECT DISTINCT CRITERE_MOT_CLE FROM LIENS order by CRITERE_MOT_CLE "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbCritereId.Recordset = rstCritere
        Set rstCritere = Nothing
        
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
        
        
        'Origine Date
        With rstDate
            sql = "SELECT DISTINCT DATE_ORIGINE_DOCUMENT FROM Documents ORDER BY DATE_ORIGINE_DOCUMENT "
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open sql, cn, , , adCmdText
        End With
        
        Set cmbStart.Recordset = rstDate
        Set cmbEnd.Recordset = rstDate
        Set rstDate = Nothing
        cmbEnd.Enabled = False
        
        
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
        With Me.cmbTypeID
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
Private Sub BuildSql()
On Error GoTo ErrHandler
    Dim typeID As String
    Dim critereID As String
    
    publicSql = ""
   
    typeID = cmbTypeID.Value
    critereID = cmbCritereId.Value
    
    publicSql = "SELECT d.NUMERO_DOCUMENT,d.CODE_NATURE_DOCUMENT," & _
                            "d.INDICE_MODIFICATION,d.LIBELLE_DOCUMENT,d.COMMENTAIRES_SUIVI," & _
                            "d.DATE_ETAT_DOCUMENT,d.DATE_ORIGINE_DOCUMENT,d.CODE_REPERE_MARQUAGE " & _
                            "FROM DOCUMENTS d " & _
                            "INNER JOIN (Select NUMERO_DOCUMENT from Liens where CODE_TYPE_CRITERE='" & typeID & "' And CRITERE_MOT_CLE ='" & critereID & "') doc_num " & _
                            "On d.NUMERO_DOCUMENT = doc_num.NUMERO_DOCUMENT ORDER BY d.NUMERO_DOCUMENT "
                            


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


Private Sub cmbStart_Change()
On Error GoTo ErrHandler
    
    If Not (cmbStart.ListIndex = -1) Then
        cmbEnd.Value = ""
        cmbEnd.Enabled = True
    Else
        cmbEnd.Value = ""
        cmbEnd.Enabled = False
    End If

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
    
End Sub

Private Sub cmdPrint_Click()
On Error GoTo ErrHandler
   
    Dim rep As Report
      
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
                CurrentDb.QueryDefs("qry_DocByCritere").sql = publicSql
                
                GB_RPT_TITLE = Nz(cmbTypeID.Value, "-") & ";" & _
                    Nz(cmbCritereId.Value, "-") & ";" & _
                    Nz(cmbNatureID.Value, "-") & ";" & _
                    Nz(cmbEtatStart.Value, "-") & ";" & _
                    Nz(cmbEtatEnd.Value, "-") & ";" & _
                    Nz(cmbStart.Value, "-") & ";" & _
                    Nz(cmbEnd.Value, "-") & ";" & _
                    Nz(cmbComment.Value, "-") & ";"
    
                DoCmd.OpenReport "rpt_DocsByCritere", acViewPreview
                'DoCmd.Minimize acForm, "frm_DocsByCritere"
                PreviousForm = ""
                PreviousForm = Me.Name
                
                DoCmd.Close acForm, "frm_DocByCritere"
    
            Else

                If GB_LANG = "EN" Then
                    MsgBox "Record Not Found!", vbOKOnly
                Else
                    MsgBox "Enregistrement non trouvé", vbOKOnly
                End If
               'Call ClearAll(Me)
               cmbTypeID.SetFocus

            End If

        'Close/Open Form?
        'DoCmd.Close acForm, "frm_DocByNature"
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

    DoCmd.OpenForm "frm_workRequest", acNormal
    DoCmd.Close acForm, "frm_DocByCritere"


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

