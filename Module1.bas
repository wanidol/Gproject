Attribute VB_Name = "Module1"
Option Compare Database

Public DocumentKey As String
Public PreviousForm As String

Public GB_tempTable As String
Public GB_RPT_TITLE As String
Public GB_LANG As String
Public GB_cDocs As Collection

Public Function fReturnRecordset(tSql As String, tLst As Variant)
On Error GoTo ErrHandler
  
    'Dim Db As Database
    Dim rst As New ADODB.Recordset
    Dim cn As New ADODB.Connection

     
    If cn.State = adStateOpen Then cn.Close
        Set cn = CurrentProject.AccessConnection
        If rst.State = adStateOpen Then rst.Close
        With rst
            .CursorType = adOpenDynamic
            .CursorLocation = adUseClient
            .LockType = adLockOptimistic
            .Open tSql, cn, , , adCmdText
        End With
        
        'Record Found
        If Not (rst.BOF And rst.EOF) Then
           
                Set tLst.Recordset = rst
                Debug.Print rst.RecordCount
                
                
        Else
           
           If GB_LANG = "EN" Then
                MsgBox "Record Not Found!", vbOKOnly
                                
           Else
                MsgBox "Enregistrement non trouvé", vbOKOnly
           End If
        
        End If
        Set rst = Nothing
        
        
Exit_Sub:
    Exit Function
ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Function


Public Sub InitGlobal()
    'Fr /Eng
    GB_LANG = "Fr"
    GB_RPT_TITLE = ""
    

End Sub

Public Sub ClearListBox(lst As ListBox)
On Error GoTo ErrHandler

    lst.RowSourceType = "Table/Query"
    lst.RowSource = ""
    lst.Requery
    lst.Enabled = False
    
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Public Sub QUITFORM()
On Error GoTo ErrHandler
    Dim a As Byte
    
    If GB_LANG = "FR" Then
        a = MsgBox("Voulez-Vous Quitter Le Programme", vbYesNo + vbQuestion, "Quitter")
    Else
        a = MsgBox("Do You Want To Exit From Program", vbYesNo + vbQuestion, "Quit")
    End If
    
    If a = vbNo Then Exit Sub
    DoCmd.Quit acQuitSaveAll
    
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub
Public Sub LockedCtl(frm As Form)
Dim ctl As Control

    For Each ctl In frm.Controls
        If ctl.Tag = "Locked" Then
            ctl.Locked = True
        End If
    Next

End Sub
Public Sub ClearAll(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
   Select Case ctl.ControlType
        Case acTextBox
           ctl.Value = ""
        Case acOptionGroup, acComboBox, acListBox
          ctl.Value = Null
        
        Case acListBox
            ctl.RowSourceType = "Table/Query"
            ctl.RowSource = ""
            ctl.Requery
            ctl.Enabled = False
            

        Case acCheckBox
         ctl.Value = False
   End Select
Next
End Sub
Public Sub setDefault(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
    Select Case ctl.ControlType
    Case acComboBox
        ctl.Value = ctl.ItemData(ctl.ListIndex + 1)
    End Select
Next

End Sub


Public Function DoesRptExist(sReportName As String) As Boolean
   Dim rpt  As Object
 
On Error GoTo Error_Handler
   'Initialize our variable
   DoesRptExist = False
 
   Set rpt = CurrentProject.AllReports(sReportName)
   
   
 
   DoesRptExist = True  'If we made it to here without triggering an error
                        'the report exists
 
Error_Handler_Exit:
   On Error Resume Next
   Set rpt = Nothing
   Exit Function
 
Error_Handler:
   If Err.Number = 2467 Then
   
    MsgBox ("Report Not Found")
      'If we are here it is because the report could not be found
   Else
      MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
      Err.Number & vbCrLf & "Error Source: DoesRptExist" & vbCrLf & "Error Description: " & _
      Err.Description, vbCritical, "An Error has Occured!"
   End If
   Resume Error_Handler_Exit
End Function

'---------------------------------------------------------------------------------------
' Procedure : CountOpenRpts
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Returns a count of the number of loaded reports (preview or design)
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites.
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2009-Oct-30                 Initial Release
' 2         2009-Oct-31                 Switched from AllReports to Reports collection
'---------------------------------------------------------------------------------------
Function CountOpenRpts()
On Error GoTo Error_Handler
 
    CountOpenRpts = Application.Reports.Count
 
Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: CountOpenRpts" & vbCrLf & "Error Description: " & _
    Err.Description, vbCritical, "An Error has Occured!"
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : ListOpenRpts
' Author    : CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Returns a list of all the loaded reports (preview or design)
'             separated by ; (ie: Report1;Report2;Report3)
' Copyright : The following may be altered and reused as you wish so long as the
'             copyright notice is left unchanged (including Author, Website and
'             Copyright).  It may not be sold/resold or reposted on other sites.
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2009-Oct-30                 Initial Release
' 2         2009-Oct-31                 Switched from AllReports to Reports collection
'---------------------------------------------------------------------------------------
Function ListOpenRpts()
On Error GoTo Error_Handler
 
    Dim DbR     As Report
    Dim DbO     As Object
    Dim Rpts    As Variant
 
    Set DbO = Application.Reports
 
    For Each DbR In DbO    'Loop all the reports
            Rpts = Rpts & ";" & DbR.Name
    Next DbR
 
    If Len(Rpts) > 0 Then
        Rpts = Right(Rpts, Len(Rpts) - 1)   'Truncate initial ;
    End If
 
    ListOpenRpts = Rpts
 
Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: ListOpenRpts" & vbCrLf & "Error Description: " & _
    Err.Description, vbCritical, "An Error has Occured!"
    Exit Function
End Function

'---------------------------------------------------------------------------------------
Function ListDbRpts() As String
On Error GoTo Error_Handler
 
    Dim DbO     As AccessObject
    Dim DbP     As Object
    Dim Rpts    As String
 
    Set DbP = Application.CurrentProject
 
    For Each DbO In DbP.AllReports
        Rpts = Rpts & ";" & DbO.Name
    Next DbO
    Rpts = Right(Rpts, Len(Rpts) - 1) 'Truncate initial ;
 
    ListDbRpts = Rpts
 
Exit Function
 
Error_Handler:
    MsgBox "MS Access has generated the following error" & vbCrLf & vbCrLf & "Error Number: " & _
    Err.Number & vbCrLf & "Error Source: ListDbRpts" & vbCrLf & "Error Description: " & _
    Err.Description, vbCritical, "An Error has Occured!"
    Exit Function
End Function
