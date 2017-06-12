Attribute VB_Name = "Module1"
Option Compare Database

Public DocumentKey As String
Public PreviousForm As String
'Global recordset
Public grst As Recordset

Public Sub ClearListBox(lst As ListBox)
On Error GoTo ErrHandler
Dim vItem
With lst
  For Each vItem In .ItemsSelected
    .Selected(vItem) = False
  Next vItem
End With

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
    a = MsgBox("Do You Want To Exit From Program", vbYesNo + vbQuestion, "Quit")
    If a = vbNo Then Exit Sub
    DoCmd.Quit acQuitSaveAll
    
Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub
End Sub

Public Sub ClearAll(frm As Form)
Dim ctl As Control

For Each ctl In frm.Controls
   Select Case ctl.ControlType
      Case acTextBox
           ctl.Value = ""
      Case acOptionGroup, acComboBox, acListBox
          ctl.Value = Null
      Case acCheckBox
         ctl.Value = False
   End Select
Next
End Sub
