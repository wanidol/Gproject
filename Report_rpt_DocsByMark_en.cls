VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Report_rpt_DocsByMark_en"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Report_Open(Cancel As Integer)
On Error GoTo ErrHandler
    Dim strOpenArgs() As String
    Dim Str As String
    'Str = Reports!Documents.OpenArgs
    'Str = Me.OpenArgs
    If Not IsNull(RPT_TITLE) Then
      strOpenArgs = Split(RPT_TITLE, ";")
      Me.lblMarkID.Caption = strOpenArgs(0)
      Me.lblNatureID.Caption = strOpenArgs(1)
      Me.lblEtatStart.Caption = strOpenArgs(2)
      Me.lblEtatStop.Caption = strOpenArgs(3)
      Me.lblDateStart.Caption = strOpenArgs(4)
      Me.lblDateStop.Caption = strOpenArgs(5)
      Me.lblComment.Caption = strOpenArgs(6)
      Me.lblCritere.Caption = strOpenArgs(7)
    End If

Exit_Sub:
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in " & _
    VBE.ActiveCodePane.CodeModule, vbOKOnly, "Error"
    
Resume Exit_Sub

    
End Sub



