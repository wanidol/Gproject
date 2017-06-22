VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsNature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Private pNatureID As String
Private pNatureDesc As String

Private pNatureRecord As Collection


Public Property Get NatureID() As String
    NatureID = pNatureID
End Property

Public Property Let NatureID(lNatureID As String)
    pDocID = lNatureID
End Property

Public Property Get NatureDesc() As String
    NatureID = pNatureDesc
End Property

Public Property Let NatureDesc(lNatureDesc As String)
    pNatureDesc = lNatureDesc
End Property

Public Property Get NatureRecords() As Collection
    Set NatureRecords = pNatureRecord
End Property

Public Property Set DocRecords(lNatureRecord As Collection)
    Set pNatureRecord = lDocRecord
End Property

