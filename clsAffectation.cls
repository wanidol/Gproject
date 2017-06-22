VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAffectation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Private pAffectID As String
Private pAffectDesc As String

Private pAffectRecord As Collection


Public Property Get AffectID() As String
    AffectID = pAffectID
End Property

Public Property Let AffectID(lAffectID As String)
    pAffectID = lAffectID
End Property

Public Property Get AffectDesc() As String
    AffectDesc = pAffectDesc
End Property

Public Property Let AffectDesc(lAffectDesc As String)
    pAffectDesc = lAffectDesc
End Property

Public Property Get AffectRecords() As Collection
    Set AffectRecords = pAffectRecord
End Property

Public Property Set AffectRecords(lAffectRecord As Collection)
    Set pAffectRecord = lAffectRecord
End Property


