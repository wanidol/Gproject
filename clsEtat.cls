VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEtat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database


Private pEtatID As String
Private pEtatDesc As String

Private pEtatRecord As Collection


Public Property Get EtatID() As String
    EtatID = pEtatID
End Property

Public Property Let EtatID(lEtatID As String)
    pEtatID = lEtatID
End Property

Public Property Get EtatDesc() As String
    EtatDesc = pEtatDesc
End Property

Public Property Let EtatDesc(lEtatDesc As String)
    pEtatDesc = lEtatDesc
End Property

Public Property Get EtatRecords() As Collection
    Set EtatRecords = pEtatRecord
End Property

Public Property Set EtatRecords(lEtatRecord As Collection)
    Set pEtatRecord = lEtatRecord
End Property


