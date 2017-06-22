VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDocuments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private pDocID As String
Private pNatureID As String
Private pIndex As String
Private pLibelle As String
Private pEtat As String
Private pOrigin As String
Private pComment As String
Private pDocRecord As Collection

'Public rep As Report

Public Property Get DocID() As String
    DocID = pDocID
End Property

Public Property Let DocID(lDocID As String)
    pDocID = lDocID
End Property

Public Property Get NatureID() As String
    NatureID = pNatureID
End Property

Public Property Let NatureID(lNatureID As String)
    pNatureID = lNatureID
End Property

Public Property Get Indice() As String
  Indice = pIndex
End Property
Public Property Let Indice(lIndex As String)
    pIndex = lIndex
End Property


Public Property Get Libelle() As String
    Libelle = pLibelle
End Property

Public Property Let Libelle(lLiebelle As String)
    pLibelle = lLiebelle
End Property

Public Property Get Etat() As String
    Etat = pEtat
    
End Property
Public Property Let Etat(lEtat As String)
    pEtat = lEtat
 
End Property

Public Property Get Origin() As String
    Origin = pOrigin
End Property

Public Property Let Origin(lOrigin As String)
    pOrigin = lOrigin
End Property

Public Property Get Commentaire() As String
    Commentaire = pComment
End Property
Public Property Let Commentaire(lComment As String)
    pComment = lComment
End Property


Public Property Get DocRecords() As Collection
    Set DocRecord = pDocRecord
End Property

Public Property Set DocRecords(lDocRecord As Collection)
    Set pDocRecord = lDocRecord
End Property

'Function PopulateArray() As Collection
'    Dim clnDoclis As Collection
'    Dim docli As DocRecords
'
'    Set clnDoclis = New Collection
'
'    Set docli = New DocRecords
'        docli.DocID = "1"
'        docli.NatureID = "001"
'        docli.Indice = "hahaha"
'    cInDolis.Add docli
'
'    Set docli = New DocRecords
'        docli.DocID = "2"
'        docli.NatureID = "002"
'        docli.Indice = "bububu"
'    cInDolis.Add docli
'
'    Set docli = New DocRecords
'        docli.DocID = "3"
'        docli.NatureID = "003"
'        docli.Indice = "kakakaka"
'    cInDolis.Add docli
'
'End Function

