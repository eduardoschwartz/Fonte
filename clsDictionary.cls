VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------------------
' clsDictionary
'--------------------------------------------------------------------------------
Option Explicit

Private mDict As Dictionary


Private Sub Class_Initialize()
    Set mDict = New Dictionary
End Sub


Public Function Exists(ByVal Key As Long) As Boolean
    Exists = mDict.Exists(Key)
End Function


Public Function Add(ByVal Key As Long, ByVal Value As Long) As Boolean
    If Exists(Key) = False Then
        mDict.Add Key, Value
        Add = True
    Else
        Add = False
    End If
End Function


Public Property Get Item(ByVal Key As Long) As Long
    If Exists(Key) = True Then
        Item = mDict(Key)
    Else
        Item = 0
    End If
End Property


Public Property Let Item(ByVal Key As Long, ByVal Value As Long)
    mDict(Key) = Value
End Property


Public Function Remove(ByVal Key As Long) As Boolean
    If Exists(Key) Then
        mDict.Remove Key
        Remove = True
    Else
        Remove = False
    End If
End Function


Public Function Count() As Long
    Count = mDict.Count
End Function



