VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PrismPrinters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Windows API Declaration
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

' ===========================================
' NOTE!
' Under Project - References, you must select
' "Windows Script Host Object Model" before
' declaring the object PR as a New WshNetwork
' ===========================================
Private PR As New WshNetwork

Public Function GetDefaultPrinter() As String
Dim def1 As String, def2 As String, def3 As String
Dim di As Long
    
   def2 = String(128, 0)
   ' Find default printer string
   di = GetPrivateProfileString("WINDOWS", "DEVICE", def1, def2, 127, def3)
   ' di = lenght of return string
   If di > 0 Then
      ' Parse string to printer name only using then comma Chr$(44)=","
      di = InStr(def2, Chr$(44)) - 1
      ' Test that di > 0
      If di Then GetDefaultPrinter = Left$(def2, di)
   End If
End Function

Public Function SetDefaultPrinter(ByVal mPrinter As String) As Boolean
Dim prnX As Printer
   
   If mPrinter = "" Then Exit Function
   
   ' Cycle through ALL printers
   For Each prnX In Printers
      ' If a printer = mPrinter then
      If StrComp(prnX.DeviceName, mPrinter, vbTextCompare) = 0 Then
         ' set system default printer
         PR.SetDefaultPrinter (mPrinter)
         ' Set VB default printer
         Set Printer = prnX
         Set PR = Nothing
         SetDefaultPrinter = True
      End If
   Next
End Function


