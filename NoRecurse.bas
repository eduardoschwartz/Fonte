Attribute VB_Name = "modNoRecurse"
Option Explicit

  Public bDoDebugPrints As Boolean
  
  Public gsDirsQueue As New Collection
  
  Public gsDirs As New Collection
  Public gsFiles As New Collection
  
  Public Const FILE_ATTRIBUTE_READONLY = 1
  Public Const FILE_ATTRIBUTE_HIDDEN = 2
  Public Const FILE_ATTRIBUTE_SYSTEM = 4
  Public Const FILE_ATTRIBUTE_DIRECTORY = 16
  Public Const FILE_ATTRIBUTE_ARCHIVE = 32
  Public Const FILE_ATTRIBUTE_ENCRYPTED = 64
  Public Const FILE_ATTRIBUTE_NORMAL = 128
  Public Const FILE_ATTRIBUTE_TEMPORARY = 256
  Public Const FILE_ATTRIBUTE_SPARSE_FILE = 512
  Public Const FILE_ATTRIBUTE_REPARSE_POINT = 1024
  Public Const FILE_ATTRIBUTE_COMPRESSED = 2048
  Public Const FILE_ATTRIBUTE_OFFLINE = 4096
  Public Const FILE_ATTRIBUTE_NOT_CONTENT_INDEXED = 8192
  
  Public Declare Function GetTickCount Lib "kernel32" () As Long
  
Function FixPath(sPath As String) As String

  If Right(sPath, 1) = "\" Then
    FixPath = Left(sPath, Len(sPath) - 1)
  Else
    FixPath = sPath
  End If
  
End Function

Public Function DirSearch(col As Collection, sDirToAdd As String) As Long

  If col.Count = 1 Then
    DirSearch = 1
    Exit Function
  End If
  
  Dim i As Long
  Dim iStart As Long
  Dim iMidPoint As Long
  
  Dim bFound As Boolean
  
  bFound = False
 'If bDoDebugPrints Then Debug.Print "There are " & col.Count & " items in the collection."
  
  iMidPoint = col.Count / 2
  If iMidPoint = 0 Then iMidPoint = 1
  
 'If bDoDebugPrints Then Debug.Print "Check midpoint item " & sDirToAdd & " > " & iMidPoint & ": " & col.Item(iMidPoint)
  
  If sDirToAdd > col.Item(iMidPoint) Then
    iStart = col.Count
  Else
    iStart = col.Count / 2 + 2
    If iStart > col.Count Then iStart = col.Count
  End If
  
  For i = iStart To 1 Step -1
   'If bDoDebugPrints Then Debug.Print "Comparing " & sDirToAdd & " > " & col.Item(i)
    If sDirToAdd > col.Item(i) Then
      bFound = True
      Exit For
    End If
  Next
  
  If bFound Then
    DirSearch = i
  Else
    DirSearch = col.Count
  End If
  
End Function

Function DirSearchB(col As Collection, sDirToAdd As String) As Long

 'Binary search now, for speed I hope.
 
 'But, first, if this is small just do a sequential search.  Save this stuff for later.
  If col.Count < 30 Then
    DirSearchB = DirSearch(col, sDirToAdd)
    Exit Function
  End If
  
  Dim iLow As Long
  Dim iHigh As Long
  Dim iPivot As Long
  Dim iAdjust As Long
  
 '1. Initialize: iHigh=high element, iLow=low element, iPivot=(iHigh - iLow + 1) \ 2
  iLow = 1: iHigh = col.Count: iPivot = (iHigh - iLow + 1) \ 2 + iLow
 '1a. Start loop and determine ending condition
  iAdjust = 1
 'Debug.Print "Top: " & "L:" & iLow & " H:" & iHigh; " P:" & iPivot & " A:" & iAdjust
  Do While iAdjust <> 0
 '2. Compare New item to Array pivot element
   'Debug.Print sDirToAdd & " > " & col.Item(iPivot)
    If StrComp(sDirToAdd, col.Item(iPivot)) = 1 Then
 '3. When the new item is HIGHER than the array element, LOW = iPivot
      iLow = iPivot
    Else
 '4. When the new item is LOWER than the array element, HIGH = iPivot
      iHigh = iPivot
    End If
 '5. Find pivot: (iHigh - iLow + 1) \ 2 + iLow
    iAdjust = (iHigh - iLow + 1) \ 2
    iPivot = iAdjust + iLow
 '6. When the adjustment value is 0, we have found the two items between which the new item is to be placed.
   'Debug.Print "End: " & "L:" & iLow & " H:" & iHigh; " P:" & iPivot & " A:" & iAdjust
    If iHigh - iLow = 1 Then
      DirSearchB = iLow
      Exit Function
    End If
  Loop
  
  DirSearchB = iPivot
 
End Function

Sub ShowCollection(sID As String, iStart As Long, iCt As Long, col As Collection)
  
  Dim i As Long
  
  Debug.Print "Start " & sID & " --------------------"
  For i = iStart To iCt
    Debug.Print col.Item(i)
  Next
  Debug.Print "End   " & sID & " --------------------"
  

End Sub
