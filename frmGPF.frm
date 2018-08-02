VERSION 5.00
Begin VB.Form frmGPF 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorer.exe"
   ClientHeight    =   3825
   ClientLeft      =   3930
   ClientTop       =   3045
   ClientWidth     =   5895
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmGPF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDetails 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmGPF.frx":08CA
      Top             =   1560
      Width           =   5655
   End
   Begin VB.CommandButton cmdDetails 
      Caption         =   "&Details >>"
      Height          =   350
      Left            =   4680
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   370
      Left            =   4680
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   240
      Picture         =   "frmGPF.frx":0921
      Top             =   960
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label lblText 
      Caption         =   "If the problem persists, contact the program vendor."
      Height          =   495
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label lblText 
      Caption         =   "This program has performed an illegal operation and will be shut down."
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image imgCritical 
      Height          =   480
      Left            =   240
      Picture         =   "frmGPF.frx":1253
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmGPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1       'Arrays starting at 1 are easier to keep track of
Option Explicit     'I always use Option Explicit

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" _
Alias "CreateToolhelp32Snapshot" _
(ByVal lFlags As Long, ByVal lProcessID As Long) As Long    'For getting list of processes

Private Const TH32CS_SNAPPROCESS As Long = 2&               'For API calls
Private Const MAX_PATH As Integer = 260                     'For API calls

Private Type PROCESSENTRY32     'A process
  dwSize As Long                'Don't know what this is
  cntUsage As Long              'Don't know what this is
  th32ProcessID As Long         'Don't know what this is
  th32DefaultHeapID As Long     'Don't know what this is
  th32ModuleID As Long          'Don't know what this is
  cntThreads As Long            'Don't know what this is
  th32ParentProcessID As Long   'Don't know what this is
  pcPriClassBase As Long        'Don't know what this is
  dwFlags As Long               'Don't know what this is
  szExeFile As String * MAX_PATH    'The path and filename
End Type

Private Declare Function ProcessFirst Lib "kernel32" _
Alias "Process32First" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long   'For getting processes

Private Declare Function ProcessNext Lib "kernel32" _
Alias "Process32Next" _
(ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long   'For getting processes

Private Declare Sub CloseHandle Lib "kernel32" _
(ByVal hPass As Long)       'When we're done, we need to close

'Private Declare Function SetSysModalWindow Lib "User32" (ByVal hWnd As Integer) As Integer 'Unused

Dim Processes() As String, processCount As Integer  'Array and counter for array

Private Sub cmdClose_Click()
    End   'End program
End Sub

Private Sub cmdDetails_Click()
    Me.Height = 4200        'Expand
    cmdDetails.Enabled = False  'Don't let the user type in the box!
End Sub

Private Sub Form_Load()
    Dim RandomNumber As Integer     'This is reused whenever we need a random number
    Dim appName As String           'The title of the app that will "illegal operation"
    Dim moduleName As String        'for "..has performed an illegal operation at XX"
    Dim extensionCutOff As Boolean  'A flag to determine if we show the extension or not
    
    'Dim SysModalVar As Integer     'Unused

    Me.Icon = imgIcon.Picture   'change the form's icon to a windows flag
    
    
    '************************************
    '   use the Windows API to
    '   read all the names of the
    '   running applications
    '   and load them into an array
    '
    '   ** I believe I got this somewhere
    '   ** on PSC but I can't remember
    '   ** where. :( If someone recognizes
    '   ** it, contact me!
    '
    '   Also, I can't exactly remember
    '   Everything about the API routine
    '   (I made this a while ago), so the
    '   comments in some places may not make
    '   much sense :)
    
    Dim strFilePath As String          'The title of the process
    Dim hSnapShot As Long              'for the API call
    Dim theProcess As PROCESSENTRY32   'for the API call
    Dim r As Long                      'for the API call
    
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)  'Getting the process(es)
    
    
    'Note that this will *never* happen since Explorer is always running..
    If hSnapShot = 0 Then               'If there are no processes running, inform the user
      MsgBox "No processes available!"  'with a simple message box
      End                               'Then exit
    End If                              'Self-explanitory
    
    theProcess.dwSize = Len(theProcess) 'Part of the API routine
    
    r = ProcessFirst(hSnapShot, theProcess) 'For the loop, etc
    
    Do While r  'Main loop to read all processes into array
        strFilePath = GetFileName(theProcess.szExeFile)     'Get the file name (i.e. "Kernel32.dll")
        r = ProcessNext(hSnapShot, theProcess)              'Get next process?
        
        processCount = processCount + 1                     'increment counter
        
        ReDim Preserve Processes(processCount)              'Add another element to array
        Processes(processCount) = strFilePath               'Put file name into array
    Loop ' loop through all the processs
    
    Call CloseHandle(hSnapShot)                 'We're done!
    
    
    '** Get a random application name from our array
    Randomize   ' Initialize random-number generator.
    RandomNumber = Int((processCount * Rnd) + 1)   ' Generate random value between 1 and the total number of processes.
    
    appName = Processes(RandomNumber)               'Get the application name (i.e. "Kernel32.dll")
    
    
    
    '** Sometimes the full name will be shown in a real GPF ("Kernel32.dll"), sometimes not
    '** ("Kernel32", does not have extension). Recreating this makes it look more realistic.
    
    'Determine if the application name will have an extension (75% chance it won't, 25% chance it will)
    Randomize   ' Initialize random-number generator.
    RandomNumber = Int((2 * Rnd) + 1)   ' Generate random value between 1 and 2.

    Select Case RandomNumber    'If the random number is a 1, the extension stays
        Case 1  'it's a 1
            appName = appName           'the extension stays
            extensionCutOff = False     'A boolean flag
        Case 2  'it's a 2
            appName = Left(appName, Len(appName) - 4) 'cut the extension off
            extensionCutOff = True      'A boolean flag
    End Select  'Self-explanitory
    
    
    
    '** This is for the Details page (X has performed an illegal operation in module ..)
    '** The module name will either be one of 8 "generic" ones OR the application's name
    '** (if it is the app's name, we want to check the extensionCutOff flag and add the
    '** extension appropriately). A very small chance for <unknown> (this *does* appear
    '** in real GPFs).
    
    'generate another random number for the module name
    Randomize   ' Initialize random-number generator.
    RandomNumber = Int((10 * Rnd) + 1)   ' Generate random value between 1 and 9.

    Select Case RandomNumber                'Simple Select Case
        Case 1: moduleName = "gdi.exe"      'generic
        Case 2: moduleName = "user.exe"     'generic
        Case 3: moduleName = "kernel.exe"   'generic
        Case 4: moduleName = "RunDLL32.exe" 'generic
        Case 5: moduleName = "RunDLL.exe"   'generic
        Case 6: moduleName = "Explorer.exe" 'generic
        Case 7: moduleName = "shell32.exe"  'generic
        Case 8: moduleName = "KERNEL32.DLL" 'generic
        Case 9  'it will be the app's name
            If extensionCutOff = True Then  'check the flag
                moduleName = appName & ".exe"   'flag is true, add extension
                
            Else                        'flag is false
                moduleName = appName    'do not add extension
            End If
        Case 10: moduleName = "<unknown>"   '1 in 10 chance of <unknown>
    End Select  'Self-explanitory
    
'    Me.Caption = appName    'The caption is the application's name
    Me.Height = 1920        'this height looks realistic
    
    '** Write random text to Details textbox
    With txtDetails     'We don't need to be typing "txtDetails." every time
        '* Generate another random number to select either "invalid page fault" or "exception"
        '* This makes it look even more realistic
        Randomize   ' Initialize random-number generator.
        RandomNumber = Int((2 * Rnd) + 1)   ' Generate random value between 1 and 2.
        
        If RandomNumber = 1 Then        'It will be "invalid page fault"
            .Text = appName & " caused an invalid page fault in module " & moduleName 'add to text
        ElseIf RandomNumber = 2 Then    'It will be "exception"
            .Text = appName & " caused an exception " & genHex(3, True) & " in module " & moduleName 'add to text
        End If
        
        Randomize   'Initialize random-number generator.
        .Text = .Text & " at " & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & Int((9 * Rnd) + 1) & ":"    'this will output " at xxxx:" (x is a random number)
        
        Randomize   'Initialize random-number generator.
        .Text = .Text & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 0) & Int((9 * Rnd) + 1) & "."    'this will output nine random numbers
        
        .Text = .Text & vbCrLf & "Registers:"   'Exactly like a real GPF
        .Text = .Text & vbCrLf & "EAX=" & genHex(8) & " " & "CS=" & genHex(4) & " " & "EIP=" & genHex(8) & " " & "EFLGS=" & genHex(8)   'hex
        .Text = .Text & vbCrLf & "EBX=" & genHex(8) & " " & "SS=" & genHex(4) & " " & "ESP=" & genHex(8) & " " & "EBP=" & genHex(8)     'hex
        .Text = .Text & vbCrLf & "ECX=" & genHex(8) & " " & "DS=" & genHex(4) & " " & "ESI=" & genHex(8) & " " & "FS=" & genHex(4)      'hex
        .Text = .Text & vbCrLf & "EDX=" & genHex(8) & " " & "ES=" & genHex(4) & " " & "EDI=" & genHex(8) & " " & "GS=" & genHex(4)      'hex
    
        .Text = .Text & vbCrLf & "Bytes at CS:EIP:" & vbCrLf & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " " & genHex(2) & " "    'looks like "Bytes at CS:EIP: f1 e6 a8 3d e5 de 32 c1 99 24 81 6f 7d 5d 1e 4f"
        
        .Text = .Text & vbCrLf & "Stack dump:" & vbCrLf & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8) & " " & genHex(8)   'lots of scary hex codes :)
    End With    'Self-explanitory
    Beep        'make a sound
    
    Me.show vbModal 'show the dialog!
    'SysModalVar = SetSysModalWindow(Me.hWnd)   'It doesn't work for some reason
End Sub
