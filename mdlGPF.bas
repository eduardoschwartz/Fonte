Attribute VB_Name = "mdlGPF"
'*************************************************************
'   This project is Copyright (c) 2002 Nathanael Barbettini
' Please contact me if you want to use this code in anything;
'     my e-mail address is nbarb99@hotmail.com. Thanks!!
'*************************************************************

Option Explicit     'I always use Option Explicit..

Public Sub MainG2()  'When the program runs it goes here first, because we can show more than 1 GPF
    Dim i As Integer    'i is for the loop
    Dim frmNewGPF As New frmGPF   ' frmNewGPF is an instance of frmGPF
    
'    Do Until i = 1  'loop
        Set frmNewGPF = New frmGPF  'Create an instance of the form
        
        frmNewGPF.Left = RndInt(7500)   'random location
        frmNewGPF.Top = RndInt(7500)    'random location
        
        frmNewGPF.show  'show it!
 '   i = i + 1: Loop 'Loop
End Sub
Public Function RndInt(upToInt As Integer)
    RndInt = Int((upToInt * Rnd) + 0)   'Generates a random number
End Function

Public Function genHex(Length As Integer, Optional AllCaps As Boolean = False)
'Generates fake hex strings
Dim theHex As String    'this is what will be returned
Dim RndInt As Integer   'for random numbers

    Do Until Len(theHex) = Length   'How long do we want it?
        Randomize   'Initialize the random number generator
        
        RndInt = Int((15 * Rnd) + 1)    'random from 1 to 15
        
        '** Explanation: We get a number from 1 to 15. If the number is less than or equal
        '** to 9, the number itself is added to the string. If it is 10 to 15, a alphabetic
        '** character is added instead.
        
        If RndInt > 9 Then  'Over nine is a hex number
            Select Case RndInt  'Self-explanitory
                Case 10 'RndInt=10
                    If AllCaps = True Then 'check the flag
                        theHex = theHex & "A"   'add uppercase A
                    Else    'flag=false
                        theHex = theHex & "a"   'lowercase a
                    End If  'Self-explanitory
                Case 11 'RndInt=11
                    If AllCaps = True Then 'check the flag
                        theHex = theHex & "B"   'add uppercase B
                    Else    'flag=false
                        theHex = theHex & "b"   'lowercase b
                    End If  'Self-explanitory
                Case 12 'RndInt=12
                    If AllCaps = True Then 'check the flag
                        theHex = theHex & "C"   'add uppercase C
                    Else    'flag=false
                        theHex = theHex & "c"   'lowercase c
                    End If  'Self-explanitory
                Case 13 'RndInt=13
                    If AllCaps = True Then 'check the flag
                        theHex = theHex & "D"   'add uppercase D
                    Else    'flag=false
                        theHex = theHex & "d"   'lowercase d
                    End If  'Self-explanitory
                Case 14 'RndInt=14
                    If AllCaps = True Then 'check the flag
                        theHex = theHex & "E"   'add uppercase E
                    Else    'flag=false
                        theHex = theHex & "e"   'lowercase e
                    End If  'Self-explanitory
                Case 15 'RndInt=15
                    If AllCaps = True Then 'check the flag
                        theHex = theHex & "F"   'add uppercase F
                    Else    'flag=false
                        theHex = theHex & "f"   'lowercase f
                    End If  'Self-explanitory
            End Select  'Self-explanitory
        Else    'RndInt is not above 9; use the number itself
            theHex = theHex & RndInt    'add the number to the string
        End If    'Self-explanitory
    Loop    'loop until we are done generating a string that is the req'd length
    
    genHex = theHex 'return the string
End Function

Public Function GetFileName(thePath As String) As String
'** I believe this function is from somewhere on PSC, but I could be wrong
Dim FilePath As String, tempstr As String, iChars As Integer, lasttempstr As String 'vars

FilePath = Trim(thePath)    'trim leading or trailing spaces
FilePath = Replace(FilePath, Chr(0), "")    'get rid of icky chars

    If InStr(FilePath, "\") = 0 Then    'if we already have a name instead of a path, exit
        GetFileName = FilePath
    Else
        iChars = 1  'for the loop
        Do While iChars < Len(FilePath) 'keep looping until we're done
            tempstr = Right(FilePath, iChars)   'temp string
            
                If InStr(tempstr, "\") <> 0 Then    'if there is a \ in it, return it -1
                    GetFileName = Right(tempstr, iChars - 1): Exit Do   'and exit
                Else    'if not..
                    lasttempstr = tempstr   'save in other var
                End If    'Self-explanitory
            
        iChars = iChars + 1: Loop   'loop until we're done
        
        GetFileName = lasttempstr   'return filename
    End If  'Self-explanitory
End Function
