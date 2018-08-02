VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmEventSecurity 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Segurança do Sistema"
   ClientHeight    =   6285
   ClientLeft      =   3615
   ClientTop       =   2685
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   11940
   Begin VB.CommandButton cmdRetorna 
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Gravar"
      Height          =   285
      Left            =   9270
      TabIndex        =   21
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00EEEEEE&
      Caption         =   "&Cancelar"
      Height          =   285
      Left            =   10260
      TabIndex        =   20
      Top             =   90
      Width           =   915
   End
   Begin MSFlexGridLib.MSFlexGrid grdResult 
      Height          =   1965
      Left            =   570
      TabIndex        =   17
      Top             =   6780
      Visible         =   0   'False
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3466
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedCols       =   0
      BackColor       =   12648447
      BackColorSel    =   64
      ForeColorSel    =   16777215
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "IdTela |Nome da Tela                     |IdEvento |Nome do Evento                      |Objeto                      |Acesso     "
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00808080&
      Caption         =   "Status Gravação"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   525
      Left            =   3990
      TabIndex        =   14
      Top             =   5685
      Width           =   4245
      Begin MSComctlLib.ProgressBar Pb 
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00808000&
      Height          =   525
      Left            =   0
      TabIndex        =   11
      Top             =   -90
      Width           =   11925
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   240
         Left            =   6975
         TabIndex        =   22
         Top             =   225
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.ComboBox cmbTela 
         Height          =   315
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   150
         Width           =   3675
      End
      Begin VB.Label lblNomeForm 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   165
         Left            =   5580
         TabIndex        =   16
         Top             =   240
         Width           =   3705
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telas do Sistema"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   210
         Width           =   1545
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00808080&
      Caption         =   "Permissões"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   3375
      Left            =   8250
      TabIndex        =   9
      Top             =   2760
      Width           =   3675
      Begin VB.ComboBox cmbEvento 
         Height          =   315
         Left            =   90
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   300
         Width           =   3285
      End
      Begin MSFlexGridLib.MSFlexGrid grdObj 
         Height          =   2625
         Left            =   60
         TabIndex        =   10
         Top             =   690
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   4630
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         ForeColorFixed  =   0
         BackColorSel    =   16777215
         ForeColorSel    =   -2147483642
         FocusRect       =   0
         FormatString    =   "Objeto                          |^S |^U |^I  |^D |^Ex   "
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808080&
      Caption         =   "Eventos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2265
      Left            =   8250
      TabIndex        =   7
      Top             =   480
      Width           =   3675
      Begin VB.ListBox lstEvento 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   60
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808080&
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   5190
      Left            =   4005
      TabIndex        =   6
      Top             =   495
      Width           =   4245
      Begin MSComctlLib.TreeView tvMenu 
         Height          =   4890
         Left            =   90
         TabIndex        =   15
         Top             =   240
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   8625
         _Version        =   393217
         Indentation     =   471
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Caption         =   "Stored Procedure"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2055
      Left            =   30
      TabIndex        =   2
      Top             =   2550
      Width           =   3945
      Begin VB.ListBox lstProc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   30
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Views"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1650
      Left            =   30
      TabIndex        =   1
      Top             =   4590
      Width           =   3945
      Begin VB.ListBox lstView 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   5
         Top             =   240
         Width           =   3825
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Tabelas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   2055
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   3945
      Begin VB.ListBox lstTable 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   60
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   240
         Width           =   3825
      End
   End
End
Attribute VB_Name = "frmEventSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetMenuStringA Lib "user32" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long

Private Const MF_BYPOSITION = &H400&

Dim RdoAux As rdoResultset
Dim Sql As String
Dim Lin As Integer, col As Integer
Dim x As Integer, y As Integer, bExec As Boolean

Private Sub cmbEvento_Click()

If Not bExec Then Exit Sub
If cmbEvento.ListIndex = -1 Then Exit Sub
For x = 1 To grdObj.Rows - 1
    For y = 1 To grdObj.Cols - 1
           grdObj.TextMatrix(x, y) = ""
    Next
Next
    
For x = 1 To grdObj.Rows - 1
    For y = 1 To grdResult.Rows - 1
        If Val(grdResult.TextMatrix(y, 0)) = cmbTela.ItemData(cmbTela.ListIndex) And Val(grdResult.TextMatrix(y, 2)) = cmbEvento.ItemData(cmbEvento.ListIndex) And UCase$(grdObj.TextMatrix(x, 0)) = UCase$(grdResult.TextMatrix(y, 4)) Then
             If InStr(1, grdResult.TextMatrix(y, 5), "S", vbBinaryCompare) > 0 Then
                  grdObj.TextMatrix(x, 1) = "X"
             End If
             If InStr(1, grdResult.TextMatrix(y, 5), "U", vbBinaryCompare) > 0 Then
                  grdObj.TextMatrix(x, 2) = "X"
             End If
             If InStr(1, grdResult.TextMatrix(y, 5), "I", vbBinaryCompare) > 0 Then
                  grdObj.TextMatrix(x, 3) = "X"
             End If
             If InStr(1, grdResult.TextMatrix(y, 5), "D", vbBinaryCompare) > 0 Then
                  grdObj.TextMatrix(x, 4) = "X"
             End If
             If InStr(1, grdResult.TextMatrix(y, 5), "E", vbBinaryCompare) > 0 Then
                  grdObj.TextMatrix(x, 5) = "X"
             End If
             Exit For
        End If
    Next
Next

End Sub

Private Sub cmbTela_Click()
If cmbTela.ListIndex = -1 Then Exit Sub
Sql = "SELECT NOMEFORM FROM SEG_TELASISTEMA WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
lblNomeForm.Caption = RdoAux!NomeForm
RdoAux.Close
grdResult.Rows = 1
LoadAtrib
Le
cmbTela.Enabled = False
cmdRetorna.Enabled = True
cmdCancel.Enabled = True

End Sub

Private Sub Le()
Dim sObj As String, sNomeMenu As String

Ocupado
cmbEvento.Clear
grdObj.Rows = 1
For x = 0 To lstTable.ListCount - 1
    lstTable.Selected(x) = False
Next
For x = 0 To lstProc.ListCount - 1
    lstProc.Selected(x) = False
Next
For x = 0 To lstView.ListCount - 1
    lstView.Selected(x) = False
Next
For x = 0 To lstEvento.ListCount - 1
    lstEvento.Selected(x) = False
Next
For x = 1 To tvMenu.Nodes.Count
    tvMenu.Nodes(x).Checked = False
Next

With grdResult
      For x = 1 To .Rows - 1
            
            For y = 0 To lstEvento.ListCount - 1
                   If Val(.TextMatrix(x, 0)) = cmbTela.ItemData(cmbTela.ListIndex) And Val(.TextMatrix(x, 2)) = lstEvento.ItemData(y) Then
                        lstEvento.Selected(y) = True
                        Exit For
                   End If
            Next
            sObj = .TextMatrix(x, 4)
            If Val(.TextMatrix(x, 0)) = cmbTela.ItemData(cmbTela.ListIndex) Then
                 If Left$(UCase$(sObj), 2) = "SP" Then
                      For y = 0 To lstProc.ListCount - 1
                             If UCase$(lstProc.List(y)) = UCase$(sObj) Then
                                  lstProc.Selected(y) = True
                                  Exit For
                             End If
                      Next
                 ElseIf Left$(UCase$(sObj), 2) = "VW" Then
                      For y = 0 To lstView.ListCount - 1
                             If UCase$(lstView.List(y)) = UCase$(sObj) Then
                                  lstView.Selected(y) = True
                                  Exit For
                             End If
                      Next
                 Else
                      For y = 0 To lstTable.ListCount - 1
                             If UCase$(lstTable.List(y)) = UCase$(sObj) Then
                                  lstTable.Selected(y) = True
                                  Exit For
                             End If
                      Next
                 End If
            End If
      Next
      lstEvento.ListIndex = 0
      lstProc.ListIndex = 0
      lstView.ListIndex = 0
      lstTable.ListIndex = 0
      If cmbEvento.ListCount > 0 Then cmbEvento.ListIndex = 0
End With

Sql = "SELECT NOMEMENU FROM SEG_MENUACESSO WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
       Do Until .EOF
           sNomeMenu = !NOMEMENU
            For x = 1 To tvMenu.Nodes.Count
                If tvMenu.Nodes(x).Key = sNomeMenu And tvMenu.Nodes(x).Children = 0 Then
                   tvMenu.Nodes(x).Checked = True
                   Call TreeCheckBoxes(tvMenu, tvMenu.Nodes(x))
                   Exit For
                End If
            Next
           .MoveNext
       Loop
End With

Liberado
End Sub

Private Sub LoadAtrib()

If cmbTela.ListIndex = -1 Then Exit Sub
Sql = "SELECT CODTELA,NOMETELA,CODEVENTO,DESCEVENTO,NOMEOBJETO,ATRIBSEG "
Sql = Sql & "FROM vwATRIBSEG  WHERE CODTELA = " & cmbTela.ItemData(cmbTela.ListIndex) & " ORDER BY NOMETELA,NOMEOBJETO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
        grdResult.AddItem !CODTELA & Chr(9) & !NOMETELA & Chr(9) & !CODEVENTO & Chr(9) & !DESCEVENTO & Chr(9) & !NOMEOBJETO & Chr(9) & !ATRIBSEG
       .MoveNext
    Loop
End With

End Sub

Private Sub cmdCancel_Click()
grdResult.Rows = 1
cmbTela.Enabled = True
cmdRetorna.Enabled = False
cmdCancel.Enabled = False

End Sub

Private Sub cmdRetorna_Click()

cmbTela.Enabled = True
cmdRetorna.Enabled = False
cmdCancel.Enabled = False

If grdResult.Rows = 1 Then
    If MsgBox("Deseja remover todos os atributos de segurança desta tela?", vbQuestion + vbYesNo, "atenção") = vbYes Then
       Sql = "DELETE FROM SEG_MENUACESSO WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
       cn.Execute Sql, rdExecDirect
       Sql = "DELETE FROM SEG_GRUPOACESSO WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
       cn.Execute Sql, rdExecDirect
       Sql = "DELETE FROM SEG_USERACESS WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
       cn.Execute Sql, rdExecDirect
       Sql = "DELETE FROM SEG_EVENTOACESSO WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
       cn.Execute Sql, rdExecDirect
    End If
    Exit Sub
End If
Screen.MousePointer = vbHourglass

Ocupado
Grava
Liberado
Screen.MousePointer = vbDefault
Pb.Value = 0
MsgBox "Atributos gravados com sucesso.", vbInformation, "Informação"

End Sub

Private Sub Grava()
Dim qd As New rdoQuery
Dim RdoS As rdoResultset
Dim x As Long

Set qd.ActiveConnection = cn

Sql = "DELETE FROM SEG_EVENTOACESSO WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
cn.Execute Sql, rdExecDirect

Pb.Value = 0
For x = 1 To grdResult.Rows - 1
     Pb.Value = Abs(x * 100 / grdResult.Rows - 1)
      Sql = "SELECT  ATRIBSEG FROM SEG_EVENTOACESSO WHERE "
      Sql = Sql & "CODTELA=" & grdResult.TextMatrix(x, 0) & " AND "
      Sql = Sql & "CODEVENTO=" & grdResult.TextMatrix(x, 2) & " AND "
      Sql = Sql & "NOMEOBJETO='" & grdResult.TextMatrix(x, 4) & "'"
      Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
      If RdoS.RowCount > 0 Then
            Sql = "UPDATE SEG_EVENTOACESSO SET ATRIBSEG='" & grdResult.TextMatrix(x, 5) & "',FLAG = 1 Where "
            Sql = Sql & "CODTELA=" & grdResult.TextMatrix(x, 0) & " AND CODEVENTO=" & grdResult.TextMatrix(x, 2) & " AND NOMEOBJETO = '" & UCase$(grdResult.TextMatrix(x, 4)) & "'"
      Else
           Sql = "INSERT SEG_EVENTOACESSO(CODTELA,CODEVENTO,NOMEOBJETO,ATRIBSEG,FLAG) values("
           Sql = Sql & grdResult.TextMatrix(x, 0) & "," & grdResult.TextMatrix(x, 2) & ",'"
           Sql = Sql & UCase$(grdResult.TextMatrix(x, 4)) & "','" & grdResult.TextMatrix(x, 5) & "',1)"
      End If
      cn.Execute Sql, rdExecDirect
Next
      
    
'GRAVA OS MENUS
Sql = "DELETE FROM SEG_MENUACESSO WHERE CODTELA=" & cmbTela.ItemData(cmbTela.ListIndex)
cn.Execute Sql, rdExecDirect

For x = 1 To tvMenu.Nodes.Count
      If tvMenu.Nodes(x).Checked = True Then
        Sql = "INSERT SEG_MENUACESSO (CODTELA,NOMEMENU) VALUES("
        Sql = Sql & cmbTela.ItemData(cmbTela.ListIndex) & ",'" & tvMenu.Nodes(x).Key & "')"
        cn.Execute Sql, rdExecDirect
      End If
Next
      
End Sub

Private Sub CarregaMenu()
Dim hMenu As Long, hSubMenu As Long, hSubMenu2 As Long, hSubMenu3 As Long, hSubMenu4 As Long
Dim lPosTopMenu As Long, lPos As Long, lPos2 As Long, lPos3 As Long, lPos4 As Long
Dim sName1 As String, sName2 As String, sName3 As String, sName4 As String
Dim sNome As String, NodX As Object

tvMenu.Nodes.Clear
hMenu = GetMenu(frmMdi.hwnd)
hSubMenu = GetSubMenu(hMenu, lPosTopMenu)

While hSubMenu
    sNome = GetMenuString(hMenu, lPosTopMenu)
    sName1 = GetMenuName(sNome)
    Set NodX = tvMenu.Nodes.Add(, , GetMenuName(sNome), sNome)
    While Len(GetMenuString(hSubMenu, lPos))
        sNome = GetMenuString(hSubMenu, lPos)
        sName2 = GetMenuName(sNome)
        Set NodX = tvMenu.Nodes.Add(sName1, tvwChild, GetMenuName(sNome), sNome)
        '**********
        hSubMenu2 = GetSubMenu(hSubMenu, lPos)
        While Len(GetMenuString(hSubMenu2, lPos2))
            sNome = GetMenuString(hSubMenu2, lPos2)
            sName3 = GetMenuName(sNome)
            Set NodX = tvMenu.Nodes.Add(sName2, tvwChild, GetMenuName(sNome), sNome)
            '****************
            hSubMenu3 = GetSubMenu(hSubMenu2, lPos2)
            While Len(GetMenuString(hSubMenu3, lPos3))
                 sNome = GetMenuString(hSubMenu3, lPos3)
                 sName4 = GetMenuName(sNome)
                 Set NodX = tvMenu.Nodes.Add(sName3, tvwChild, GetMenuName(sNome), sNome)
                 '************
                 hSubMenu4 = GetSubMenu(hSubMenu3, lPos3)
                 While Len(GetMenuString(hSubMenu4, lPos4))
                    sNome = GetMenuString(hSubMenu4, lPos4)
                    Set NodX = tvMenu.Nodes.Add(sName4, tvwChild, GetMenuName(sNome), sNome)
                    lPos4 = lPos4 + 1
                 Wend
                 lPos4 = 0
                 '************
                 lPos3 = lPos3 + 1
            Wend
            lPos3 = 0
            '****************
            lPos2 = lPos2 + 1
        Wend
        lPos2 = 0
        '**********
        lPos = lPos + 1
    Wend
    lPosTopMenu = lPosTopMenu + 1
    hSubMenu = GetSubMenu(hMenu, lPosTopMenu)
    lPos = 0
Wend

End Sub

Private Sub CarregaMenu2()
Dim x As Integer, NodX As Object

With frmMdi.m_cMenuPrincipal
    Set NodX = tvMenu.Nodes.Add(, , "cmdPrincipal", "Principal")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdPrincipal", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

With frmMdi.m_cMenuParam
    Set NodX = tvMenu.Nodes.Add(, , "cmdParametros", "Parâmetros")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdParametros", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

With frmMdi.m_cMenuImob
    Set NodX = tvMenu.Nodes.Add(, , "cmdImobiliario", "Imobiliário")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdImobiliario", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

With frmMdi.m_cMenuMob
    Set NodX = tvMenu.Nodes.Add(, , "cmdMobiliario", "Mobiliário")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdMobiliario", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

With frmMdi.m_cMenuAtende
    Set NodX = tvMenu.Nodes.Add(, , "cmdAtende", "Atendimento")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdAtende", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

With frmMdi.m_cMenuTrib
    Set NodX = tvMenu.Nodes.Add(, , "cmdTributo", "Tributário")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdTributo", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

With frmMdi.m_cMenuProt
    Set NodX = tvMenu.Nodes.Add(, , "cmdProtocolo", "Protocólo")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdProtocolo", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

With frmMdi.m_cMenuOutro
    Set NodX = tvMenu.Nodes.Add(, , "cmdOutros", "Outros")
    For x = 1 To .Count
        If .ItemParentIndex(x) = 0 Then
            Set NodX = tvMenu.Nodes.Add("cmdOutros", tvwChild, .ItemKey(x), .Caption(x))
        Else
            Set NodX = tvMenu.Nodes.Add(.ItemKey(.ItemParentIndex(x)), tvwChild, .ItemKey(x), .Caption(x))
        End If
    Next
End With

End Sub

Private Sub Command1_Click()
Dim x As Integer

For x = 0 To lstTable.ListCount - 1
    Sql = "GRANT SELECT,INSERT,UPDATE,DELETE ON " & lstTable.List(x) & " TO gtisys"
    cn.Execute Sql, rdExecDirect
Next
For x = 0 To lstView.ListCount - 1
    Sql = "GRANT SELECT ON " & lstView.List(x) & " TO gtisys"
    cn.Execute Sql, rdExecDirect
Next
For x = 0 To lstProc.ListCount - 1
    Sql = "GRANT EXEC ON " & lstProc.List(x) & " TO gtisys"
    cn.Execute Sql, rdExecDirect
Next

MsgBox "fim"

End Sub

Private Sub Form_Load()

cmdRetorna.Enabled = False
cmdCancel.Enabled = False
Ocupado
Screen.MousePointer = vbHourglass
Centraliza Me

Sql = "SELECT CODTELA,NOMETELA "
Sql = Sql & "FROM SEG_TELASISTEMA ORDER BY NOMETELA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
       Do Until .EOF
             cmbTela.AddItem !NOMETELA
             cmbTela.ItemData(cmbTela.NewIndex) = !CODTELA
            .MoveNext
       Loop
       cmbTela.ListIndex = -1
      .Close
End With


CarregaLista

'LoadAtrib
cmbTela_Click
Pb.Value = 0
Screen.MousePointer = vbDefault

bExec = True
Liberado

End Sub

Private Sub CarregaLista()

Dim x As Integer
On Error Resume Next
Dim NodX As Object
Dim sParent As String, LastItem As Integer
Dim n1 As Integer, n2 As Integer, n3 As Integer, d1 As String

'tvMenu.ImageList = frmMdi.ilsIcons

'Monta o Menu
CarregaMenu2


For x = 1 To tvMenu.Nodes.Count
   tvMenu.Nodes(x).EnsureVisible
Next
tvMenu.Nodes(1).EnsureVisible
'tvMenu.ExpandAll

Sql = "SELECT CODEVENTO,DESCEVENTO "
Sql = Sql & "FROM SEG_EVENTO ORDER BY DESCEVENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
       Do Until .EOF
             lstEvento.AddItem !DESCEVENTO
             lstEvento.ItemData(lstEvento.NewIndex) = !CODEVENTO
            .MoveNext
       Loop
      .Close
End With

For x = 1 To cn.rdoTables.Count
    If cn.rdoTables(x).Type = "TABLE" Then
        lstTable.AddItem cn.rdoTables(x).Name
    ElseIf cn.rdoTables(x).Type = "VIEW" Then
        If UCase(Left(cn.rdoTables(x).Name, 2)) = "VW" Then
            lstView.AddItem cn.rdoTables(x).Name
        End If
    End If
Next

lstProc.AddItem "spDADOSDEUMIMOVEL"
lstProc.AddItem "spEXTRATO"
lstProc.AddItem "spEXTRATONEW"
lstProc.AddItem "spGRAVABAIXATMP"
lstProc.AddItem "spGRAVAMOBILIARIO"
lstProc.AddItem "spGRAVAPARAMPARCELA"
lstProc.AddItem "spGRAVAPROCESSO"
lstProc.AddItem "spRELDEVEDOR"
lstProc.AddItem "spRELDEVEDORREPARCELAMENTO"

Exit Sub

d1 = "Tributacao"
'd1 = Mid(cn.Connect, n1, n3)
With oSQLServer.Databases(d1)
     For x = 1 To .Tables.Count
           If Not .Tables(x).SystemObject Then
                lstTable.AddItem .Tables(x).Name
                lstTable.ItemData(lstTable.NewIndex) = .Tables(x).ID
           End If
     Next
     For x = 1 To .StoredProcedures.Count
           If Not .StoredProcedures(x).SystemObject Then
                lstProc.AddItem .StoredProcedures(x).Name
                lstProc.ItemData(lstProc.NewIndex) = .StoredProcedures(x).ID
           End If
     Next
     For x = 1 To .Views.Count
           If Not .Views(x).SystemObject And .Views(x).Name <> "View2" Then
                lstView.AddItem .Views(x).Name
                lstView.ItemData(lstView.NewIndex) = .Views(x).ID
           End If
     Next
End With
d1 = "protocolo"
'd1 = Mid(cn.Connect, n1, n3)
With oSQLServer.Databases(d1)
     For x = 1 To .Tables.Count
           If Not .Tables(x).SystemObject Then
                lstTable.AddItem .Tables(x).Name
                lstTable.ItemData(lstTable.NewIndex) = .Tables(x).ID
           End If
     Next
     For x = 1 To .StoredProcedures.Count
           If Not .StoredProcedures(x).SystemObject Then
                lstProc.AddItem .StoredProcedures(x).Name
                lstProc.ItemData(lstProc.NewIndex) = .StoredProcedures(x).ID
           End If
     Next
     For x = 1 To .Views.Count
           If Not .Views(x).SystemObject Then
                lstView.AddItem .Views(x).Name
                lstView.ItemData(lstView.NewIndex) = .Views(x).ID
           End If
     Next
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oSQLServer = Nothing
' Pressing the End button or selecting End from the Run menu without
' unhooking causes an Invalid Page Fault and closes Microsoft
' Visual Basic. Call procedure to stop intercepting the messages for
' this window
'Unhook

End Sub

Private Sub grdObj_Click()
If cmbEvento.ListIndex = -1 Then
     MsgBox "Selecione um Evento.", vbExclamation, "Atenção"
     cmbEvento.SetFocus
     Exit Sub
End If

With grdObj
     grdObj_EnterCell
   If .TextMatrix(Lin, col) = "" Then
         If Left$(.TextMatrix(Lin, 0), 2) = "sp" Then
              If col = 5 Then
                  .TextMatrix(Lin, col) = "X"
                   AdicionaAcesso
              End If
         Else
              If col < 5 Then
                   .TextMatrix(Lin, col) = "X"
                   AdicionaAcesso
              End If
         End If
    Else
        .TextMatrix(Lin, col) = ""
         RemoveAcesso
    End If
End With

End Sub

Private Sub tvMenu_NodeCheck(ByVal Node As MSComctlLib.Node)

Call TreeCheckBoxes(tvMenu, Node)

End Sub

Public Sub TreeCheckBoxes(TR As TreeView, CurrentNode As Node)
    'This code is copyright (c)2002 by Scott
    '     Durrett - All Rights Reserved
    'No changes are allow without written ap
    '     proval from the Author.
    Dim liNodeIndex As Integer
    Dim lbDirty As Boolean
    Dim lParentNode As Node
    Dim lChildNode As Node
    lbDirty = False
    liNodeIndex = CurrentNode.Index


    If CurrentNode.Checked = True Then 'node is checked
        'Children Check/UnCheck


        If Not TR.Nodes.Item(CurrentNode.Index).Child Is Nothing Then
            Set lParentNode = TR.Nodes.Item(liNodeIndex).Child.FirstSibling


            Do While Not lParentNode Is Nothing
                lParentNode.Checked = CurrentNode.Checked


                If Not lParentNode.Child Is Nothing Then
                    Set lChildNode = lParentNode.Child


                    Do While Not lChildNode Is Nothing
                        lChildNode.Checked = CurrentNode.Checked


                        If Not lChildNode.Next Is Nothing Then
                            Set lChildNode = lChildNode.Next
                        Else
                            Set lChildNode = lChildNode.Child
                        End If
                    Loop
                End If
                Set lParentNode = lParentNode.Next
            Loop
        End If
        '=======================================
        '     =====================
        'Check all parent nodes


        Do While Not TR.Nodes.Item(liNodeIndex).Parent Is Nothing
            TR.Nodes.Item(liNodeIndex).Parent.Checked = CurrentNode.Checked
            liNodeIndex = TR.Nodes.Item(liNodeIndex).Parent.Index
        Loop
        '===========================
    ElseIf CurrentNode.Checked = False Then 'node is unchecked
        'Children Check/UnCheck


        If Not TR.Nodes.Item(CurrentNode.Index).Child Is Nothing Then
            Set lParentNode = TR.Nodes.Item(liNodeIndex).Child.FirstSibling


            Do While Not lParentNode Is Nothing
                lParentNode.Checked = CurrentNode.Checked


                If Not lParentNode.Child Is Nothing Then
                    Set lChildNode = lParentNode.Child


                    Do While Not lChildNode Is Nothing
                        lChildNode.Checked = CurrentNode.Checked


                        If Not lChildNode.Next Is Nothing Then
                            Set lChildNode = lChildNode.Next
                        Else
                            Set lChildNode = lChildNode.Child
                        End If
                    Loop
                End If
                Set lParentNode = lParentNode.Next
            Loop
        End If
        '=======================================
        '     =====================
        Set lParentNode = Nothing
        Set lChildNode = Nothing


        If Not CurrentNode.Parent Is Nothing Then
            Set lParentNode = CurrentNode.Parent.Child


            Do While Not lParentNode Is Nothing
                Set lChildNode = lParentNode.FirstSibling


                Do While Not lChildNode Is Nothing


                    If lChildNode.Checked = True Then
                        lbDirty = True
                        Exit Do
                    End If
                    'If Not lChildNode.Next Is Nothing Then
                    Set lChildNode = lChildNode.Next
                    'End If
                Loop


                If lbDirty = False Then


                    If Not lParentNode.Parent Is Nothing Then
                        lParentNode.Parent.Checked = False
                        lbDirty = False
                    End If
                Else
                    Exit Do
                End If


                If Not lParentNode.Parent Is Nothing Then
                    Set lParentNode = lParentNode.Parent
                Else
                    Set lParentNode = lParentNode.Parent
                End If
            Loop
        End If
    End If
    Set CurrentNode = Nothing
    Set lParentNode = Nothing
    Set lChildNode = Nothing
End Sub

Private Sub RemoveAcesso()
Dim LinhaUpdate As Integer

With grdResult
        For x = 1 To .Rows - 1
              If UCase$(.TextMatrix(x, 4)) = UCase$(grdObj.TextMatrix(Lin, 0)) And .TextMatrix(x, 2) = cmbEvento.ItemData(cmbEvento.ListIndex) Then
                   Exit For
              End If
        Next
        
        LinhaUpdate = x
        
        If col = 1 Then
            RemoveLetra LinhaUpdate, "S"
        ElseIf col = 2 Then
            RemoveLetra LinhaUpdate, "U"
        ElseIf col = 3 Then
            RemoveLetra LinhaUpdate, "I"
        ElseIf col = 4 Then
            RemoveLetra LinhaUpdate, "D"
        ElseIf col = 5 Then
            RemoveLetra LinhaUpdate, "E"
        End If
        If .TextMatrix(LinhaUpdate, 5) = "" Then
             If .Rows > 2 Then
                  .RemoveItem LinhaUpdate
             Else
                  .Rows = 1
             End If
        End If

End With

End Sub

Private Sub RemoveLetra(nLin As Integer, sLetra As String)
Dim sPalavra As String
Dim sNovaPalavra As String
On Error Resume Next
sPalavra = grdResult.TextMatrix(nLin, 5)
sNovaPalavra = ""
For x = 1 To Len(sPalavra)
       If Mid(sPalavra, x, 1) <> sLetra Then
            sNovaPalavra = sNovaPalavra & Mid(sPalavra, x, 1)
       End If
Next
grdResult.TextMatrix(nLin, 5) = sNovaPalavra
        
End Sub

Private Sub AdicionaAcesso()
Dim Existe As Boolean
Dim LinhaUpdate As Integer

Existe = False
With grdResult
        For x = 1 To .Rows - 1
              If Val(.TextMatrix(x, 0)) = cmbTela.ItemData(cmbTela.ListIndex) And .TextMatrix(x, 4) = grdObj.TextMatrix(Lin, 0) And .TextMatrix(x, 2) = cmbEvento.ItemData(cmbEvento.ListIndex) Then
                   Existe = True
                   Exit For
              End If
        Next
        
        If Not Existe Then
             .AddItem cmbTela.ItemData(cmbTela.ListIndex) & Chr(9) & cmbTela.Text & Chr(9) & cmbEvento.ItemData(cmbEvento.ListIndex) & Chr(9) & cmbEvento.Text & Chr(9) & grdObj.TextMatrix(Lin, 0)
              LinhaUpdate = .Rows - 1
        Else
              LinhaUpdate = x
        End If
        
        If col = 1 Then
            .TextMatrix(LinhaUpdate, 5) = .TextMatrix(LinhaUpdate, 5) & "S"
        ElseIf col = 2 Then
            .TextMatrix(LinhaUpdate, 5) = .TextMatrix(LinhaUpdate, 5) & "U"
        ElseIf col = 3 Then
            .TextMatrix(LinhaUpdate, 5) = .TextMatrix(LinhaUpdate, 5) & "I"
        ElseIf col = 4 Then
            .TextMatrix(LinhaUpdate, 5) = .TextMatrix(LinhaUpdate, 5) & "D"
        ElseIf col = 5 Then
            .TextMatrix(LinhaUpdate, 5) = .TextMatrix(LinhaUpdate, 5) & "E"
        End If
End With

End Sub

Private Sub grdObj_EnterCell()
Lin = grdObj.Row
col = grdObj.col
End Sub

Private Sub lstEvento_ItemCheck(Item As Integer)
On Error Resume Next
Dim DelItem As Integer
bExec = False
If lstEvento.Selected(Item) Then
     cmbEvento.AddItem lstEvento.Text
     cmbEvento.ItemData(cmbEvento.NewIndex) = lstEvento.ItemData(Item)
Else
    For x = 0 To cmbEvento.ListCount - 1
          cmbEvento.ListIndex = x
          If cmbEvento.Text = lstEvento.Text Then
               DelItem = cmbEvento.ItemData(cmbEvento.ListIndex)
               cmbEvento.RemoveItem (x)
               Exit For
          End If
     Next
     If DelItem > 0 Then
        For x = 1 To grdResult.Rows - 1
            If grdResult.TextMatrix(x, 2) = DelItem Then
               If grdResult.Rows > 2 Then
                  grdResult.RemoveItem (x)
               Else
                  grdResult.Rows = 1
               End If
'               Exit For
            End If
        Next
     End If
End If
bExec = True
End Sub

Private Sub lstProc_ItemCheck(Item As Integer)
If lstProc.Selected(Item) Then
     grdObj.AddItem lstProc.Text
Else
    For x = 1 To grdObj.Rows - 1
          If grdObj.TextMatrix(x, 0) = lstProc.Text Then
               If grdObj.Rows > 2 Then
                    grdObj.RemoveItem (x)
               Else
                    grdObj.Rows = 1
               End If
               Exit For
          End If
    Next
End If

End Sub

Private Sub lstTable_ItemCheck(Item As Integer)

If Not lstTable.Selected(Item) Then
     For x = 1 To grdResult.Rows - 1
           If Val(grdResult.TextMatrix(x, 0)) = cmbTela.ItemData(cmbTela.ListIndex) And UCase$(grdResult.TextMatrix(x, 4)) = UCase$(lstTable.Text) Then
                MsgBox "Exclua todos os acessos ao Objeto antes de remove-lo.", vbExclamation, "Atenção"
                lstTable.Selected(Item) = True
                Exit Sub
           End If
     Next
End If

If lstTable.Selected(Item) Then
     grdObj.AddItem lstTable.Text
Else
    For x = 1 To grdObj.Rows - 1
          If grdObj.TextMatrix(x, 0) = lstTable.Text Then
               If grdObj.Rows > 2 Then
                    grdObj.RemoveItem (x)
               Else
                    grdObj.Rows = 1
               End If
               Exit For
          End If
    Next
End If

End Sub

Private Sub lstView_ItemCheck(Item As Integer)
If lstView.Selected(Item) Then
     grdObj.AddItem lstView.Text
Else
    For x = 1 To grdObj.Rows - 1
          If grdObj.TextMatrix(x, 0) = lstView.Text Then
               If grdObj.Rows > 2 Then
                    grdObj.RemoveItem (x)
               Else
                    grdObj.Rows = 1
               End If
               Exit For
          End If
    Next
End If
End Sub

Private Function GetMenuString(ByVal hMenu As Long, ByVal POS As Long) As String
Dim sBuf As String

    sBuf = String(100, 0)
    sBuf = Left$(sBuf, GetMenuStringA(hMenu, POS, sBuf, Len(sBuf), MF_BYPOSITION))
    GetMenuString = sBuf
End Function

Private Function GetMenuName(sMenuCaption As String) As String
Dim Ct As Control

For Each Ct In frmMdi.Controls
    If TypeOf Ct Is Menu Then
        If Ct.Caption = sMenuCaption Then
            Exit For
        End If
    End If
Next
GetMenuName = Ct.Name
End Function
