VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHistSenha 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico senhas"
   ClientHeight    =   3555
   ClientLeft      =   9720
   ClientTop       =   2130
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvNF 
      Height          =   3540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   6244
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Senha"
         Object.Width           =   1412
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Guiche"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Hora Cham."
         Object.Width           =   2187
      EndProperty
   End
End
Attribute VB_Name = "frmHistSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim itmX As ListItem, z As Long
Ocupado
z = SendMessage(lvNF.hwnd, LVM_DELETEALLITEMS, 0, 0)
On Error Resume Next
Sql = "SELECT senha, guiche,datachamada,horachamada From sspac WHERE dataentrada='" & Format(Now, "mm/dd/yyyy") & " ' ORDER BY senha"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Set itmX = lvNF.ListItems.Add(, , Format(!SENHA, "000"))
       itmX.SubItems(1) = Format(!GUICHE, "00")
       If Not IsNull(!datachamada) Then
        'itmX.SubItems(2) = Format(!datachamada, "dd/mm/yyyy")
        itmX.SubItems(2) = !horachamada
        End If
      .MoveNext
    Loop
   .Close
End With
Liberado

End Sub
