VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAlertaAnexo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2715
   ClientLeft      =   4800
   ClientTop       =   5640
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   8835
   ShowInTaskbar   =   0   'False
   Begin Tributacao.jcFrames pnlAnexo 
      Height          =   2715
      Left            =   0
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   4789
      FrameColor      =   192
      BackColor       =   128
      FillColor       =   64
      Caption         =   "O processo que você está recebendo contêm o(s) anexo(s) abaixo, você confirma o recebimento?"
      TextBoxHeight   =   18
      TextColor       =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ThemeColor      =   5
      ColorFrom       =   255
      ColorTo         =   255
      Begin MSComctlLib.ListView lvAnexo 
         Height          =   1695
         Left            =   60
         TabIndex        =   0
         Top             =   420
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   2990
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Processo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Numero"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Assunto"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Requerente"
            Object.Width           =   4306
         EndProperty
      End
      Begin prjChameleon.chameleonButton cmdConfirma 
         Height          =   345
         Left            =   6060
         TabIndex        =   1
         ToolTipText     =   "Confirma o recebimento do(s) anexo(s)"
         Top             =   2250
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Confirmar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmAlertaAnexo.frx":0000
         PICN            =   "frmAlertaAnexo.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNaoConfirma 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   7350
         TabIndex        =   2
         ToolTipText     =   "Não confirma o recebimento dos anexos"
         Top             =   2250
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "&Cancelar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmAlertaAnexo.frx":0176
         PICN            =   "frmAlertaAnexo.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
End
Attribute VB_Name = "frmAlertaAnexo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNaoConfirma_Click()
Unload Me
End Sub
