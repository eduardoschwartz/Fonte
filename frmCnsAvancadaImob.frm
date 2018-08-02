VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmCnsAvancadaImob 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta avançada de imóveis e gerador de correspondência"
   ClientHeight    =   5985
   ClientLeft      =   2970
   ClientTop       =   2730
   ClientWidth     =   11400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5985
   ScaleWidth      =   11400
   Begin Tributacao.jcFrames fr2 
      Height          =   5415
      Left            =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   9551
      FrameColor      =   8388608
      TextBoxColor    =   11595760
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Lista das empresas selecionadas pelos critérios"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.Frame frDDList1 
         BackColor       =   &H00EEEEEE&
         Height          =   375
         Left            =   2925
         TabIndex        =   25
         Top             =   315
         Width           =   2580
         Begin VB.ListBox lstDDList1 
            Height          =   3210
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   26
            Top             =   405
            Width           =   2490
         End
         Begin prjChameleon.chameleonButton cmdDDList1 
            Height          =   240
            Left            =   315
            TabIndex        =   27
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0000
            PICN            =   "frmCnsAvancadaImob.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDD1All 
            Height          =   240
            Left            =   1710
            TabIndex        =   28
            ToolTipText     =   "Selecionar todos"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "+"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0176
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdDD1None 
            Height          =   240
            Left            =   2070
            TabIndex        =   29
            ToolTipText     =   "Manter apenas o código"
            Top             =   0
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   423
            BTYPE           =   3
            TX              =   "-"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   12632256
            BCOLO           =   12632256
            FCOL            =   12582912
            FCOLO           =   12582912
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0192
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
      Begin vbAcceleratorSGrid6.vbalGrid grdMain 
         Height          =   4470
         Left            =   90
         TabIndex        =   30
         Top             =   855
         Width           =   11130
         _ExtentX        =   19632
         _ExtentY        =   7885
         NoHorizontalGridLines=   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         HighlightForeColor=   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   2
         DisableIcons    =   -1  'True
         GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
      End
      Begin prjChameleon.chameleonButton cmdGroup 
         Height          =   315
         Left            =   9900
         TabIndex        =   62
         ToolTipText     =   "Avançar para a próxima tela"
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Agrupar"
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
         MICON           =   "frmCnsAvancadaImob.frx":01AE
         PICN            =   "frmCnsAvancadaImob.frx":01CA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Campos a serem exibidos..:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   31
         Top             =   405
         Width           =   2760
      End
   End
   Begin Tributacao.jcFrames fr1 
      Height          =   5415
      Left            =   45
      Top             =   45
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   9551
      FrameColor      =   8388608
      TextBoxColor    =   11595760
      Style           =   2
      RoundedCornerTxtBox=   -1  'True
      Caption         =   "Seleção de Critérios"
      TextBoxHeight   =   18
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.Frame frUT 
         BackColor       =   &H00EEEEEE&
         Caption         =   "     Uso do terreno "
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   180
         TabIndex        =   7
         Top             =   630
         Width           =   1905
         Begin VB.CheckBox chkUT 
            BackColor       =   &H00EEEEEE&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   50
            Top             =   10
            Width           =   195
         End
         Begin VB.ListBox lstUT 
            Height          =   1635
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   8
            Top             =   360
            Width           =   1815
         End
         Begin prjChameleon.chameleonButton cmdUT 
            Height          =   240
            Left            =   1530
            TabIndex        =   9
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0324
            PICN            =   "frmCnsAvancadaImob.frx":0340
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame frPE 
         BackColor       =   &H00EEEEEE&
         Caption         =   "     Pedologia"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   180
         TabIndex        =   44
         Top             =   2430
         Width           =   1905
         Begin VB.CheckBox chkPE 
            BackColor       =   &H00EEEEEE&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   55
            Top             =   10
            Width           =   195
         End
         Begin VB.ListBox lstPE 
            Height          =   1635
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   45
            Top             =   360
            Width           =   1815
         End
         Begin prjChameleon.chameleonButton cmdPE 
            Height          =   240
            Left            =   1530
            TabIndex        =   46
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":049A
            PICN            =   "frmCnsAvancadaImob.frx":04B6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame frSI 
         BackColor       =   &H00EEEEEE&
         Caption         =   "     Situação"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   180
         TabIndex        =   41
         Top             =   2070
         Width           =   1905
         Begin VB.CheckBox chkSI 
            BackColor       =   &H00EEEEEE&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   54
            Top             =   10
            Width           =   195
         End
         Begin VB.ListBox lstSI 
            Height          =   1635
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   42
            Top             =   360
            Width           =   1815
         End
         Begin prjChameleon.chameleonButton cmdSI 
            Height          =   240
            Left            =   1530
            TabIndex        =   43
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0610
            PICN            =   "frmCnsAvancadaImob.frx":062C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame frCP 
         BackColor       =   &H00EEEEEE&
         Caption         =   "     Categ.Prop."
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   180
         TabIndex        =   38
         Top             =   1710
         Width           =   1905
         Begin VB.CheckBox chkCP 
            BackColor       =   &H00EEEEEE&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   53
            Top             =   10
            Width           =   195
         End
         Begin VB.ListBox lstCP 
            Height          =   1635
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   39
            Top             =   360
            Width           =   1815
         End
         Begin prjChameleon.chameleonButton cmdCP 
            Height          =   240
            Left            =   1530
            TabIndex        =   40
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0786
            PICN            =   "frmCnsAvancadaImob.frx":07A2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame frTO 
         BackColor       =   &H00EEEEEE&
         Caption         =   "     Topografia"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   180
         TabIndex        =   35
         Top             =   1350
         Width           =   1905
         Begin VB.CheckBox chkTO 
            BackColor       =   &H00EEEEEE&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   52
            Top             =   10
            Width           =   195
         End
         Begin VB.ListBox lstTO 
            Height          =   1635
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   36
            Top             =   360
            Width           =   1815
         End
         Begin prjChameleon.chameleonButton cmdTO 
            Height          =   240
            Left            =   1530
            TabIndex        =   37
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":08FC
            PICN            =   "frmCnsAvancadaImob.frx":0918
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame frBE 
         BackColor       =   &H00EEEEEE&
         Caption         =   "     Benfeitoria"
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   180
         TabIndex        =   32
         Top             =   990
         Width           =   1905
         Begin VB.CheckBox chkBE 
            BackColor       =   &H00EEEEEE&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   90
            TabIndex        =   51
            Top             =   10
            Width           =   195
         End
         Begin VB.ListBox lstBE 
            Height          =   1635
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   33
            Top             =   360
            Width           =   1815
         End
         Begin prjChameleon.chameleonButton cmdBE 
            Height          =   240
            Left            =   1530
            TabIndex        =   34
            ToolTipText     =   "Exibir Lista"
            Top             =   0
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   423
            BTYPE           =   14
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   0   'False
            BCOL            =   14869218
            BCOLO           =   14869218
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0A72
            PICN            =   "frmCnsAvancadaImob.frx":0A8E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   -1  'True
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Importação de Arquivos"
         ForeColor       =   &H00800000&
         Height          =   1455
         Left            =   90
         TabIndex        =   10
         Top             =   3375
         Width           =   4110
         Begin VB.TextBox txtArq 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Height          =   285
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   12
            Top             =   315
            Width           =   3390
         End
         Begin VB.TextBox txtDelimiter 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1035
            MaxLength       =   1
            TabIndex        =   11
            Text            =   ","
            Top             =   675
            Width           =   330
         End
         Begin prjChameleon.chameleonButton cmdOpen 
            Height          =   315
            Left            =   3555
            TabIndex        =   13
            ToolTipText     =   "Localizar arquivo texto"
            Top             =   315
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0BE8
            PICN            =   "frmCnsAvancadaImob.frx":0C04
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdImportar 
            Height          =   315
            Left            =   1395
            TabIndex        =   14
            ToolTipText     =   "Importar o arquivo selecionado"
            Top             =   675
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   "Importar"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0C8B
            PICN            =   "frmCnsAvancadaImob.frx":0CA7
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdLimpar 
            Height          =   315
            Left            =   3555
            TabIndex        =   15
            ToolTipText     =   "Limpar texto"
            Top             =   675
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":0EBA
            PICN            =   "frmCnsAvancadaImob.frx":0ED6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdPreview 
            Height          =   315
            Left            =   3105
            TabIndex        =   16
            ToolTipText     =   "Visualizar arquivo"
            Top             =   675
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":10F3
            PICN            =   "frmCnsAvancadaImob.frx":110F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Delimitador.:"
            Height          =   240
            Index           =   0
            Left            =   90
            TabIndex        =   19
            Top             =   735
            Width           =   870
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Imóveis localizados no arquivo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   18
            Top             =   1125
            Width           =   2895
         End
         Begin VB.Label lblTotImp 
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   3285
            TabIndex        =   17
            Top             =   1125
            Width           =   420
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Critérios Diversos"
         ForeColor       =   &H00800000&
         Height          =   2940
         Left            =   90
         TabIndex        =   6
         Top             =   315
         Width           =   4110
         Begin VB.ComboBox cmbQuadra 
            Height          =   315
            ItemData        =   "frmCnsAvancadaImob.frx":1269
            Left            =   3060
            List            =   "frmCnsAvancadaImob.frx":126B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   2520
            Width           =   825
         End
         Begin VB.ComboBox cmbSetor 
            Height          =   315
            ItemData        =   "frmCnsAvancadaImob.frx":126D
            Left            =   1710
            List            =   "frmCnsAvancadaImob.frx":126F
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   2520
            Width           =   780
         End
         Begin VB.ComboBox cmbDist 
            Height          =   315
            ItemData        =   "frmCnsAvancadaImob.frx":1271
            Left            =   450
            List            =   "frmCnsAvancadaImob.frx":1281
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   2520
            Width           =   780
         End
         Begin VB.ComboBox cmbIA 
            Height          =   315
            ItemData        =   "frmCnsAvancadaImob.frx":1293
            Left            =   2160
            List            =   "frmCnsAvancadaImob.frx":1295
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   1125
            Width           =   1725
         End
         Begin VB.ComboBox cmbTI 
            Height          =   315
            Left            =   2160
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   765
            Width           =   1725
         End
         Begin VB.ComboBox cmbFI 
            Height          =   315
            ItemData        =   "frmCnsAvancadaImob.frx":1297
            Left            =   2160
            List            =   "frmCnsAvancadaImob.frx":1299
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   405
            Width           =   1725
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Q.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   4
            Left            =   2700
            TabIndex        =   58
            Top             =   2565
            Width           =   240
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "S.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   3
            Left            =   1395
            TabIndex        =   57
            Top             =   2565
            Width           =   240
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "D.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   56
            Top             =   2565
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00EEEEEE&
         ForeColor       =   &H00000080&
         Height          =   4515
         Left            =   4275
         TabIndex        =   63
         Top             =   315
         Width           =   6945
         Begin VB.Frame frCond 
            BackColor       =   &H00EEEEEE&
            Caption         =   "     Condomínios"
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   135
            TabIndex        =   78
            Top             =   1350
            Width           =   6675
            Begin VB.CheckBox chkCond 
               BackColor       =   &H00EEEEEE&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   80
               Top             =   10
               Width           =   195
            End
            Begin VB.ListBox lstCond 
               Height          =   1860
               Left            =   45
               Style           =   1  'Checkbox
               TabIndex        =   79
               Top             =   360
               Width           =   6585
            End
            Begin prjChameleon.chameleonButton cmdCond 
               Height          =   240
               Left            =   1620
               TabIndex        =   81
               ToolTipText     =   "Exibir Lista"
               Top             =   0
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   423
               BTYPE           =   14
               TX              =   ""
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   0   'False
               BCOL            =   14869218
               BCOLO           =   14869218
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   13026246
               MPTR            =   1
               MICON           =   "frmCnsAvancadaImob.frx":129B
               PICN            =   "frmCnsAvancadaImob.frx":12B7
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   0   'False
            End
            Begin prjChameleon.chameleonButton cmdCondAll 
               Height          =   240
               Left            =   2430
               TabIndex        =   82
               ToolTipText     =   "Selecionar todos"
               Top             =   0
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   423
               BTYPE           =   3
               TX              =   "+"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   0   'False
               BCOL            =   12632256
               BCOLO           =   12632256
               FCOL            =   12582912
               FCOLO           =   12582912
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmCnsAvancadaImob.frx":1411
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prjChameleon.chameleonButton cmdCondNone 
               Height          =   240
               Left            =   2790
               TabIndex        =   83
               ToolTipText     =   "Manter apenas o código"
               Top             =   0
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   423
               BTYPE           =   3
               TX              =   "-"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   0   'False
               BCOL            =   12632256
               BCOLO           =   12632256
               FCOL            =   12582912
               FCOLO           =   12582912
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmCnsAvancadaImob.frx":142D
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
         Begin VB.Frame frBairro 
            BackColor       =   &H00EEEEEE&
            Caption         =   "     Bairros "
            ForeColor       =   &H00000080&
            Height          =   330
            Left            =   135
            TabIndex        =   72
            Top             =   945
            Width           =   6675
            Begin VB.ListBox lstBairro 
               Height          =   1860
               Left            =   45
               Style           =   1  'Checkbox
               TabIndex        =   74
               Top             =   360
               Width           =   6585
            End
            Begin VB.CheckBox chkBairro 
               BackColor       =   &H00EEEEEE&
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   90
               TabIndex        =   73
               Top             =   10
               Width           =   195
            End
            Begin prjChameleon.chameleonButton cmdBairro 
               Height          =   240
               Left            =   1620
               TabIndex        =   75
               ToolTipText     =   "Exibir Lista"
               Top             =   0
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   423
               BTYPE           =   14
               TX              =   ""
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   0   'False
               BCOL            =   14869218
               BCOLO           =   14869218
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   13026246
               MPTR            =   1
               MICON           =   "frmCnsAvancadaImob.frx":1449
               PICN            =   "frmCnsAvancadaImob.frx":1465
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   -1  'True
               VALUE           =   0   'False
            End
            Begin prjChameleon.chameleonButton cmdBairroAll 
               Height          =   240
               Left            =   2430
               TabIndex        =   76
               ToolTipText     =   "Selecionar todos"
               Top             =   0
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   423
               BTYPE           =   3
               TX              =   "+"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   0   'False
               BCOL            =   12632256
               BCOLO           =   12632256
               FCOL            =   12582912
               FCOLO           =   12582912
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmCnsAvancadaImob.frx":15BF
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin prjChameleon.chameleonButton cmdBairroNone 
               Height          =   240
               Left            =   2790
               TabIndex        =   77
               ToolTipText     =   "Manter apenas o código"
               Top             =   0
               Width           =   315
               _ExtentX        =   556
               _ExtentY        =   423
               BTYPE           =   3
               TX              =   "-"
               ENAB            =   0   'False
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   2
               FOCUSR          =   0   'False
               BCOL            =   12632256
               BCOLO           =   12632256
               FCOL            =   12582912
               FCOLO           =   12582912
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "frmCnsAvancadaImob.frx":15DB
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
         Begin prjChameleon.chameleonButton cmdDelProp 
            Height          =   315
            Left            =   6390
            TabIndex        =   70
            ToolTipText     =   "Limpar Campo"
            Top             =   585
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":15F7
            PICN            =   "frmCnsAvancadaImob.frx":1613
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prjChameleon.chameleonButton cmdBuscaProp 
            Height          =   315
            Left            =   5940
            TabIndex        =   71
            ToolTipText     =   "Busca proprietário"
            Top             =   585
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            BTYPE           =   5
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   0   'False
            BCOL            =   15658734
            BCOLO           =   15658734
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   13026246
            MPTR            =   1
            MICON           =   "frmCnsAvancadaImob.frx":1830
            PICN            =   "frmCnsAvancadaImob.frx":184C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtProp 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Height          =   285
            Left            =   1530
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   585
            Width           =   4350
         End
         Begin VB.TextBox txtNomeLogr 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   65
            Top             =   225
            Width           =   4350
         End
         Begin VB.TextBox txtCodLogr 
            Appearance      =   0  'Flat
            BackColor       =   &H00EEEEEE&
            Height          =   285
            Left            =   5940
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   225
            Width           =   855
         End
         Begin VB.ListBox lstNomeLog 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            Height          =   1590
            ItemData        =   "frmCnsAvancadaImob.frx":19A6
            Left            =   1530
            List            =   "frmCnsAvancadaImob.frx":19A8
            TabIndex        =   66
            Top             =   225
            Visible         =   0   'False
            Width           =   4380
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Proprietário:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   135
            TabIndex        =   69
            Top             =   615
            Width           =   1410
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Logradouro..:"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   67
            Top             =   270
            Width           =   1365
         End
      End
   End
   Begin Tributacao.jcFrames tb2 
      Height          =   465
      Left            =   45
      Top             =   5490
      Visible         =   0   'False
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   820
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin prjChameleon.chameleonButton cmdVoltar 
         Height          =   315
         Left            =   9990
         TabIndex        =   20
         ToolTipText     =   "Voltar a tela anterior"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Voltar"
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
         MICON           =   "frmCnsAvancadaImob.frx":19AA
         PICN            =   "frmCnsAvancadaImob.frx":19C6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdTxt 
         Height          =   315
         Left            =   5670
         TabIndex        =   21
         ToolTipText     =   "Gerar em arquivo TXT"
         Top             =   90
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Gerar em Excel"
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
         MICON           =   "frmCnsAvancadaImob.frx":1B20
         PICN            =   "frmCnsAvancadaImob.frx":1B3C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdEtiqueta 
         Height          =   315
         Left            =   1980
         TabIndex        =   22
         ToolTipText     =   "Gerar etiquetas para mala direta"
         Top             =   90
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Gerar Etiquetas"
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
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCnsAvancadaImob.frx":1BC9
         PICN            =   "frmCnsAvancadaImob.frx":1BE5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdTela 
         Height          =   315
         Left            =   135
         TabIndex        =   23
         ToolTipText     =   "Gerar etiquetas para mala direta"
         Top             =   90
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Gerar na Tela"
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
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnsAvancadaImob.frx":1C96
         PICN            =   "frmCnsAvancadaImob.frx":1CB2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCarta 
         Height          =   315
         Left            =   3825
         TabIndex        =   24
         ToolTipText     =   "Gerar etiquetas para mala direta"
         Top             =   90
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Gerar Cartas"
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
         FOCUSR          =   -1  'True
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmCnsAvancadaImob.frx":1D52
         PICN            =   "frmCnsAvancadaImob.frx":1D6E
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
   Begin Tributacao.jcFrames tb1 
      Height          =   465
      Left            =   45
      Top             =   5490
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   820
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin prjChameleon.chameleonButton cmdConsultar 
         Height          =   315
         Left            =   7380
         TabIndex        =   0
         ToolTipText     =   "Consultar as empresas"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Pesquisar"
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
         MICON           =   "frmCnsAvancadaImob.frx":1DD0
         PICN            =   "frmCnsAvancadaImob.frx":1DEC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdNext 
         Height          =   315
         Left            =   8685
         TabIndex        =   1
         ToolTipText     =   "Avançar para a próxima tela"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Continuar"
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
         MICON           =   "frmCnsAvancadaImob.frx":1F1B
         PICN            =   "frmCnsAvancadaImob.frx":1F37
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Tributacao.XP_ProgressBar PBar 
         Height          =   240
         Left            =   4320
         TabIndex        =   2
         Top             =   135
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   12632064
         Scrolling       =   1
         ShowText        =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdSair 
         Height          =   315
         Left            =   9990
         TabIndex        =   3
         ToolTipText     =   "Sair da Tela"
         Top             =   90
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Sair"
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
         MICON           =   "frmCnsAvancadaImob.frx":2091
         PICN            =   "frmCnsAvancadaImob.frx":20AD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblTot 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   2835
         TabIndex        =   5
         Top             =   90
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Números de imóveis:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         TabIndex        =   4
         Top             =   90
         Width           =   2805
      End
   End
   Begin VB.Menu mnuTipoEtiqueta 
      Caption         =   "Tipo de Etiqueta"
      Visible         =   0   'False
      Begin VB.Menu mnuEtNormal 
         Caption         =   "Endereço de Entrega"
      End
      Begin VB.Menu mnuEtNot 
         Caption         =   "Endereço do Imóvel"
      End
   End
End
Attribute VB_Name = "frmCnsAvancadaImob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xImovel As clsImovel
Private Type Imovel
    nCodigo As Long
    sProprietario As String
    nProprietario As Long
    sLogradouro As String
    nNumero As Integer
    sComplemento  As String
    sBairro As String
    sUsoTerreno As String
    sBenfeitoria As String
    sPedologia As String
    sTopografia As String
    sSituacao As String
    sCategProp As String
    nFracaoIdeal As Double
    sAtivo As String
    sInscricao As String
    nDistrito As Integer
    nSetor As Integer
    nQuadra As Integer
    nLote As Integer
    nFace As Integer
    sTipoImovel As String
    nAreaPredial As Double
    nAreaTerreno As Double
    nTipoEndereco As Integer
    sLogradouroEnt As String
    nNumeroEnt As Integer
    sBairroEnt As String
    sCidadeEnt As String
    sUFEnt As String
    sCEPEnt As String
    sComplementoEnt As String
    sQuadras As String
    sLotes As String
    sNomeCondominio As String
    nVVP As Double
    nVVT As Double
    nVVI As Double
    nCodLogr As Long
    nMatricula As Long
    nTranscricao As Long
End Type

Private Type Area
    nCodigo As Long
    nArea As Double
End Type

Dim aCodigos() As Imovel, aCodigosImp() As Long, strCodigos As String
Dim aCodigoComArea() As Long, aArea() As Area

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If
Me.Refresh
If cGetInputState() <> 0 Then DoEvents
End Sub

Private Sub Chkcond_Click()
cmdCond.Enabled = chkCond.value
cmdCondAll.Enabled = chkCond.value
cmdCondNone.Enabled = chkCond.value
If chkCond.value = vbChecked Then
    frCond.Height = 2265
    cmdCond.value = True
    frCond.ZOrder 0
Else
    frCond.Height = 330
End If

End Sub

Private Sub chkBairro_Click()
cmdBairro.Enabled = chkBairro.value
cmdBairroAll.Enabled = chkBairro.value
cmdBairroNone.Enabled = chkBairro.value
If chkBairro.value = vbChecked Then
    frBairro.Height = 2265
    cmdBairro.value = True
    frBairro.ZOrder 0
Else
    frBairro.Height = 330
End If

End Sub

Private Sub chkBE_Click()
HideList
cmdBE.Enabled = chkBE.value
If chkBE.value = vbChecked Then
    frBE.Height = 2040
    cmdBE.value = True
    frBE.ZOrder 0
Else
    frBE.Height = 330
End If
End Sub

Private Sub chkCP_Click()
HideList
cmdCP.Enabled = chkCP.value
If chkCP.value = vbChecked Then
    frCP.Height = 2040
    cmdCP.value = True
    frCP.ZOrder 0
Else
    frCP.Height = 330
End If
End Sub

Private Sub chkPE_Click()
HideList
cmdPE.Enabled = chkPE.value
If chkPE.value = vbChecked Then
    frPE.Height = 2040
    cmdPE.value = True
    frPE.ZOrder 0
Else
    frPE.Height = 330
End If
End Sub

Private Sub chkSI_Click()
HideList
cmdSI.Enabled = chkSI.value
If chkSI.value = vbChecked Then
    frSI.Height = 2040
    cmdSI.value = True
    frSI.ZOrder 0
Else
    frSI.Height = 330
End If
End Sub

Private Sub chkTO_Click()
HideList
cmdTO.Enabled = chkTO.value
If chkTO.value = vbChecked Then
    frTO.Height = 2040
    cmdTO.value = True
    frTO.ZOrder 0
Else
    frTO.Height = 330
End If
End Sub

Private Sub chkUT_Click()
HideList
cmdUT.Enabled = chkUT.value
If chkUT.value = vbChecked Then
    frUT.Height = 2040
    cmdUT.value = True
    frUT.ZOrder 0
Else
    frUT.Height = 330
End If
End Sub

Private Sub cmbDist_Click()
Dim RdoAux As rdoResultset, Sql As String
cmbSetor.Clear: cmbQuadra.Clear

If cmbDist.ListIndex < 1 Then Exit Sub
cmbSetor.AddItem " "
Sql = "SELECT CODSETOR FROM SETOR WHERE CODDISTRITO=" & Val(cmbDist.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbSetor.AddItem Format(!CODSETOR, "00")
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmbSetor_Click()
Dim RdoAux As rdoResultset, Sql As String
cmbQuadra.Clear

If cmbSetor.ListIndex < 1 Then Exit Sub
cmbQuadra.AddItem " "
Sql = "SELECT DISTINCT CODQUADRA FROM FACEQUADRA WHERE CODDISTRITO=" & Val(cmbDist.Text) & " AND CODSETOR=" & Val(cmbSetor.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbQuadra.AddItem Format(!CODQUADRA, "0000")
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdBairro_Click()
If cmdBairro.value = True Then
    frBairro.Height = 2265
    frBairro.ZOrder 0
Else
    frBairro.Height = 330
End If
End Sub

Private Sub cmdBairroAll_Click()
Dim x As Integer
For x = 0 To lstBairro.ListCount - 1
    lstBairro.Selected(x) = True
Next
End Sub

Private Sub cmdBairroNone_Click()
Dim x As Integer
For x = 0 To lstBairro.ListCount - 1
    lstBairro.Selected(x) = False
Next
lstBairro.Selected(0) = True
End Sub

Private Sub cmdBE_Click()
If cmdBE.value = True Then
    frBE.Height = 2040
    frBE.ZOrder 0
Else
    frBE.Height = 330
End If
End Sub

Private Sub cmdBuscaProp_Click()
Dim frm As Object
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frmCnsCidadao.show
End Sub

Private Sub cmdCond_Click()
If cmdCond.value = True Then
    frCond.Height = 2265
    frCond.ZOrder 0
Else
    frCond.Height = 330
End If

End Sub

Private Sub cmdCondAll_Click()
Dim x As Integer
For x = 0 To lstCond.ListCount - 1
    lstCond.Selected(x) = True
Next

End Sub

Private Sub cmdCondNone_Click()
Dim x As Integer
For x = 0 To lstCond.ListCount - 1
    lstCond.Selected(x) = False
Next

End Sub

Private Sub cmdCP_Click()
If cmdCP.value = True Then
    frCP.Height = 2040
    frCP.ZOrder 0
Else
    frCP.Height = 330
End If
End Sub

Private Sub cmdDDList1_Click()
If cmdDDList1.value = True Then
    frDDList1.Height = 3660
Else
    frDDList1.Height = 375
    HideColumns
End If
End Sub

Private Sub cmdDelProp_Click()
txtProp.Text = ""
End Sub

Private Sub cmdEtiqueta_Click()
PopupMenu mnuTipoEtiqueta, tb1.Top, cmdEtiqueta.Left

End Sub

Private Sub cmdGroup_Click()
grdMain.AllowGrouping = cmdGroup.value
End Sub

Private Sub cmdImportar_Click()
Dim strLinha As String, z As Variant, x As Integer, nCodigo As Long
lblTotImp.Caption = 0
If txtDelimiter.Text = "" Then
    MsgBox "Especifique um delimitador", vbCritical, "Erro"
    Exit Sub
End If
If txtArq.Text = "" Then
    MsgBox "Selecione um arquivo", vbCritical, "Erro"
    Exit Sub
End If


ReDim aCodigosImp(0): strCodigos = ""
Open txtArq.Text For Input As #1
   Do While Not EOF(1)
        Line Input #1, strLinha
        z = Split(strLinha, txtDelimiter.Text)
        For x = 0 To UBound(z)
            If Not IsNumeric(z(x)) Then
               GoTo proximo
            End If
            nCodigo = CLng(z(x))
            If nCodigo > 100000 Then
               GoTo Erro
            End If
            ReDim Preserve aCodigosImp(UBound(aCodigosImp) + 1)
            aCodigosImp(UBound(aCodigosImp)) = nCodigo
            strCodigos = strCodigos & nCodigo & ","
        Next
proximo:
   Loop
Close #1
strCodigos = Chomp(strCodigos, chomp_righT, 1)
lblTotImp.Caption = UBound(aCodigosImp)

Exit Sub
Erro:
MsgBox "Arquivo inválido !!!", vbCritical, "Erro de importação"
Close #1

End Sub

Private Sub cmdLimpar_Click()
txtArq.Text = "": lblTotImp.Caption = 0
End Sub

Private Sub cmdNext_Click()
If Val(lblTot.Caption) > 0 Then
    If Not Valida Then Exit Sub
    cmdNext.Enabled = False
    Ocupado
    If cGetInputState() <> 0 Then DoEvents
    CarregaCampos
    CarregaLista
    HideColumns
    Liberado
    tb2.Visible = True
    tb1.Visible = False
    fr2.Visible = True
    fr1.Visible = False
    grdMain.SetFocus
    grdMain.SelectedRow = 1
    cmdNext.Enabled = True
Else
    MsgBox "Nenhum imóvel possui os critérios selecionados ou não foi gerada consulta.", vbExclamation, "Atenção"
End If

End Sub

Private Sub cmdOpen_Click()
Dim fName As String, cc As cCommonDlg

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Documento de Texto|*.txt;*.csv|Todos os Arquivos|*.*", , App.Path & "\Bin", "Selecione um arquivo texto", , Me.hwnd, OFN_HIDEREADONLY, False
txtArq.Text = fName
End Sub

Private Sub cmdPE_Click()
If cmdPE.value = True Then
    frPE.Height = 2040
    frPE.ZOrder 0
Else
    frPE.Height = 330
End If
End Sub

Private Sub cmdPreview_Click()
If (txtArq.Text) <> "" Then
    z = Shell(App.Path & "\NOTEPAD2" & " " & txtArq.Text, vbNormalFocus)
End If
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSI_Click()
If cmdSI.value = True Then
    frSI.Height = 2040
    frSI.ZOrder 0
Else
    frSI.Height = 330
End If
End Sub

Private Sub cmdTela_Click()
Dim ax As String, z As Long, x As Integer, y As Integer

Open sPathBin & "\TEMPIMOB.TXT" For Output As #1

With grdMain
    ax = ""
    For y = 1 To .Columns
        If .ColumnVisible(y) = True Then
            ax = ax & FillSpace(.ColumnHeader(y), Val(Right(.ColumnKey(y), 3))) & vbTab
        End If
    Next
    Print #1, ax
    For x = 1 To .Rows
        ax = ""
        For y = 1 To .Columns
            If .ColumnVisible(y) = True Then
                ax = ax & FillSpace(.cell(x, y).Text, Val(Right(.ColumnKey(y), 3))) & vbTab
            End If
        Next
        Print #1, ax
    Next

End With

Close #1

z = Shell(App.Path & "\NOTEPAD2" & " " & sPathBin & "\TEMPIMOB.TXT", vbNormalFocus)
End Sub

Private Sub cmdTO_Click()
If cmdTO.value = True Then
    frTO.Height = 2040
    frTO.ZOrder 0
Else
    frTO.Height = 330
End If
End Sub

Private Sub cmdTxt_Click()
'Dim ax As String, z As Long, x As Integer, y As Integer, sChar As String
'
'If txtSep.Text = "" Then
'    sChar = " "
'Else
'    sChar = txtSep.Text
'End If
'Ocupado
'Open sPathBin & "\RELATIMOB.TXT" For Output As #1
'
'With grdMain
'    ax = ""
'    For y = 1 To .Columns
'        If .ColumnVisible(y) = True Then
'            ax = ax & FillSpace(.ColumnHeader(y), Val(Right(.ColumnKey(y), 3))) & sChar
'        End If
'    Next
'    ax = Chomp(ax, chomp_righT, 1)
'    Print #1, ax
'    For x = 1 To .Rows
'        If cGetInputState() <> 0 Then DoEvents
'        ax = ""
'        For y = 1 To .Columns
'            If .ColumnVisible(y) = True Then
'                ax = ax & FillSpace(.cell(x, y).Text, Val(Right(.ColumnKey(y), 3))) & sChar
'            End If
'        Next
'        If sChar <> " " Then ax = Chomp(ax, chomp_righT, 1)
'        Print #1, ax
'    Next
'
'End With
'
'Close #1
'Liberado
'MsgBox "O arquivo foi salvo em " & sPathBin & "\RELATIMOB.TXT"

Dim x As Long, y As Long, ax As String, Scr_hdc As Long, z As Long
Dim cnExcel As ADODB.Connection, Rs As ADODB.Recordset, nCont As Integer, sFile As String
Scr_hdc = GetDesktopWindow()
         
Set cnExcel = New ADODB.Connection
sFile = "Rel" & Format(Now, "ddmmyyyyhhmmss") & ".xls"
cnExcel.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" & sPathBin & "\" & sFile & "; Extended Properties=""Excel 8.0;HDR=YES"""
cnExcel.Open

ax = ""
For y = 1 To grdMain.Columns
    If grdMain.ColumnVisible(y) = True Then
        ax = ax & RemoveSpace(grdMain.ColumnHeader(y)) & " char(255), "
    End If
Next
ax = Left(ax, Len(ax) - 2)
cnExcel.Execute "Create Table Table1(" & ax & ")"

Set Rs = New ADODB.Recordset
Rs.Open "[Table1$]", cnExcel, adOpenDynamic, adLockOptimistic, adCmdTable


For x = 1 To grdMain.Rows
    Rs.AddNew
    nCont = 0
    For y = 1 To grdMain.Columns
        If grdMain.ColumnVisible(y) = True Then
            Rs.Fields(nCont).value = grdMain.cell(x, y).Text
            nCont = nCont + 1
        End If
        
    Next
    Rs.Update
Next


 cnExcel.Close
Set Rs = Nothing
Set cnExcel = Nothing

z = ShellExecute(Scr_hdc, "Open", sFile, "", sPathBin, SW_SHOWNORMAL)

End Sub

Private Sub cmdUT_Click()
If cmdUT.value = True Then
    frUT.Height = 2040
    frUT.ZOrder 0
Else
    frUT.Height = 330
End If
End Sub

Private Sub cmdVoltar_Click()
tb2.Visible = False
tb1.Visible = True
fr2.Visible = False
fr1.Visible = True
End Sub

Private Sub Form_Load()
Centraliza Me

Set xImovel = New clsImovel
PBar.Color = vbWhite
Init
GridHeader
End Sub

Private Sub Init()
Dim RdoAux As rdoResultset, Sql As String, x As Integer
Ocupado
DoEvents
Sql = "SELECT * FROM USOTERRENO WHERE CODUSOTERRENO<>999 ORDER BY DESCUSOTERRENO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstUT.AddItem !DescUsoTerreno
        lstUT.ItemData(lstUT.NewIndex) = !CODUSOTERRENO
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstUT.ListCount - 1: lstUT.Selected(x) = True: Next
lstUT.ListIndex = 0

Sql = "SELECT * FROM TOPOGRAFIA WHERE CODTOPOGRAFIA<>999 ORDER BY DESCTOPOGRAFIA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstTO.AddItem !DescTopografia
        lstTO.ItemData(lstTO.NewIndex) = !CODTOPOGRAFIA
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstTO.ListCount - 1: lstTO.Selected(x) = True: Next
lstTO.ListIndex = 0

Sql = "SELECT * FROM BENFEITORIA WHERE CODBENFEITORIA<>999 ORDER BY DESCBENFEITORIA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstBE.AddItem !DescBenfeitoria
        lstBE.ItemData(lstBE.NewIndex) = !CODBENFEITORIA
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstBE.ListCount - 1: lstBE.Selected(x) = True: Next
lstBE.ListIndex = 0

Sql = "SELECT * FROM PEDOLOGIA WHERE CODPEDOLOGIA<>999 ORDER BY DESCPEDOLOGIA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstPE.AddItem !DescPedologia
        lstPE.ItemData(lstPE.NewIndex) = !CODPEDOLOGIA
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstPE.ListCount - 1: lstPE.Selected(x) = True: Next
lstPE.ListIndex = 0

Sql = "SELECT * FROM CATEGPROP WHERE CODCATEGPROP<>999 ORDER BY DESCCATEGPROP"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstCP.AddItem !DescCategProp
        lstCP.ItemData(lstCP.NewIndex) = !CODCATEGPROP
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstCP.ListCount - 1: lstCP.Selected(x) = True: Next
lstCP.ListIndex = 0

Sql = "SELECT * FROM SITUACAO WHERE CODSITUACAO<>999 ORDER BY DESCSITUACAO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstSI.AddItem !DescSituacao
        lstSI.ItemData(lstSI.NewIndex) = !Codsituacao
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstSI.ListCount - 1: lstSI.Selected(x) = True: Next
lstSI.ListIndex = 0

Sql = "SELECT * FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND CODBAIRRO<>999 ORDER BY DESCBAIRRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        If !DescBairro <> "" Then
            lstBairro.AddItem !DescBairro
            lstBairro.ItemData(lstBairro.NewIndex) = !CodBairro
        End If
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstBairro.ListCount - 1: lstBairro.Selected(x) = True: Next
lstBairro.ListIndex = 0

Sql = "SELECT CD_CODIGO,CD_NOMECOND FROM CONDOMINIO ORDER BY CD_NOMECOND"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        lstCond.AddItem !cd_nomecond
        lstCond.ItemData(lstCond.NewIndex) = !CD_CODIGO
       .MoveNext
    Loop
   .Close
End With
For x = 0 To lstCond.ListCount - 1: lstCond.Selected(x) = True: Next
lstCond.ListIndex = 0


cmbFI.AddItem "Com/Sem Fração"
cmbFI.AddItem "Somente com fração"
cmbFI.AddItem "Somente sem fração"
cmbFI.ListIndex = 0

cmbTI.AddItem "Predial/Territorial"
cmbTI.AddItem "Somente predial"
cmbTI.AddItem "Somente territorial"
cmbTI.ListIndex = 0

cmbIA.AddItem "Ativo/Inativo"
cmbIA.AddItem "Somente ativo"
cmbIA.AddItem "Somente inativo"
cmbIA.ListIndex = 1

ReDim aCodigoComArea(0)
Sql = "SELECT DISTINCT codreduzido From Areas ORDER BY CODREDUZIDO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aCodigoComArea(UBound(aCodigoComArea) + 1)
        aCodigoComArea(UBound(aCodigoComArea)) = !CODREDUZIDO
       .MoveNext
    Loop
   .Close
End With

ReDim aArea(0)
Sql = "SELECT codreduzido, SUM(areaconstr) AS SOMAAREA From Areas "
Sql = Sql & "GROUP BY codreduzido ORDER BY codreduzido"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
        ReDim Preserve aArea(UBound(aArea) + 1)
        aArea(UBound(aArea)).nCodigo = !CODREDUZIDO
        aArea(UBound(aArea)).nArea = !SOMAAREA
       .MoveNext
    Loop
   .Close
End With
Liberado
End Sub

Private Sub cmdConsultar_Click()
Dim Sql As String, RdoAux As rdoResultset, nTotal As Long, RdoAux2 As rdoResultset, cn2 As rdoConnection, nTot As Long, nPos As Long
Dim s As Integer, lResult As Long, sUT As String, sBE As String, sCP As String, sTO As String, sSI As String, sPE As String, sBairro As String
Dim sTipoImovel As String, nArea As Double, x As Integer, nTipoEnd As Integer, sLogradouro As String, nNumero As Integer, sCidade As String
Dim sBairroE As String, sUF As String, sCep As String, sCompl As String, sCondominio As String

If Not Valida Then Exit Sub

HideList
lblTot.Caption = 0
'Set cn2 = en.OpenConnection(dsname:="odbcTributacao", Prompt:=rdDriverNoPrompt, Connect:="uid=" & UL & ";PWD=" & UP & ";driver={SQL Server};")
   Conn$ = "UID=gtisys;PWD=everest;" _
    & "DATABASE=tributacao;" _
    & "SERVER=" & IPServer & ";" _
    & "DRIVER={SQL SERVER};DSN='';"
    Set cn2 = en.OpenConnection(dsname:="", Prompt:=rdDriverNoPrompt, Connect:=Conn$)

Ocupado
DoEvents

nTotal = 0: ReDim aCodigos(0): nPos = 1
Sql = "SELECT DISTINCT codreduzido, Ativo, INSCRICAO, nomecidadao, CPF, CNPJ, rg, LOGRADOURO, li_num, li_compl, descbairro, li_quadras, li_lotes, li_codbairro, codlogr, inativo,"
Sql = Sql & "dt_areaterreno, dt_codusoterreno, dt_codbenf, dt_codtopog, dt_codcategprop, dt_codsituacao, dt_codpedol, dt_numagua, dt_fracaoideal, dc_qtdeedif, dc_qtdepav,"
Sql = Sql & "ee_tipoend, distrito, setor, quadra, lote, seq, unidade, subunidade, li_uf, li_codcidade, descbenfeitoria, descusoterreno, desctopografia, desccategprop, descsituacao,"
Sql = Sql & "descpedologia, codcidadao, NOMELOGRADOURO2, abrevtitlog, abrevtipolog, nomelogradouro, numimovel, complemento, DESCBAIRROP, siglauf, codlogradouro,"
Sql = Sql & "ee_codlog, ee_nomelog, ee_numimovel, ee_complemento, BairroEE, CidadeEE, ee_uf, ee_cep, ee_descbairro, AbrevTipoLogEE, AbrevTitLogEE, cd_nomecond,"
Sql = Sql & "codcondominio, datainclusao, codagrupa, desccidade, ee_cidade, ee_bairro, nomecidade, codcidade, cep, telefone, email, desccidade, codbairro, nummat,"
Sql = Sql & "TipoMat FROM vwFULLIMOVEL2 WHERE "
If Val(lblTotImp.Caption) = 0 Then
    Sql = Sql & "CODREDUZIDO < 100000 "
Else
    Sql = Sql & "CODREDUZIDO in (" & strCodigos & ")"
End If

If chkUT.value = vbChecked Then
    sUT = ""
    For x = 0 To lstUT.ListCount - 1
        If lstUT.Selected(x) = True Then
            sUT = sUT & lstUT.ItemData(x) & ","
        End If
    Next
    If sUT = "" Then
        Liberado
        MsgBox "Selecione ao menos um uso do terreno.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sUT = Chomp(sUT, chomp_righT, 1)
    Sql = Sql & " AND DT_CODUSOTERRENO in (" & sUT & ")"
End If

If chkBE.value = vbChecked Then
    sBE = ""
    For x = 0 To lstBE.ListCount - 1
        If lstBE.Selected(x) = True Then
            sBE = sBE & lstBE.ItemData(x) & ","
        End If
    Next
    If sBE = "" Then
        Liberado
        MsgBox "Selecione ao menos uma benfeitoria.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sBE = Chomp(sBE, chomp_righT, 1)
    Sql = Sql & " AND DT_CODBENF in (" & sBE & ")"
End If

If chkPE.value = vbChecked Then
    sPE = ""
    For x = 0 To lstPE.ListCount - 1
        If lstPE.Selected(x) = True Then
            sPE = sPE & lstPE.ItemData(x) & ","
        End If
    Next
    If sPE = "" Then
        Liberado
        MsgBox "Selecione ao menos uma pedologia.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sPE = Chomp(sPE, chomp_righT, 1)
    Sql = Sql & " AND DT_CODPEDOL in (" & sPE & ")"
End If

If chkSI.value = vbChecked Then
    sSI = ""
    For x = 0 To lstSI.ListCount - 1
        If lstSI.Selected(x) = True Then
            sSI = sSI & lstSI.ItemData(x) & ","
        End If
    Next
    If sSI = "" Then
        Liberado
        MsgBox "Selecione ao menos uma situação.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sSI = Chomp(sSI, chomp_righT, 1)
    Sql = Sql & " AND DT_CODSITUACAO in (" & sSI & ")"
End If

If chkCP.value = vbChecked Then
    sCP = ""
    For x = 0 To lstCP.ListCount - 1
        If lstCP.Selected(x) = True Then
            sCP = sCP & lstCP.ItemData(x) & ","
        End If
    Next
    If sCP = "" Then
        Liberado
        MsgBox "Selecione ao menos uma categoria da propriedade.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sCP = Chomp(sCP, chomp_righT, 1)
    Sql = Sql & " AND DT_CODCATEGPROP in (" & sCP & ")"
End If

If chkTO.value = vbChecked Then
    sTO = ""
    For x = 0 To lstTO.ListCount - 1
        If lstTO.Selected(x) = True Then
            sTO = sTO & lstTO.ItemData(x) & ","
        End If
    Next
    If sTO = "" Then
        Liberado
        MsgBox "Selecione ao menos uma topografia.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sTO = Chomp(sTO, chomp_righT, 1)
    Sql = Sql & " AND DT_CODTOPOG in (" & sTO & ")"
End If

If cmbFI.ListIndex > 0 Then
    If cmbFI.ListIndex = 1 Then
        Sql = Sql & " AND DT_FRACAOIDEAL>0"
    Else
        Sql = Sql & " AND DT_FRACAOIDEAL=0"
    End If
End If

If cmbIA.ListIndex > 0 Then
    If cmbIA.ListIndex = 1 Then
        Sql = Sql & " AND INATIVO=0"
    Else
        Sql = Sql & " AND INATIVO=1"
    End If
End If

If Val(cmbDist.Text) > 0 Then
    Sql = Sql & " AND DISTRITO=" & Val(cmbDist.Text)
End If

If Val(cmbSetor.Text) > 0 Then
    Sql = Sql & " AND SETOR=" & Val(cmbSetor.Text)
End If

If Val(cmbQuadra.Text) > 0 Then
    Sql = Sql & " AND QUADRA=" & Val(cmbQuadra.Text)
End If

If txtProp.Text <> "" Then
    Sql = Sql & " AND CODCIDADAO=" & Val(Left(txtProp.Text, 6))
End If

If Val(txtCodLogr.Text) > 0 Then
    Sql = Sql & " AND CODLOGR=" & Val(txtCodLogr.Text)
End If

If chkBairro.value = vbChecked Then
    sBairro = ""
    For x = 0 To lstBairro.ListCount - 1
        If lstBairro.Selected(x) = True Then
            sBairro = sBairro & lstBairro.ItemData(x) & ","
        End If
    Next
    If sBairro = "" Then
        Liberado
        MsgBox "Selecione ao menos um bairro.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sBairro = Chomp(sBairro, chomp_righT, 1)
    Sql = Sql & " AND LI_CODBAIRRO in (" & sBairro & ")"
End If

If chkCond.value = vbChecked Then
    sCondominio = ""
    For x = 0 To lstCond.ListCount - 1
        If lstCond.Selected(x) = True Then
            sCondominio = sCondominio & lstCond.ItemData(x) & ","
        End If
    Next
    If sCondominio = "" Then
        Liberado
        MsgBox "Selecione ao menos um condomínio.", vbCritical, "Erro de validação"
        Exit Sub
    End If
    sCondominio = Chomp(sCondominio, chomp_righT, 1)
    Sql = Sql & " AND CODCONDOMINIO in (" & sCondominio & ")"
End If

Sql = Sql & " ORDER BY CODREDUZIDO"
cmdConsultar.Enabled = False
Ocupado
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        If nPos Mod 50 = 0 Then
            CallPb nPos, nTot
        End If
        lResult = BinarySearchLong(aCodigoComArea(), !CODREDUZIDO)
        If lResult = -1 Then
            sTipoImovel = "Territorial"
            nArea = 0
        Else
            sTipoImovel = "Predial"
            nArea = aArea(lResult).nArea
        End If
        If cmbTI.ListIndex > 0 Then
            If (cmbTI.ListIndex = 1 And lResult = -1) Or (cmbTI.ListIndex = 2 And lResult > -1) Then
                GoTo proximo
            End If
        End If
        
        'ENDEREÇO DE ENTREGA
        nTipoEnd = !Ee_TipoEnd
        If nTipoEnd = 0 Then 'Endereço do imóvel
            xImovel.RetornaEndereco !CODREDUZIDO, Imobiliario, Localizacao
        ElseIf nTipoEnd = 1 Then 'Endereço do prop
            xImovel.RetornaEndereco !CODREDUZIDO, Imobiliario, cadastrocidadao
        ElseIf nTipoEnd = 2 Then 'Endereço de entrega
            xImovel.RetornaEndereco !CODREDUZIDO, Imobiliario, Entrega
        End If
        
        sLogradouro = xImovel.Endereco
        nNumero = Val(SubNull(xImovel.Numero))
        sCompl = xImovel.Complemento
        sBairroE = xImovel.Bairro
        sCidade = xImovel.Cidade
        sCep = Format(xImovel.Cep, "00000-000")
        sUF = SubNull(xImovel.UF)

        ReDim Preserve aCodigos(UBound(aCodigos) + 1)
        aCodigos(UBound(aCodigos)).nCodigo = !CODREDUZIDO
        aCodigos(UBound(aCodigos)).sProprietario = SubNull(!nomecidadao)
        aCodigos(UBound(aCodigos)).nProprietario = Val(SubNull(!CodCidadao))
        aCodigos(UBound(aCodigos)).sLogradouro = !Logradouro
        aCodigos(UBound(aCodigos)).nNumero = !Li_Num
        aCodigos(UBound(aCodigos)).sComplemento = SubNull(!Li_Compl)
        aCodigos(UBound(aCodigos)).sBairro = SubNull(!DescBairro)
        aCodigos(UBound(aCodigos)).sUsoTerreno = !DescUsoTerreno
        aCodigos(UBound(aCodigos)).sBenfeitoria = !DescBenfeitoria
        aCodigos(UBound(aCodigos)).sCategProp = !DescCategProp
        aCodigos(UBound(aCodigos)).sSituacao = !DescSituacao
        aCodigos(UBound(aCodigos)).sTopografia = !DescTopografia
        aCodigos(UBound(aCodigos)).sPedologia = !DescPedologia
        aCodigos(UBound(aCodigos)).nFracaoIdeal = !Dt_FracaoIdeal
        aCodigos(UBound(aCodigos)).sAtivo = IIf(!Inativo = True, "Não", "Sim")
        aCodigos(UBound(aCodigos)).sInscricao = !Inscricao
        aCodigos(UBound(aCodigos)).nDistrito = !Distrito
        aCodigos(UBound(aCodigos)).nSetor = !Setor
        aCodigos(UBound(aCodigos)).nQuadra = !Quadra
        aCodigos(UBound(aCodigos)).nLote = !Lote
        aCodigos(UBound(aCodigos)).nFace = !Seq
        aCodigos(UBound(aCodigos)).sTipoImovel = sTipoImovel
        aCodigos(UBound(aCodigos)).nAreaPredial = nArea
        aCodigos(UBound(aCodigos)).nAreaTerreno = !Dt_AreaTerreno
        aCodigos(UBound(aCodigos)).nTipoEndereco = nTipoEnd
        aCodigos(UBound(aCodigos)).sLogradouroEnt = sLogradouro
        aCodigos(UBound(aCodigos)).nNumeroEnt = nNumero
        aCodigos(UBound(aCodigos)).sComplementoEnt = sCompl
        aCodigos(UBound(aCodigos)).sBairroEnt = sBairroE
        aCodigos(UBound(aCodigos)).sCidadeEnt = sCidade
        aCodigos(UBound(aCodigos)).sUFEnt = sUF
        aCodigos(UBound(aCodigos)).sCEPEnt = sCep
        aCodigos(UBound(aCodigos)).sQuadras = SubNull(!Li_Quadras)
        aCodigos(UBound(aCodigos)).sLotes = SubNull(!Li_Lotes)
        aCodigos(UBound(aCodigos)).sNomeCondominio = SubNull(!cd_nomecond)
        aCodigos(UBound(aCodigos)).nCodLogr = !CodLogr
        If IsNull(!TipoMat) Then
            aCodigos(UBound(aCodigos)).nMatricula = 0
            aCodigos(UBound(aCodigos)).nTranscricao = 0
        Else
            If !TipoMat = "M" Then
                aCodigos(UBound(aCodigos)).nMatricula = SubNull(!NumMat)
                aCodigos(UBound(aCodigos)).nTranscricao = 0
            Else
                aCodigos(UBound(aCodigos)).nMatricula = 0
                aCodigos(UBound(aCodigos)).nTranscricao = SubNull(!NumMat)
            End If
        End If
        
        Sql = "SELECT VVT,VVC,VVI FROM LASERIPTU WHERE ANO=" & Year(Now) & " AND CODREDUZIDO=" & !CODREDUZIDO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If .RowCount > 0 Then
                aCodigos(UBound(aCodigos)).nVVT = !VVT
                aCodigos(UBound(aCodigos)).nVVP = !vvc
                aCodigos(UBound(aCodigos)).nVVI = !VVI
            Else
                aCodigos(UBound(aCodigos)).nVVT = 0
                aCodigos(UBound(aCodigos)).nVVP = 0
                aCodigos(UBound(aCodigos)).nVVI = 0
            End If
           .Close
        End With
        
        nTotal = nTotal + 1
proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
    lblTot.Caption = nTotal
   .Close
End With
PBar.value = 0: PBar.Color = vbWhite
cn2.Close
cmdConsultar.Enabled = True
Liberado
If nTotal = 0 Then
    MsgBox "Nenhum imóvel possui os critérios especificados.", vbInformation, "Informação"
End If
End Sub

Private Function Valida() As Boolean
Valida = True
End Function

Private Sub CarregaCampos()
Dim x As Integer

With lstDDList1
    .Clear
    .AddItem "Codigo"
    .ItemData(.NewIndex) = 0
    .AddItem "Proprietário"
    .ItemData(.NewIndex) = 1
    .AddItem "Logradouro"
    .ItemData(.NewIndex) = 2
    .AddItem "Número"
    .ItemData(.NewIndex) = 3
    .AddItem "Complemento"
    .ItemData(.NewIndex) = 4
    .AddItem "Bairro"
    .ItemData(.NewIndex) = 5
    .AddItem "Uso do Terreno"
    .ItemData(.NewIndex) = 6
    .AddItem "Categ.Propriedade"
    .ItemData(.NewIndex) = 7
    .AddItem "Topografia"
    .ItemData(.NewIndex) = 8
    .AddItem "Situação"
    .ItemData(.NewIndex) = 9
    .AddItem "Benfeitoria"
    .ItemData(.NewIndex) = 10
    .AddItem "Pedologia"
    .ItemData(.NewIndex) = 11
    .AddItem "Fração Ideal"
    .ItemData(.NewIndex) = 12
    .AddItem "Ativo"
    .ItemData(.NewIndex) = 13
    .AddItem "Inscrição"
    .ItemData(.NewIndex) = 14
    .AddItem "Distrito"
    .ItemData(.NewIndex) = 15
    .AddItem "Setor"
    .ItemData(.NewIndex) = 16
    .AddItem "Quadra"
    .ItemData(.NewIndex) = 17
    .AddItem "Lote"
    .ItemData(.NewIndex) = 18
    .AddItem "face"
    .ItemData(.NewIndex) = 19
    .AddItem "Tipo de imóvel"
    .ItemData(.NewIndex) = 20
    .AddItem "Área Predial"
    .ItemData(.NewIndex) = 21
    .AddItem "Área Territorial"
    .ItemData(.NewIndex) = 22
    .AddItem "Tipo de Endereço"
    .ItemData(.NewIndex) = 23
    .AddItem "Logradouro de Entrega"
    .ItemData(.NewIndex) = 24
    .AddItem "Número de Entrega"
    .ItemData(.NewIndex) = 25
    .AddItem "Complemento de Entrega"
    .ItemData(.NewIndex) = 26
    .AddItem "Bairro de Entrega"
    .ItemData(.NewIndex) = 27
    .AddItem "Cidade de Entrega"
    .ItemData(.NewIndex) = 28
    .AddItem "UF de Entrega"
    .ItemData(.NewIndex) = 29
    .AddItem "CEP de Entrega"
    .ItemData(.NewIndex) = 30
    .AddItem "Quadra Original"
    .ItemData(.NewIndex) = 31
    .AddItem "Lote Original"
    .ItemData(.NewIndex) = 32
    .AddItem "Nome do Condomínio"
    .ItemData(.NewIndex) = 33
    .AddItem "V.V.Territorial"
    .ItemData(.NewIndex) = 34
    .AddItem "V.V.Predial"
    .ItemData(.NewIndex) = 35
    .AddItem "V.V.Imóvel"
    .ItemData(.NewIndex) = 36
    .AddItem "Código Logradouro"
    .ItemData(.NewIndex) = 37
    .Selected(0) = True
    .Selected(1) = True
    .Selected(2) = True
    .Selected(3) = True
    .Selected(5) = True
    .Selected(13) = True
    .ItemData(.NewIndex) = 38
    .AddItem "Matrícula"
    .ItemData(.NewIndex) = 39
    .AddItem "Transcrição"
End With

'For X = 0 To lstDDList1.ListCount - 1
'    cmbCampos.AddItem lstDDList1.List(X)
'Next
'cmbCampos.ListIndex = 0

End Sub

Private Sub GridHeader()

With grdMain
    .HeaderFlat = True
    .HeaderHeight = 18
    .DefaultRowHeight = 17
    .GridFillLineColor = vbWhite
    .RowMode = True
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
        
    .AddColumn "kCodigo006", "Código", ecgHdrTextALignCentre, , 50
    .AddColumn "kProprietario080", "Proprietário", ecgHdrTextALignLeft, , 200
    .AddColumn "kLogradouro080", "Logradouro", ecgHdrTextALignLeft, , 200
    .AddColumn "kNumero004", "Num", ecgHdrTextALignRight, , 40
    .AddColumn "kComplemento040", "Complemento", ecgHdrTextALignLeft, , 100
    .AddColumn "kBairro040", "Bairro", ecgHdrTextALignLeft, , 100
    .AddColumn "kUsoTerreno025", "UsoTerreno", ecgHdrTextALignLeft, , 100
    .AddColumn "kCategProp025", "Cat.Prop.", ecgHdrTextALignLeft, , 100
    .AddColumn "kTopografia025", "Topografia", ecgHdrTextALignLeft, , 100
    .AddColumn "kSituacao025", "Situacao", ecgHdrTextALignLeft, , 100
    .AddColumn "kBenfeitoria025", "Benfeitoria", ecgHdrTextALignLeft, , 100
    .AddColumn "kPedologia025", "Pedologia", ecgHdrTextALignLeft, , 100
    .AddColumn "kFracao010", "Fr.ideal", ecgHdrTextALignRight, , 50
    .AddColumn "kAtivo004", "Ativo", ecgHdrTextALignCentre, , 40
    .AddColumn "kInscricao030", "Inscrição", ecgHdrTextALignLeft, , 150
    .AddColumn "kDistrito004", "Distrito", ecgHdrTextALignCentre, , 50
    .AddColumn "kSetor004", "Setor", ecgHdrTextALignCentre, , 50
    .AddColumn "kQuadra004", "Quadra", ecgHdrTextALignCentre, , 50
    .AddColumn "kLote005", "Lote", ecgHdrTextALignCentre, , 50
    .AddColumn "kFace004", "Face", ecgHdrTextALignCentre, , 50
    .AddColumn "kTipo012", "Tipo", ecgHdrTextALignLeft, , 60
    .AddColumn "kArea012", "Área Predial", ecgHdrTextALignRight, , 70
    .AddColumn "kAreaT012", "Área Terreno", ecgHdrTextALignRight, , 80
    .AddColumn "kTE004", "TEnd", ecgHdrTextALignCentre, , 40
    .AddColumn "kLogradouroEnt080", "Logradouro Entrega", ecgHdrTextALignLeft, , 200
    .AddColumn "kNumeroEnt004", "Num", ecgHdrTextALignRight, , 40
    .AddColumn "kComplementoEnt040", "Complem.Ent", ecgHdrTextALignLeft, , 100
    .AddColumn "kBairroEnt040", "Bairro Entrega", ecgHdrTextALignLeft, , 100
    .AddColumn "kCidadeEnt040", "Cidade Entrega", ecgHdrTextALignLeft, , 100
    .AddColumn "kUFEnt004", "UF", ecgHdrTextALignCentre, , 40
    .AddColumn "kCEPEnt012", "CEP", ecgHdrTextALignCentre, , 80
    .AddColumn "kQuadras012", "Qdr.Orig", ecgHdrTextALignLeft, , 80
    .AddColumn "kLotes012", "Lote.Orig", ecgHdrTextALignLeft, , 80
    .AddColumn "kCondominio040", "Nome Condomínio", ecgHdrTextALignLeft, , 120
    .AddColumn "kVVT", "V.V.Territorial", ecgHdrTextALignRight, , 70
    .AddColumn "kVVP", "V.V.Predial", ecgHdrTextALignRight, , 70
    .AddColumn "kVVI", "V.V.Imóvel", ecgHdrTextALignRight, , 70
    .AddColumn "kCodLogr", "Código Logradouro", ecgHdrTextALignRight, , 60
    .AddColumn "kMatr", "Matrícula", ecgHdrTextALignRight, , 60
    .AddColumn "kTrans", "Transcrição", ecgHdrTextALignRight, , 60
    
End With

End Sub

Private Sub CarregaLista()
Dim itmX As ListItem, sDoc As String
Dim x As Long
Ocupado
grdMain.Redraw = False
grdMain.Clear
grdMain.Redraw = True
grdMain.Redraw = False

For x = 1 To UBound(aCodigos)
    If cGetInputState() <> 0 Then DoEvents
    grdMain.AddRow
    grdMain.CellDetails grdMain.Rows, 1, Format(aCodigos(x).nCodigo, "000000"), DT_CENTER
    grdMain.CellDetails grdMain.Rows, 2, aCodigos(x).sProprietario
    grdMain.CellDetails grdMain.Rows, 3, aCodigos(x).sLogradouro
    grdMain.CellDetails grdMain.Rows, 4, aCodigos(x).nNumero, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 5, aCodigos(x).sComplemento
    grdMain.CellDetails grdMain.Rows, 6, aCodigos(x).sBairro
    grdMain.CellDetails grdMain.Rows, 7, aCodigos(x).sUsoTerreno
    grdMain.CellDetails grdMain.Rows, 8, aCodigos(x).sCategProp
    grdMain.CellDetails grdMain.Rows, 9, aCodigos(x).sTopografia
    grdMain.CellDetails grdMain.Rows, 10, aCodigos(x).sSituacao
    grdMain.CellDetails grdMain.Rows, 11, aCodigos(x).sBenfeitoria
    grdMain.CellDetails grdMain.Rows, 12, aCodigos(x).sPedologia
    grdMain.CellDetails grdMain.Rows, 13, FormatNumber(aCodigos(x).nFracaoIdeal, 2), DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 14, aCodigos(x).sAtivo, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 15, aCodigos(x).sInscricao, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 16, Format(aCodigos(x).nDistrito, "00"), DT_CENTER
    grdMain.CellDetails grdMain.Rows, 17, Format(aCodigos(x).nSetor, "00"), DT_CENTER
    grdMain.CellDetails grdMain.Rows, 18, Format(aCodigos(x).nQuadra, "0000"), DT_CENTER
    grdMain.CellDetails grdMain.Rows, 19, Format(aCodigos(x).nLote, "00000"), DT_CENTER
    grdMain.CellDetails grdMain.Rows, 20, Format(aCodigos(x).nFace, "00"), DT_CENTER
    grdMain.CellDetails grdMain.Rows, 21, aCodigos(x).sTipoImovel
    grdMain.CellDetails grdMain.Rows, 22, FormatNumber(aCodigos(x).nAreaPredial, 2), DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 23, FormatNumber(aCodigos(x).nAreaTerreno, 2), DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 24, aCodigos(x).nTipoEndereco, DT_CENTER
    grdMain.CellDetails grdMain.Rows, 25, aCodigos(x).sLogradouroEnt
    grdMain.CellDetails grdMain.Rows, 26, aCodigos(x).nNumeroEnt, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 27, aCodigos(x).sComplementoEnt
    grdMain.CellDetails grdMain.Rows, 28, aCodigos(x).sBairroEnt
    grdMain.CellDetails grdMain.Rows, 29, aCodigos(x).sCidadeEnt
    grdMain.CellDetails grdMain.Rows, 30, aCodigos(x).sUFEnt
    grdMain.CellDetails grdMain.Rows, 31, aCodigos(x).sCEPEnt
    grdMain.CellDetails grdMain.Rows, 32, aCodigos(x).sQuadras
    grdMain.CellDetails grdMain.Rows, 33, aCodigos(x).sLotes
    grdMain.CellDetails grdMain.Rows, 34, aCodigos(x).sNomeCondominio
    grdMain.CellDetails grdMain.Rows, 35, FormatNumber(aCodigos(x).nVVT, 2), DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 36, FormatNumber(aCodigos(x).nVVP, 2), DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 37, FormatNumber(aCodigos(x).nVVI, 2), DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 38, aCodigos(x).nCodLogr, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 39, aCodigos(x).nMatricula, DT_RIGHT
    grdMain.CellDetails grdMain.Rows, 40, aCodigos(x).nTranscricao, DT_RIGHT
    
Next
Liberado
grdMain.Redraw = True
End Sub

Private Sub cmdDD1All_Click()
Dim x As Integer
For x = 0 To lstDDList1.ListCount - 1
    lstDDList1.Selected(x) = True
Next
End Sub

Private Sub cmdDD1None_Click()
Dim x As Integer
For x = 0 To lstDDList1.ListCount - 1
    lstDDList1.Selected(x) = False
Next
lstDDList1.Selected(0) = True
End Sub

Private Sub HideColumns()
Dim x As Integer, y As Integer, bAchou As Boolean

bAchou = False
For x = 0 To lstDDList1.ListCount - 1
    If lstDDList1.Selected(x) = True Then
        bAchou = True
        Exit For
    End If
Next
If Not bAchou Then
    MsgBox "Voce deve selecionar pelo menos um campo para exibição.", vbExclamation, "Atenção"
    Exit Sub
End If

For x = 0 To lstDDList1.ListCount - 1
     grdMain.ColumnVisible(x + 1) = lstDDList1.Selected(x)
Next

End Sub

Private Sub HideList()
cmdUT.value = False: cmdUT_Click
cmdCP.value = False: cmdCP_Click
cmdPE.value = False: cmdPE_Click
cmdTO.value = False: cmdTO_Click
cmdSI.value = False: cmdSI_Click
cmdBE.value = False: cmdBE_Click
cmdBairro.value = False: cmdBairro_Click
cmdCond.value = False: cmdCond_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set xImovel = Nothing
End Sub

Private Sub grdMain_ColumnClick(ByVal lcol As Long)

Dim sTag As String
Dim iSortIndex As Long
      
   With grdMain.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lcol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lcol
      End If
   
      ' Determine which sort order to apply:
      sTag = grdMain.ColumnTag(lcol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdMain.ColumnTag(lcol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = grdMain.ColumnSortType(lcol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   grdMain.Sort
   Screen.MousePointer = vbDefault

End Sub


Private Sub mnuEtNormal_Click()
Dim x As Integer, RdoAux As rdoResultset
Ocupado
If cGetInputState() <> 0 Then DoEvents
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
    
   
    
With grdMain
    For x = 1 To .Rows
        If cGetInputState() <> 0 Then DoEvents
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,PROCESSO) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & x & ",'" & .CellText(x, 1) & "','" & Mask(.CellText(x, 2)) & "','"
        Sql = Sql & Mask(Left(.CellText(x, 25) & " " & .CellText(x, 26) & " " & .CellText(x, 27), 60)) & "','" & Mask(.CellText(x, 28) & " - " & .CellText(x, 29) & "/" & .CellText(x, 30)) & "','"
        Sql = Sql & .CellText(x, 31) & "','" & Mask(Left((.CellText(x, 25) & " " & .CellText(x, 26) & " " & .CellText(x, 27)), 60)) & "')"
        cn.Execute Sql, rdExecDirect
    Next
End With
    
        
Liberado
If cGetInputState() <> 0 Then DoEvents
frmReport.ShowReport "ETIQUETAIPTU", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub mnuEtNot_Click()
Dim x As Integer
Ocupado
If cGetInputState() <> 0 Then DoEvents
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
    
With grdMain
    For x = 1 To .Rows
        If cGetInputState() <> 0 Then DoEvents
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & x & ",'" & "" & "','" & Mask(.CellText(x, 2)) & "','"
        Sql = Sql & Mask(Left(.CellText(x, 3), 60)) & " " & .CellText(x, 4) & "','" & .CellText(x, 6) & " - " & "JABOTICABAL" & "/" & "SP" & "','" & RetornaCEP(Val(.CellText(x, 38)), Val(.CellText(x, 4))) & "')"
        cn.Execute Sql, rdExecDirect
    Next
End With
Liberado
If cGetInputState() <> 0 Then DoEvents
frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub txtNomeLogr_Change()
If Trim$(txtNomeLogr) = "" Then
   txtCodLogr.Text = 0
End If
End Sub

Private Sub txtNomeLogr_GotFocus()
txtNomeLogr.SelStart = 0
txtNomeLogr.SelLength = Len(txtNomeLogr)
End Sub

Private Sub txtNomeLogr_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLogr.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLogr) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder (0)
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro não encontrado.", vbInformation, "Atenção"
             lstNomeLog.Visible = False
             txtNomeLogr.SetFocus
          End If
      End With
   End If
Else
   txtCodLogr.Text = 0
End If

End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
   txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   txtCodLogr_LostFocus
   lstNomeLog.Visible = False
   txtNumImovel.SetFocus
End If

End Sub

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
   txtNomeLogr.SetFocus
End If

End Sub

Private Sub txtCodLogr_LostFocus()
If Val(txtCodLogr.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLogr.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          txtNomeLogr.Text = ""
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtCodLogr.SetFocus
       End If
      .Close
   End With
End If

End Sub

Private Function FillLeft(sTexto As String, nTamanho As Integer) As String

FillLeft = Space(nTamanho - Len(sTexto)) & sTexto

End Function

Private Function FillSpace(sPalavra As String, nTamanho As Integer) As String
Dim sTmp As String
If nTamanho = 0 Then
   sTmp = sPalavra
Else
If Len(sPalavra) > nTamanho Then sPalavra = Left(sPalavra, nTamanho)
sTmp = sPalavra & Space(nTamanho - Len(sPalavra))
End If
FillSpace = sTmp

End Function

