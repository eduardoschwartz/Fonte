VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3255
   ClientLeft      =   2280
   ClientTop       =   3150
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Timer RedrawTimer 
      Interval        =   300
      Left            =   3000
      Top             =   3360
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
