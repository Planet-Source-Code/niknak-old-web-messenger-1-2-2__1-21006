VERSION 5.00
Begin VB.Form frm_help 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Help"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lbl_help 
      Alignment       =   2  'Center
      Caption         =   "<HELP STRING>"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4395
   End
   Begin VB.Image img_help 
      Height          =   480
      Left            =   2040
      Picture         =   "frm_help.frx":0000
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frm_help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
