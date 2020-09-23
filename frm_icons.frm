VERSION 5.00
Begin VB.Form frm_icons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Available Icons"
   ClientHeight    =   4830
   ClientLeft      =   150
   ClientTop       =   705
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt_avicons 
      Height          =   4695
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   5595
   End
   Begin VB.Menu men_help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frm_icons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    load_settings
    read_avicons
End Sub

Private Sub read_avicons()
    For av_icons = 0 To noof_icons - 1
        txt_avicons.Text = txt_avicons.Text & msg_icons(av_icons).icon_recogstr & "     Gets replaced by a " & msg_icons(av_icons).icon_description & vbCrLf
    Next av_icons
End Sub

Private Sub load_settings()
    With frm_icons
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    save_window Me.Caption, Me.Top, Me.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
End Sub

Private Sub men_help_Click()
    frm_help.Show
    frm_help.Caption = "Help-" & Me.Caption
    frm_help.lbl_help.Caption = help_icons
End Sub

