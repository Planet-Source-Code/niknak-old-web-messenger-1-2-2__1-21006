VERSION 5.00
Begin VB.Form frm_colours 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Colour Settings"
   ClientHeight    =   4710
   ClientLeft      =   150
   ClientTop       =   705
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   7
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4380
      Width           =   255
   End
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   6
      Left            =   780
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4140
      Width           =   255
   End
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   5
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4380
      Width           =   255
   End
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   4
      Left            =   540
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4140
      Width           =   255
   End
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   3
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4380
      Width           =   255
   End
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   2
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4140
      Width           =   255
   End
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   1
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4380
      Width           =   255
   End
   Begin VB.OptionButton opt_colour 
      Height          =   255
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4140
      Width           =   255
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Height          =   555
      Left            =   4560
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   3360
      TabIndex        =   1
      Top             =   4080
      Width           =   1095
   End
   Begin VB.ListBox lst_colvars 
      Height          =   3960
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5595
   End
   Begin VB.Menu men_help 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frm_colours"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_ok_Click()
    For s_colvar = 0 To noof_colvars - 1
        wm_colvars(s_colvar).save_vars
    Next s_colvar
    Unload Me
End Sub

Private Sub Form_Load()
    load_settings
    refresh_colvars
    refresh_palette
End Sub

Private Sub refresh_colvars()
    For r_colvars = 0 To noof_colvars
        With wm_colvars(r_colvars)
            lst_colvars.AddItem .variable_description
        End With
    Next r_colvars
End Sub

Private Sub refresh_palette()
    For r_swatch = 0 To noof_colours - 1
        opt_colour(r_swatch).BackColor = wm_colours(r_swatch).win_colour
    Next r_swatch
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
End Sub

Private Sub lst_colvars_Click()
    opt_colour(wm_colvars(lst_colvars.ListIndex).variable_colour) = True
End Sub

Private Sub men_help_Click()
    frm_help.Show
    frm_help.Caption = "Help-" & Me.Caption
    frm_help.lbl_help.Caption = help_colours
End Sub

Private Sub load_settings()
    With frm_colours
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    save_window Me.Caption, Me.Top, Me.Left
End Sub

Private Sub opt_colour_Click(Index As Integer)
    If lst_colvars.ListIndex > -1 Then
        wm_colvars(lst_colvars.ListIndex).variable_colour = Index
    End If
End Sub
