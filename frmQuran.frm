VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmQuran 
   BorderStyle     =   0  'None
   Caption         =   "ÈÓã Çááå"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14775
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   14775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "B"
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   2760
      Width           =   11415
   End
   Begin VB.CommandButton cmdALBaghareh 
      Caption         =   "AL Baghareh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   43
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdMaryam 
      Caption         =   "Maryan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   42
      Top             =   6480
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "frmQuran.frx":0000
      Left            =   8400
      List            =   "frmQuran.frx":015A
      TabIndex        =   41
      Top             =   4200
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CDlg 
      Left            =   120
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      DialogTitle     =   "Open File"
      Filter          =   "*.txt"
   End
   Begin VB.ListBox lstAbjad 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5490
      Index           =   6
      ItemData        =   "frmQuran.frx":041A
      Left            =   12840
      List            =   "frmQuran.frx":0472
      TabIndex        =   38
      Top             =   2760
      Width           =   495
   End
   Begin VB.ListBox lstBase3 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1080
      ItemData        =   "frmQuran.frx":04FC
      Left            =   120
      List            =   "frmQuran.frx":04FE
      MultiSelect     =   2  'Extended
      TabIndex        =   37
      Top             =   5880
      Width           =   6135
   End
   Begin VB.TextBox txtBase3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   36
      Top             =   1920
      Width           =   14655
   End
   Begin VB.ListBox lstBase2 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1080
      ItemData        =   "frmQuran.frx":0500
      Left            =   120
      List            =   "frmQuran.frx":0502
      MultiSelect     =   2  'Extended
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4800
      Width           =   6135
   End
   Begin VB.ListBox lstBase 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1080
      ItemData        =   "frmQuran.frx":0504
      Left            =   120
      List            =   "frmQuran.frx":0506
      MultiSelect     =   2  'Extended
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3720
      Width           =   6135
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "set Pad Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      TabIndex        =   35
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
      Height          =   300
      Left            =   13920
      TabIndex        =   34
      Top             =   0
      Width           =   855
   End
   Begin VB.TextBox lblIndx2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   240
      TabIndex        =   33
      Text            =   "0"
      Top             =   7080
      Width           =   1695
   End
   Begin VB.TextBox lblIndx1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   4680
      TabIndex        =   32
      Text            =   "0"
      Top             =   7080
      Width           =   1695
   End
   Begin VB.ListBox lstAbjad 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2760
      Index           =   5
      ItemData        =   "frmQuran.frx":0508
      Left            =   9855
      List            =   "frmQuran.frx":0536
      TabIndex        =   29
      Top             =   4200
      Width           =   1560
   End
   Begin VB.ListBox lstAbjad 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5490
      Index           =   2
      ItemData        =   "frmQuran.frx":05E3
      Left            =   14085
      List            =   "frmQuran.frx":063B
      TabIndex        =   28
      Top             =   2760
      Width           =   495
   End
   Begin VB.ListBox lstAbjad 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2760
      Index           =   3
      ItemData        =   "frmQuran.frx":06CB
      Left            =   11400
      List            =   "frmQuran.frx":06F9
      TabIndex        =   27
      Top             =   4200
      Width           =   855
   End
   Begin VB.ListBox lstAbjad 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2760
      Index           =   4
      ItemData        =   "frmQuran.frx":0743
      Left            =   9435
      List            =   "frmQuran.frx":0771
      TabIndex        =   26
      Top             =   4200
      Width           =   495
   End
   Begin VB.ListBox lstAbjad 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5490
      Index           =   1
      ItemData        =   "frmQuran.frx":07B2
      Left            =   13830
      List            =   "frmQuran.frx":080A
      TabIndex        =   25
      Top             =   2760
      Width           =   255
   End
   Begin VB.ListBox lstAbjad 
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5490
      Index           =   0
      ItemData        =   "frmQuran.frx":0862
      Left            =   13335
      List            =   "frmQuran.frx":08BA
      TabIndex        =   24
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox txtAbjdN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   10920
      TabIndex        =   23
      Top             =   8280
      Width           =   375
   End
   Begin VB.TextBox txtAbjdN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      HelpContextID   =   3
      Index           =   3
      Left            =   10560
      TabIndex        =   22
      Top             =   8280
      Width           =   375
   End
   Begin VB.TextBox txtAbjdN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      HelpContextID   =   2
      Index           =   2
      Left            =   10200
      TabIndex        =   21
      Top             =   8280
      Width           =   375
   End
   Begin VB.TextBox txtAbjdN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   9840
      TabIndex        =   20
      Top             =   8280
      Width           =   375
   End
   Begin VB.TextBox txtAbjdN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   9480
      TabIndex        =   19
      Top             =   8280
      Width           =   375
   End
   Begin VB.TextBox txtAbjd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   4
      Left            =   10920
      TabIndex        =   18
      Top             =   8040
      Width           =   375
   End
   Begin VB.TextBox txtAbjd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   3
      Left            =   10560
      TabIndex        =   17
      Top             =   8040
      Width           =   375
   End
   Begin VB.TextBox txtAbjd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   2
      Left            =   10200
      TabIndex        =   16
      Top             =   8040
      Width           =   375
   End
   Begin VB.TextBox txtAbjd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   1
      Left            =   9840
      TabIndex        =   15
      Top             =   8040
      Width           =   375
   End
   Begin VB.TextBox txtAbjd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Index           =   0
      Left            =   9480
      TabIndex        =   14
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton cmdAbjad_Calc 
      Caption         =   "Abjad_Calc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   13
      Top             =   7560
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   12
      Top             =   5280
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7080
      TabIndex        =   11
      Top             =   4080
      Width           =   255
   End
   Begin VB.TextBox txtBase2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   1080
      Width           =   14655
   End
   Begin VB.CommandButton cmdOpenFile2 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   8
      Top             =   5160
      Width           =   735
   End
   Begin VB.CommandButton cmdClearText2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   7
      Top             =   5400
      Width           =   735
   End
   Begin VB.TextBox txtWord 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   855
      Left            =   4560
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7500
      Width           =   4335
   End
   Begin VB.CommandButton cmdClearText 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   2
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdOpenFile 
      Caption         =   "Open"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6360
      TabIndex        =   1
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtBase 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   240
      Width           =   14655
   End
   Begin VB.TextBox txtOut 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00787878&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   7500
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.PictureBox picBFrmQuran 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   8775
      Left            =   0
      ScaleHeight     =   8745
      ScaleWidth      =   14760
      TabIndex        =   30
      Top             =   0
      Width           =   14790
      Begin VB.TextBox lblWordCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   7920
         TabIndex        =   39
         Text            =   "0"
         Top             =   8280
         Width           =   975
      End
      Begin VB.CheckBox chkAlphaEnable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   3720
         TabIndex        =   31
         Top             =   9960
         Width           =   225
      End
   End
   Begin VB.Label lblListCount 
      Caption         =   "0"
      Height          =   255
      Left            =   9240
      TabIndex        =   4
      Top             =   8160
      Width           =   1335
   End
End
Attribute VB_Name = "frmQuran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************************
'
'   Copyright(C) 2010 By Kaveh Abdollahi.   kavehplus@gmail.com
'   Time Engine
'   June 2010
'
'******************************************************************************************

Option Explicit
Private Path As String


Private Sub cmdClearText_Click()
    txtBase.Text = ""
    lstBase.Clear
    txtWord.Text = ""
End Sub

Private Sub cmdClearText2_Click()
    txtBase2.Text = ""
    lstBase2.Clear
End Sub

Private Sub cmdOpenFile_Click()
Dim intf As Integer, ln As Long
Dim s As String

    CDlg.ShowOpen
    Path = CDlg.filename
    
    If Path = "" Then Exit Sub
    cmdOpenFile.Enabled = False
    
    intf = FreeFile
    Open Path For Input As #intf
    On Error GoTo err
    
    While Not EOF(intf)
        Input #intf, s
        lstBase.AddItem Trim(s)
        DoEvents
    Wend

err:
    Close #intf

    cmdOpenFile.Enabled = True
    lblListCount = lstBase.ListCount
End Sub

Private Sub cmdOpenFile2_Click()
Dim intf As Integer, ln As Long
Dim s As String

    CDlg.ShowOpen
    Path = CDlg.filename
    
    If Path = "" Then Exit Sub
    cmdOpenFile2.Enabled = False
    
    intf = FreeFile
    Open Path For Input As #intf
    On Error GoTo err
    
    While Not EOF(intf)
        Input #intf, s
        lstBase2.AddItem Trim(s)
        DoEvents
    Wend

err:
    Close #intf

    cmdOpenFile2.Enabled = True
    lblListCount = lstBase.ListCount

End Sub

Public Sub cmdAbjad_Calc_Click()
Dim x As Integer, s As String, s2 As String, i As Integer, sS() As String, y As Integer
sS = Split(lstBase.List(lstBase.ListIndex), " ", , vbTextCompare)

Text1.Text = AbjadClC(lstBase.List(lstBase.ListIndex))

End Sub

Private Sub cmdMaryam_Click()
    lstBase.Clear
    lstBase2.Clear
    lstBase3.Clear
    LoadMaryan
End Sub

Private Sub Command2_Click()
Dim x As Integer
    txtOut.Visible = True
    txtOut.Text = ""
    txtOut.ZOrder 0
    
    For x = 0 To lstBase.ListCount - 1
        txtOut.Text = txtOut.Text & lstBase.List(x) & vbCrLf
    Next x
    
End Sub

Private Sub Command3_Click()
Dim x As Integer
    txtOut.Visible = True
    txtOut.Text = ""
    txtOut.ZOrder 0
    
    For x = 0 To lstBase2.ListCount - 1
        txtOut.Text = txtOut.Text & lstBase2.List(x) & vbCrLf
    Next x
End Sub

Private Sub Command4_Click()
    Unload Me
End Sub

Public Sub cmdSet_Click()
Dim s As String, sS As String, s2() As String, s3() As String, s4() As String, i As Integer, j As Integer
s = txtBase2.Text
    If Len(s) > 130 Then
        s2 = Split(s, " ", , vbTextCompare)
        s3 = Split(txtBase.Text, " ", , vbTextCompare)
        s4 = Split(txtBase3.Text, " ", , vbTextCompare)
        
        txtBase.Text = ""
        txtBase2.Text = "": j = 1
        txtBase3.Text = ""
            For i = 0 To UBound(s2)
                txtBase2.Text = txtBase2.Text & " " & s2(i)
                If Len(txtBase2.Text) > 130 * j Then txtBase2.Text = txtBase2.Text & vbCrLf: j = j + 1
            Next i
            j = 1
            For i = 0 To UBound(s4)
                txtBase3.Text = txtBase3.Text & " " & s4(i)
                If Len(txtBase3.Text) > 130 * j Then txtBase3.Text = txtBase3.Text & vbCrLf: j = j + 1
            Next i
            For i = 0 To UBound(s3)
                txtBase.Text = txtBase.Text & " " & s3(i)
                If Int(UBound(s3) * 0.66) = i And j > 2 Then
                    sS = ""
                    txtBase.Text = sS & txtBase.Text & sS & vbCrLf
                End If
                If Int(UBound(s3) * 0.33) = i And j > 3 Then
                    sS = ""
                    txtBase.Text = sS & txtBase.Text & sS & vbCrLf
                End If
            Next i
      
      Else
        
    End If
    
    frmBase.txtNumber.Text = lstBase.ListIndex + 1
    frmBase.txtPad(0).Text = "" & txtBase.Text & "" & vbCrLf
    frmBase.txtPad(1).Text = "" & txtBase2.Text & "" & vbCrLf
    frmBase.txtPad(2).Text = "" & txtBase3.Text & "" & vbCrLf
    frmBase.txtPad(3).Text = "" & Text1.Text & "" & vbCrLf
End Sub

Private Sub cmdALBaghareh_Click()
    lstBase.Clear
    lstBase2.Clear
    lstBase3.Clear
    LoadALBaghare
End Sub

Private Sub Form_Load()
Dim intf As Integer, ln As Long, i As Integer
Dim s As String, s2 As String, z1, z2
    
    On Error GoTo err

    Path = App.Path & "\Quran\ALBaghare.txt"
    intf = FreeFile
    Open Path For Input As #intf
    
    While Not EOF(intf)
        Input #intf, s
        z1 = InStr(1, s, "«", vbTextCompare)
        lstBase.AddItem Trim(Left(s, z1 - 1))
        DoEvents
    Wend

    Close #intf

    Path = App.Path & "\Quran\ALBaghare_English.txt"
    intf = FreeFile
    Open Path For Input As #intf
    s2 = ""
    While Not EOF(intf)
        Input #intf, s
        s2 = s2 & s
        If Right(s2, 1) = "." Then
            lstBase3.AddItem Mid(s2, InStr(1, s2, ".") + 1, Len(s2) - InStr(1, s2, ".") - 1) & " ": s2 = ""
        End If
        DoEvents
    Wend

    Close #intf

    lblListCount = lstBase.ListCount
    Path = App.Path & "\Quran\ALBaghare_Persian.txt"
    intf = FreeFile
    Open Path For Input As #intf
    
    While Not EOF(intf)
        Input #intf, s
        z1 = 1: s2 = ""
        While z1 <> 0
            z1 = InStr(1, s, "(", vbTextCompare)
            If z1 <> 0 Then
                z2 = InStr(z1, s, ")", vbTextCompare)
                s2 = Mid(s, z1, z2 - z1 + 1)
                s = Replace(s, s2, "")
                s = Replace(s, ";", "")
                s = Replace(s, ".", "")
                s = Replace(s, "!", "")
                s = Replace(s, "ÓþÐááøå", "")
                s = Replace(s, "ÓþÎþááøå", "")
            End If
        Wend
        lstBase2.AddItem Trim(s)
        DoEvents
    Wend

    Close #intf

    lblListCount = lstBase.ListCount
    lstBase.ListIndex = lstBase.ListCount - 1
    lstBase.SetFocus
    
err:
    DoEvents
End Sub
Public Sub LoadALBaghare()
Dim intf As Integer, ln As Long, i As Integer
Dim s As String, s2 As String, z1, z2
    On Error GoTo err

    Path = App.Path & "\Quran\ALBaghare.txt"
    intf = FreeFile
    Open Path For Input As #intf
    
    While Not EOF(intf)
        Input #intf, s
        z1 = InStr(1, s, "«", vbTextCompare)
        lstBase.AddItem Trim(Left(s, z1 - 1))
        DoEvents
    Wend

    Close #intf

    Path = App.Path & "\Quran\ALBaghare_English.txt"
    intf = FreeFile
    Open Path For Input As #intf
    s2 = ""
    While Not EOF(intf)
        Input #intf, s
        s2 = s2 & s
        If Right(s2, 1) = "." Then
            lstBase3.AddItem Mid(s2, InStr(1, s2, ".") + 1, Len(s2) - InStr(1, s2, ".") - 1) & " ": s2 = ""
        End If
        DoEvents
    Wend

    Close #intf

    lblListCount = lstBase.ListCount
    Path = App.Path & "\Quran\ALBaghare_Persian.txt"
    intf = FreeFile
    Open Path For Input As #intf
    
    While Not EOF(intf)
        Input #intf, s
        z1 = 1: s2 = ""
        While z1 <> 0
            z1 = InStr(1, s, "(", vbTextCompare)
            If z1 <> 0 Then
                z2 = InStr(z1, s, ")", vbTextCompare)
                s2 = Mid(s, z1, z2 - z1 + 1)
                s = Replace(s, s2, "")
                s = Replace(s, ";", "")
                s = Replace(s, ".", "")
                s = Replace(s, "!", "")
                s = Replace(s, "ÓþÐááøå", "")
                s = Replace(s, "ÓþÎþááøå", "")
            End If
        Wend
        lstBase2.AddItem Trim(s)
        DoEvents
    Wend

    Close #intf

    lblListCount = lstBase.ListCount
    lstBase.ListIndex = lstBase.ListCount - 1
    lstBase.SetFocus
    
err:
    DoEvents


End Sub
Public Sub LoadMaryan()
Dim intf As Integer, ln As Long, i As Integer
Dim s As String, s2 As String, z1, z2
    On Error GoTo err
    If frmQuran.Tag = "M" Then frmQuran.Tag = "B" ': LoadMaryan: Exit Sub
    Path = App.Path & "\Quran\Maryam.txt"
    intf = FreeFile
    Open Path For Input As #intf
    
    While Not EOF(intf)
        Input #intf, s
        z1 = InStr(1, s, "«", vbTextCompare)
        lstBase.AddItem Trim(Left(s, z1 - 1))
        DoEvents
    Wend

    Close #intf

    Path = App.Path & "\Quran\Maryam_English.txt"
    intf = FreeFile
    Open Path For Input As #intf
    s2 = ""
    While Not EOF(intf)
        Input #intf, s
        s2 = s2 & s
        If Right(s2, 1) = "." Then
            lstBase3.AddItem Mid(s2, InStr(1, s2, ".") + 1, Len(s2) - InStr(1, s2, ".") - 1) & " ": s2 = ""
        End If
        DoEvents
    Wend

    Close #intf

    lblListCount = lstBase.ListCount
    Path = App.Path & "\Quran\Maryam_Persian.txt"
    intf = FreeFile
    Open Path For Input As #intf
    
    While Not EOF(intf)
        Input #intf, s
        z1 = 1: s2 = ""
        While z1 <> 0
            z1 = InStr(1, s, "(", vbTextCompare)
            If z1 <> 0 Then
                z2 = InStr(z1, s, ")", vbTextCompare)
                s2 = Mid(s, z1, z2 - z1 + 1)
                s = Replace(s, s2, "")
                s = Replace(s, ";", "")
                s = Replace(s, ".", "")
                s = Replace(s, "!", "")
                s = Replace(s, "ÓþÐááøå", "")
                s = Replace(s, "ÓþÎþááøå", "")
            End If
        Wend
        lstBase2.AddItem Trim(s)
        DoEvents
    Wend

    Close #intf

    lblListCount = lstBase.ListCount
    lstBase.ListIndex = lstBase.ListCount - 1
    lstBase.SetFocus
    
err:
    DoEvents

End Sub
Private Sub Form_Unload(Cancel As Integer)
frP = True
    frmBase.Enabled = True
    frmQuran.Visible = False
End Sub

Public Sub lstBase_Click()
Dim s() As String, str As String, x As Integer, chrLen As Single
    
    txtBase.Text = lstBase.List(lstBase.ListIndex)
    txtBase3.Text = lstBase3.List(lstBase.ListIndex)
    cmdAbjad_Calc_Click
    s = Split(txtBase.Text, " ")
    txtWord.Text = ""
    chrLen = 0
    
    For x = 0 To UBound(s)
        str = str & s(x) & " - "
    Next x
    
    x = InStr(1, str, "«", vbTextCompare)
    If x = 0 Then x = InStr(1, str, "(", vbTextCompare)
    If x = 0 Then x = Len(str) + 1
    
    lstBase2.ListIndex = lstBase.ListIndex
    lblIndx1.Text = lstBase.ListIndex + 1
End Sub

Private Sub lstBase2_Click()
Dim s() As String, str As String, x As Integer, chrLen As Single
    If lstBase.ListIndex <> lstBase2.ListIndex Then lstBase.ListIndex = lstBase2.ListIndex
    txtBase2.Text = lstBase2.List(lstBase2.ListIndex)
    lblIndx2.Text = lstBase2.ListIndex + 1
    DoEvents
End Sub

