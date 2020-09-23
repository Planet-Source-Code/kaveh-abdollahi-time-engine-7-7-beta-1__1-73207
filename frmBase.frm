VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBase 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Liquid Skyes "
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Dialog Light"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "frmBase.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmBase.frx":030A
   ScaleHeight     =   11520
   ScaleMode       =   0  'User
   ScaleWidth      =   15360
   Begin VB.Frame fraProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7455
      Left            =   0
      TabIndex        =   260
      Tag             =   "1"
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   20
         Left            =   2100
         TabIndex        =   351
         Text            =   "1"
         ToolTipText     =   "Agg Value"
         Top             =   6120
         Width           =   780
      End
      Begin VB.CheckBox chkBGQ 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   2
         Left            =   2415
         TabIndex        =   501
         ToolTipText     =   "orbt (2.3) Set Agg For Step"
         Top             =   5880
         Width           =   225
      End
      Begin VB.CheckBox chkBGQ 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   3
         Left            =   2670
         TabIndex        =   502
         ToolTipText     =   "orbt (3.3) Set Agg For Step"
         Top             =   5880
         Width           =   225
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DDD3E0&
         Caption         =   "Cam"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   496
         Top             =   523
         Width           =   570
      End
      Begin VB.CheckBox chkAvalue 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   1
         Left            =   2030
         TabIndex        =   482
         Top             =   1080
         Width           =   190
      End
      Begin VB.CheckBox chkAvalue 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   0
         Left            =   2970
         TabIndex        =   469
         Top             =   1080
         Value           =   1  'Checked
         Width           =   190
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   2175
         TabIndex        =   468
         Text            =   "0"
         Top             =   1065
         Width           =   825
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00B9CBAF&
         Caption         =   "Set By"
         Height          =   225
         Index           =   0
         Left            =   2175
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   467
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   11
         Left            =   2220
         MaxLength       =   9
         TabIndex        =   466
         Text            =   "1"
         Top             =   1440
         Width           =   735
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Index           =   28
         Left            =   2340
         Locked          =   -1  'True
         TabIndex        =   465
         Text            =   " Time"
         Top             =   1245
         Width           =   495
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00B9CBAF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   1920
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   464
         Top             =   1425
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00B9CBAF&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   2970
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   463
         Top             =   1440
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   3900
         TabIndex        =   434
         Text            =   "1"
         Top             =   1425
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Set Start ByTime"
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   5010
         MaskColor       =   &H000000FF&
         TabIndex        =   436
         Top             =   5955
         Width           =   840
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   260
         Index           =   5
         Left            =   3900
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   435
         Text            =   "Time +="
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   26
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   433
         Top             =   1200
         Width           =   300
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   4905
         TabIndex        =   432
         Text            =   "1"
         Top             =   1545
         Width           =   945
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   431
         Top             =   1680
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   24
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   430
         Top             =   1440
         Width           =   300
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set By"
         Height          =   225
         Index           =   1
         Left            =   4905
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   429
         Top             =   1305
         UseMaskColor    =   -1  'True
         Width           =   945
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   26
         Left            =   4530
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   428
         Top             =   1215
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   4530
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   427
         Top             =   1650
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   4530
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   426
         Top             =   1440
         Width           =   300
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   3630
         ItemData        =   "frmBase.frx":BF95B
         Left            =   5040
         List            =   "frmBase.frx":BF9A7
         TabIndex        =   407
         Top             =   1920
         Width           =   810
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         ForeColor       =   &H00000000&
         Height          =   3630
         ItemData        =   "frmBase.frx":BFA44
         Left            =   4245
         List            =   "frmBase.frx":BFA90
         TabIndex        =   404
         Top             =   1920
         Width           =   810
      End
      Begin VB.CheckBox chkLogP 
         BackColor       =   &H00000000&
         Caption         =   "Set Logs To pic"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   4440
         TabIndex        =   402
         Top             =   7200
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.TextBox txtAutoAGG 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS SystemEx"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Left            =   1800
         TabIndex        =   401
         Text            =   "1"
         ToolTipText     =   "Interval"
         Top             =   6600
         Width           =   450
      End
      Begin VB.CommandButton cmdShot 
         BackColor       =   &H0000FF00&
         Caption         =   "Shot"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5115
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   319
         Top             =   530
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdShock 
         Caption         =   "Shock"
         Height          =   255
         Left            =   3480
         TabIndex        =   398
         Top             =   6960
         Width           =   615
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   3630
         ItemData        =   "frmBase.frx":BFB1D
         Left            =   3600
         List            =   "frmBase.frx":BFB69
         MultiSelect     =   2  'Extended
         TabIndex        =   350
         Top             =   1920
         Width           =   285
      End
      Begin VB.CheckBox chkZx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   2
         Left            =   1080
         TabIndex        =   392
         Top             =   7200
         Width           =   225
      End
      Begin VB.CheckBox chkZx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   1
         Left            =   780
         TabIndex        =   391
         Top             =   7200
         Width           =   225
      End
      Begin VB.CheckBox chkZx 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   210
         Index           =   0
         Left            =   480
         TabIndex        =   390
         Top             =   7200
         Width           =   225
      End
      Begin VB.CheckBox chkP 
         BackColor       =   &H002D061B&
         Caption         =   "1"
         ForeColor       =   &H00F1E4F3&
         Height          =   180
         Index           =   1
         Left            =   490
         MaskColor       =   &H000000FF&
         TabIndex        =   389
         Top             =   6960
         Width           =   400
      End
      Begin VB.CommandButton cmdPad 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pad"
         Height          =   240
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   382
         Top             =   530
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   27
         Left            =   1860
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   272
         Top             =   4080
         Width           =   300
      End
      Begin VB.TextBox txtGR2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EFD6E7&
         BeginProperty Font 
            Name            =   "MS SystemEx"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2160
         TabIndex        =   275
         Text            =   "0.1"
         Top             =   1800
         Width           =   690
      End
      Begin VB.CheckBox chkP 
         BackColor       =   &H002D061B&
         Caption         =   "0"
         ForeColor       =   &H00F1E4F3&
         Height          =   180
         Index           =   0
         Left            =   75
         MaskColor       =   &H000000FF&
         TabIndex        =   381
         Top             =   6960
         Width           =   400
      End
      Begin VB.CheckBox chkP 
         BackColor       =   &H002D061B&
         Caption         =   "3"
         ForeColor       =   &H00F1E4F3&
         Height          =   180
         Index           =   3
         Left            =   1320
         MaskColor       =   &H000000FF&
         TabIndex        =   380
         ToolTipText     =   "Magic"
         Top             =   6960
         Width           =   400
      End
      Begin VB.CheckBox chkBGQ 
         BackColor       =   &H00000000&
         Caption         =   "Finish index"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   1
         Left            =   3600
         TabIndex        =   349
         Top             =   6270
         Width           =   1305
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Level 1 Str"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   376
         Top             =   1845
         Width           =   1500
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00000000&
         Caption         =   "Auto Clear"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   3240
         MaskColor       =   &H000000FF&
         TabIndex        =   375
         Top             =   7200
         Value           =   1  'Checked
         Width           =   1140
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   27
         Left            =   2955
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   273
         Top             =   4080
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   19
         Left            =   2130
         TabIndex        =   262
         Text            =   "1"
         Top             =   4110
         Width           =   855
      End
      Begin VB.Frame fraMsgBx 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1800
         TabIndex        =   368
         Tag             =   "x"
         Top             =   7320
         Visible         =   0   'False
         Width           =   3855
         Begin VB.TextBox txtMsgBx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            CausesValidation=   0   'False
            Height          =   175
            Left            =   120
            TabIndex        =   372
            Top             =   600
            Width           =   375
         End
         Begin VB.CommandButton cmdMsgBx 
            BackColor       =   &H00FFFFFF&
            Caption         =   "New Rec"
            Height          =   275
            Index           =   2
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   371
            Top             =   570
            Width           =   975
         End
         Begin VB.CommandButton cmdMsgBx 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Overwrite"
            Height          =   275
            Index           =   1
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   370
            Top             =   570
            Width           =   975
         End
         Begin VB.CommandButton cmdMsgBx 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1EDED&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   3625
            MaskColor       =   &H000040C0&
            Style           =   1  'Graphical
            TabIndex        =   369
            TabStop         =   0   'False
            Top             =   30
            UseMaskColor    =   -1  'True
            Width           =   210
         End
         Begin VB.TextBox TextLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            DragMode        =   1  'Automatic
            ForeColor       =   &H00FFFFFF&
            Height          =   900
            Index           =   33
            Left            =   50
            Locked          =   -1  'True
            TabIndex        =   373
            Text            =   "Owerwrite Record ?  or  Save  New Record ?"
            Top             =   45
            Width           =   3760
         End
      End
      Begin VB.CheckBox chkAutoMax 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Set Finish ByTime"
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   5010
         MaskColor       =   &H000000FF&
         TabIndex        =   367
         Top             =   6480
         Width           =   840
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   1290
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   315
         Top             =   3645
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "ê"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   23
         Left            =   2925
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   366
         Top             =   5580
         Width           =   450
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   23
         Left            =   2925
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   365
         Top             =   5040
         Width           =   450
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   23
         Left            =   2925
         Locked          =   -1  'True
         TabIndex        =   364
         Text            =   "384"
         Top             =   5325
         Width           =   450
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "è"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   22
         Left            =   2445
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   363
         Top             =   5040
         Width           =   450
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   22
         Left            =   2445
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   362
         Top             =   5580
         Width           =   450
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   22
         Left            =   2445
         Locked          =   -1  'True
         TabIndex        =   361
         Text            =   "512"
         Top             =   5325
         Width           =   450
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Level 3"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   90
         TabIndex        =   358
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Level 2"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   357
         Top             =   2370
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.ListBox lst11Pows 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   630
         ItemData        =   "frmBase.frx":BFBC4
         Left            =   2280
         List            =   "frmBase.frx":BFBD7
         TabIndex        =   356
         ToolTipText     =   "Double Click"
         Top             =   6600
         Width           =   690
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0038E493&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   28
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   303
         Top             =   5955
         Width           =   300
      End
      Begin VB.CommandButton cmdSetLev 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Set Layer"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5010
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   353
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   840
      End
      Begin VB.ListBox lstBStp 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   3630
         ItemData        =   "frmBase.frx":BFBF0
         Left            =   3870
         List            =   "frmBase.frx":BFC3C
         TabIndex        =   355
         Top             =   1920
         Width           =   390
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   28
         Left            =   3780
         TabIndex        =   300
         Text            =   "1"
         Top             =   5955
         Width           =   900
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   2895
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   299
         Top             =   6075
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   302
         Top             =   6075
         Width           =   300
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   2040
         TabIndex        =   352
         Text            =   "Agg"
         Top             =   5880
         Width           =   450
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0038E493&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   21
         Left            =   4695
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   298
         Top             =   6555
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0038E493&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   21
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   304
         Top             =   6555
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   21
         Left            =   3780
         TabIndex        =   301
         Text            =   "10000"
         Top             =   6600
         Width           =   900
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0038E493&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   28
         Left            =   4695
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   297
         Top             =   5955
         Width           =   300
      End
      Begin VB.CheckBox chkBGQ 
         BackColor       =   &H00000000&
         Caption         =   "Start index"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Index           =   0
         Left            =   3600
         TabIndex        =   348
         Top             =   5640
         Width           =   1305
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "4"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   720
         MaskColor       =   &H000000FF&
         TabIndex        =   347
         Top             =   1080
         Value           =   1  'Checked
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "3"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   360
         MaskColor       =   &H000000FF&
         TabIndex        =   346
         Top             =   1080
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "2"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1080
         MaskColor       =   &H000000FF&
         TabIndex        =   345
         Top             =   840
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   720
         MaskColor       =   &H000000FF&
         TabIndex        =   344
         Top             =   840
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "5"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   1080
         MaskColor       =   &H000000FF&
         TabIndex        =   343
         Top             =   1080
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   360
         MaskColor       =   &H000000FF&
         TabIndex        =   342
         Top             =   840
         Width           =   345
      End
      Begin VB.TextBox txtR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1815
         Locked          =   -1  'True
         TabIndex        =   341
         Text            =   "1"
         Top             =   5325
         Width           =   495
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Level 1 Str"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   339
         Top             =   1590
         Width           =   1500
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Dr Ellipse"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   338
         Top             =   1320
         Width           =   1500
      End
      Begin VB.CheckBox chkP 
         BackColor       =   &H002D061B&
         Caption         =   "2"
         ForeColor       =   &H00F1E4F3&
         Height          =   180
         Index           =   2
         Left            =   905
         MaskColor       =   &H000000FF&
         TabIndex        =   330
         Top             =   6960
         Width           =   400
      End
      Begin VB.CheckBox chkPant 
         BackColor       =   &H00404040&
         Caption         =   "3D (Oblique) View"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   280
         Index           =   2
         Left            =   75
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   318
         Top             =   6325
         Width           =   1620
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00D696A2&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   2970
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   294
         Top             =   4515
         Width           =   350
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D696A2&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   33
         Left            =   1815
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   289
         Top             =   4515
         Width           =   350
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmBase.frx":BFCA5
         Left            =   90
         List            =   "frmBase.frx":BFCD9
         Style           =   2  'Dropdown List
         TabIndex        =   317
         Top             =   3000
         Width           =   1500
      End
      Begin VB.CheckBox chkAlpha 
         BackColor       =   &H00473842&
         Caption         =   "Alpha"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Left            =   90
         TabIndex        =   316
         Top             =   3360
         Width           =   1500
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00BCB89A&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   314
         Top             =   3645
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   7
         Left            =   390
         TabIndex        =   313
         Text            =   "255"
         Top             =   3645
         Width           =   900
      End
      Begin VB.CheckBox chkPant 
         BackColor       =   &H00404040&
         Caption         =   "4D View"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   280
         Index           =   1
         Left            =   75
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   312
         Top             =   6600
         Value           =   1  'Checked
         Width           =   1620
      End
      Begin VB.CheckBox chkPant 
         BackColor       =   &H00404040&
         Caption         =   "2D ( Polar ) View "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   280
         Index           =   0
         Left            =   75
         MaskColor       =   &H0000FF00&
         Style           =   1  'Graphical
         TabIndex        =   311
         Top             =   6050
         Width           =   1620
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00808080&
         Caption         =   "Str  L2 L3"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   310
         Top             =   2115
         Width           =   1500
      End
      Begin VB.CheckBox chkAutoFix 
         BackColor       =   &H00473842&
         Caption         =   "Auto"
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   2175
         TabIndex        =   284
         Top             =   4800
         Width           =   780
      End
      Begin VB.TextBox txtGR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS SystemEx"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   2400
         TabIndex        =   296
         Text            =   "1"
         ToolTipText     =   "+ Agg"
         Top             =   6360
         Width           =   570
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         HideSelection   =   0   'False
         Index           =   33
         Left            =   2145
         TabIndex        =   295
         Text            =   "10"
         Top             =   4590
         Width           =   855
      End
      Begin VB.CommandButton cmdNextS 
         BackColor       =   &H0000FFFF&
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   945
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   288
         Top             =   4380
         UseMaskColor    =   -1  'True
         Width           =   645
      End
      Begin VB.CommandButton cmdPrevius 
         BackColor       =   &H0000FFFF&
         Caption         =   "Previus"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   287
         Top             =   4380
         UseMaskColor    =   -1  'True
         Width           =   645
      End
      Begin VB.TextBox txtRecCo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   286
         Text            =   "0"
         Top             =   4200
         Width           =   375
      End
      Begin VB.TextBox txtRecCo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   1
         Left            =   1215
         TabIndex        =   285
         Text            =   "0"
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   31
         Left            =   1860
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   271
         Top             =   2280
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   30
         Left            =   1860
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   269
         Top             =   2880
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   29
         Left            =   1860
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   267
         Top             =   3480
         Width           =   300
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   11
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   282
         Text            =   "C *"
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   270
         Index           =   16
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   281
         Text            =   "CC *"
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   14
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   280
         Text            =   "E *"
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00FF8080&
         Caption         =   "Set By 1"
         Height          =   225
         Index           =   3
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   279
         Top             =   2640
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Set By 1"
         Height          =   225
         Index           =   4
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   278
         Top             =   3240
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   15
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   277
         Text            =   "M *"
         Top             =   2640
         Width           =   375
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   31
         Left            =   2955
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   274
         Top             =   2280
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   30
         Left            =   2955
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   270
         Top             =   2880
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   29
         Left            =   2955
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   268
         Top             =   3480
         Width           =   300
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Set By 1"
         Height          =   225
         Index           =   2
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   266
         Top             =   2040
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Set By 1"
         Height          =   225
         Index           =   5
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   265
         Top             =   3870
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   16
         Left            =   2130
         TabIndex        =   264
         Text            =   "1"
         Top             =   2325
         Width           =   855
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   17
         Left            =   2130
         TabIndex        =   263
         Text            =   "1"
         Top             =   2910
         Width           =   855
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0FF&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   18
         Left            =   2130
         TabIndex        =   261
         Text            =   "1"
         Top             =   3510
         Width           =   855
      End
      Begin VB.CheckBox chkAutoSampl 
         BackColor       =   &H00000000&
         Caption         =   "Auto Change"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   283
         Tag             =   "0"
         Top             =   4762
         Width           =   1425
      End
      Begin VB.CommandButton cmdLoadP 
         BackColor       =   &H0080FFFF&
         Caption         =   "Load"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   290
         Top             =   5400
         Width           =   615
      End
      Begin VB.CommandButton cmdSavePa 
         BackColor       =   &H0080FF80&
         Caption         =   "Save"
         Height          =   330
         Left            =   975
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   291
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   615
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   9
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   90
         Style           =   2  'Dropdown List
         TabIndex        =   292
         Top             =   5040
         Width           =   1500
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   315
         Index           =   31
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   293
         Text            =   "Samples"
         Top             =   4140
         Width           =   1425
      End
      Begin VB.CheckBox chkA_Aggr 
         BackColor       =   &H00000000&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   1920
         TabIndex        =   305
         ToolTipText     =   "Auto Set Agg+"
         Top             =   6360
         Width           =   465
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H0080FFFF&
         Height          =   270
         Index           =   19
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   309
         Text            =   "Size"
         Top             =   4440
         Width           =   825
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1EDED&
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4200
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   329
         TabStop         =   0   'False
         Top             =   530
         UseMaskColor    =   -1  'True
         Width           =   570
      End
      Begin VB.CommandButton cmdOpenTelo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "±"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4455
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   325
         ToolTipText     =   "Camera"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   5520
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   324
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdMiniMize 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   4815
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   322
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdLogs 
         BackColor       =   &H00E4E0E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   321
         ToolTipText     =   "Logs"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdCtrl 
         BackColor       =   &H00E4E0E0&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   320
         ToolTipText     =   "Logs"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdMax 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   5175
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   323
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdMini 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   5175
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   328
         Top             =   30
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.PictureBox picBProcs 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   8055
         Left            =   1
         ScaleHeight     =   8025
         ScaleWidth      =   5880
         TabIndex        =   332
         Top             =   7200
         Width           =   5910
         Begin VB.CheckBox chkAlphaEnable 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00473842&
            ForeColor       =   &H00E0E0E0&
            Height          =   225
            Left            =   3720
            TabIndex        =   333
            Top             =   9960
            Width           =   225
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H0061272D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   525
         Left            =   1
         TabIndex        =   276
         ToolTipText     =   "Powerd By Kaveh Abdollahi"
         Top             =   0
         Width           =   8790
      End
      Begin VB.Label lblLQSky 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Skies"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   525
         Left            =   120
         TabIndex        =   327
         ToolTipText     =   "Powered By Kaveh Abdollahi"
         Top             =   -30
         Width           =   3855
      End
      Begin VB.Label cmdAbout 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   120
         TabIndex        =   334
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "  Time Engine"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   230
         Left            =   1920
         TabIndex        =   326
         Top             =   165
         Width           =   1455
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00E0E0E0&
         X1              =   0
         X2              =   8760
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame fraDrive 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   11175
      Left            =   15
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   3930
      Begin VB.CommandButton Command9 
         Caption         =   "Del *.gpg"
         Height          =   255
         Left            =   1920
         TabIndex        =   503
         Top             =   7800
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00787878&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00FFFFFF&
         Height          =   980
         Left            =   0
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   360
         Text            =   "frmBase.frx":BFD9A
         Top             =   5280
         Width           =   2895
      End
      Begin VB.TextBox txtLQT2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00787878&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00FFFFFF&
         Height          =   980
         Left            =   2880
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   359
         Text            =   "frmBase.frx":BFE2E
         Top             =   5280
         Width           =   1010
      End
      Begin VB.CheckBox chkShotM 
         BackColor       =   &H00000000&
         Caption         =   "Scr or Bk"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   840
         TabIndex        =   340
         Top             =   7800
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   3600
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   337
         TabStop         =   0   'False
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox chkRGB_mu 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Set RGB By Music"
         ForeColor       =   &H00E0E0E0&
         Height          =   735
         Left            =   3000
         TabIndex        =   336
         Top             =   4200
         Width           =   705
      End
      Begin VB.CommandButton cmdInvertPage 
         BackColor       =   &H00FFFF00&
         Caption         =   "Inverse Screen"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1800
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   335
         Top             =   2730
         UseMaskColor    =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   12
         Left            =   420
         MousePointer    =   1  'Arrow
         TabIndex        =   308
         Text            =   "1"
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   307
         Top             =   2400
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   12
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   306
         Top             =   2400
         Width           =   300
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   7
         Left            =   30
         TabIndex        =   259
         Text            =   "25796"
         Top             =   3240
         Width           =   945
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H000000FF&
         Caption         =   "Set By"
         Height          =   260
         Index           =   7
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   258
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   945
      End
      Begin VB.CheckBox chkBW 
         BackColor       =   &H00000000&
         Caption         =   "Back White"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   960
         TabIndex        =   257
         Top             =   10920
         Width           =   1185
      End
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H00000000&
         Caption         =   "Motion"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   30
         TabIndex        =   256
         Top             =   10920
         Width           =   825
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   32
         Left            =   1245
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   207
         Top             =   4065
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         HideSelection   =   0   'False
         Index           =   32
         Left            =   330
         TabIndex        =   208
         Text            =   "100"
         Top             =   4065
         Width           =   900
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   32
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   4065
         Width           =   300
      End
      Begin VB.Frame fraShoter 
         Appearance      =   0  'Flat
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   30
         TabIndex        =   225
         Tag             =   "x"
         Top             =   8160
         Width           =   3855
         Begin VB.TextBox TextLabel 
            Appearance      =   0  'Flat
            BackColor       =   &H002D061B&
            DragMode        =   1  'Automatic
            ForeColor       =   &H00F1E4F3&
            Height          =   225
            Index           =   35
            Left            =   2325
            Locked          =   -1  'True
            TabIndex        =   252
            Text            =   "File Saved"
            Top             =   180
            Width           =   855
         End
         Begin VB.CheckBox chkSetDateTimeDir 
            BackColor       =   &H00000080&
            Caption         =   "Set Save Directory By Time"
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   150
            TabIndex        =   251
            ToolTipText     =   "Auto Shot"
            Top             =   510
            Value           =   1  'Checked
            Width           =   1695
         End
         Begin VB.TextBox txtShotCount 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00512D4B&
            BorderStyle     =   0  'None
            ForeColor       =   &H00FFFFFF&
            Height          =   200
            Left            =   1860
            TabIndex        =   229
            Text            =   "0"
            Top             =   180
            Width           =   465
         End
         Begin VB.CommandButton cmdSF 
            BackColor       =   &H0000FF00&
            Caption         =   "Shot a Pic"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   150
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   227
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1695
         End
         Begin VB.CommandButton cmdChandSaveFolder 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Change Save Folder"
            Height          =   330
            Left            =   2190
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   232
            Tag             =   "0"
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   1575
         End
         Begin VB.TextBox txtPath 
            Appearance      =   0  'Flat
            BackColor       =   &H00CFE4E0&
            DragMode        =   1  'Automatic
            Height          =   345
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   228
            Top             =   840
            Width           =   2415
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00F275B0&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   12
               Charset         =   1
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   150
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   245
            Top             =   1800
            Width           =   300
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00F275B0&
            Caption         =   "+"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   12
               Charset         =   1
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   8
            Left            =   3465
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   244
            Top             =   1800
            Width           =   300
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00F275B0&
            Caption         =   "+"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   12
               Charset         =   1
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   10
            Left            =   1575
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   237
            Top             =   1800
            Width           =   300
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00F275B0&
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   12
               Charset         =   1
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   8
            Left            =   1920
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   236
            Top             =   1800
            Width           =   300
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00B4E0CE&
            ForeColor       =   &H001E1E1E&
            Height          =   225
            Left            =   3120
            TabIndex        =   246
            Text            =   "Sec"
            ToolTipText     =   "JPG Quality"
            Top             =   2010
            Width           =   345
         End
         Begin VB.TextBox txtQua 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0C2DA&
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   375
            TabIndex        =   238
            Text            =   "45"
            ToolTipText     =   "JPG Quality"
            Top             =   2010
            Width           =   1215
         End
         Begin VB.TextBox txtspm 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0C2DA&
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   8
            Left            =   2175
            MaxLength       =   4
            OLEDragMode     =   1  'Automatic
            TabIndex        =   240
            Text            =   "2"
            ToolTipText     =   "Auto Shot Interval"
            Top             =   2010
            Width           =   975
         End
         Begin VB.TextBox TextLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00F8AFA9&
            BorderStyle     =   0  'None
            ForeColor       =   &H001E1E1E&
            Height          =   285
            Index           =   3
            Left            =   375
            Locked          =   -1  'True
            TabIndex        =   239
            Text            =   "Photo Quality"
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00F8AFA9&
            BorderStyle     =   0  'None
            ForeColor       =   &H001E1E1E&
            Height          =   285
            Left            =   2175
            Locked          =   -1  'True
            TabIndex        =   241
            Text            =   "Interval"
            ToolTipText     =   "JPG Quality"
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtMaxShot 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00665766&
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   3180
            MousePointer    =   1  'Arrow
            TabIndex        =   243
            Text            =   "100"
            Top             =   1500
            Width           =   500
         End
         Begin VB.TextBox txtMaxShot 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00665766&
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   1
            Left            =   3180
            MousePointer    =   1  'Arrow
            TabIndex        =   242
            Text            =   "300"
            Top             =   1200
            Width           =   500
         End
         Begin VB.CheckBox chkBreakShot 
            BackColor       =   &H00000080&
            Caption         =   "New Dir On "
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   234
            ToolTipText     =   "Auto Shot"
            Top             =   1500
            Width           =   1845
         End
         Begin VB.CheckBox chkBreakShot 
            BackColor       =   &H00000080&
            Caption         =   "Break On"
            ForeColor       =   &H0000FFFF&
            Height          =   270
            Index           =   1
            Left            =   1920
            TabIndex        =   230
            ToolTipText     =   "Auto Shot"
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1845
         End
         Begin VB.CheckBox chkAutoShot 
            BackColor       =   &H00000080&
            Caption         =   "&Auto Shot"
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   150
            TabIndex        =   231
            ToolTipText     =   "Auto Shot"
            Top             =   1200
            Width           =   3525
         End
         Begin VB.CheckBox chkShotAll 
            BackColor       =   &H00000080&
            Caption         =   "&Shot All Frame"
            ForeColor       =   &H0000FFFF&
            Height          =   285
            Left            =   150
            TabIndex        =   235
            ToolTipText     =   "Auto Shot"
            Top             =   1500
            Value           =   1  'Checked
            Width           =   3525
         End
         Begin VB.CommandButton cmdDriectory 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add Driectory"
            Height          =   345
            Left            =   2580
            Style           =   1  'Graphical
            TabIndex        =   233
            Tag             =   "0"
            Top             =   840
            Width           =   1185
         End
         Begin VB.CommandButton cmdShotBx 
            Appearance      =   0  'Flat
            BackColor       =   &H00F1EDED&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3195
            MaskColor       =   &H000040C0&
            Style           =   1  'Graphical
            TabIndex        =   226
            TabStop         =   0   'False
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   570
         End
         Begin VB.TextBox TextLabel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            DragMode        =   1  'Automatic
            ForeColor       =   &H00FFFFFF&
            Height          =   2220
            Index           =   34
            Left            =   60
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   247
            Top             =   60
            Width           =   3760
         End
      End
      Begin VB.CommandButton cmdCls 
         BackColor       =   &H0000FF00&
         Caption         =   "&Clear Screen"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   855
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   155
         Top             =   10560
         UseMaskColor    =   -1  'True
         Width           =   1500
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dr LAST Pot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   210
         Index           =   6
         Left            =   120
         TabIndex        =   224
         Top             =   4680
         Width           =   1260
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FFFF&
         Height          =   0
         Index           =   23
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   210
         Text            =   "Alpha"
         Top             =   0
         Width           =   0
      End
      Begin VB.CheckBox chkPause 
         BackColor       =   &H00000000&
         Caption         =   "Pause"
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   120
         TabIndex        =   158
         Top             =   585
         Width           =   855
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H0000FF00&
         Caption         =   "&Restart"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   10560
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   6
         Left            =   1980
         TabIndex        =   153
         Text            =   "10"
         Top             =   2400
         Width           =   945
      End
      Begin VB.CheckBox chkLastP 
         BackColor       =   &H002D061B&
         Caption         =   "Draw Last Value"
         ForeColor       =   &H00F1E4F3&
         Height          =   315
         Left            =   30
         MaskColor       =   &H000000FF&
         TabIndex        =   209
         Top             =   3765
         Width           =   1545
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00F8AFA9&
         Caption         =   "Set By"
         Height          =   225
         Index           =   6
         Left            =   1980
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   150
         Top             =   2160
         UseMaskColor    =   -1  'True
         Width           =   945
      End
      Begin VB.CommandButton cmdMaxPoints 
         BackColor       =   &H008080FF&
         Caption         =   "Set Finish 148900"
         Height          =   255
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   3480
         Width           =   1520
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   25
         Left            =   3328
         Locked          =   -1  'True
         TabIndex        =   164
         Text            =   "1"
         Top             =   1575
         Width           =   275
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "®"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   163
         Top             =   1515
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   162
         Top             =   1515
         Width           =   330
      End
      Begin MSComctlLib.Slider slrCol 
         Height          =   975
         Index           =   0
         Left            =   2160
         TabIndex        =   377
         Top             =   3960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   3
         Min             =   1
         Max             =   15
         SelStart        =   3
         TickStyle       =   3
         Value           =   3
      End
      Begin MSComctlLib.Slider slrCol 
         Height          =   975
         Index           =   1
         Left            =   2400
         TabIndex        =   378
         Top             =   3960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   3
         Min             =   1
         Max             =   15
         SelStart        =   3
         TickStyle       =   3
         Value           =   3
      End
      Begin MSComctlLib.Slider slrCol 
         Height          =   975
         Index           =   2
         Left            =   2640
         TabIndex        =   379
         Top             =   3960
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   1720
         _Version        =   393216
         Orientation     =   1
         LargeChange     =   3
         Min             =   1
         Max             =   15
         SelStart        =   3
         TickStyle       =   3
         Value           =   3
         TextPosition    =   1
      End
   End
   Begin VB.Frame fraLogs 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "( X ) len of data "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5265
      Left            =   6360
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   8610
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   615
         Left            =   2280
         TabIndex        =   500
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton cmdCopy2Excel 
         Caption         =   "Send To Excel"
         Height          =   390
         Left            =   2220
         Style           =   1  'Graphical
         TabIndex        =   499
         Top             =   4560
         Width           =   735
      End
      Begin VB.CommandButton cmdLogClr 
         BackColor       =   &H000000FF&
         Caption         =   "Clear List"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   498
         Top             =   2520
         Width           =   615
      End
      Begin VB.ListBox lstL 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   4680
         Index           =   3
         ItemData        =   "frmBase.frx":BFE41
         Left            =   7560
         List            =   "frmBase.frx":BFFB6
         TabIndex        =   422
         Top             =   360
         Width           =   1065
      End
      Begin VB.ListBox lstL 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   4680
         Index           =   1
         ItemData        =   "frmBase.frx":C0447
         Left            =   6486
         List            =   "frmBase.frx":C04A8
         TabIndex        =   405
         Top             =   360
         Width           =   1065
      End
      Begin VB.ListBox lstL 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   4680
         Index           =   2
         ItemData        =   "frmBase.frx":C05B2
         Left            =   5414
         List            =   "frmBase.frx":C0613
         TabIndex        =   406
         Top             =   360
         Width           =   1065
      End
      Begin VB.ListBox lstL 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   4680
         Index           =   0
         ItemData        =   "frmBase.frx":C06FF
         Left            =   5062
         List            =   "frmBase.frx":C0760
         TabIndex        =   403
         Top             =   360
         Width           =   345
      End
      Begin VB.TextBox txtLCo 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2280
         TabIndex        =   400
         Text            =   "1000"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtLogs 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00CFE4E0&
         Height          =   4680
         Left            =   3000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   399
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdGetLog2 
         BackColor       =   &H00C9B6D1&
         Caption         =   "Get Log >"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   354
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton cmdGetLog 
         BackColor       =   &H00C9B6D1&
         Caption         =   "< Get Log"
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   9
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   189
         Top             =   15
         Width           =   255
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   15
         Width           =   255
      End
      Begin VB.CheckBox chkALog 
         BackColor       =   &H00000000&
         Caption         =   "Auto"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2280
         MaskColor       =   &H000000FF&
         TabIndex        =   187
         Top             =   1800
         Width           =   735
      End
      Begin VB.ListBox lstLogs 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H00CFE4E0&
         Height          =   4680
         Left            =   120
         TabIndex        =   176
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton cmdHideLogs 
         BackColor       =   &H00E0E0E0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8280
         TabIndex        =   26
         Top             =   45
         Width           =   375
      End
      Begin VB.PictureBox picBLogs 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   5145
         Left            =   0
         ScaleHeight     =   5115
         ScaleWidth      =   8565
         TabIndex        =   24
         Top             =   5040
         Width           =   8595
      End
      Begin VB.Label lblLogs 
         Alignment       =   2  'Center
         BackColor       =   &H002D061B&
         Caption         =   "Logs"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   8490
      End
   End
   Begin VB.Frame fraPad 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   9165
      Left            =   3600
      TabIndex        =   383
      Top             =   1800
      Visible         =   0   'False
      Width           =   7005
      Begin VB.CheckBox chkTxt2PicAN 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   5640
         TabIndex        =   497
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtMailTo 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1440
         TabIndex        =   494
         Top             =   6600
         Width           =   3255
      End
      Begin VB.CommandButton cmdSendMail 
         Caption         =   "Send Mail To :"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   485
         Top             =   6600
         Width           =   1215
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   260
         Index           =   7
         Left            =   4800
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   483
         Text            =   "0"
         Top             =   5160
         Width           =   615
      End
      Begin VB.CommandButton cmdTxtS 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   450
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtA 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   3
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   449
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00EFBCC7&
         Caption         =   "A B J A D"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   479
         Top             =   7080
         Width           =   1095
      End
      Begin VB.TextBox txt4AbjadNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EFBCC7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   478
         Top             =   7560
         Width           =   5055
      End
      Begin VB.TextBox txt4Abjad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EFBCC7&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Left            =   1320
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   477
         Top             =   7080
         Width           =   3855
      End
      Begin VB.CommandButton cmdRemovelstPRG 
         Caption         =   "Remove"
         Height          =   240
         Left            =   6000
         TabIndex        =   476
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txtMain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   475
         Top             =   5160
         Width           =   4575
      End
      Begin VB.CommandButton cmdPrimeIndex 
         BackColor       =   &H00D678CA&
         Caption         =   "Prime (Index)?"
         Height          =   375
         Left            =   5400
         Style           =   1  'Graphical
         TabIndex        =   473
         Top             =   7080
         Width           =   1335
      End
      Begin VB.TextBox txtPrimeIndex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EED0EE&
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   5400
         TabIndex        =   474
         Text            =   "1"
         Top             =   7500
         Width           =   1335
      End
      Begin VB.TextBox txtPrimeIndex 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00EED0EE&
         ForeColor       =   &H00000000&
         Height          =   1215
         Index           =   1
         Left            =   5280
         MultiLine       =   -1  'True
         TabIndex        =   472
         Text            =   "frmBase.frx":C07E8
         Top             =   7800
         Width           =   1575
      End
      Begin VB.ListBox lstPRG 
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
         Height          =   1200
         ItemData        =   "frmBase.frx":C07EA
         Left            =   4800
         List            =   "frmBase.frx":C07EC
         TabIndex        =   471
         Top             =   5400
         Width           =   2055
      End
      Begin VB.TextBox txtTextSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   5
         Left            =   6360
         TabIndex        =   461
         Text            =   "28"
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton cmdTxtS 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   460
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtPad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   5
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   459
         Top             =   4440
         Width           =   5955
      End
      Begin VB.CheckBox chkTxt2Pic 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   5
         Left            =   6120
         TabIndex        =   458
         Top             =   4680
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtA 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   0
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   457
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtA 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   1
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   445
         Top             =   2550
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtA 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   2
         Left            =   6615
         Style           =   1  'Graphical
         TabIndex        =   447
         Top             =   3270
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtS 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   444
         Top             =   4680
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtA 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   5
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   443
         Top             =   4680
         Width           =   255
      End
      Begin VB.TextBox txtPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Andalus"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   411
         Text            =   "frmBase.frx":C07EE
         Top             =   960
         Width           =   5955
      End
      Begin VB.CheckBox chkTxt2Pic 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   441
         Top             =   1830
         Width           =   255
      End
      Begin VB.CheckBox chkTxt2Pic 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   6120
         TabIndex        =   440
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkTxt2Pic 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   6120
         TabIndex        =   439
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox chkTxt2Pic 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   2
         Left            =   6120
         TabIndex        =   438
         Top             =   3270
         Width           =   255
      End
      Begin VB.CheckBox chkTxt2Pic 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   437
         Top             =   2550
         Width           =   255
      End
      Begin VB.TextBox txtNumber 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   3120
         TabIndex        =   421
         Text            =   "0"
         Top             =   533
         Width           =   615
      End
      Begin VB.TextBox txtPad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   3
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   419
         Top             =   3720
         Width           =   5955
      End
      Begin VB.TextBox txtPad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   416
         Top             =   3000
         Width           =   5955
      End
      Begin VB.Timer Timer_AutoNext 
         Interval        =   10000
         Left            =   120
         Tag             =   "0"
         Top             =   0
      End
      Begin VB.CheckBox chkAutoNext 
         BackColor       =   &H00000000&
         Caption         =   "Auto Next"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   415
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtPad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   1
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   414
         Top             =   2280
         Width           =   5955
      End
      Begin VB.CheckBox chkText2Pic 
         BackColor       =   &H00000000&
         Caption         =   "Enable Set To Screen"
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   409
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H0000FF00&
         Caption         =   "ÞÑÂä"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   388
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.CommandButton cmdTextSet 
         Caption         =   "Add2List"
         Height          =   240
         Left            =   4800
         TabIndex        =   387
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txtPad 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   0
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   386
         Top             =   1560
         Width           =   5955
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1EDED&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   6645
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   384
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdTxtS 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   446
         Top             =   2550
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtS 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   448
         Top             =   3270
         Width           =   255
      End
      Begin VB.TextBox txtTextSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   1
         Left            =   6360
         TabIndex        =   453
         Text            =   "28"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtTextSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   0
         Left            =   6360
         TabIndex        =   425
         Text            =   "40"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtTextSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   2
         Left            =   6360
         TabIndex        =   454
         Text            =   "28"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtTextSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   3
         Left            =   6360
         TabIndex        =   455
         Text            =   "28"
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton cmdTxtS 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   452
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton cmdTxtA 
         Appearance      =   0  'Flat
         BackColor       =   &H009BD3D5&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   4
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   451
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtTextSize 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   225
         Index           =   4
         Left            =   6360
         TabIndex        =   456
         Text            =   "30"
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         Caption         =   "Pad"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   495
         Left            =   0
         TabIndex        =   385
         ToolTipText     =   "Powerd By Kaveh Abdollahi"
         Top             =   0
         Width           =   7095
      End
   End
   Begin VB.Frame fraTelo 
      Appearance      =   0  'Flat
      BackColor       =   &H002D061B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00512D4B&
      Height          =   6210
      Left            =   8520
      TabIndex        =   160
      Top             =   0
      Visible         =   0   'False
      Width           =   5010
      Begin VB.CheckBox chkImg2Pic2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4575
         TabIndex        =   493
         Top             =   5760
         Width           =   255
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ê"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   37
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   489
         Top             =   5790
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   37
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   490
         Top             =   5160
         Width           =   330
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   14
         Left            =   120
         MousePointer    =   1  'Arrow
         TabIndex        =   491
         Text            =   "64"
         Top             =   5497
         Width           =   570
      End
      Begin VB.CommandButton cmdHideScop 
         BackColor       =   &H00E0E0E0&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4650
         TabIndex        =   484
         Top             =   15
         Width           =   375
      End
      Begin VB.CheckBox chkImg2Pic 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4575
         TabIndex        =   481
         Top             =   5520
         Width           =   255
      End
      Begin VB.CommandButton cmdOpenPic 
         Caption         =   "Load Image"
         Height          =   430
         Left            =   3720
         TabIndex        =   480
         Top             =   5565
         Width           =   855
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "è"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   36
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   205
         Top             =   5474
         Width           =   330
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   34
         Left            =   4080
         MousePointer    =   1  'Arrow
         TabIndex        =   198
         Text            =   "512"
         Top             =   5160
         Width           =   570
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0082ADD5&
         Caption         =   "è"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   34
         Left            =   4680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   200
         Top             =   5160
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0082ADD5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   34
         Left            =   3720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   199
         Top             =   5160
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   36
         Left            =   720
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   206
         Top             =   5474
         Width           =   330
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   36
         Left            =   1020
         MousePointer    =   1  'Arrow
         TabIndex        =   204
         Text            =   "64"
         Top             =   5497
         Width           =   690
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "ê"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   35
         Left            =   3105
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   203
         Top             =   5670
         Width           =   570
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   35
         Left            =   3105
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   5160
         Width           =   570
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   15
         Width           =   255
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   6
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   15
         Width           =   255
      End
      Begin VB.CheckBox chkBox 
         BackColor       =   &H00784138&
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   27.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   5640
         Width           =   495
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   35
         Left            =   3105
         MousePointer    =   1  'Arrow
         TabIndex        =   201
         Text            =   "384"
         Top             =   5430
         Width           =   570
      End
      Begin VB.PictureBox picTele 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   4800
         Left            =   75
         ScaleHeight     =   4876.19
         ScaleMode       =   0  'User
         ScaleWidth      =   4876.19
         TabIndex        =   168
         Top             =   360
         Width           =   4800
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0099E1B3&
         Caption         =   "Camera 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   15
         TabIndex        =   161
         ToolTipText     =   "Powerd By Kaveh Abdollahi"
         Top             =   15
         Width           =   5010
      End
   End
   Begin VB.ComboBox DevicesBox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8160
      Style           =   2  'Dropdown List
      TabIndex        =   331
      Top             =   11000
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame fraBlur 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   5295
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H00000000&
         Caption         =   "Blur 1"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   0
         Left            =   2640
         TabIndex        =   152
         Top             =   1680
         Width           =   795
      End
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H00000000&
         Caption         =   "Type 4"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   3
         Left            =   2640
         TabIndex        =   118
         Top             =   1440
         Width           =   795
      End
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H00000000&
         Caption         =   "Type 3"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   2
         Left            =   2640
         TabIndex        =   117
         Top             =   1200
         Width           =   795
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   27
         Left            =   2130
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   115
         Text            =   "Cpu 0"
         Top             =   3720
         Width           =   720
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   26
         Left            =   2835
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   114
         Text            =   "Cpu 1"
         Top             =   3720
         Width           =   720
      End
      Begin VB.TextBox txtProcess1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2835
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   110
         Text            =   "00"
         Top             =   3915
         Width           =   720
      End
      Begin VB.TextBox txtProcess0 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2130
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   109
         Text            =   "00"
         Top             =   3915
         Width           =   720
      End
      Begin VB.TextBox txtProcessSum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   2160
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   116
         Text            =   "00"
         Top             =   3495
         Width           =   1425
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   112
         Text            =   "WIN Cpu Usage"
         Top             =   3270
         Width           =   1425
      End
      Begin VB.ListBox lstPsent 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   2880
         Left            =   2055
         TabIndex        =   100
         Top             =   360
         Width           =   550
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   150
         Index           =   17
         Left            =   1695
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   107
         Text            =   "ms"
         Top             =   3315
         Width           =   250
      End
      Begin VB.TextBox txtEFRM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005B425A&
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1170
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   105
         Text            =   "10.00"
         Top             =   3240
         Width           =   540
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   150
         Index           =   9
         Left            =   1695
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   104
         Text            =   "ms"
         Top             =   3525
         Width           =   250
      End
      Begin VB.TextBox txtDoEvSleep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005B425A&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1170
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   103
         Text            =   "10.00"
         Top             =   3465
         Width           =   540
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   18
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   108
         Text            =   "                Sum : "
         Top             =   3240
         Width           =   2370
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   6
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   111
         Text            =   "Others"
         Top             =   3465
         Width           =   2490
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   113
         Text            =   "Timers"
         Top             =   3720
         Width           =   2475
      End
      Begin VB.CheckBox chkSortP 
         BackColor       =   &H00665766&
         Caption         =   "Sort Lists"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   0
         TabIndex        =   106
         Top             =   4320
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.ListBox lstFunctions 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   2880
         Left            =   120
         TabIndex        =   102
         Top             =   360
         Width           =   1425
      End
      Begin VB.ListBox lstProcess 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   2880
         Left            =   1530
         TabIndex        =   101
         Top             =   360
         Width           =   540
      End
      Begin VB.TextBox txtProcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0FFFF&
         Height          =   168
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   14
         Text            =   "00.00"
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox txtTimeP2Sky 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005B425A&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2640
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   98
         Text            =   "10.00"
         Top             =   4080
         Width           =   540
      End
      Begin VB.Timer Timer_Sky 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1920
         Top             =   4200
      End
      Begin VB.PictureBox picBBlur 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         ForeColor       =   &H80000008&
         Height          =   4260
         Left            =   0
         ScaleHeight     =   4230
         ScaleWidth      =   3465
         TabIndex        =   13
         Top             =   360
         Width           =   3500
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   6
         X1              =   120
         X2              =   2760
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblBlur 
         Alignment       =   2  'Center
         BackColor       =   &H002D061B&
         Caption         =   "Process Time"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   15
         TabIndex        =   10
         Top             =   30
         Width           =   3465
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   2400
         Y1              =   6725
         Y2              =   6725
      End
      Begin VB.Label lblFileName 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "------------"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   9
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame fraFullScr 
      Appearance      =   0  'Flat
      BackColor       =   &H001B171C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      TabIndex        =   190
      Top             =   11325
      Visible         =   0   'False
      Width           =   15240
      Begin VB.TextBox txtFrm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   7800
         MaxLength       =   5
         TabIndex        =   424
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   0
         Width           =   480
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   29
         Left            =   8400
         MousePointer    =   1  'Arrow
         TabIndex        =   423
         Text            =   "fps"
         Top             =   0
         Width           =   255
      End
      Begin VB.Label lVideoInfo 
         Alignment       =   2  'Center
         BackColor       =   &H001B171C&
         Caption         =   "-"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   8760
         TabIndex        =   418
         Top             =   0
         Width           =   2295
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   5
         Left            =   4560
         TabIndex        =   196
         Top             =   15
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   4
         Left            =   3600
         TabIndex        =   195
         Top             =   15
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   3
         Left            =   2640
         TabIndex        =   194
         Top             =   15
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   2
         Left            =   1200
         TabIndex        =   193
         Top             =   15
         Width           =   1215
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   192
         Top             =   15
         Width           =   855
      End
      Begin VB.Label lblFullscr 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "© 2010 Kaveh Abdollahi"
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   13440
         TabIndex        =   250
         Top             =   0
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   8
         Left            =   6600
         TabIndex        =   249
         Top             =   0
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "Auto Shot Off"
         ForeColor       =   &H00C0C0C0&
         Height          =   180
         Index           =   7
         Left            =   5520
         TabIndex        =   248
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label lblFullscr 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Sky 7.7.639  "
         ForeColor       =   &H0080FFFF&
         Height          =   180
         Index           =   0
         Left            =   11880
         TabIndex        =   191
         Top             =   15
         Visible         =   0   'False
         Width           =   1575
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4005
      Locked          =   -1  'True
      TabIndex        =   143
      Text            =   "ETC"
      Top             =   11160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4005
      Locked          =   -1  'True
      TabIndex        =   142
      Text            =   "Blur"
      Top             =   11160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Left            =   4215
      Locked          =   -1  'True
      TabIndex        =   141
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   8
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   140
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   6
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   139
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   5
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   138
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   4
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   137
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   3
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   136
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   1
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   135
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   2
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   134
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtEtcSumD 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3990
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   133
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtWaveSumD 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3990
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   132
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtETCT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   131
      Text            =   "00.00"
      Top             =   11280
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   10
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   130
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   11
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   129
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   12
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   128
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   13
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   127
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   14
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   126
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   15
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   125
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   16
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   124
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   17
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   123
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   18
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   122
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   19
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   121
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   20
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   120
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.Timer Timer_AutoSave 
      Interval        =   100
      Left            =   6960
      Top             =   11040
   End
   Begin VB.Timer Timer_AutoLog 
      Interval        =   200
      Left            =   6600
      Top             =   11040
   End
   Begin VB.Timer Timer_Process 
      Interval        =   100
      Left            =   7680
      Top             =   11040
   End
   Begin VB.PictureBox picViewEE 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1.72800e5
      Left            =   0
      ScaleHeight     =   11520
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   85
      Top             =   0
      Width           =   2.30400e5
      Begin MSComDlg.CommonDialog CDlg 
         Left            =   6120
         Top             =   10920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer Timer_Seconds 
         Interval        =   1000
         Left            =   7320
         Top             =   11040
      End
   End
   Begin VB.PictureBox picBuffEE 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   11520
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   86
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox picBuffEE2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   2  'Dot
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   11520
      ScaleMode       =   0  'User
      ScaleWidth      =   11520
      TabIndex        =   79
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
      Begin VB.Timer Timer_AHeight 
         Interval        =   5
         Left            =   7320
         Top             =   11040
      End
      Begin VB.Line Line3 
         X1              =   90
         X2              =   1260
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   408
      Top             =   10320
      Visible         =   0   'False
      Width           =   2880
   End
   Begin VB.PictureBox picText 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Index           =   1
      Left            =   0
      ScaleHeight     =   2055
      ScaleWidth      =   9000
      TabIndex        =   410
      Top             =   9120
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Andalus"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1500
      Index           =   0
      Left            =   0
      ScaleHeight     =   1500
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   412
      Top             =   8040
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1500
      Index           =   1
      Left            =   0
      ScaleHeight     =   1500
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   413
      Top             =   9000
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1500
      Index           =   2
      Left            =   0
      ScaleHeight     =   1500
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   417
      Top             =   9720
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   8180
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3350
      Begin VB.CheckBox chkCM 
         BackColor       =   &H00512D4B&
         Caption         =   " C = M"
         Enabled         =   0   'False
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Left            =   720
         TabIndex        =   222
         Top             =   1440
         Width           =   1140
      End
      Begin VB.CheckBox chkLock 
         BackColor       =   &H00512D4B&
         Caption         =   " M = E"
         Enabled         =   0   'False
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Left            =   720
         MaskColor       =   &H000000FF&
         TabIndex        =   221
         Top             =   1200
         Width           =   1140
      End
      Begin VB.CheckBox chkAGrow 
         BackColor       =   &H00512D4B&
         Caption         =   "Grow"
         Enabled         =   0   'False
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Left            =   720
         TabIndex        =   220
         Top             =   1680
         Width           =   1140
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   219
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   218
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   19
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   216
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   19
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   1080
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00B9D3B8&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         HideSelection   =   0   'False
         Index           =   10
         Left            =   1080
         MousePointer    =   1  'Arrow
         TabIndex        =   197
         Text            =   "1"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkInverse 
         BackColor       =   &H0000FFFF&
         Caption         =   "Coloring"
         ForeColor       =   &H001E1E1E&
         Height          =   225
         Left            =   1680
         MaskColor       =   &H000000FF&
         TabIndex        =   154
         Top             =   3720
         Width           =   1005
      End
      Begin VB.CheckBox chkFallCol 
         BackColor       =   &H00000000&
         Caption         =   "Fall Colors"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Left            =   240
         MaskColor       =   &H000000FF&
         TabIndex        =   99
         Top             =   3780
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CommandButton cmdSnCGr 
         BackColor       =   &H00FF8080&
         Caption         =   "B"
         Height          =   250
         Index           =   2
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3195
         Width           =   680
      End
      Begin VB.CommandButton cmdSnCGr 
         BackColor       =   &H0080FF80&
         Caption         =   "G"
         Height          =   250
         Index           =   1
         Left            =   1275
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3195
         Width           =   680
      End
      Begin VB.CommandButton cmdSnCGr 
         BackColor       =   &H000000FF&
         Caption         =   "R"
         Height          =   250
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3195
         Width           =   680
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   0
         Left            =   600
         MousePointer    =   1  'Arrow
         TabIndex        =   19
         Text            =   "2"
         Top             =   3480
         Width           =   675
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   2
         Left            =   1920
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Text            =   "2"
         Top             =   3480
         Width           =   675
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   1
         Left            =   1260
         MousePointer    =   1  'Arrow
         TabIndex        =   20
         Text            =   "2"
         Top             =   3480
         Width           =   675
      End
      Begin VB.CommandButton cmdRGB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set RGB With 0"
         Height          =   250
         Index           =   10
         Left            =   600
         TabIndex        =   15
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox txtMinC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "0"
         Top             =   2880
         Width           =   680
      End
      Begin VB.TextBox txtMinC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   915
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "0"
         Top             =   2880
         Width           =   680
      End
      Begin VB.TextBox txtMinC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   2880
         Width           =   680
      End
      Begin VB.TextBox txtMaxC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "255"
         Top             =   2685
         Width           =   680
      End
      Begin VB.TextBox txtMaxC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   915
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "255"
         Top             =   2685
         Width           =   680
      End
      Begin VB.TextBox txtMaxC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "255"
         Top             =   2685
         Width           =   680
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   20
         Left            =   235
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   11
         Text            =   "RGB Limiter"
         Top             =   2490
         Width           =   2040
      End
      Begin VB.PictureBox picBCol 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         ForeColor       =   &H80000008&
         Height          =   3730
         Left            =   0
         ScaleHeight     =   3705
         ScaleWidth      =   3315
         TabIndex        =   12
         Top             =   360
         Width           =   3345
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00C0C0C0&
            Height          =   250
            Index           =   4
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   148
            Top             =   1680
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00473842&
            Height          =   250
            Index           =   3
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   147
            Top             =   1410
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00665766&
            Height          =   250
            Index           =   2
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   146
            Top             =   1140
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00FFFFFF&
            Height          =   250
            Index           =   1
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   145
            Top             =   870
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00000000&
            Height          =   250
            Index           =   0
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   144
            Top             =   600
            Width           =   680
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   3330
         X2              =   3330
         Y1              =   -4545
         Y2              =   -475
      End
      Begin VB.Label lblColorSet 
         Alignment       =   2  'Center
         BackColor       =   &H00322732&
         Caption         =   "Colors"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   1
         Top             =   15
         Width           =   3330
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   3
         X1              =   338
         X2              =   2978
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1500
      Index           =   3
      Left            =   0
      ScaleHeight     =   1500
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   420
      Top             =   8040
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Andalus"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1500
      Index           =   4
      Left            =   0
      ScaleHeight     =   1500
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   442
      Top             =   9120
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Andalus"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1500
      Index           =   5
      Left            =   0
      ScaleHeight     =   1500
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   462
      Top             =   9000
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Index           =   6
      Left            =   11400
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   3840
      TabIndex        =   470
      Top             =   11160
      Visible         =   0   'False
      Width           =   3840
   End
   Begin VB.PictureBox picStore 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawMode        =   14  'Copy Pen
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   9000
      ScaleWidth      =   9000
      TabIndex        =   374
      Top             =   0
      Visible         =   0   'False
      Width           =   9000
   End
   Begin VB.Frame fraAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   7560
      TabIndex        =   156
      Top             =   5400
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdCloseAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1EDED&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   3600
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   78
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   6585
         Left            =   2160
         Picture         =   "frmBase.frx":C0814
         ScaleHeight     =   6555
         ScaleWidth      =   1800
         TabIndex        =   487
         Top             =   -1680
         Width           =   1830
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00322732&
            Caption         =   "Liquid Skies"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FF80&
            Height          =   465
            Left            =   -2160
            TabIndex        =   488
            ToolTipText     =   "Powerd By Kaveh Abdollahi"
            Top             =   1680
            Width           =   3975
         End
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         Caption         =   "Liquid Skies"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   465
         Left            =   0
         TabIndex        =   486
         ToolTipText     =   "Powerd By Kaveh Abdollahi"
         Top             =   0
         Width           =   3975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   0
         X2              =   4080
         Y1              =   4965
         Y2              =   4965
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   0
         X2              =   4080
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrighjt© 2010  Kaveh Abdollahi"
         ForeColor       =   &H00CFE4E0&
         Height          =   330
         Index           =   5
         Left            =   2520
         TabIndex        =   255
         Top             =   5040
         Width           =   1335
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "kavehplus@gmail.com"
         ForeColor       =   &H00CFE4E0&
         Height          =   165
         Index           =   4
         Left            =   120
         TabIndex        =   254
         Top             =   5205
         Width           =   1695
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "HiPerP.com"
         ForeColor       =   &H00CFE4E0&
         Height          =   135
         Index           =   3
         Left            =   120
         TabIndex        =   253
         Top             =   5040
         Width           =   1335
      End
   End
   Begin VB.PictureBox picStore2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      DrawMode        =   14  'Copy Pen
      ForeColor       =   &H80000008&
      Height          =   8385
      Left            =   0
      Picture         =   "frmBase.frx":C6043
      ScaleHeight     =   8385
      ScaleWidth      =   7500
      TabIndex        =   492
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
   Begin VB.PictureBox pic2Text 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   540
      Index           =   7
      Left            =   120
      ScaleHeight     =   540
      ScaleMode       =   0  'User
      ScaleWidth      =   4560
      TabIndex        =   495
      Top             =   120
      Visible         =   0   'False
      Width           =   4560
   End
   Begin VB.Frame fraStart 
      BackColor       =   &H00000000&
      Height          =   1815
      Left            =   5040
      TabIndex        =   393
      Top             =   3180
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2160
         TabIndex        =   395
         Text            =   "Waiting ..............."
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtStart 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1140
         TabIndex        =   394
         Text            =   "."
         Top             =   1200
         Width           =   2880
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Between    1   -    150,000,000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   397
         Top             =   480
         Width           =   3720
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Generating Prime Numbers"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   375
         Index           =   0
         Left            =   720
         TabIndex        =   396
         Top             =   160
         Width           =   3720
      End
   End
   Begin VB.Frame fraControls 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "( X ) len of data "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9120
      Left            =   11640
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   3480
      Begin VB.CommandButton CmdDefault 
         BackColor       =   &H0080FFFF&
         Caption         =   "Set By Defaults Setting"
         Height          =   345
         Left            =   1080
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   8400
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0099E1B3&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   360
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   5955
         Width           =   1065
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D696A2&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   360
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   6390
         Width           =   1065
      End
      Begin VB.TextBox txtFpS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Cordia New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   360
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   174
         Text            =   "30"
         Top             =   6120
         Width           =   825
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   1095
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   173
         Top             =   6150
         Width           =   330
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0070616C&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   8
         Left            =   360
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   172
         Text            =   "Speed Set on             FPS"
         Top             =   5760
         Width           =   1905
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   9
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   171
         Text            =   "100"
         Top             =   5760
         Width           =   435
      End
      Begin VB.CommandButton cmdNormalSize 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "#"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   159
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00665766&
         Caption         =   "Draw"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   0
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   119
         Top             =   3960
         Width           =   915
      End
      Begin VB.CheckBox chktest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         Caption         =   "High Light"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   2520
         TabIndex        =   80
         Top             =   8760
         Width           =   210
      End
      Begin VB.CheckBox chkScript 
         BackColor       =   &H00000000&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   180
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   8460
         Width           =   780
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         DragMode        =   1  'Automatic
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   12
         Left            =   720
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   96
         Top             =   4185
         Width           =   480
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         DragMode        =   1  'Automatic
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   13
         Left            =   720
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   95
         Top             =   4425
         Width           =   480
      End
      Begin VB.CheckBox chkP4Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   2
         Left            =   3240
         MaskColor       =   &H000000FF&
         TabIndex        =   93
         Top             =   4440
         Width           =   195
      End
      Begin VB.CheckBox chkP4Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   1
         Left            =   2820
         MaskColor       =   &H000000FF&
         TabIndex        =   92
         Top             =   4440
         Width           =   435
      End
      Begin VB.CheckBox chkP4Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   0
         Left            =   2520
         MaskColor       =   &H000000FF&
         TabIndex        =   89
         Top             =   4440
         Width           =   315
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00665766&
         Caption         =   "Clr Draw"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   7
         Left            =   1200
         MaskColor       =   &H000000FF&
         TabIndex        =   87
         Top             =   4440
         Width           =   1275
      End
      Begin VB.CheckBox chkP3Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   2
         Left            =   3240
         MaskColor       =   &H000000FF&
         TabIndex        =   91
         Top             =   4200
         Width           =   195
      End
      Begin VB.CheckBox chkP3Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   1
         Left            =   2820
         MaskColor       =   &H000000FF&
         TabIndex        =   90
         Top             =   4200
         Width           =   435
      End
      Begin VB.CheckBox chkP3Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   0
         Left            =   2520
         MaskColor       =   &H000000FF&
         TabIndex        =   74
         Top             =   4200
         Width           =   315
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00473842&
         Caption         =   "P3"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   1
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   94
         Top             =   4200
         Width           =   1035
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00473842&
         Caption         =   "Clr Draw"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   5
         Left            =   1200
         MaskColor       =   &H000000FF&
         TabIndex        =   78
         Top             =   4200
         Width           =   1275
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00665766&
         Caption         =   "P4"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   6
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   88
         Top             =   4440
         Width           =   1035
      End
      Begin VB.TextBox txtLScr 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFC0&
         Height          =   165
         Left            =   447
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "00.00"
         Top             =   2040
         Width           =   425
      End
      Begin VB.TextBox txtBR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   3090
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   84
         Text            =   "2"
         Top             =   1635
         Width           =   200
      End
      Begin VB.TextBox txtBLR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2257
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   83
         Text            =   "2"
         Top             =   1635
         Width           =   240
      End
      Begin VB.TextBox txtBL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   1470
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   82
         Text            =   "2"
         Top             =   1635
         Width           =   200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test"
         Height          =   255
         Left            =   2505
         TabIndex        =   81
         Top             =   8760
         Width           =   855
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   7575
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7380
         Width           =   300
      End
      Begin VB.CheckBox chkClrAlter 
         BackColor       =   &H00800080&
         Caption         =   "Alternate Clear"
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   56
         Top             =   5160
         Width           =   1395
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2110
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2513
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   6
         Left            =   480
         MaxLength       =   5
         MousePointer    =   1  'Arrow
         TabIndex        =   38
         Text            =   "1"
         Top             =   2230
         Width           =   675
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1140
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2230
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   195
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2230
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2520
         MaxLength       =   5
         MousePointer    =   1  'Arrow
         TabIndex        =   50
         Text            =   "15"
         Top             =   3360
         Width           =   555
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FAF1F3&
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   2400
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   37
         Text            =   "1"
         Top             =   2520
         Width           =   675
      End
      Begin VB.CheckBox chkABalance 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         Height          =   210
         Left            =   2040
         TabIndex        =   34
         Top             =   3352
         Width           =   210
      End
      Begin VB.CheckBox chkAHeight 
         BackColor       =   &H00473842&
         Caption         =   "Auto Balance Height"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Left            =   180
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   3357
         Width           =   2055
      End
      Begin VB.TextBox txtBL2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   1710
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   72
         Text            =   "2"
         Top             =   1485
         Width           =   200
      End
      Begin VB.TextBox txtBLR2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2257
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   71
         Text            =   "2"
         Top             =   1485
         Width           =   240
      End
      Begin VB.TextBox txtBLR3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2257
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   70
         Text            =   "2"
         Top             =   1335
         Width           =   240
      End
      Begin VB.TextBox txtBand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   1950
         MousePointer    =   1  'Arrow
         TabIndex        =   69
         Text            =   "2"
         Top             =   495
         Width           =   200
      End
      Begin VB.TextBox txtBand3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   1950
         MousePointer    =   1  'Arrow
         TabIndex        =   68
         Text            =   "2"
         Top             =   825
         Width           =   200
      End
      Begin VB.TextBox txtBandAvg2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   67
         Text            =   "2"
         Top             =   660
         Width           =   195
      End
      Begin VB.TextBox txtBandAvg1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   66
         Text            =   "2"
         Top             =   495
         Width           =   195
      End
      Begin VB.TextBox txtBandAvg3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Text            =   "2"
         Top             =   825
         Width           =   195
      End
      Begin VB.TextBox txtBand2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   1950
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Text            =   "2"
         Top             =   660
         Width           =   200
      End
      Begin VB.TextBox txtBL3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   1950
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   63
         Text            =   "2"
         Top             =   1335
         Width           =   200
      End
      Begin VB.TextBox txtBR3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2595
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   62
         Text            =   "2"
         Top             =   1335
         Width           =   200
      End
      Begin VB.TextBox txtBR2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2835
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   61
         Text            =   "2"
         Top             =   1485
         Width           =   200
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2790
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   59
         Text            =   "Midl"
         Top             =   660
         Width           =   300
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2940
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   58
         Text            =   "Bass"
         Top             =   480
         Width           =   360
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2595
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   57
         Text            =   "Treble"
         Top             =   840
         Width           =   435
      End
      Begin VB.TextBox txtProcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0FFFF&
         Height          =   168
         Index           =   9
         Left            =   1395
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   46
         Text            =   "00.00"
         Top             =   390
         Width           =   410
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   54
         Text            =   "2"
         Top             =   7380
         Width           =   795
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   33
         Text            =   "4"
         Top             =   7575
         Width           =   795
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2400
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   36
         Text            =   "256"
         Top             =   2760
         Width           =   675
      End
      Begin VB.TextBox txtProcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0FFFF&
         Height          =   168
         Index           =   7
         Left            =   2472
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   30
         Text            =   "00.00"
         Top             =   7800
         Width           =   410
      End
      Begin VB.CheckBox chkAdjFreq 
         BackColor       =   &H00473842&
         Caption         =   "Adjustment (X Level)"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   73
         Top             =   7365
         Width           =   1875
      End
      Begin VB.CheckBox chkAdjFreq 
         BackColor       =   &H00473842&
         Caption         =   "Adjustment (Z Level)"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   60
         Top             =   7590
         Width           =   1875
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   25
         Left            =   195
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   55
         Text            =   "Sec of Last Freq  Match In Screen "
         Top             =   2025
         Width           =   3150
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3360
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3360
         Width           =   300
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   24
         Left            =   180
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   48
         Text            =   "Base Heigth of Scope  "
         Top             =   2745
         Width           =   1920
      End
      Begin VB.CheckBox chkInc 
         BackColor       =   &H00665766&
         Caption         =   "Increase Scope Height"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   47
         Top             =   2985
         Width           =   3195
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   2
         Left            =   180
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   45
         Text            =   "Scopes Count  "
         Top             =   2505
         Width           =   1920
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2760
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2115
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2760
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2513
         Width           =   300
      End
      Begin VB.CheckBox chkTransparent 
         BackColor       =   &H00473842&
         Caption         =   "Transparent All Panels"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   1560
         TabIndex        =   35
         Top             =   5160
         Width           =   1935
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7575
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   7380
         Width           =   300
      End
      Begin VB.PictureBox PicFFT 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H001E1E1E&
         DrawWidth       =   2
         FillColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00665766&
         Height          =   1470
         Left            =   1350
         Negotiate       =   -1  'True
         ScaleHeight     =   117.517
         ScaleMode       =   0  'User
         ScaleWidth      =   40.306
         TabIndex        =   75
         Top             =   390
         Width           =   2010
      End
      Begin VB.PictureBox picBCtrl 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         ForeColor       =   &H80000008&
         Height          =   8760
         Left            =   0
         ScaleHeight     =   8730
         ScaleWidth      =   3450
         TabIndex        =   76
         Top             =   360
         Width           =   3475
         Begin VB.CheckBox chkATime 
            BackColor       =   &H00800000&
            Caption         =   "LQ Time 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   186
            Top             =   6240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "+1"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   185
            Top             =   6240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   184
            Top             =   6240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   183
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "100"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   182
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "1000"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   181
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "1000"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   180
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtspm 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00F8AFA9&
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   13
            Left            =   1800
            MaxLength       =   6
            MousePointer    =   1  'Arrow
            TabIndex        =   179
            Text            =   "1"
            Top             =   6000
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtLQT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0070616C&
            DragMode        =   1  'Automatic
            ForeColor       =   &H00E0E0E0&
            Height          =   705
            Left            =   1800
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   178
            Top             =   6240
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CheckBox chkPR 
            BackColor       =   &H00004040&
            Caption         =   "View All Available Threads"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   2
            Left            =   1800
            MaskColor       =   &H000000FF&
            TabIndex        =   177
            Top             =   6240
            Visible         =   0   'False
            Width           =   1635
         End
      End
      Begin VB.Label lblControls 
         Alignment       =   2  'Center
         BackColor       =   &H00665766&
         Caption         =   "   Controls"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   77
         Top             =   30
         Width           =   3450
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         X1              =   1200
         X2              =   3180
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   2
         X1              =   15
         X2              =   3465
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   4
         X1              =   15
         X2              =   3470
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   5
         X1              =   360
         X2              =   3000
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   7
         X1              =   360
         X2              =   3000
         Y1              =   4200
         Y2              =   4200
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   -120
      X2              =   -120
      Y1              =   -2640
      Y2              =   -840
   End
End
Attribute VB_Name = "frmBase"
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
Private Wave As WaveHdr, te As Long, te2 As Long


Public Sub KeyPrss()
    If GetAsyncKeyState(vbKeyS) <> 0 Then cmdSF_Click
    If GetAsyncKeyState(vbKeyC) <> 0 Then KaCls
    If GetAsyncKeyState(vbKeyR) <> 0 Then cmd0_Click 10
    
    If GetAsyncKeyState(vbKeyDivide) <> 0 Then cmdPrevius_Click: Sleep (50)
    If GetAsyncKeyState(vbKeyMultiply) <> 0 Then cmdNextS_Click: Sleep (50)
    If GetAsyncKeyState(vbKeyAdd) <> 0 And chkText2Pic.Value = 0 Then cmdLarger_Click 33: Sleep (50)
    If GetAsyncKeyState(vbKeySubtract) <> 0 And chkText2Pic.Value = 0 Then cmdSmaler_Click 33: Sleep (50)
  
    If (GetAsyncKeyState(vbKeyPageUp) <> 0 Or GetAsyncKeyState(vbKeyAdd) <> 0) And chkText2Pic.Value = 1 Then
        
        If frmQuran.lstBase.ListIndex < frmQuran.lstBase.ListCount - 1 Then
            frmQuran.lstBase.ListIndex = frmQuran.lstBase.ListIndex + 1
           Else
            frmQuran.lstBase.ListIndex = 0
        End If
'        DoEvents
        frmQuran.cmdSet_Click
    End If
    If (GetAsyncKeyState(vbKeyPageDown) <> 0 Or GetAsyncKeyState(vbKeySubtract) <> 0) And chkText2Pic.Value = 1 Then
        
        If frmQuran.lstBase.ListIndex > 0 Then
            frmQuran.lstBase.ListIndex = frmQuran.lstBase.ListIndex - 1
           Else
            frmQuran.lstBase.ListIndex = frmQuran.lstBase.ListCount - 1
        End If
'        DoEvents
        frmQuran.cmdSet_Click
    End If
    
    If GetAsyncKeyState(vbKeyPageUp) <> 0 And chkText2Pic.Value = 0 Then cmdLarger_Click 20
    If GetAsyncKeyState(vbKeyPageDown) <> 0 And chkText2Pic.Value = 0 Then cmdSmaler_Click 20
    
    If GetAsyncKeyState(vbKeyV) <> 0 Then ChkDraw_Click 4
    
    If GetAsyncKeyState(vbKeyF4) <> 0 Then chkAvalue(0).Value = chkAvalue(0).Value Xor 1: Sleep (50) 'time++
    If GetAsyncKeyState(vbKeyF2) <> 0 Then chkAutoShot.Value = chkAutoShot.Value Xor 1: Sleep (50)
    If GetAsyncKeyState(vbKeyF3) <> 0 Then chkAvalue(0).Value = chkAvalue(0).Value Xor 1: chkAvalue(1).Value = chkAvalue(0).Value Xor 1: Sleep (50) 'time++ to 'time--
    If GetAsyncKeyState(vbKeyF5) <> 0 Then chkAutoShot.Value = chkAutoShot.Value Xor 1: chkAvalue(0).Value = chkAutoShot.Value: Sleep (50)
 
    If GetAsyncKeyState(vbKeyEscape) <> 0 And cmdMini.Tag = "1" Then Unload Me
    If GetAsyncKeyState(vbKeyEscape) <> 0 Then cmdMini.Tag = "1": cmdMini_Click
    
   
'   DoEvents
End Sub


Private Sub chkAutoMax_Click()
    If chkAutoMax Then
        cmdMaxPoints.Enabled = False
        txtspm(21).Enabled = False
        cmdLarger(21).Enabled = False
        cmdSmaler(21).Enabled = False
        txtRST(7).Enabled = False
        cmd0(7).Enabled = False
      Else
        cmdMaxPoints.Enabled = True
        txtspm(21).Enabled = True
        cmdLarger(21).Enabled = True
        cmdSmaler(21).Enabled = True
        txtRST(7).Enabled = True
        cmd0(7).Enabled = True
     End If
End Sub

Private Sub chkAutoShot_Click()
    If chkAutoShot Then
        Timer_AutoSave.Enabled = True
        lblFullscr(7).Caption = "Auto Shot On"
        lblFullscr(7).ForeColor = &HFFFF&
        lblFullscr(7).FontBold = True
    Else
        Timer_AutoSave.Enabled = False
        lblFullscr(7).Caption = "Auto Shot Off"
        lblFullscr(7).ForeColor = &HC0C0C0
        lblFullscr(7).FontBold = False
    End If
End Sub

Private Sub chkAvalue_Click(Index As Integer)
    If chkAvalue(1).Value = 1 And chkAvalue(0).Value = 1 Then chkAvalue(Index Xor 1) = 0
End Sub

Private Sub chkBGQ_Click(Index As Integer)
    If chkBGQ(1).Value = 1 And Index = 1 Then chkBGQ(0).Value = 0
    If chkBGQ(0).Value = 1 And Index = 0 Then chkBGQ(1).Value = 0
End Sub

Private Sub chkClrAlter_Click()
    minY = 0: maxY = 768
    kaAltrCls
End Sub

Private Sub chkCM_Click()
    If chkCM Then txtspm(18) = txtspm(17)
End Sub

Private Sub chkCol_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        CDlg.ShowColor
        coB = CDlg.Color
        chkCol(0).BackColor = coB
        chkCol(1).BackColor = coB
        chkCol(2).BackColor = coB
        chkCol(3).BackColor = coB
        chkCol(4).BackColor = coB
        chkCol(5).BackColor = coB
    End If
    coFl = True
End Sub

Private Sub chkCol_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim A As Byte
    If chkCol(Index) = 1 Then
        For A = 0 To chkCol.count - 1
           chkCol(A) = 0
        Next A
     chkCol(Index) = 1
    End If
End Sub

Private Sub ChkDraw_Click(Index As Integer)
Dim x As Integer
    If Index = 0 Or Index = 1 Or Index = 6 And ChkDraw(Index) = 0 Then ' cmdCls_Click
         picTmp.ForeColor = vbBlack
        For x = 1 To 20
            Polyline picTmp.hdc, PtL(0, x), 512
        Next x
    End If
End Sub

Private Sub chkDrawCntr_Click(Index As Integer)
    
    Select Case Index
        
        Case 0:
        
        Case 1:
        
        Case 3:
        
        Case 4:
    
    End Select

End Sub


Private Sub chkInc_Click()
    If chkInc Then
      chkInc.ForeColor = &HFFFF&
      chkInc.BackColor = &HFF&
     Else
      chkInc.ForeColor = &HE0E0E0
      chkInc.BackColor = &H0&
    End If
End Sub
'

Private Sub chkLock_Click()
    If chkLock Then txtspm(16) = 1 / txtspm(17)
End Sub

Private Sub chkP1_Click()
    
End Sub

Private Sub chkP_Click(Index As Integer)
Dim x
    If chkP(Index).Value = 0 Then Exit Sub
    For x = 0 To 3
        If x <> Index Then chkP(x).Value = 0
    Next x
End Sub

Private Sub chkP3Opt_Click(Index As Integer)
    If Index = 2 Then
      chkP3Opt(0) = chkP3Opt(0) * -1 + 1
      chkP3Opt(1) = chkP3Opt(1) * -1 + 1
    End If
End Sub

Private Sub chkP4Opt_Click(Index As Integer)
    If Index = 2 Then
      chkP4Opt(0) = chkP4Opt(0) * -1 + 1
      chkP4Opt(1) = chkP4Opt(1) * -1 + 1
    End If
End Sub

Private Sub chkPant_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 And chkPant(0) Then
        chkPant(1).Value = 0: chkPant(2).Value = 0
    ElseIf Index = 1 And chkPant(1) Then
        chkPant(0).Value = 0: chkPant(2).Value = 0
    ElseIf Index = 2 And chkPant(2) Then
        chkPant(0).Value = 0: chkPant(1).Value = 0
    Else
        chkPant(Index).Value = 1
    End If
End Sub

Private Sub chkTimeEnable_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 4 Then chkTimeEnable(0).Value = 0 ': chkTimeEnable(2).Value = 0
'    If Index = 2 Then chkTimeEnable(0).Value = 0: chkTimeEnable(4).Value = 0
    If Index = 0 Then chkTimeEnable(4).Value = 0
End Sub

Private Sub chkTransparent_Click()
 
 If chkTransparent Then Exit Sub
    
      picBCtrl.Cls
      picBCol.Cls
      picBBlur.Cls
      picBProcs.Cls
      picBLogs.Cls
 
End Sub


Private Sub cmd0_Click(Index As Integer)
    
    If Index = 0 Then LQT2 = txtRST(0): txtspm(11) = txtRST(0)
    If Index = 1 Then txtspm(2) = txtRST(1)
    If Index = 2 Then txtspm(16) = 1 ' txtRST(2)
    If Index = 3 Then txtspm(17) = 1 ' txtRST(3)
    If Index = 4 Then txtspm(18) = 1 ' txtRST(4)
    If Index = 5 Then txtspm(19) = 1 ' txtRST(5)
    If Index = 6 Then txtspm(20) = txtRST(6)
    If Index = 7 Then txtspm(21) = txtRST(7)
    
    If Index = 10 Then
        KaCls
        txtspm(13) = 0: LQT = 0
        txtspm(11) = txtRST(0): LQT2 = txtRST(0)
    End If

End Sub

Private Sub cmdAbout_Click()
   If fraAbout.Visible = False Then
       fraAbout.Visible = True
       fraAbout.Left = 15
       fraAbout.Top = 15
       fraAbout.ZOrder 0
    Else
       fraAbout.Visible = False
   End If
End Sub

Private Sub cmdBackCol_Click(Index As Integer)
    picTmp.BackColor = cmdBackCol(Index).BackColor
End Sub


Private Sub cmdChandSaveFolder_Click()
Dim s As String
    If sPath <> "" Then
        s = BrowseForFolder(bPath, Me.hWnd, "Select Folder For Save Image")  'App.Path & "\"
    Else
        s = BrowseForFolder("c:\", Me.hWnd, "Select Folder For Save Image") 'App.Path & "\"
    End If
    If s <> "" Then sPath = s: bPath = s
txtPath = sPath
End Sub

Private Sub cmdCloseAbout_Click()
    fraAbout.Visible = False
End Sub

Private Sub cmdCls_Click()
    KaCls
End Sub

Private Sub cmdCtrl_Click()
    If fraControls.Visible = False Then
        fraControls.Visible = True
        fraBlur.Visible = True
    Else
        fraControls.Visible = False
        fraBlur.Visible = False
    End If
End Sub
'

Private Sub cmdDriectory_Click()
On Error Resume Next
    If cmdDriectory.Tag = "0" Then
        MkDir bPath & "\" & Trim(Date$)
        sPath = bPath & "\" & Trim(Date$)
        bPath = sPath
        cmdDriectory.Tag = "1"
    End If
    
        MkDir bPath & "\" & Left(Replace(Time$, ":", "-"), 5)
        sPath = bPath & "\" & Left(Replace(Time$, ":", "-"), 5)

    txtPath = sPath
End Sub

Private Sub cmdExit_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdGetLog_Click()
Dim x As Long, dist As Currency, d As Currency
    lstLogs.Clear
    lstLogs.AddItem PXY1(1, 0) & " , " & PXY1(1, 1)
    For x = 2 To Sz1_b - 1
        d = GetDistance(PXY1(x, 0), PXY1(x, 1), PXY1(x - 1, 0), PXY1(x - 1, 1))
        dist = dist + d
        lstLogs.AddItem PXY1(x, 0) & " , " & PXY1(x, 1) & " , " & d * txtspm(33) & " , " & dist * txtspm(33)
    Next x
End Sub

Private Sub cmdGetLog2_Click()
    Dim x As Integer, s As String
    txtLogs.Visible = False
    cmdGetLog2.Visible = False
    txtLogs = ""
    txtLCo = Val(txtLCo.Text)
    DoEvents
        For x = 1 To txtLCo
            s = s & PrK(2, x) & vbCrLf
        Next x
    txtLogs = s
    txtLogs.Visible = True
    cmdGetLog2.Visible = True
End Sub

Private Sub cmdHideLogs_Click()
    cmdLogs_Click
End Sub

Private Sub cmdBlureOpen_Click()
Dim CommonDialog1 As OSDialog
Set CommonDialog1 = New OSDialog

' Examples:-
  Dim title$, Filt$, InDir$, FileSpec$, CurrPath$
  Dim FIndex As Long

'  LOAD egs
   title$ = "Blur Files"
   Filt$ = "Blur Files (*.BLR)|*.BLR|All files (*.*)|*.*" '"Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
'   Filt$ = "Open vbp (*.vbp)|*.vbp|All files (*.*)|*.*"
   FileSpec$ = ""
   InDir$ = CurrPath$ 'Pathspec$
'   Set CommonDialog1 = New OSDialog

   CommonDialog1.ShowOpen FileSpec$, title$, Filt$, InDir$, "", Me.hWnd, FIndex
'   FIndex = 1 bmp
'   FIndex = 2 jpg
'   etc

   Set CommonDialog1 = Nothing

'  SAVE eg
'   Title$ = "Save Mask as 2-color bmp"
'   Filt$ = "Save bmp|*.bmp"
'   InDir$ = CurrPath$ 'Pathspec$
'   FileSpec$=""
'   Set CommonDialog1 = New OSDialog
'   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd
'   Set CommonDialog1 = Nothing
'
'   Len(FileSpec$)=0 for cancel
'''''''''''''''''''''''''''''''''''''''''
  Dim intf, x As Integer
  Dim s As String
'
'    CommonDialog1.Filter = "Blur Files (*.BLR)|*.BLR"
'    CommonDialog1.CancelError = True
'    On Error GoTo ErrHandler
'    CommonDialog1.ShowOpen
'
'    If CommonDialog1.FileName <> "" Then
'        MousePointer = 11
'        '        LoadNewFile (CommonDialog1.FileName)
'        intf = FreeFile
'        If Trim(FileSpec$) = "" Then Exit Sub
'        Open FileSpec$ For Input As #intf
'        lstFa.Clear
'        lstFaName.Clear
'        While Not EOF(intf)
'            Input #intf, S
'            lstFa.AddItem Trim(S)
'            Input #intf, S
'            lstFaName.AddItem Trim(S)
'        Wend
'        Close #intf
'        S = ""
'        For x = Len(CommonDialog1.FileName) To 1 Step -1
'            If Mid(CommonDialog1.FileName, x, 1) = "\" Then Exit For
'            DoEvents
'        Next x
'        lblFileName.Caption = Mid(CommonDialog1.FileName, x + 1, Len(CommonDialog1.FileName) - x - 4)
'        picBuff.Cls
'        MousePointer = 0
'        Exit Sub '---> Bottom
'    End If
'ErrHandler:
'
'Exit Sub

End Sub

Private Sub cmdHideScop_Click()
    fraTelo.Visible = False
    fraTelo.ZOrder 0
    chkBox.Value = 0
End Sub

Public Sub cmdInvertPage_Click()
    KaInvert 0, 0, ResX, ResY
    reAl = True
End Sub

Public Sub cmdLarger_Click(Index As Integer)
On Error Resume Next
Dim ct As Integer
    If Index = 31 Then txtspm(16) = txtspm(16) + Val(txtGR2)
    If Index = 30 Then txtspm(17) = txtspm(17) + Val(txtGR2)
    If Index = 29 Then txtspm(18) = txtspm(18) + Val(txtGR2)
    If Index = 27 Then txtspm(19) = txtspm(19) + Val(txtGR2)
    
    If Index = 32 Then txtspm(32) = txtspm(32) + 1
   
    If Index = 33 And txtspm(33) < 3 Then
        txtspm(33) = txtspm(33) + 0.1: txtspm(33).Refresh
      ElseIf Index = 33 Then
        txtspm(33) = txtspm(33) + 1: txtspm(33).Refresh
    End If
    

    
    If Index = 34 Then txtspm(34) = txtspm(34) + 8
    If Index = 35 Then txtspm(35) = txtspm(35) + 8
    If Index = 36 Then txtspm(36) = txtspm(36) + 4
    If Index = 37 Then txtspm(14) = txtspm(14) + 4
    
    If Index = 28 Then txtspm(28) = txtspm(28) + 1
    
    If Index = 20 Then txtspm(20) = txtspm(20) + Val(txtGR)
      
    If Index = 21 Then txtspm(21) = txtspm(21) + 1: txtspm(21).Refresh
    If Index = 22 Then txtspm(22) = txtspm(22) + 16: txtspm(22).Refresh
    If Index = 23 Then txtspm(23) = txtspm(23) + 16: txtspm(23).Refresh
    If Index = 25 Then txtspm(25) = txtspm(25) + 0.1: txtspm(25).Refresh
    
    If Index = 0 And Val(txtspm(0).Text) < 768 Then txtspm(0).Text = Val(txtspm(0).Text) + 32
    If Index = 1 And txtspm(1) < 10 Then txtspm(1).Text = Val(txtspm(1).Text) + 1
    If Index = 2 Then txtspm(2) = txtspm(2) + 0.0001
    If Index = 24 Then txtspm(2) = txtspm(2) + 0.01
    If Index = 26 Then txtspm(2) = txtspm(2) + 0.1
    If Index = 3 Then txtspm(3) = txtspm(3) + 1
    If Index = 4 Then txtspm(4) = txtspm(4) + 1
    If Index = 5 And Val(txtspm(5).Text) < 50 Then txtspm(5).Text = Val(txtspm(5).Text) + 0.15 '* ((11 - txtspm(5)) \ 10 + 1)
    If Index = 6 And txtspm(6) < 4 Then txtspm(6) = txtspm(6) + 0.1:    txtspm(6).Refresh
    If Index = 7 And txtspm(7) < 255 Then txtspm(7).Text = Val(txtspm(7).Text) + 1
    If Index = 8 And txtspm(8) < 100 Then txtspm(8).Text = Val(txtspm(8).Text) + 0.1
    If Index = 9 Then txtspm(9) = txtspm(9) + 1
    If Index = 10 And txtQua < 100 Then txtQua = txtQua + 5
    If Index = 11 Then LQT2 = LQT2 + txtspm(2): txtspm(11) = LQT2
    
    If Index = 12 And txtspm(12) < 16 Then txtspm(12) = txtspm(12) + 0.1
    If Index = 15 And LQT < 148931 Then LQT = LQT + 1:  txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 14 And LQT < 148931 - 1000 Then LQT = LQT + 1000: txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 13 And LQT < 148931 - 100 Then
       LQT = LQT + 100
       txtspm(13) = LQT
       txtspm(13).Refresh: frmBase.txtLQT.Refresh
'       DoEvents
   End If

End Sub

Private Sub cmdLarger_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    idxL = Index
    DoClickL = True
    DoL = 0
End Sub

Private Sub cmdLarger_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    DoClickL = False
End Sub

Private Sub cmdLoadSig_Click()

Dim x As Long, y As Long, intf As Integer, By As Byte, BBy As Byte
    intf = FreeFile
    Open txtPath & "\" & txtMaxShot(0) & ".kpi" For Random As #intf Len = 1
    BBy = 2
    For x = 1 To FileLen(txtPath & "\" & txtMaxShot(0) & ".kpi")
        Get #intf, x, By
        PrK(2, x) = By
        If By > BBy Then BBy = By
        PrK(3, x) = BBy
    Next x

    Close #intf
End Sub

Private Sub cmdLogClr_Click()
    lstLogs.Clear
    txtLogs = ""
End Sub

Private Sub cmdLogs_Click()
    If fraLogs.Visible = True Then
            fraLogs.Visible = False
    Else
        fraLogs.Top = 360
        fraLogs.Visible = True
        fraLogs.ZOrder 0
    End If
End Sub

Private Sub cmdMaxPoints_Click()
    txtspm(21) = 148900
End Sub

Private Sub cmdMiniMize_Click()

    frmBase.WindowState = 1

End Sub

Private Sub cmdMax_Click()
    cmdMax.Visible = False
    cmdMini.Visible = True
    fraProcess.Visible = False
    fraLogs.Visible = False
    fraPad.Visible = False
End Sub

Private Sub cmdMini_Click()
    cmdMax.Visible = True
    cmdMini.Visible = False
    fraProcess.Visible = True
End Sub

Private Sub cmdMsgBx_Click(Index As Integer)
    If Index = 0 Then fraMsgBx.Visible = False: fraMsgBx.Tag = "x"
    If Index = 1 Then fraMsgBx.Tag = "o": cmdSavePa_Click
    If Index = 2 Then fraMsgBx.Tag = "n": cmdSavePa_Click
End Sub

Private Sub cmdNav_Click(Index As Integer)
    If Index = 6 Then fraTelo.Left = Screen.Width - fraTelo.Width: fraTelo.Top = 0
    If Index = 7 Then fraTelo.Left = Screen.Width - fraTelo.Width * 2: fraTelo.Top = 0
    
    If Index = 9 Then fraLogs.Left = Screen.Width - fraLogs.Width: fraLogs.Top = 360
    If Index = 8 Then fraLogs.Left = 0: fraLogs.Top = 360
End Sub


Private Sub cmdNextS_Click()
    If Combo2.ListIndex < txtRecCo(0) Then
        Combo2.ListIndex = Combo2.ListIndex + 1
    Else
        Combo2.ListIndex = 0
    End If
    cmdLoadP_Click
    Shock = True
End Sub

Private Sub cmdNormalSize_Click()
If cmdNormalSize.Tag <> "0" Then
    frmBase.Width = frmBase.Width - 3000
    frmBase.Height = frmBase.Height - 2000
    
    picViewEE.Width = frmBase.Width
    picViewEE.Height = frmBase.Height
    picViewEE.Left = 0
    picViewEE.Top = 0
    
    fraControls.Left = fraControls.Left - 3000
    fraColors.Left = fraColors.Left - 2000
    fraBlur.Left = fraBlur.Left - 1000
    fraControls.Top = picViewEE.Top
    fraColors.Top = picViewEE.Top
    fraBlur.Top = picViewEE.Top
    fraProcess.Top = 0 ' picViewEE.Top
    
    cmdNormalSize.Tag = "0"
Else
    frmBase.Width = frmBase.Width + 3000
    frmBase.Height = frmBase.Height + 2000
    picViewEE.Width = frmBase.Width
    picViewEE.Height = frmBase.Height
    picViewEE.Left = 0
    picViewEE.Top = 0
    
    fraControls.Left = fraControls.Left + 3000
    fraColors.Left = fraColors.Left + 2000
    fraBlur.Left = fraBlur.Left + 1000
    fraControls.Top = picViewEE.Top
    fraColors.Top = picViewEE.Top
    fraBlur.Top = picViewEE.Top
    fraProcess.Top = 0 ' picViewEE.Top
    
'    txtFrm.Left = txtFrm.Left + 3000
'    txtFrm.Top = picViewEE.Top
    cmdNormalSize.Tag = "1"
End If
 
End Sub


Private Sub cmdOpenShoter_Click()
    fraShoter.Top = 9000
    fraShoter.Left = 45
    fraShoter.Visible = True
    fraShoter.ZOrder 0
End Sub

Private Sub cmdOpenTelo_Click()
    Nvg(1) = 286: Nvg(2) = 15: Nvg(3) = 738
    If fraTelo.Visible = False Then
        fraTelo.Left = Screen.Width - fraTelo.Width
        fraTelo.Visible = True
        fraTelo.ZOrder 0
     Else
        fraTelo.Visible = False
        chkBox.Value = 0
    End If
    
End Sub

Private Sub cmdPad_Click()
    If fraPad.Visible = False Then
        If chkAutoShot.Value = 1 Then chkAutoShot.Tag = "2"
        chkAutoShot.Value = 0
        
        fraPad.Visible = True
        fraPad.Left = 60
        fraPad.Height = 9000
        fraPad.Top = 1000
        fraPad.ZOrder 0
    Else
        fraPad.Visible = False
        If chkAutoShot.Tag = "2" Then chkAutoShot.Tag = "0": chkAutoShot.Value = 1
    End If
End Sub

Private Sub cmdPrevius_Click()
    If Combo2.ListIndex > 0 Then
        Combo2.ListIndex = Combo2.ListIndex - 1
    Else
        Combo2.ListIndex = txtRecCo(0)
    End If
    cmdLoadP_Click
    
    Shock = True
End Sub

Private Sub cmdPrimeIndex_Click()
On Error Resume Next
    txtPrimeIndex(1) = Primes(txtPrimeIndex(0))
    txt4AbjadNum = txt4AbjadNum & txtPrimeIndex(1) & vbCrLf
End Sub

Private Sub cmdRemovelstPRG_Click()
Dim s As String, intf As Integer, x As Integer, Path As String
    If lstPRG.ListIndex < 0 Then Exit Sub
    lstPRG.RemoveItem lstPRG.ListIndex
    
    Path = App.Path & "\PRG.txt"
    intf = FreeFile
    Open Path For Output As #intf
    For x = 0 To lstPRG.ListCount - 1
        Print #intf, lstPRG.List(x)
        DoEvents
    Next x
    Close #intf
End Sub

Private Sub cmdRGB_Click(Index As Integer)
    chRGB (Index)
End Sub

Private Sub cmdRGB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    fraColors.ZOrder 0
End Sub

Private Sub cmdRColSt_Click()
Dim A As Byte
    For A = 0 To chkCol.count - 1
       chkCol(A) = 0
    Next A
    Colv_R = Primes(LQT) Mod 256
    Colv_G = Primes(LQT - 1) Mod 256
    Colv_B = Primes(LQT - 2) Mod 256
    frmBase.chkCol(Rnd * 5) = 1
    
    A = 0
    Do While cS(0) = 0 Or cS(1) = 0 Or cS(2) = 0
        cS(0) = Rnd * 1: cS(1) = Rnd * 1: cS(2) = Rnd * 1
        A = A + 1
        If A > 200 Then cS(0) = 1: cS(1) = 1: cS(2) = 1: Exit Do
    Loop

End Sub

Private Sub cmdRstCol_Click()

End Sub

Private Sub cmdSaveSig_Click()
Dim x As Long, y As Long, intf As Integer, By As Byte
    intf = FreeFile
    Open txtPath & "\" & txtMaxShot(0) & ".kpi" For Random As #intf Len = 1
     
    For x = 1 To 100000
        By = PrK(2, x)
        Put #intf, x, By
    Next x
       
    Close #intf
    
End Sub
'
'Private Sub cmdSaveAVI_Click()
'On Error Resume Next
'    DoEvents
'    BUILD_AVI App.Path & "\tmp\", 16, Me.hWnd, lVideoInfo
'    DoEvents
'
'    Kill App.Path & "\tmp\*.*"
'    Frame = 0
'End Sub

Private Sub cmdSendMail_Click()
    SendMail txtMain, txt4AbjadNum, txtMailTo
End Sub

Private Sub cmdSetLev_Click()

If List1.ListIndex < 1 Then List1.ListIndex = 1

    If List1.ListIndex < List1.ListCount - 1 Then
        List1.ListIndex = List1.ListIndex + 1
        List1_DblClick
    Else
        List1.ListIndex = 1
        List1_DblClick
    End If
    
End Sub

Private Sub cmdShot_Click()
    cmdSF_Click
End Sub

Private Sub cmdShotBx_Click()
    fraShoter.Visible = False
End Sub

Public Sub cmdSmaler_Click(Index As Integer)
Dim cd As Integer
On Error Resume Next
    If Index = 31 Then txtspm(16) = txtspm(16) - Val(txtGR2)
    If Index = 30 Then txtspm(17) = txtspm(17) - Val(txtGR2)
    If Index = 29 Then txtspm(18) = txtspm(18) - Val(txtGR2)
    If Index = 27 Then txtspm(19) = txtspm(19) - Val(txtGR2)
  
    If Index = 33 And txtspm(33) > 0.1 Then
        If txtspm(33) <= 3 Then
             txtspm(33) = txtspm(33) - 0.1: txtspm(33).Refresh
          Else
             txtspm(33) = txtspm(33) - 1: txtspm(33).Refresh
        End If
    End If
    
    
    If Index = 32 And txtspm(32) > 0 Then txtspm(32) = txtspm(32) - 1
    
    If Index = 34 Then txtspm(34) = txtspm(34) - 8
    If Index = 35 Then txtspm(35) = txtspm(35) - 8
    If Index = 36 Then txtspm(36) = txtspm(36) - 4
    If Index = 37 Then txtspm(14) = txtspm(14) - 4
  
    If Index = 28 And txtspm(28) > 1 Then txtspm(28) = txtspm(28) - 1
    
    If Index = 20 And txtspm(20) - Val(txtGR) >= 1 Then txtspm(20) = txtspm(20) - Val(txtGR)
    
    If Index = 21 Then txtspm(21) = txtspm(21) - 1: txtspm(21).Refresh
    If Index = 22 Then txtspm(22) = txtspm(22) - 16: txtspm(22).Refresh
    If Index = 23 Then txtspm(23) = txtspm(23) - 16: txtspm(23).Refresh
    If Index = 25 Then txtspm(25) = txtspm(25) - 0.1: txtspm(25).Refresh
    
    If Index = 0 And Val(txtspm(0).Text) > 32 Then txtspm(0).Text = Val(txtspm(0).Text) - 32: txtspm(0).Refresh
    If Index = 1 And Val(txtspm(1).Text) > 1 Then txtspm(1).Text = Val(txtspm(1).Text) - 1
    If Index = 2 Then txtspm(2) = txtspm(2) - 0.0001
    If Index = 24 Then txtspm(2) = txtspm(2) - 0.01
    If Index = 26 Then txtspm(2) = txtspm(2) - 0.1
    If Index = 3 And txtspm(3) > 1 Then txtspm(3).Text = Val(txtspm(3).Text) - 1
    If Index = 4 And txtspm(4) > 1 Then txtspm(4).Text = Val(txtspm(4).Text) - 1
    If Index = 5 And txtspm(5) > 0.15 Then txtspm(5).Text = Val(txtspm(5).Text) - 0.15 '* (txtspm(5) \ 10 + 1)
    If Index = 6 And txtspm(6) > 0 Then txtspm(6) = txtspm(6) - 0.1: txtspm(6).Refresh
    If Index = 7 And txtspm(7) Then txtspm(Index).Text = Val(txtspm(Index).Text) - 1
    If Index = 8 And txtspm(8) > 0.1 Then txtspm(8).Text = Val(txtspm(8).Text) - 0.1
    If Index = 9 And txtspm(9) > 1 Then txtspm(9) = txtspm(9) - 1
    If Index = 10 And txtQua > 5 Then txtQua = txtQua - 5
    If Index = 11 Then LQT2 = LQT2 - txtspm(2): txtspm(11) = LQT2
    
    If Index = 12 And txtspm(12) > 0.1 Then txtspm(12) = txtspm(12) - 0.1
    If Index = 15 And LQT > 1 Then LQT = LQT - 1:  txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 14 And LQT > 1001 Then LQT = LQT - 1000: txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 13 And LQT > 101 Then
        LQT = LQT - 100:  txtspm(13) = LQT
        txtspm(13).Refresh: frmBase.txtLQT.Refresh
'        DoEvents
    End If
End Sub

Private Sub cmdSmaler_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    idxS = Index
    DoClickS = True
    DoS = 0
End Sub

Private Sub cmdSmaler_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    DoClickS = False
End Sub

Public Sub cmdSF_Click()
Dim s As String, b As Boolean, x As Integer
Dim s2 As String, s3 As String
    cmdSF.BackColor = vbWhite
    cmdSF.Enabled = False
    txtShotCount = Val(txtShotCount) + 1
    
    s2 = txtspm(19) & "-" & txtspm(18) & "-" & txtspm(17) & _
          "-" & txtspm(16) & "-" & Int(LQT2)
        
    If txtShotCount Mod txtMaxShot(0).Text = 0 And chkBreakShot(0) Then
        cmdDriectory_Click
    End If
    SaveCount = SaveCount + 1
    
    If Trim(sPath) = "" Then sPath = BrowseForFolder("e:\", Me.hWnd, "Select Folder For Save Image")  'App.Path & "\"
    If Len(Trim(sPath)) < 2 Then Exit Sub
    
    s = sPath & "\HPP-" & CStr(SaveCount) & "-" & s2 & ".jpg"
    
    chkPause.Value = 1
    
    
    If chkShotM Then
        SaveJpeg s, txtQua, picView
    Else
        SaveJpeg s, txtQua, picBuffEE
    End If
    
    chkPause.Value = 0
    
    SaveSetting "KV_M_B", "kvvisulation", "SaveCount", SaveCount
    SaveSetting "KV_M_B", "kvvisulation", "sPath", sPath
    
    
    cmdSF.Enabled = True
    cmdSF.BackColor = &HFF00&
    If chkBreakShot(1) And chkAutoShot And Val(txtMaxShot(1)) <= Val(txtShotCount) Then
        chkAutoShot.Value = 0
        txtShotCount = "0"
    End If
End Sub


Private Sub cmdTextSet_Click()
Dim s As String, intf As Integer, x As Integer, Path As String
   
    lstPRG.AddItem txtMain.Text
    Path = App.Path & "\PRG.txt"
    intf = FreeFile
    Open Path For Output As #intf
    For x = 0 To lstPRG.ListCount - 1
        Print #intf, lstPRG.List(x)
        DoEvents
    Next x
    Close #intf

End Sub

Private Sub cmdTxtA_Click(Index As Integer)
    txtTextSize(Index) = txtTextSize(Index) + 1
End Sub

Private Sub cmdTxtS_Click(Index As Integer)
    If txtTextSize(Index) > 0 Then txtTextSize(Index) = txtTextSize(Index) - 1
End Sub

Private Sub Combo1_Click()
    Combo1_Validate True
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex > 0 Then picBuff.DrawMode = Combo1.ListIndex + 1
End Sub
Private Sub Combo2_Click()
    Combo2_Validate True
End Sub

Public Sub cmdCopy2Excel_Click()
Dim x As Integer, z1 As Double, z2 As Double                '''' only for test KRandom()
  If lstLogs.ListCount > 0 Then
   ReDim ExelArray(0 To lstLogs.ListCount, 0 To 6)
  Else
   ReDim ExelArray(0 To 55000, 0 To 6)
  End If
    For x = 0 To UBound(ExelArray) - 1
        ExelArray(x, 0) = lstLogs.List(x)
    Next x

  ExcelSaveArray
End Sub

Private Sub cmdLoadP_Click()
Dim i As Integer, intf As Integer, A As Integer, b As Integer
On Error Resume Next
chkPant(0).Value = 0: chkPant(1).Value = 0: chkPant(2).Value = 0
    
    A = Combo2.ListIndex
        
        For b = 0 To 5
            chkCol(b).Value = Smp(A).Chk(b)
        Next b  '6
        
        For b = 0 To 7
             chkTimeEnable(b).Value = Smp(A).Chk(b + 6)
        Next b  '14
        
        chkShotM.Value = Smp(A).Chk(14)
        
        i = 16
        chkAlphaEnable.Value = Smp(A).Chk(i): i = i + 1
        chkAlpha.Value = Smp(A).Chk(i): i = i + 1
        chkAutoMax.Value = Smp(A).Chk(i): i = i + 1
        chkLastP.Value = Smp(A).Chk(i): i = i + 1
        chkAutoFix.Value = Smp(A).Chk(i): i = i + 1
        chkPant(0).Value = Smp(A).Chk(i): i = i + 1
        chkPant(1).Value = Smp(A).Chk(i): i = i + 1
        Combo1.ListIndex = Smp(A).Chk(i): i = i + 1
        ChkDraw(4).Value = Smp(A).Chk(i): i = i + 1
        chkZx(0).Value = Smp(A).Chk(i): i = i + 1
        chkZx(1).Value = Smp(A).Chk(i): i = i + 1
        chkZx(2).Value = Smp(A).Chk(i): i = i + 1
        chkRGB_mu.Value = Smp(A).Chk(i): i = i + 1
        chkP(0).Value = Smp(A).Chk(i): i = i + 1
        chkP(1).Value = Smp(A).Chk(i): i = i + 1
        chkP(2).Value = Smp(A).Chk(i): i = i + 1
        chkP(3).Value = Smp(A).Chk(i): i = i + 1
        chkA_Aggr.Value = Smp(A).Chk(i): i = i + 1
        chkBGQ(2).Value = Smp(A).Chk(i): i = i + 1
        chkBGQ(3).Value = Smp(A).Chk(i): i = i + 1
        
        txtspm(11) = Smp(A).sp(0): LQT = Smp(A).sp(0)
        txtspm(2) = Smp(A).sp(1)
        txtspm(20) = Smp(A).sp(2)
        txtspm(21) = Smp(A).sp(3)
        txtspm(28) = Smp(A).sp(4)
        txtspm(32) = Smp(A).sp(5)
        txtspm(16) = Smp(A).sp(6)
        txtspm(17) = Smp(A).sp(7)
        txtspm(18) = Smp(A).sp(8)
        txtspm(19) = Smp(A).sp(9)
        txtspm(7) = 255 ' Smp(A).sp(10)
        txtspm(22) = ResX \ 2 'Smp(A).sp(11)
        txtspm(23) = ResY \ 2 ' Smp(A).sp(12)
        txtspm(25) = Smp(A).sp(13)
        txtspm(33) = Smp(A).sp(14)
        txtspm(12) = Smp(A).sp(22)
        
        txtRST(0) = Smp(A).sp(15)
        txtRST(1) = Smp(A).sp(16)
        txtRST(6) = Smp(A).sp(17)
        txtRST(7) = Smp(A).sp(18)
        
        txtGR = Smp(A).sp(23)
        txtGR2 = Smp(A).sp(24)
        
        slrCol(0).Value = Smp(A).sp(19): slrCol(0).Refresh
        slrCol(1).Value = Smp(A).sp(20): slrCol(1).Refresh
        slrCol(2).Value = Smp(A).sp(21): slrCol(2).Refresh
        
        If chkPant(0).Value + chkPant(1).Value + chkPant(2).Value = 0 Then chkPant(2).Value = 1
        
        cmd0_Click (0)
        cmdCls_Click
End Sub

Private Sub cmdSavePa_Click()
Dim A As Integer, b As Integer, i As Integer, intf As Integer
    
    txtMsgBx = Combo2.ListIndex
    fraMsgBx.ZOrder 0
    fraMsgBx.Top = 4440
    fraMsgBx.Left = 45
    fraMsgBx.Visible = True
    If fraMsgBx.Tag = "x" Then Exit Sub

    If fraMsgBx.Tag = "o" Then A = Combo2.ListIndex
    If fraMsgBx.Tag = "n" Then
        A = Combo2.ListIndex
        intf = FreeFile
        Open App.Path & "\" & "PSamp.bin" For Random As intf Len = Len(Smp(0))
            Put #intf, FileLen(App.Path & "\" & "PSamp.bin") \ Len(Smp(0)) + 1, Smp(A)
            txtRecCo(0) = Combo2.ListCount - 1
            Combo2.ListIndex = Combo2.ListCount - 1
            Combo2.AddItem A & "-*-" & Round(Smp(A).sp(6), 2) & "," & Round(Smp(A).sp(7), 2) & "," & Round(Smp(A).sp(8), 2) & "," & Round(Smp(A).sp(9), 2)
            txtRecCo(0) = Combo2.ListCount - 1
            Combo2.ListIndex = Combo2.ListCount - 1
            fraMsgBx.Tag = "x"
        Close #intf
    End If
      
On Error Resume Next
'    A = Combo2.ListIndex
    i = 0
    For b = 0 To 5
        Smp(A).Chk(b) = chkCol(b).Value
    Next b
    For b = 0 To 7
        Smp(A).Chk(b + 6) = chkTimeEnable(b).Value
    Next b
    Smp(A).Chk(14) = chkShotM.Value
    
    i = 16
    If Combo1.ListIndex < 0 Then Combo1.ListIndex = 12
    
    Smp(A).Chk(i) = chkAlphaEnable.Value: i = i + 1
    Smp(A).Chk(i) = chkAlpha.Value: i = i + 1
    Smp(A).Chk(i) = chkAutoMax.Value: i = i + 1
    Smp(A).Chk(i) = chkLastP.Value: i = i + 1
    Smp(A).Chk(i) = chkAutoFix.Value: i = i + 1
    Smp(A).Chk(i) = chkPant(0).Value: i = i + 1
    Smp(A).Chk(i) = chkPant(1).Value: i = i + 1
    Smp(A).Chk(i) = Combo1.ListIndex: i = i + 1
    Smp(A).Chk(i) = ChkDraw(4).Value: i = i + 1
    Smp(A).Chk(i) = chkZx(0).Value: i = i + 1
    Smp(A).Chk(i) = chkZx(1).Value: i = i + 1
    Smp(A).Chk(i) = chkZx(2).Value: i = i + 1
    Smp(A).Chk(i) = chkRGB_mu.Value: i = i + 1
    Smp(A).Chk(i) = chkP(0).Value: i = i + 1
    Smp(A).Chk(i) = chkP(1).Value: i = i + 1
    Smp(A).Chk(i) = chkP(2).Value: i = i + 1
    Smp(A).Chk(i) = chkP(3).Value: i = i + 1
    Smp(A).Chk(i) = chkA_Aggr.Value: i = i + 1
    Smp(A).Chk(i) = chkBGQ(2).Value: i = i + 1
    Smp(A).Chk(i) = chkBGQ(3).Value: i = i + 1
    
    Smp(A).sp(0) = Round(txtspm(11), 16)
    Smp(A).sp(1) = Round(txtspm(2), 16)
    Smp(A).sp(2) = Round(txtspm(20), 16)
    Smp(A).sp(3) = Round(txtspm(21), 16)
    Smp(A).sp(4) = Round(txtspm(28), 16)
    Smp(A).sp(5) = Round(txtspm(32), 16)
    Smp(A).sp(6) = Round(txtspm(16), 16)
    Smp(A).sp(7) = Round(txtspm(17), 16)
    Smp(A).sp(8) = Round(txtspm(18), 16)
    Smp(A).sp(9) = Round(txtspm(19), 16)
    Smp(A).sp(10) = 255 ' Round(txtspm(7), 16)
    Smp(A).sp(11) = Round(txtspm(22), 16)
    Smp(A).sp(12) = Round(txtspm(23), 16)
    Smp(A).sp(13) = Round(txtspm(25), 16)
    Smp(A).sp(14) = Round(txtspm(33), 16)
    
    Smp(A).sp(15) = Round(txtRST(0), 16)
    Smp(A).sp(16) = Round(txtRST(1), 16)
    Smp(A).sp(17) = Round(txtRST(6), 16)
    Smp(A).sp(18) = Round(txtRST(7), 16)
    
    Smp(A).sp(19) = slrCol(0)
    Smp(A).sp(20) = slrCol(1)
    Smp(A).sp(21) = slrCol(2)
    Smp(A).sp(22) = txtspm(12)
    
    Smp(A).sp(23) = txtGR
    Smp(A).sp(24) = txtGR2
    
    intf = FreeFile
    Open App.Path & "\" & "PSamp.bin" For Random As intf Len = Len(Smp(0))
    For A = 0 To FileLen(App.Path & "\" & "PSamp.bin") \ Len(Smp(0)) - 1 '50
        Put #intf, A + 1, Smp(A)
        Combo2.List(A) = A & "-*-" & Round(Smp(A).sp(6), 2) & "," & Round(Smp(A).sp(7), 2) & "," & Round(Smp(A).sp(8), 2) & "," & Round(Smp(A).sp(9), 2)
    Next A
    Close #intf
    
    TextLabel(31) = "Saved"
    TextLabel(31).Tag = "10"

    fraMsgBx.Visible = False
    fraMsgBx.Tag = "x"
End Sub
Private Sub LoadFPara()
    Dim A As Integer, b As Integer, i As Integer, intf As Integer
    Dim sn As String
    
'    CDlg.ShowOpen
'    sn = CDlg.filename
    sn = App.Path & "\" & "PSamp.bin"
    intf = FreeFile
    Combo2.Clear

    Open sn For Random As intf Len = Len(Smp(0))
    For A = 1 To FileLen(sn) \ Len(Smp(0)) - 1
        Get #intf, A, Smp(A)
        Combo2.AddItem A & "-*-" & Round(Smp(A).sp(6), 2) & "," & Round(Smp(A).sp(7), 2) & "," & Round(Smp(A).sp(8), 2) & "," & Round(Smp(A).sp(9), 2)
    Next A
    Close #intf
    Combo2.Refresh
End Sub
Private Sub Combo2_Validate(Cancel As Boolean)
    Dim A As Integer, b As Integer, i As Integer, intf As Integer
    intf = FreeFile
    Open App.Path & "\" & "PSamp.bin" For Random As intf Len = Len(Smp(0))
    For A = 0 To FileLen(App.Path & "\" & "PSamp.bin") \ Len(Smp(0)) - 1
        Get #intf, A + 1, Smp(A)
        Combo2.List(A) = A & "-*-" & Round(Smp(A).sp(6), 2) & "," & Round(Smp(A).sp(7), 2) & "," & Round(Smp(A).sp(8), 2) & "," & Round(Smp(A).sp(9), 2)
    Next A
    Close #intf
    Combo2.Refresh
    txtRecCo(1) = Combo2.ListIndex
End Sub

Private Sub Command2_Click()
'    frmCam.Show 1, frmBase
End Sub

Private Sub Command3_Click()
    fraDrive.Visible = True
    fraProcess.Visible = False
End Sub

Private Sub Command4_Click()
    fraDrive.Visible = False
    fraDrive.Left = 15: fraDrive.Top = 0
    fraProcess.Visible = True
End Sub

Private Sub Command5_Click()
    cmdPad_Click
End Sub


Private Sub cmdShock_Click()
    Shock = True
End Sub

Private Sub Command6_Click()
Dim s As String
    txt4AbjadNum = AbjadClC(txt4Abjad)
End Sub

Private Sub Command7_Click()
    frmBase.Enabled = False
    frmQuran.Visible = True
End Sub

Private Sub cmdOpenPic_Click()
Dim s As String
CDlg.ShowOpen
s = CDlg.filename

    If chkImg2Pic Then
        picStore.Picture = LoadPicture(s)
    ElseIf chkImg2Pic2 Then
        picStore2.Picture = LoadPicture(s)
    End If
End Sub

Private Sub Command8_Click()
    Dim x As Integer, s As String, y As Integer
    txtLogs.Visible = False
    cmdGetLog2.Visible = False
    On Error Resume Next
    txtLCo = Val(txtLCo.Text)
    DoEvents
        z = 1
        For x = 1 To txtLCo
            z = (x)
            s = s & z & " , "
            For y = 1 To txtLCo
                If z >= 9737333 Then Exit For
                s = s & Primes(z) & " , "
                z = (Primes(z))
            Next y
            s = s & vbCrLf
        Next x
        
    txtLogs = s
    txtLogs.Visible = True
    cmdGetLog2.Visible = True
End Sub

Private Sub Command9_Click()
    On Error Resume Next
    Kill sPath & "\*.*"
    txtShotCount = "0"
End Sub

Private Sub Form_Activate()
    
  Static waveFormat As WaveFormatEx

    With waveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    waveInOpen DevHandle, DevicesBox.ListIndex, VarPtr(waveFormat), 0, 0, 0
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!") ' 
        Exit Sub
    End If
    Call waveInStart(DevHandle)
    Inited = True
    DoEv = True
    fraStart.Visible = True
    fraStart.ZOrder 0
    DoEvents
    frmQuran.lstBase.ListIndex = CInt(txtNumber.Text):  frmQuran.cmdSet_Click
    BaseSub

End Sub

Private Sub Form_DblClick()
    If fraProcess.Visible = True Then
        cmdMax.Visible = False
        cmdMini.Visible = True
        
        fraProcess.Visible = False
        fraLogs.Visible = False

      Else
        cmdMax.Visible = True
        cmdMini.Visible = False
    
        fraProcess.Visible = True
        
    End If
    BoardF1 = 25
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    KeyPrss
End Sub

Private Sub Form_Resize()
    picViewEE.Width = frmBase.Width
    picViewEE.Height = frmBase.Height
    fraControls.Left = frmBase.Width - fraControls.Width - 360
    fraColors.Left = fraControls.Left - fraColors.Width
    fraBlur.Left = fraColors.Left - fraBlur.Width
End Sub

Private Sub Label1_Click()
    lblLQSky_Click
End Sub

Private Sub Label10_Click()
    lblLQSky_Click
End Sub

Private Sub lblClick_Click()
    lblLQSky_Click
End Sub

Private Sub lblControls_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraControls.ZOrder 0
End Sub

Public Sub txtFocus_GotFocus()
'    DoEvents
End Sub

Private Sub txtFormula_Click()
    chkScript.Value = 0
End Sub

Private Sub Form_Load()

  
    picBuffEE.Width = Screen.Width
    picBuffEE.Height = Screen.Height
    picBuffEE2.Width = Screen.Width
    picBuffEE2.Height = Screen.Height
    frmBase.Width = Screen.Width
    frmBase.Height = Screen.Height
    ResX = Screen.Width \ Screen.TwipsPerPixelX  '- 1
    ResY = Screen.Height \ Screen.TwipsPerPixelY  '- 1
    fraFullScr.Top = Screen.Height - 180
    fraFullScr.Width = frmBase.Width
    txtspm(22) = ResX \ 2:    txtspm(23) = ResY \ 2
    picStore.BackColor = vbBlack
    picStore.AutoSize = True

    pic2Text(0).Width = frmBase.Width
    pic2Text(1).Width = frmBase.Width
    pic2Text(2).Width = frmBase.Width
    pic2Text(3).Width = frmBase.Width
    pic2Text(4).Width = frmBase.Width
    pic2Text(5).Width = frmBase.Width
    
    lblFullscr(0).Left = frmBase.Width - 3720
    lblFullscr(9).Left = frmBase.Width - 2160
    picStore.Width = frmBase.Width
    picStore.Height = frmBase.Height
    picStore.Left = 0
    picStore.Top = 0
    picStore2.Width = frmBase.Width
    picStore2.Height = frmBase.Height
    picStore2.Left = 0
    picStore2.Top = 0
    
    Set picView = picViewEE
    Set picBuff = picBuffEE
    Set picBuffSe = picBuffEE
    Set picBuffSe2 = picBuffEE
    Set picTmp = picBuffEE2
    
    picView.Refresh
    picView.Visible = True
    picView.ZOrder 1
    
    initSetData
    lblFullscr(0).Caption = "Liquid Skies " & App.Major & "." & App.Minor & "." & App.Revision
     
    SsPtr = 0
    txP = 1: tyP = 1
    Set Clk = New cCpuClk            'Create the CpuClk instance
    DoEvents
    Call QueryPerformanceFrequency(cCycles)
    
    stFirst = 1
    dRFlag = 1
    
    fraDrive.Top = 0
    fraDrive.Left = 15
    fraProcess.Top = 0
    fraProcess.Left = 15
    fraProcess.Height = 5535
    fraBlur.Height = 350
    fraControls.Height = 350
    picBProcs.Height = 7400
    picBProcs.Top = 720
    picBLogs.Top = 360
    
    
    FlgBlur = 1
    Set_Process
    
    xColStp = 1
    yCol = 5
    xCol = 5
    
    LoadREG
    txtPath = sPath
    ReDim Smp(0 To FileLen(App.Path & "\" & "PSamp.bin") \ Len(Smp(0)) + 1)
    txtRecCo(0) = FileLen(App.Path & "\" & "PSamp.bin") \ Len(Smp(0)) - 1 ' Combo2.ListCount
    LoadFPara
    Combo2.ListIndex = 0
    cmdLoadP_Click
    DoEvents
    LQT2 = frmBase.txtspm(11)
    LQT = frmBase.txtspm(11)
    
Dim s As String, intf As Integer, x As Integer, Pth As String
    Pth = App.Path & "\PRG.txt"
    intf = FreeFile
    Open Pth For Input As #intf
    While Not EOF(intf)
        Input #intf, s
        lstPRG.AddItem s
        DoEvents
    Wend
    Close #intf
End Sub

Private Sub initSetData()
Dim caps As WAVEINCAPS, Which As Long
Dim x As Integer
     
    MVolu = 1
    Fst = True
    DoEv = 11
   
    ABass = 1
    AMidl = 1
    ATreb = 1
    AFreq = 1
    ABass2 = 1
    AMidl2 = 1
    ATreb2 = 1
    AFreq2 = 1
    Randomize (Timer)
    RV = CDbl(Rnd(2) * 255)
    GV = CDbl(Rnd(3) * 255)
    BV = CDbl(Rnd(5) * 255)
    Randomize (Timer)
    RN = CDbl(Rnd(2) * 255)
    GN = CDbl(Rnd(3) * 255)
    BN = CDbl(Rnd(5) * 255)
    
    PiTAdd1 = 8.72664625997165E-03
    PiTAdd2 = 1.74532925199433E-02
    
    For x = 0 To 2
        MaxC(x) = Val(txtMaxC(x).Text)
        MinC(x) = Val(txtMinC(x).Text)
    Next x
    
    Ox = 0
    Oy = 128
    Ox2 = 256
    Oy2 = 128
    BlurNum = 0
    
    tx = 1
    ty = 1

    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(caps), Len(caps)) ' 
        If caps.Formats And WAVE_FORMAT_1S08 Then 'Now is 1S08 -- Check for devices that can do stereo 8-bit 11kHz
            Call DevicesBox.AddItem(StrConv(caps.ProductName, vbUnicode), Which) ' 
        End If
    Next ' Repeat For-Variable: WHICH
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End ' There are better ways to terminate
    End If
    DevicesBox.ListIndex = 0

    ColSt(0) = &H0&     'Black 0+0+0
    ColSt(1) = &HFF&    'Red
    ColSt(2) = &HFF00&  'Green
    ColSt(3) = &HFF0000 'Blue
    ColSt(4) = &HFFFF00 'Cyan B+G
    ColSt(5) = &HFF00FF 'Maginta B+R
    ColSt(6) = &HFFFF&  'Yelow G+R
    ColSt(7) = &H7F7F7F 'Gray  127+127+127
    ColSt(8) = &HFFFFFF 'White 255+255+255
    ColSt(9) = ColSt(8) Xor ColSt(2)
    ColSt(10) = ColSt(5) Xor ColSt(4)

End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If DevHandle <> 0 Then
    Call waveInReset(DevHandle) ' 
    Call waveInClose(DevHandle) ' 
    DoEvents
    DevHandle = 0
    
    End If
    SaveREG

    End
    
End Sub

Private Sub fraBlur_Click()
    fraBlur.ZOrder (0)
End Sub

Private Sub fraBlur_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = True
End Sub

Private Sub fraBlur_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraBlur.ZOrder (0)
End Sub

Private Sub fraBlur_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = False
End Sub

Private Sub fraColors_Click()
    fraColors.ZOrder (0)
End Sub

Private Sub fraColors_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = True
End Sub

Private Sub fraColors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = False
End Sub

Private Sub fraControls_Click()
    fraControls.ZOrder (0)
End Sub


Private Sub fraControls_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = True
End Sub

Private Sub fraControls_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = False
End Sub

Private Sub fraProcess_Click()
    fraProcess.ZOrder (0)
End Sub


Private Sub fraProcess_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = True
End Sub

Private Sub fraProcess_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEv = False
End Sub

Private Sub lblBlur_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraBlur.ZOrder 0
End Sub

Private Sub lblColorSet_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraColors.ZOrder 0
End Sub

Private Sub lblLogs_Click()
    If fraLogs.Height > 400 Then
        fraLogs.Height = 350
      Else
        fraLogs.Height = 4905
        fraLogs.ZOrder 0
    End If
End Sub

Private Sub lblColorSet_Click()
    If fraColors.Height > 400 Then
        fraColors.Height = 350
      Else
        fraColors.Height = 4095
        fraColors.ZOrder (0)
    End If
End Sub

Private Sub lblControls_Click()
    If fraControls.Height > 400 Then
        fraControls.Height = 350
      Else
        fraControls.Height = 9120
        fraControls.ZOrder (0)
    End If
End Sub

Private Sub lblBlur_Click()
    If fraBlur.Height > 400 Then
        fraBlur.Height = 350
      Else
        fraBlur.Height = 4620
        fraBlur.ZOrder (0)
    End If
End Sub

Private Sub lblLogs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraLogs.ZOrder 0
End Sub

Private Sub lblLQSky_Click()
    If fraProcess.Height > 525 Then
        fraProcess.Height = 525
        fraMsgBx.Visible = False
      Else
        fraProcess.Height = 7455
        fraProcess.ZOrder (0)
    End If
End Sub

Private Sub lblVol_Click()
    lblControls_Click
End Sub

Private Sub cmdSnCGr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If cS(Index) > 0 Then
        cS(Index) = -1
    ElseIf cS(Index) < 0 Then
        cS(Index) = 0
    Else
        cS(Index) = 1
    End If
End Sub

Private Sub List1_DblClick()
    
    If chkBGQ(0).Value = 1 Then
        txtspm(28) = List1.List(List1.ListIndex)
    ElseIf chkBGQ(1).Value = 1 Then
        txtspm(21) = List1.List(List1.ListIndex) - 1
    End If
    
    If (chkBGQ(0).Value = 1 And chkBGQ(1).Value = 1) Or (chkBGQ(0).Value = 0 And chkBGQ(1).Value = 0) Then
        If List1.ListIndex < 1 Then List1.ListIndex = 1
       
        txtspm(28) = List1.List(List1.ListIndex - 1)
        List2.ListIndex = (List1.ListIndex - 1)
        txtspm(21) = List1.List(List1.ListIndex) - 1
        
    End If
   
Dim a1 As Integer, a2 As Integer, z As Integer
    a1 = List1.ListIndex
    a2 = Val(List1.Tag)
    
    
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraLogs.ZOrder 0
End Sub

Private Sub List2_DblClick()
    List1.ListIndex = List2.ListIndex
    List1_DblClick
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraLogs.ZOrder 0
End Sub


Private Sub List3_DblClick()
    List1.ListIndex = List3.ListIndex
    List1_DblClick
End Sub

Private Sub lst11Pows_DblClick()
    If lst11Pows.ListIndex < 5 Then
        txtGR = lst11Pows.List(lst11Pows.ListIndex)
'        txtspm(20) = txtGR
    End If
End Sub

Private Sub lstBStp_DblClick()
    List1.ListIndex = lstBStp.ListIndex
    List1_DblClick
End Sub

Private Sub lstL_Click(Index As Integer)
    txtspm(28) = lstL(3).List(lstL(3).ListIndex) * (chkBGQ(0) Xor 1)
    If lstL(3).ListIndex + 1 < lstL(3).ListCount Then txtspm(21) = lstL(3).List(lstL(3).ListIndex + 1) * (chkBGQ(1) Xor 1)
End Sub

Private Sub lstPRG_Click()
    txtMain = lstPRG.Text
    cmdPrimeIndex_Click
End Sub

Private Sub lstPRG_Validate(Cancel As Boolean)
    TextLabel(7).Text = lstPRG.ListCount
End Sub

Private Sub picBBlur_Click()
    fraBlur.ZOrder 0
End Sub

Private Sub picBBlur_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraBlur.ZOrder 0
End Sub

Private Sub picBCol_Click()
    fraColors.ZOrder 0
End Sub

Private Sub picBCol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraColors.ZOrder 0
End Sub

Private Sub picBCtrl_Click()
    fraControls.ZOrder 0
End Sub

Private Sub picBCtrl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraControls.ZOrder 0
End Sub

Private Sub picBLogs_Click()
    fraLogs.ZOrder 0
End Sub

Private Sub picBLogs_DblClick()
    Form_DblClick
End Sub

Private Sub picBLogs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    fraLogs.ZOrder 0
End Sub

Private Sub picBProcs_Click()
    fraProcess.ZOrder 0
End Sub

Private Sub picBProcs_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    fraProcess.ZOrder 0
End Sub

Private Sub picViewEE_DblClick()
    Call Form_DblClick
End Sub

Private Sub Timer_AHeight_Timer()
On Error Resume Next
    Timer_AHeight.Tag = Val(Timer_AHeight.Tag) + 1
    If Val(Timer_AHeight.Tag) > 1 Then
        Timer_AHeight.Enabled = False
        chkAHeight.Value = 0
        Timer_AHeight.Tag = "0"
    End If
    If (Abs(maxY - minY) < Val(txtspm(0).Text * 1.1)) And (Abs(maxY - minY) > Val(txtspm(0).Text * 0.9)) Then
        Timer_AHeight.Enabled = False
        chkAHeight.Value = 0
        Timer_AHeight.Tag = "0"
    End If
End Sub



Private Sub Timer_AutoNext_Timer()

If chkAutoNext.Value = 0 Then Exit Sub

   If frmQuran.lstBase.ListCount - 1 > frmQuran.lstBase.ListIndex Then
        frmQuran.lstBase.ListIndex = frmQuran.lstBase.ListIndex + 1
   Else
        frmQuran.lstBase.ListIndex = 0
   End If
   frmQuran.lstBase.Refresh
'   DoEvents
   frmQuran.cmdSet_Click
   Timer_AutoNext.Interval = Len(frmQuran.txtBase2.Text) * 90 + 1000
End Sub

Private Sub Timer_Process_Timer()
    Dim hIcon As Long, pdhStatus As Long
    Dim dbl As Double

        If AvgUsageCount = 0 Then
            AvgCpu0 = 0
            AvgCpu1 = 0
        End If
        If AvgUsageCount > 100 Then
            AvgCpu0 = AvgCpu0 / AvgUsageCount
            AvgCpu1 = AvgCpu1 / AvgUsageCount
            AvgUsageCount = 1
        End If
    
        PdhCollectQueryData (hQuery)
    
        dbl = PdhVbGetDoubleCounterValue(Counters(0).hCounter, pdhStatus)
        If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
            AvgCpu0 = AvgCpu0 + dbl
        End If
        
        If NumOfCores = 2 Then
            dbl = PdhVbGetDoubleCounterValue(Counters(1).hCounter, pdhStatus)
            If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
                AvgCpu1 = AvgCpu1 + dbl
            End If
        End If
    
     
        AvgUsageCount = AvgUsageCount + 1
    
    txtProcess0 = CStr(Int(AvgCpu0 / AvgUsageCount)) + " %"
    txtProcess1 = CStr(Int(AvgCpu1 / AvgUsageCount)) + " %"
    txtProcessSum.Text = CStr(Int(((AvgCpu0 / AvgUsageCount) + (AvgCpu1 / AvgUsageCount)) / 2)) + " %"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Sub


Private Sub Timer_AutoSave_Timer()

On Error Resume Next

 TiS = TiS + 1
    If TiS >= txtspm(8) * 10 And chkAutoShot And chkShotAll.Value <> 1 Then
        TiS = 0
        cmdSF_Click
        txtspm(8).BackColor = txtspm(8).BackColor Xor vbBlue
        txtspm(8).ForeColor = txtspm(8).ForeColor Xor vbRed
        TiS = 0
    End If

End Sub

Private Sub Timer_Qu_Timer()

End Sub

Private Sub Timer_Seconds_Timer()
    If txtAutoAGG = "0" Then txtAutoAGG = "1"
    If chkA_Aggr And St_Time Mod Val(txtAutoAGG) = 0 Then txtspm(20) = Val(txtspm(20)) + Val(txtGR)
    
    St_Time = St_Time + 1
    If chkAutoSampl Then
        chkAutoSampl.Tag = Val(chkAutoSampl.Tag) + 1
        If Val(chkAutoSampl.Tag) >= 30 Then cmdNextS_Click: chkAutoSampl.Tag = "0"
    End If
    If St_Time = 10 Then chkText2Pic.Value = 0
    
End Sub


Private Sub txt2Pic_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
End Sub


Private Sub txtMinC_Change(Index As Integer)
    MinC(Index) = Val(txtMinC(Index).Text)
End Sub

Private Sub txtMsgBx_Change()
    If Val(txtMsgBx) > Val(txtRecCo(0)) Then txtMsgBx = txtRecCo(0)
End Sub

Private Sub txtPT_Size_DblClick()
End Sub

Private Sub txtNumber_Click()
    txtNumber_KeyDown 13, 1
End Sub

Private Sub txtNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then frmQuran.lstBase.ListIndex = CInt(txtNumber.Text):  frmQuran.cmdSet_Click ': DoEvents
End Sub

Private Sub txtPrimeIndex_Change(Index As Integer)
    If Index = 1 Then txtPrimeIndex(0) = txtPrimeIndex(1)
End Sub

Private Sub txtR_Click()
    txtR_DblClick
End Sub

Private Sub txtR_DblClick()
    If txtR = "1" Then
        txtR = "2": txtR.BackColor = &HFFC0FF
    Else
        txtR = "1": txtR.BackColor = &HFFC0C0
    End If
End Sub

Private Sub txtShotCount_Change()
    lblFullscr(8) = txtShotCount
End Sub

Private Sub txtShotCount_DblClick()
    txtShotCount = 0
End Sub

Private Sub txtspm_Change(Index As Integer)
    If Not IsNumeric(txtspm(Index)) Then txtspm(Index) = Val(txtspm(Index)) + 1
    If Index = 0 Then chkAHeight.Value = 1
    If Index = 20 And Val(txtspm(20)) < 1 Then txtspm(20) = txtspm(20) + 1
    If Index = 28 And Val(txtspm(28)) < 1 Then txtspm(28) = 1
    Sleep Int(FpS * 10)
    txtspm(Index).Refresh
'    DoEvents
End Sub

Private Sub txtspm_DblClick(Index As Integer)
    If Index = 22 Then
        txtspm(22) = ResX \ 2
    End If
    If Index = 23 Then
        txtspm(23) = ResY \ 2
    End If
    
End Sub

Private Sub txtTextSize_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        CDlg.ShowColor
        txtTextSize(Index).BackColor = CDlg.Color
    End If
End Sub
