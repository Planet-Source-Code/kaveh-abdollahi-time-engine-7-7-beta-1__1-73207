VERSION 5.00
Begin VB.Form frmCam 
   AutoRedraw      =   -1  'True
   Caption         =   "TJ's Webcam Input Thresholding Program"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   426
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDrawPix 
      Caption         =   "&Show thresholded pixels in green."
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   2835
   End
   Begin VB.CheckBox chkMouse 
      Caption         =   "Control &mouse cursor."
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6180
      Width           =   2055
   End
   Begin VB.HScrollBar hscStep 
      Height          =   255
      LargeChange     =   4
      Left            =   1380
      Max             =   32
      Min             =   1
      TabIndex        =   11
      Top             =   5820
      Value           =   3
      Width           =   3375
   End
   Begin VB.HScrollBar hscThres 
      Height          =   255
      LargeChange     =   16
      Left            =   1380
      Max             =   255
      TabIndex        =   9
      Top             =   5460
      Value           =   16
      Width           =   3375
   End
   Begin VB.HScrollBar hscBlue 
      Height          =   255
      LargeChange     =   16
      Left            =   1380
      Max             =   255
      TabIndex        =   8
      Top             =   5100
      Width           =   3375
   End
   Begin VB.HScrollBar hscGreen 
      Height          =   255
      LargeChange     =   16
      Left            =   1380
      Max             =   255
      TabIndex        =   6
      Top             =   4740
      Width           =   3375
   End
   Begin VB.HScrollBar hscRed 
      Height          =   255
      LargeChange     =   16
      Left            =   1380
      Max             =   255
      TabIndex        =   4
      Top             =   4380
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6000
      Top             =   6240
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   960
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   240
      Width           =   4800
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   720
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   1
      Top             =   60
      Width           =   4800
   End
   Begin VB.Label Label6 
      Caption         =   "&Performance:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   5820
      Width           =   1155
   End
   Begin VB.Label Label5 
      Caption         =   "&Threshold:"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5460
      Width           =   1155
   End
   Begin VB.Label Label4 
      Caption         =   "&Blue:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5100
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "&Green:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4740
      Width           =   1155
   End
   Begin VB.Label Label2 
      Caption         =   "&Red:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4380
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Threshold Settings:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4080
      Width           =   1575
   End
End
Attribute VB_Name = "frmCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program: Webcam Input Thresholding Program
'Author: Chear Tze Jian
'Date: Sept 7 2006
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PrintWindow Lib "user32" (ByVal hWnd As Long, ByVal hdcBlt As Long, ByVal nFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private mCapHwnd As Long

Private Const CONNECT As Long = 1034
Private Const DISCONNECT As Long = 1035
Private Const GET_FRAME As Long = 1084
Private Const WS_CHILD As Long = &H40000000
Private Const WS_VISIBLE As Long = &H10000000

Private r_t As Long, g_t As Long, b_t As Long
Private threshold As Long
Private lStep As Long
Private mX As Long, mY As Long

Private pntMouse As POINTAPI
Private deltaMouse As POINTAPI
Private pntMid As POINTAPI
Private pichDC As Long

Private Sub chkMouse_Click()
If chkMouse.Value = 1 Then
    pntMid.x = mX
    pntMid.y = mY
End If
End Sub

Private Sub Form_Load()
threshold = hscThres.Value
lStep = hscStep.Value
pichDC = Picture1.hdc
mCapHwnd = capCreateCaptureWindow("BetterCam", WS_CHILD Or WS_VISIBLE, 0, 0, 320, 240, Picture2.hWnd, 0)
DoEvents
SendMessage mCapHwnd, CONNECT, 0, 0
Timer1.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Timer1.Enabled = False
SendMessage mCapHwnd, DISCONNECT, 0, 0
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim c As Long
c = Picture1.Point(x, y)
r_t = c Mod 256
g_t = (c / 256) Mod 256
b_t = (c / 256 / 256) Mod 256
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
hscRed.Value = r_t
hscGreen.Value = g_t
hscBlue.Value = b_t
End Sub

Private Sub Timer1_Timer()
SendMessage mCapHwnd, GET_FRAME, 0, 0
Call PrintWindow(mCapHwnd, Picture1.hdc, 0)
'Picture1.Picture = Picture1.Image
Dim x As Long, y As Long, c As Long
Dim r As Long, g As Long, b As Long
Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
Dim pntMse As POINTAPI

Dim pnts() As POINTAPI, i As Long, j As Long, aX As Long, aY As Long
i = 0
aX = 0
aY = 0
ReDim pnts(0)

x1 = Picture1.Width
y1 = Picture1.Height
x2 = -1
y2 = -1
For y = 0 To Picture1.Height - 1 Step lStep
    For x = 0 To Picture1.Width - 1 Step lStep
        'c = Picture1.Point(x, y)
        c = GetPixel(pichDC, x, y)
        r = c Mod 256
        g = (c / 256) Mod 256
        b = (c / 256 / 256) Mod 256
        If IsWithinThreshold(r, g, b) Then
            If x < x1 Then x1 = x
            If x > x2 Then x2 = x
            If y < y1 Then y1 = y
            If y > y2 Then y2 = y
            pnts(i).x = x
            pnts(i).y = y
            i = i + 1
            ReDim Preserve pnts(i)
            'Picture1.PSet (x, y), vbGreen
            'If chkDrawPix.Value > 0 Then
                SetPixel pichDC, x, y, vbGreen
            'End If
        End If
    Next
Next
For j = 0 To i - 1
    aX = aX + pnts(j).x
    aY = aY + pnts(j).y
Next
If i > 0 Then
    mX = aX / i
    mY = aY / i
End If
Picture1.Line (mX - 10, mY)-(mX + 10, mY), vbGreen
Picture1.Line (mX, mY - 10)-(mX, mY + 10), vbGreen
Picture1.Line (x1, y1)-(x2, y2), vbRed, B

If chkMouse.Value = 1 Then
    GetCursorPos pntMse
    pntMse.x = pntMse.x - 2 * (mX - pntMid.x) 'Invert horizontal
    pntMse.y = pntMse.y + 2 * (mY - pntMid.y)
    SetCursorPos pntMse.x, pntMse.y
End If

pntMid.x = mX
pntMid.y = mY

End Sub

Private Function IsWithinThreshold(r As Long, g As Long, b As Long) As Boolean
IsWithinThreshold = ((Abs(r - r_t) <= threshold) And (Abs(g - g_t) <= threshold) And (Abs(b - b_t) <= threshold))
End Function

Private Function IsAboveRGB(r As Long, g As Long, b As Long) As Boolean
IsAboveRGB = ((r >= r_t) And (g >= g_t) And (b >= b_t))
End Function

Private Function IsBelowRGB(r As Long, g As Long, b As Long) As Boolean
IsBelowRGB = ((r <= r_t) And (g <= g_t) And (b <= b_t))
End Function

'Supress G and B components
Private Function IsAboveRBelowGB(r As Long, g As Long, b As Long) As Boolean
IsAboveRBelowGB = ((r >= r_t) And (g <= g_t) And (b <= b_t))
End Function

Private Sub hscBlue_Change()
b_t = hscBlue.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscBlue_Scroll()
b_t = hscBlue.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscGreen_Change()
g_t = hscGreen.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscGreen_Scroll()
g_t = hscGreen.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscRed_Change()
r_t = hscRed.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscRed_Scroll()
r_t = hscRed.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscStep_Change()
lStep = hscStep.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscStep_Scroll()
lStep = hscStep.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscThres_Change()
threshold = hscThres.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub

Private Sub hscThres_Scroll()
threshold = hscThres.Value
Me.Caption = "RGB(" & r_t & "," & g_t & "," & b_t & ") ; Threshold: " & threshold & " ;Step: " & lStep
End Sub
