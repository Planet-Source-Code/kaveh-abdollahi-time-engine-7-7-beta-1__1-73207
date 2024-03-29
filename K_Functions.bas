Attribute VB_Name = "K_Functions"
'******************************************************************************************
'
'   Copyright(C) 2010 By Kaveh Abdollahi.   kavehplus@gmail.com
'   Time Engine
'   June 2010
'
'******************************************************************************************


Option Explicit

Private Const PI = 3.14159265358979, Rad = 0.017453292519943

Public P3P As Mstp, P4P As Mstp, p()  As Mstp, Polindx As Integer
Public lTi As Long
Public Type Mstp
   Tmz  As Single
   TM2  As Single
   Tm4  As Single
   TM1  As Single
   Tm3  As Single
   mX    As Single
   mY    As Single
   x    As Single
   y    As Single
   z    As Single
   Xx   As Single
   Yy   As Single
   Zz   As Single
   Zm   As Single
   ColV   As Single
   Col    As Single
   Pow  As Single
   str  As String * 100
   Pt(0 To 511) As POINTAPI
   PtL(0 To 511) As POINTAPI
End Type
Public Sub PicGrey()
    'Affects only the Left panel
    Dim x As Integer, y As Integer, Col As Integer
    Dim N1 As Integer, N2 As Integer, N3 As Integer 'buffers to prevent overflow
    
    GetBitmapBits picView, Pic.bmWidthBytes * Pic.bmHeight * 3, Bytes(0, 0)
    
    'Finds average color value and makes it grayscale
    For x = 0 To 255
        For y = 0 To 255
            N1 = Bytes(x, y).Red
            N2 = Bytes(x, y).Green
            N3 = Bytes(x, y).Blue
            Col = (N1 + N2 + N3) / 3
            Bytes(x, y).Red = Col
            Bytes(x, y).Green = Col
            Bytes(x, y).Blue = Col
        Next y
    Next x
    SetBitmapBits picView, Pic.bmWidthBytes * Pic.bmHeight * 3, Bytes(0, 0)
'    picView.Cls
End Sub

Public Function Sine(Degrees_Arg)
   Sine = Sin(Degrees_Arg * Atn(1) / 45)
End Function

Public Function Cosine(Degrees_Arg)
   Cosine = Cos(Degrees_Arg * Atn(1) / 45)
End Function


Public Sub kaAltrCls()
Dim xt As Integer, fo As POINTAPI
    picBuff.ForeColor = 0
    If reAl Or (maxY > 767 Or minY < 1) Then minY = 0: maxY = 768: reAl = False
    
     minY = 0: maxY = 768
    If minY Mod 2 = 1 Then minY = minY - 3
    For xt = minY To maxY * 1.1 Step 2
        MoveToEx picBuff.hdc, 0, xt, fo
        LineTo picBuff.hdc, ResX, xt
    Next xt
End Sub

Public Sub KaCls()
     BitBlt picBuff.hdc, 0, 0, ResX, ResY, picBuff.hdc, 0, 0, vbBlackness
     BitBlt picTmp.hdc, 0, 0, ResX, ResY, picTmp.hdc, 0, 0, vbBlackness
     StretchBlt picView.hdc, 0, 0, ResX, ResY _
        , picBuff.hdc, 0, 0, ResX, ResY, vbSrcCopy
End Sub
Private Sub ColSet()
  Dim x As Integer, tes As Double, xCu As Long
    
    
    xCu = frmBase.txtspm(1)

    If Colv_R > (MaxC(0) - xCu) Or Colv_R < MinC(0) Then cS(0) = -cS(0)
    If Colv_R > (MaxC(0) - xCu) Then Colv_R = MaxC(0) - xCu
    If Colv_R < MinC(0) Then Colv_R = MinC(0)

    If Colv_G > (MaxC(1) - xCu) Or Colv_G < MinC(1) Then cS(1) = -cS(1)
    If Colv_G > (MaxC(1) - xCu) Then Colv_G = MaxC(1) - xCu
    If Colv_G < MinC(1) Then Colv_G = MinC(1)

    If Colv_B > (MaxC(2) - xCu) Or Colv_B < MinC(2) Then cS(2) = -cS(2)
    If Colv_B > (MaxC(2) - xCu) Then Colv_B = MaxC(2) - xCu
    If Colv_B < MinC(2) Then Colv_B = MinC(2)

  ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    If frmBase.chkCol(0) Then
        Colv_R = Colv_R + cS(0) * 5 / 19 * Cos(Rad * (Primes(LQT) - Primes(LQT - 1)))
        Colv_G = Colv_G + cS(1) * 3 / 13 * Cos(Rad * (Primes(LQT) - Primes(LQT - 1)))
        Colv_B = Colv_B + cS(2) * 2 / 11 * Cos(Rad * (Primes(LQT) - Primes(LQT - 1)))
    End If
    If frmBase.chkCol(1) Then
        Colv_R = Colv_R + cS(0) * 5 / 19 * Log(Rad * (Primes(LQT) - Primes(LQT - 1)))
        Colv_G = Colv_G + cS(1) * 3 / 13 * Log(Rad * (Primes(LQT) - Primes(LQT - 1)))
        Colv_B = Colv_B + cS(2) * 2 / 11 * Log(Rad * (Primes(LQT) - Primes(LQT - 1)))
    End If
    If frmBase.chkCol(2) Then
        Colv_R = Colv_R + cS(0) * (PI / 3) * Cos(Rad * LQ_Pr_Mod57)  '''''''' ok
        Colv_G = Colv_G + cS(1) * (PI / 5) * Cos(Rad * LQ_Pr_Mod57)
        Colv_B = Colv_B + cS(2) * (PI / 7) * Cos(Rad * LQ_Pr_Mod57)
    End If
    If frmBase.chkCol(3) Then
        Colv_R = Colv_R + cS(0) * (5 / 17) * PI * Atn(Primes(LQT))
        Colv_G = Colv_G + cS(1) * (3 / 17) * PI * Atn(Primes(LQT))
        Colv_B = Colv_B + cS(2) * (2 / 17) * PI * Atn(Primes(LQT))
    End If
    If frmBase.chkCol(4) Then
        Colv_R = Colv_R + cS(0) * Cos((Bass * Bass) / 2)
        Colv_G = Colv_G + cS(1) * Cos((Midl * Midl) / 2)
        Colv_B = Colv_B + cS(2) * Cos((Treb * Treb) / 2)
    End If
    
        
    If frmBase.chkFallCol Then
        For x = 1 To 20
           If Not frmBase.chkInverse Then ColTn(21 - x) = RGB(Colv_R + x * 8, Colv_G + x * 8, Colv_B + x * 8)
           If frmBase.chkInverse Then ColTn(21 - x) = RGB(256 - Colv_R + x * 8, 256 - Colv_G + x * 8, 256 - Colv_B + x * 8)
        Next x
     Else
        For x = 1 To 20
           If Not frmBase.chkInverse Then ColTn(x) = RGB(Colv_R + x * 8, Colv_G + x * 8, Colv_B + x * 8)
           If frmBase.chkInverse Then ColTn(x) = RGB(256 - Colv_R + x * 8, 256 - Colv_G + x * 8, 256 - Colv_B + x * 8)
        Next x
    End If

    ColPR = ColPR * ColPRsgn + 0.5 * Sin(LQT) + 0.17
    ColPG = ColPG * ColPGsgn + 0.5 * Sin(LQT) + 0.13
    ColPB = ColPB * ColPBsgn + 0.5 * Sin(LQT) + 0.11
    If ColPR >= 255 Then ColPRsgn = -1: ColPR = 255
    If ColPG >= 255 Then ColPGsgn = -1: ColPG = 255
    If ColPB >= 255 Then ColPBsgn = -1: ColPB = 255
    If ColPR <= 1 Then ColPRsgn = 1: ColPR = 1
    If ColPG <= 1 Then ColPGsgn = 1: ColPG = 1
    If ColPB <= 1 Then ColPBsgn = 1: ColPB = 1


    ColP = RGB(Abs(ColPR), Abs(ColPG), Abs(ColPB))
    For x = 1 To 127
        ColTp(x) = RGB((128 - x + ColPR), (128 - x + ColPG), (128 - x + ColPB)) And ColTn(x \ 2 + 1)
        ColTp(x) = ColTp(x) And ColSt(x Mod 6 + 1)
    Next x


    ColB = vbBlack
End Sub
Public Sub DelCtrls()
Dim Control As Control

    For Each Control In frmBase
        If TypeOf Control Is Label Then Control.Caption = ""
    Next Control

End Sub
'
'Public Sub DrawP2()
'Dim LAvg As Long, RAvg As Long
'    If Fst Then InitK: Cof_X = 1: iH = 384    'initial at first time
'
'    ''''''''''''' set parameters variables '''''''''''''
'
'    CycleST
'
'        Cu = frmBase.txtspm(1) + 1
'        Xtmp = frmBase.txtspm(6)
'        z11 = vsY * 4
'        z22 = vsY * 4 '* (Bass / 20 - Treb / 10)
'        z33 = frmBase.txtspm(5) * 4
'        z44 = vsX ' / 2 '* 8
'        zvF = 0: zvT = 1
'        zV = z33   ''''' If frmbase.chkInc is not set Then zV = 1 and not use in loop _
'                          '     Else use  >> Ss?O(x, d) * zV
'        z33 = z33 / 2
'        If frmBase.chkInc.Value <> 0 Then zvF = 1: zvT = 0
'
'    ''''''''''''''''''
'    '''''''''''''''''' set y points of master polyline with data . Pt( , 1) is Master polyline ''''''''''''''''''
'    '''''''''''''''''' x points only set in load in InitK Sub in first time '
''        x=frmBase.txtspm(0)*
'    x2 = (384 - frmBase.txtspm(0) / 2) + frmBase.txtspm(0) / 2 '/ 2
'
'     For x = 0 To 255
'        d = SsPtr - x * Xtmp
'        zV = zvF * (((x + 1) / 32) * z33) + (zvT * zV)
'
'        Pt(255 - x, 1).y = (SsLO(x, d) * zV) - (67 * zV) + x2
'        Pt(256 + x, 1).y = (SsRO(x, d) * zV) - (67 * zV) + x2
'     Next x
'
'    CycleED
'    Process(12, 1) = Round(tFa, 2)
'
'    '''''''''''''''''' set y points in other polylines with Master polyline ''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''
'
'    CycleST
'  ''''''''''''''''''''''''''
'    bf4 = z22
'    For x2 = 2 To Cu  ' Cu is Count of Scopes
'       z22 = (z22 * 0.95)
'        For x = 0 To 255
'            Pt(255 - x, x2).y = Pt(255 - x, 1).y + z22
'            Pt(256 + x, x2).y = Pt(256 + x, 1).y + z22
'        Next x
'    Next x2
'    ''''''''''''''''''''''''''
'    CycleED
'    Process(13, 1) = Round(tFa, 2)
'
'
'    '''''''''''''''''' find  minY and maxY points of Master polyline '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''
'    CycleST
'
'    minLY = minY: maxLY = maxY
'    minY = 768: maxY = 0
'
'    For x = 0 To 511
'        If maxY < Pt(x, 2).y Then maxY = Pt(x, 2).y
'        If minY > Pt(x, Cu).y Then minY = Pt(x, Cu).y
'    Next x
'    If maxY < 1 Then maxY = 1
'    If minY > 768 Then minY = 768
'
'    ''''''''''''''''''''''''''''''''''''''''''
'    '''''''''''''''''' Set BaseSub Height for scopes '''''''''''''''''''''''''''''''''''''''
'    ''''''''''''''''''''''''''''''''''
'
'    If frmBase.chkABalance And ((maxY - minY) < (frmBase.txtspm(0) / 2) Or (maxY - minY) > (frmBase.txtspm(0) * 2)) Then frmBase.chkAHeight.Value = 1
'
'    If frmBase.chkAHeight Then
'        If (maxY - minY) > (frmBase.txtspm(0) * 2) Then
'            frmBase.cmdSmaler_Click (5): frmBase.txtspm(5).Refresh
'        ElseIf (maxY - minY) < (frmBase.txtspm(0) / 2) Then
'            frmBase.cmdLarger_Click (5): frmBase.txtspm(5).Refresh
'        End If
'    End If
'
'    minY = minY - (maxY - minY) / 8 - 64
'    maxY = maxY + (maxY - minY) / 8 + 64
'
'
'    Call CycleED
'    Process(15, 1) = Round(tFa, 2)
'
'
'    '''''''''''''''' Clear last polyline if chkCls1 is checked ''''''''''''''''
'    CycleST
'
''        If frmBase.chkCls1 Then BitBlt picBuff.hdc, 0, 0, 1024, 768, picBuff.hdc, 0, 0, vbBlackness
'
'        If frmBase.ChkDraw(4) Then   '''''   last polyline clear
'            picBuff.ForeColor = vbBlack
'            picBuff.FillStyle = vbSolid
'            picBuff.FillColor = vbBlack
'
'           If frmBase.ChkDraw(2) Then
'                For x = 2 To Cu
'                 Polygon picBuff.hdc, PtL(1, x), 255
'                 Polygon picBuff.hdc, PtL(256, x), 255
'                Next x
'           Else
'                For x = 2 To Cu
'                 Polyline picBuff.hdc, PtL(1, x), 255
'                 Polyline picBuff.hdc, PtL(256, x), 255
'                Next x
'           End If
'        End If
'
'    '''''''''''''''' draw Polylines ''''''''''''''''
'        For x = 2 To Cu
'             picBuff.ForeColor = ColTn(x)
'             picBuff.FillStyle = vbSolid
'             picBuff.FillColor = ColTn(x)
'
'             If frmBase.ChkDraw(2) Then
'                Polygon picBuff.hdc, Pt(1, x), 255
'                Polygon picBuff.hdc, Pt(256, x), 255
'             Else
'                Polyline picBuff.hdc, Pt(1, x), 255
'                Polyline picBuff.hdc, Pt(256, x), 255
'             End If
'        Next x
'
'    '''''''''''''''''' store data for use in next polylines in next frame ''''''''''''''''''
'
'        For x = 2 To Cu
'            CopyMemory PtL(0, x).x, Pt(0, x).x, 4096
'        Next x
'
'    Call CycleED
'    Process(14, 1) = Round(tFa, 2)
'
'End Sub

Public Function KRand(Optional r As Double = 1)
    KRand = (Right(123 * LQ_Pr_Mod57 / 2, 4)) / 9999.9 * r
End Function

Public Sub InitK()
Dim A As Long, b As Long, c As Double, InFile$, x As Integer, TM1 As Byte
ReDim p(1 To 100)
    
    Vsgn = -1: Hsgn = 1
    Set picTmp = picBuff
    ColPRsgn = 1
    ColPGsgn = 1
    ColPBsgn = 1
    p(1).ColV = 1
    Randomize (Timer)
    PrimeBase
    LQT = frmBase.txtspm(11)     '     set LQT start value
    St_Time = 0
    
    ResX = Screen.Width / Screen.TwipsPerPixelX
    ResY = Screen.Height / Screen.TwipsPerPixelY
    Colv_R = Primes(LQT) Mod 196
    Colv_G = Primes(LQT) Mod 128
    Colv_B = Primes(LQT) Mod 256
     
    For b = 1 To 1000
        PN(b) = Primes(b)
    Next b
    PiN = 2
    
    Do While cS(0) = 0 Or cS(1) = 0 Or cS(2) = 0
        cS(0) = Rnd * 1: cS(1) = Rnd * 1: cS(2) = Rnd * 1
        A = A + 1
        If A > 200 Then cS(0) = 1: cS(1) = 1: cS(2) = 1: Exit Do
    Loop
    
    
    ''' set one polyline x 0 to 511  ' Pt(x, 1).x '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     For x = 0 To 511
        Pt(x, 1).x = (x) * 2 + 1 '- 1
        p(1).Pt(x).x = x * 2 + 1
'        Pt(511 - x, 1).x = 1023 - (x + 1) * 2 '+ 1
     Next x
    
    ''' set other polylines copy from first polyline ( all polylines x are equal ) '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
     For x = 2 To 100
        CopyMemory Pt(0, x).x, Pt(0, 1).x, 4096
     Next x
    
    reAl = True
    minY = 768: maxY = 0
    minLY = 0: maxLY = 768
     
    z = 1
    
'''''''''''''''''''''''''''''
    Dim lp As POINTAPI, co(0 To 100) As Long
    maxViewP = 2
    lastSTP = 1
    
    With AO
        .AlphaOption = AC_SRC_OVER
        .AlphaFlags = 0
        .SourceConstantAlpha = 0
        .AlphaFormat = 0
    End With
    RtlMoveMemory newAO, AO, 4
    AlphaBlend picView.hdc, 0, 0, ResX, ResY, frmBase.picBuffEE.hdc, 0, 0, ResX, ResY, newAO
    AlphaIncrease = True
    
    For x = 1 To 24
        A = RGB(x * 13, x * 13, x * 13)
        COLevels(x) = A
        frmBaseSUP.chkN(x).BackColor = A
    Next x
    
    Jtppp
   
End Sub
 
Public Sub KaInvert(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
  Dim hRgn As Long
    hRgn = CreateRectRgn(x1, y1, x2, y2)
    InvertRgn picBuff.hdc, hRgn
    AlphaBlend picView.hdc, 0, 0, ResX, ResY, picBuff.hdc, 0, 0, ResX, ResY, BlendPtr
    
End Sub


Private Sub Config()

    If PrK(2, LQT) > maxViewP Then
        maxViewP = PrK(2, LQT)
    End If

With frmBase

    If stFirst = 1 Then
        stFirst = 0
        GoTo lb1
    End If
    CycleED
        
        tFa = 1000 / (CDbl(cCycles) / (c2 - clCpu3))
        FPrcS = ((FPrcS + tFa) / 2)
        .txtEFRM = Format$(Round((FPrcS - Process(20, 1)), 2), "0#.#0")
        FpS = (FpS + FPrcS / 1000) / 2
        .txtFpS = Format$(Round(FpS, 2), "###.#0")
        
    CycleST
    If LQT Mod (FpS \ 8 + 1) = 0 Then
    
      If Process(1, 2) <> "Blur" Then
        Process(1, 2) = "Blur"
        Process(2, 2) = "Config Data"
        Process(3, 2) = "Shift Data Arrays"
        Process(4, 2) = "FH R L"
        Process(5, 2) = "Buffer To Screen"
        Process(6, 2) = "Clear Alternate"
        Process(7, 2) = "This Timer"
        Process(8, 2) = "Adjustment Freq"
        Process(9, 2) = "FFT Calculate"
        Process(10, 2) = "TransParent Wins"
        Process(11, 2) = "Color Generate"
        Process(12, 2) = "Set Master Polyline "
        Process(13, 2) = "P2 Set"
        Process(14, 2) = "P2 Draw"
        Process(15, 2) = "Find minY & maxY"
        Process(16, 2) = "P3 Set"
        Process(17, 2) = "P4 Draw"
        Process(18, 2) = "P3 Draw"
        Process(19, 2) = "Sleep"
        Process(20, 2) = ""
     End If
        
        aC = 0
        For xC = 1 To 19
            aC = aC + Process(xC, 1)
        Next xC
        Process(20, 1) = Round(FPrcS - aC, 2)
        aC = aC + Process(20, 1)
        aC = 100 / aC
        aa = Int(Process(20, 1) * aC)
            
        If .chkSortP Then
            For xC = 1 To 19
                For yC = 1 To 19
                    If Process(yC, 1) < Process(xC, 1) Then
                        
                        Process(0, 1) = Process(xC, 1)
                        Process(xC, 1) = Process(yC, 1)
                        Process(yC, 1) = Process(0, 1)
                        Process(0, 2) = Process(xC, 2)
                        Process(xC, 2) = Process(yC, 2)
                        Process(yC, 2) = Process(0, 2)
                        
                    End If
                Next yC
            Next xC
          
          Else
            
            For xC = 1 To 19
                For yC = 1 To 19
                    If Left$(Process(yC, 2), 1) > Left$(Process(xC, 2), 1) Then
                        
                        Process(0, 1) = Process(xC, 1)
                        Process(xC, 1) = Process(yC, 1)
                        Process(yC, 1) = Process(0, 1)
                        Process(0, 2) = Process(xC, 2)
                        Process(xC, 2) = Process(yC, 2)
                        Process(yC, 2) = Process(0, 2)
                        
                    End If
                Next yC
            Next xC
        End If
        
            .lstProcess.Clear
            .lstFunctions.Clear
            .lstPsent.Clear

            For xC = 1 To 19
                .lstProcess.AddItem Process(xC, 1)
                .lstFunctions.AddItem Process(xC, 2)
                .lstPsent.AddItem CStr(Round(Process(xC, 1) * aC, 2)) + "%"
            Next xC
        
        .txtDoEvSleep = Process(20, 1)
        .TextLabel(10) = Format$(FpS / (Val(.TextLabel(5)) + 1), "###%")
        .TextLabel(12) = .txtProcess(17)
        .TextLabel(13) = .txtProcess(18)
    
    End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For xC = 1 To 20
      Process(20, 1) = 0
    Next xC
    
    CycleED
    Process(7, 1) = Round(tFa, 2)
        
lb1:
    
    CycleST
        clCpu2 = C1
        clCpu3 = C1
        FHL = 0: FHR = 0
        FLL = 256: FLR = 256
        sLC = 0: src = 0
    
    For xC = 0 To 255
        SsL(xC) = InData(xC * 2) / 2
        Stro(xC).L = SsL(xC)
        SsR(xC) = InData(xC * 2 + 1) / 2
        Stro(xC).r = SsR(xC)
        If SsL(xC) > FHL Then FHL = SsL(xC)
        If SsR(xC) > FHR Then FHR = SsR(xC)
        If SsL(xC) < FLL Then FLL = SsL(xC)
        If SsR(xC) < FLR Then FLR = SsR(xC)
        sLC = sLC + SsL(xC)
        src = src + SsR(xC)
    Next xC
        
        sLC = (sLC / 256)
        src = (src / 256)
        
        MVolu = (((FHL - FLL) + (FHR - FLR)) / 256)
    
    CycleED
    Process(4, 1) = Round(tFa, 2)
   
    
    CycleST
    
    FFTAudio Stro, FFTStro
    If LQT Mod (FpS \ 8 + 1) = 0 Then
        .PicFFT.Cls
        DrawFFT
        DoEvents
    End If
    
    Call CycleED
    Process(9, 1) = Round(tFa, 2)

    '''''''''''''''''''''' shift '''''''''''''''''''''''''''''''''''''''''''''''''''''

    Call CycleST
      If frmBase.txtspm(6) >= 3 Then
        CopyMemory SsLO(0, SsPtr - 1024), SsLO(0, SsPtr - 1023), 1048576
        CopyMemory SsRO(0, SsPtr - 1024), SsRO(0, SsPtr - 1023), 1048576
      ElseIf frmBase.txtspm(6) >= 2 Then
        CopyMemory SsLO(0, SsPtr - 768), SsLO(0, SsPtr - 767), 786432
        CopyMemory SsRO(0, SsPtr - 768), SsRO(0, SsPtr - 767), 786432
      ElseIf frmBase.txtspm(6) >= 1 Then
        CopyMemory SsLO(0, SsPtr - 512), SsLO(0, SsPtr - 511), 524288
        CopyMemory SsRO(0, SsPtr - 512), SsRO(0, SsPtr - 511), 524288
      Else
        CopyMemory SsLO(0, SsPtr - 256), SsLO(0, SsPtr - 255), 262144
        CopyMemory SsRO(0, SsPtr - 256), SsRO(0, SsPtr - 255), 262144
      End If
    
    Call CycleED
    Process(3, 1) = Round(tFa, 2)
    

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''FixFreq xC''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''FixFreq xC''''''''''''''''''''''''''''''''''''''
    Call CycleST

   If .chkAdjFreq(0) Then
        iC = Val(.txtspm(3).Text)
        For xC = 0 To 255 - iC
            
            bC = 0: b2C = 0
            For aa = 1 To iC
                bC = bC + SsL(xC + aa)
                b2C = b2C + SsR(xC + aa)
            Next aa
            bC = bC / (iC)
            b2C = b2C / (iC)

            For aa = 0 To iC
                SsL(xC + aa) = (SsL(xC + aa) + bC) / 2
                SsR(xC + aa) = (SsR(xC + aa) + b2C) / 2
            Next aa
        
        Next xC
       
         aa = 0
         xC = (SsL(aa) + SsR(aa)) * 0.5
         iC = (SsL(aa) + SsR(aa) + SsL(aa + 1) + SsR(aa + 1)) * 0.25
         SsL(aa) = xC: SsL(aa + 1) = xC
         SsR(aa) = xC: SsR(aa + 1) = iC
    
    End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FixFreq Z '''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FixFreq Z '''
    
    If .chkAdjFreq(1) Then
        iC = Val(.txtspm(4).Text)
        For xC = 0 To 255
            bC = 1: ixC = 1
                
                For aC = iC To 1 Step -1
                       SsL(xC) = (SsL(xC) + SsLO(xC, SsPtr - aC) * ixC)
                       SsR(xC) = (SsR(xC) + SsRO(xC, SsPtr - aC) * ixC)
                    bC = bC + aC
                    ixC = ixC + 1
                Next aC
            
            SsL(xC) = SsL(xC) / bC
            SsR(xC) = SsR(xC) / bC
        Next xC
    End If

    ''''''''''''''''''''''''''
    
    aC = 0: bC = 256
    For iC = 0 To 255
        If SsL(iC) + SsR(iC) > aC Then aC = SsL(iC) + SsR(iC)
        If SsL(iC) + SsR(iC) < bC Then bC = SsL(iC) + SsR(iC)
    Next iC
    xC = Abs(aC - bC)
    FrqVbr = xC
   
   End With

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        CopyMemory SsLO(0, SsPtr), SsL(0), 1024
        CopyMemory SsRO(0, SsPtr), SsR(0), 1024
        
        If SsPtr = 1000 Then SsPtr = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    CycleED
       Process(8, 1) = Round(tFa, 2)
   

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CycleST
    
        ty = ((BassR - BassL) * Bass * 0.1) / (Bass * Midl * 0.1 + 1)
        Vsyy = (Vsyy + ty) * 0.95
        
        tx = ((TrebL - TrebR) * Treb * 0.1) / (Treb * Midl * 0.1 + 1)
        Vsxx = (Vsxx + tx) * 0.95
        
        vsX = Vsxx
        vsY = Vsyy
        
        vsTx = (vsTx * 9 + Vsxx) \ 10
        vsTy = (vsTy * 9 + Vsyy) \ 10
        
        tyP = tyP + FrqVbr * 0.02 * Sgn(BassL - BassR) + Bass - LBass
        txP = txP + FrqVbr * 0.04 * Sgn(TrebL - TrebR)
        vsTx = vsTx + (Sgn(tyP) + Bass * 0.01) + Bass - LBass
        vsTy = vsTy + (Sgn(txP) + Treb * 0.01)
        
        If vsTy > 31 Or vsTy < -31 Then vsTy = vsTy * 0.98
        If vsTx > 31 Or vsTx < -31 Then vsTx = vsTx * 0.98
    
    
'''''''''''''' set freq boxes '''''''''''''''''
    
    With frmBase
      If LQT Mod 50 = 0 Then
        .txtBand = (Int(Bass))
        .txtBandAvg1 = (Int(ABass))
        .txtBR = (Int(BassR))
        .txtBL = (Int(BassL))
        .txtBand2 = (Int(Midl))
        .txtBandAvg2 = (Int(AMidl))
        .txtBR2 = (Int(MidlR))
        .txtBL2 = (Int(MidlL))
        .txtBand3 = (Int(TrebL))
        .txtBandAvg3 = (Int(ATreb))
        .txtBR3 = (Int(TrebR))
        .txtBL3 = (Int(TrebL))
        .txtBLR = (Int(BassR - BassL))
        .txtBLR2 = (Int(MidlR - MidlL))
        .txtBLR3 = (Int(TrebR - TrebL))
       End If
        
        For xC = 0 To 2
                If cS(xC) = 1 Then
                      .txtRGB(xC).BackColor = vbBlue
                  Else
                      .txtRGB(xC).BackColor = &H4359F8
                End If
            If LQT Mod 50 = 0 Then .txtRGB(xC).Refresh       ' this line only decrease txtRGB() refresh speed
        Next xC
        
        .txtRGB(0) = Int(Colv_R) & " + " & cS(0)
        .txtRGB(1) = Int(Colv_G) & " + " & cS(1)
        .txtRGB(2) = Int(Colv_B) & " + " & cS(2)
        

    End With
    CycleED
    Process(2, 1) = Round(tFa, 2)
    
   '''''''' txtLScr ''''''''''
    
    frmBase.txtLScr = Round((frmBase.txtspm(6) * 256 / frmBase.txtFpS), 1)
    frmBase.txtLScr.Refresh

End Sub

    ''''''''''''''''''''''''''''''''''' Loger '''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub Loger()
Dim s As String, s2 As String, s3 As String
Dim P1 As String, p2 As Single

    
    s = LQT2 & " , " & Primes(LQT2) & " , " & PrK(3, LQT2) & " , " & PrK(2, LQT2)
    
   
    If frmBase.lstLogs.ListCount > 32000 Then Exit Sub ' frmBase.lstLogs.Clear
    frmBase.lstLogs.AddItem s
    frmBase.lstLogs.ListIndex = frmBase.lstLogs.ListCount - 1
 
End Sub

Public Sub BaseSub()

Static Wave As WaveHdr, te As Long, te2 As Long
Dim a1 As Integer, a2 As Integer, i As Double
    
    Wave.lpData = VarPtr(InData(0))
    Wave.dwBufferLength = 512
    Wave.dwFlags = 0
    If Fst Then
        InitK
        Cof_X = 1
        iH = 384
        frmBase.fraStart.Visible = False
     End If
    DoEvents
  On Error Resume Next
    Do  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
      Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
      Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
      If DevHandle = 0 Then
          Exit Do
      End If
  

      LQ_ModTime = ((LQT Mod 2) * (LQT Mod 3) + (LQT Mod 5) + (LQT Mod 7))
            
       If frmBase.chkPause = False Then     ' if app not set puase
                
            CycleST
                If frmBase.chkClrAlter Then kaAltrCls
            CycleED
            Process(6, 1) = Round(tFa, 2)
           
           ''''''''''''' Calculated output data before draw ''''''''''''''
            
            Config
               
           ''''''''''''' set colors ''''''''''''''''''''''''''''''''''''''
            
            CycleST
'                ColSet
            CycleED
            Process(11, 1) = Round(tFa, 2)
            
            ''''''''''''''''''''''''''''''''''''''''''''''''''''
            Call CycleST
'               Blur
            Call CycleED
            Process(1, 1) = Round(tFa, 2)
           
           ''''''''''''''' Draw Data To Buffer '''''''''''''''
            
                           
                If frmBase.ChkDraw(1) Then DrawP3
            
           
           ''''''''''''''' Draw Data To Buffer '''''''''''''''
            
            CycleST
               
             With frmBase
               
                If .chkP(0) Then
                    DrawP1
                ElseIf .chkP(1) Then
                    DrawP2
                ElseIf .chkP(2) Then
                    DrawP3
                ElseIf .chkP(3) Then
                    DrawP4
                Else
                    DrawP2
                End If
            
            CycleED
            Process(17, 1) = Round(tFa, 2)
           
            CycleST
             
            If LQT Mod (FpS \ 64 + 1) = 0 Then
                If .fraControls.Height > 360 And .chkTransparent Then
                      StretchBlt .picBCtrl.hdc, 0, 0, .picBCtrl.Width, .picBCtrl.Height _
                    , picBuff.hdc, .fraControls.Left \ 15, 26, .picBCtrl.Width, .picBCtrl.Height, vbSrcCopy
                End If
                If .fraColors.Height > 360 And .chkTransparent Then
                      StretchBlt .picBCol.hdc, 0, 0, .picBCol.Width, .picBCol.Height _
                    , picBuff.hdc, .fraColors.Left \ 15, 24, .picBCol.Width, .picBCol.Height, vbSrcCopy
                End If
                If .fraBlur.Height > 360 And .chkTransparent Then
                     StretchBlt .picBBlur.hdc, 0, 0, .picBBlur.Width, .picBBlur.Height _
                    , picBuff.hdc, .fraBlur.Left \ 15 + 3, 24, .picBBlur.Width, .picBBlur.Height, vbSrcCopy
                End If
                If .fraProcess.Height > 360 And .chkTransparent Then
                    StretchBlt .picBProcs.hdc, 0, 0, .picBProcs.Width, .picBProcs.Height _
                    , picBuff.hdc, 1, 48, .picBProcs.Width, .picBProcs.Height, vbSrcCopy
                End If
                If .picBLogs.Height > 360 And .chkTransparent Then
                    StretchBlt .picBLogs.hdc, 0, 0, .picBLogs.Width, .picBLogs.Height _
                    , picBuff.hdc, .fraLogs.Left \ 15, .fraLogs.Top \ 15 + 25, .picBLogs.Width, .picBLogs.Height, vbSrcCopy
                End If
                
                If frmBase.Enabled Then
                    StretchBlt frmQuran.picBFrmQuran.hdc, 0, 0, frmQuran.picBFrmQuran.Width, frmQuran.picBFrmQuran.Height _
                    , picBuff.hdc, frmQuran.Left \ 15, frmQuran.Top \ 15, frmQuran.picBFrmQuran.Width, frmQuran.picBFrmQuran.Height, vbSrcCopy
                  
                End If
 
            End If
            
            End With
            
            CycleED
            Process(10, 1) = Round(tFa, 2)
            
            '''''''''''''''''''''''''''''''''''''
             
             If GetInputState <> 0 Then
                  frmBase.txtspm(2).Refresh
                  DoEvents
             End If
             
             If LQT Mod (FpS \ 8 + 1) = 0 Then
                 frmBase.txtFpS.Refresh
             End If
                  
       Else
            
            Sleep 10
            DoEvents         '''''' if Pause checked
            
       End If
        
        If DoClickS And DoS > FpS \ 8 Then frmBase.cmdSmaler_Click (idxS): DoS = 0: If idxS <> 14 And idxS <> 15 Then frmBase.txtspm(idxS).Refresh
        If DoClickL And DoL > FpS \ 8 Then frmBase.cmdLarger_Click (idxL): DoL = 0: If idxL <> 14 And idxL <> 15 Then frmBase.txtspm(idxL).Refresh
        DoL = DoL + 1: DoS = DoS + 1
        If DoL > 100 Then DoL = 0
        If DoS > 100 Then DoS = 0
   ''''''''''''''''''''''''''''''''''''''''
        CycleST
        
        
        If frmBase.TextLabel(31) <> "Samples" Then
            frmBase.TextLabel(31).Tag = CStr(Val(frmBase.TextLabel(31).Tag - 1))
            If frmBase.TextLabel(31).Tag = "1" Then frmBase.TextLabel(31) = "Samples"
        End If
        Call frmBase.KeyPrss
        CycleED
        Process(19, 1) = Round(tFa, 2)
        If LQT2 Mod 10 = 1 Then Call frmBase.KeyPrss: DoEvents
         
        frmBase.txtFrm = Round(1000 / ((GetTickCount - lTi) / 1), 2)
        lTi = GetTickCount
    If FpS Mod 5 = 1 Then frmBase.cmdMini.Tag = ""
    Loop While DevHandle <> 0
End Sub


Private Sub Blur()
Dim x As Integer, y As Integer
Dim s  As Double
Dim ts  As Double
Dim te As Integer, te2 As Integer
'On Error Resume Next
    
With frmBase
    ts = 1
    
''''''''''''''''''''''''''''''''''''''''''''''
    SetStretchBltMode picBuff.hdc, 2

''''''''''''''''''''''''''''''''''''''''''''''
    
    If .chkBlur(2) Then  'And Rnd * 1000 > (990 - Rnd * 30)
    
        te = 5 + Rnd * 15
        BitBlt picBuff.hdc, 0, 0, ResX, te, picBuff.hdc, 0, 0, vbBlackness
        BitBlt picBuff.hdc, 0, ResY - te, ResX, te, picBuff.hdc, 0, 0, vbBlackness
        BitBlt picBuff.hdc, 0, 0, te, ResY, picBuff.hdc, 0, 0, vbBlackness
        BitBlt picBuff.hdc, ResX - te, 0, ResX, ResY, picBuff.hdc, 0, 0, vbBlackness
    End If

''''''''''''''''''''''''''''''''''''''''''''''
    If .chkBlur(3) And Rnd * 1000 > (999 - Rnd * (Bass * Bass / 8)) Then
        
        StretchBlt picBuff.hdc, 0, 0, 256, 192 _
              , picBuff.hdc, 0, 0, ResX, ResY, vbSrcCopy
        
            
         For x = 0 To 3
            For y = 0 To 3
                te = x * 256
                te2 = y * 192
                StretchBlt picBuff.hdc, te, te2, 256, 192 _
                  , picBuff.hdc, 0, 0, 256, 192, vbSrcCopy
            Next y
         Next x
   
    End If

''''''''''''''''''''''''''''''''''''''''''''''
    If .chkBlur(2) And LQT Mod (FpS / 64 + 1) = 0 Then         'And LQT Mod (FpS / 32 + 1) = 0
        StretchBlt picBuffSe.hdc, 0, 0, ResX, ResY, _
                   picBuff.hdc, p(1).Zm + vsX, p(1).Zm + vsY _
            , 1024 - vsX * 8 * p(1).Zm, 768 - vsY * 8 * p(1).Zm, vbSrcCopy
'             vsX, vsY, 1024 + vsX * 2, 768 + vsY * 2
    End If
''''''''''''''''''''''''''''''''''''''''''''''

    If .chkBlur(1) Then
        te = 1 '  -Abs(vsX)
        te2 = -2 ' -Abs(vsY) * 2 - 20
        StretchBlt picBuff.hdc, 0, 0, ResX, ResY _
                 , picBuff.hdc, -1, -2, ResX, ResY, vbSrcCopy
    End If
    
''''''''''''''''''''''''''''''''''''''''''''''
    If .chkBlur(0) Then
        StretchBlt picBuff.hdc, 10, 10, 1014 - 10, 758 - 10 _
                 , picBuff.hdc, vsX, vsY, 1024 + vsX * 2, 768 + vsY * 2, vbSrcCopy
    End If
    BlrF = Not BlrF

'''''''''''''''''''''''''''''''''''''''''''''''
End With


End Sub

Public Sub chRGB(Index As Integer)
 Dim A As Integer, b As Integer, c As Long
 Const Rad = 0.0174532925199
    
    If Index = 10 Then Colv_B = 1: Colv_R = 1: Colv_G = 1
    
    If Index = 0 Then cS(0) = 1: cS(1) = 1: cS(2) = 1: _
                      Colv_R = Colv_R * 1.096: Colv_G = Colv_G * 1.096: Colv_B = Colv_B * 1.096
    If Index = 1 Then cS(0) = -1: cS(1) = -1: cS(2) = -1: _
                      Colv_R = Colv_R * 0.9: Colv_G = Colv_G * 0.9: Colv_B = Colv_B * 0.9
    If Index = 2 Then c = Colv_R: Colv_R = Colv_G: Colv_G = Colv_B: Colv_B = c

    If Index = 3 Then c = Colv_R / 2: Colv_R = Colv_G / 2: Colv_G = Colv_B / 2: Colv_B = c / 2
    If Index = 4 Then c = Colv_B / 2: Colv_G = Colv_B / 2: Colv_B = Colv_R / 2: Colv_R = c / 2
    If Index = 5 Then c = Colv_B / 2: Colv_B = Colv_R / 2: Colv_R = Colv_G / 2: Colv_G = c / 2
    
    If Index = 6 Then cS(0) = -cS(0)
    If Index = 7 Then cS(1) = -cS(1)
    If Index = 8 Then cS(2) = -cS(2)
    
End Sub



Public Function RGBRed(RGBCol As Long) As Long
'Return the Red component from an RGB Color
RGBRed = RGBCol And &HFF
End Function


Public Function RGBGreen(RGBCol As Long) As Long
'Return the Green component from an RGB Color
RGBGreen = ((RGBCol And &H100FF00) \ &H100)
End Function


Public Function RGBBlue(RGBCol As Long) As Long
'Return the Blue component from an RGB Color
RGBBlue = (RGBCol And &HFF0000) \ &H10000
End Function

Public Function AbjadClC(str As String) As String
Dim pr As Long, x As Integer, s As String, i As Integer, sS() As String, y As Integer
Dim s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, s6 As String, s7 As String, s8 As String, s9 As String, s10 As String
sS = Split(str, " ", , vbTextCompare)


    For x = 0 To UBound(sS)
        For y = 1 To Len(sS(x))
          i = (Asc(Mid(sS(x), y, 1)))
          If i = 194 Or i = 197 Or i = 195 Then i = 199
          If i = 236 Or i = 198 Then i = 237
          
          If i <> 248 And i <> 243 Then
            s1 = s1 & (Fnts(i, 1)) & "."
            s2 = s2 & (Fnts(i, 2)) & "."
            s3 = s3 & (Fnts(i, 3)) & "."
            s4 = s4 & (Fnts(i, 4)) & "."
            s5 = s5 & (Fnts(i, 5)) & "."
            s6 = s6 & (Fnts(i, 6)) & "."
            s7 = s7 & (Fnts(i, 7)) & "."
          End If
        Next y
        
        s1 = s1 & " * "
        s2 = s2 & " * "
        s3 = s3 & " * "
        s4 = s4 & " * "
        s5 = s5 & " * "
        s6 = s6 & " * "
        s7 = s7 & " * "
    
    Next x
   
   s1 = Replace(s1, ".0.", ".", 1, Len(s1), vbTextCompare)
   s2 = Replace(s2, ".0.", ".", 1, Len(s2), vbTextCompare)
   s3 = Replace(s3, ".0.", ".", 1, Len(s3), vbTextCompare)
   s4 = Replace(s4, ".0.", ".", 1, Len(s4), vbTextCompare)
   s5 = Replace(s5, ".0.", ".", 1, Len(s5), vbTextCompare)
   s6 = Replace(s6, ".0.", ".", 1, Len(s6), vbTextCompare)
   s7 = Replace(s7, ".0.", ".", 1, Len(s7), vbTextCompare)

On Error Resume Next
   s1 = Mid(s1, 1, Len(s1) - 4)
   s2 = Mid(s2, 1, Len(s2) - 4)
   s3 = Mid(s3, 1, Len(s3) - 4)
   s4 = Mid(s4, 1, Len(s4) - 4)
   s5 = Mid(s5, 1, Len(s5) - 4)
   s6 = Mid(s6, 1, Len(s6) - 4)
   s7 = Mid(s7, 1, Len(s7) - 4)
   s = s1 & vbCrLf & s2 & vbCrLf & s3 & vbCrLf & s4 & vbCrLf & s5 & vbCrLf & s6 & vbCrLf & s7
   AbjadClC = s ' Replace(s, ".", vbTab)
End Function



Public Function SendMail(msgBody As String, Subj As String, MailTo As String)

    Dim lobj_cdomsg As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = "smtp.gmail.com"
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = 465
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = True
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = 1
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = ""
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = ""
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = 2
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = MailTo ' ""
    lobj_cdomsg.From = ""
    lobj_cdomsg.Subject = Subj
    lobj_cdomsg.TextBody = msgBody
    'lobj_cdomsg.AddAttachment ("filepath")
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
End Function


