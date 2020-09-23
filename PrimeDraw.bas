Attribute VB_Name = "PrimeDraw"
'******************************************************************************************
'
'   Copyright(C) 2010 By Kaveh Abdollahi.   kavehplus@gmail.com
'   Time Engine
'   June 2010
'
'******************************************************************************************

Option Explicit

Private A As Single, b As Single, coc(0 To 9) As Long, cocS(0 To 255) As Long
Private Stp As Long, BStp As Long, NStp As Long, x As Single, y As Single, yen As Long, yst As Long
Private Rad As Single, SI1 As Single, SI2 As Single, SI3 As Single, SI4 As Single, SI5 As Single, SI6 As Single
Private SI61 As Single, SI62 As Single, SI63 As Single, x2 As Single, y2 As Single
Private pd As Single, cTim As Single, r1, g1, b1
Private xTm As Single, yTm As Single, xTm1 As Single, yTm1 As Single, xTm2 As Single, yTm2 As Single, xTm3 As Single, yTm3 As Single
Private E As Single, m As Single, m2 As Single, cc As Single, c As Single, Ti2 As Single, Ti3 As Single, Ti4 As Single, Ti5 As Single
Private xCntr As Long, yCntr As Long, W(1 To 10) As Single
Private bBol(1 To 10) As Byte, Tmp1 As Single, tmp2 As Single, tmp3 As Single, tmp4 As Single
Public pot As POINTAPI, co As Long, co1 As Long, co2 As Long, co3 As Long, co4 As Long, co5 As Long, co6 As Long
Private cA As Long, cR As Long, cG As Long, cb As Long, s As String
Private PCol(0 To 256, 0 To 2) As Long
Public sz As Currency, Sz1 As Currency, Sz1_b As Currency, Sz2 As Currency, Sz2_b As Currency, Sz3 As Currency, Sz3_b As Currency
Public Ch1 As Double
Public BoardF1 As Long, BoardF2 As Long, BoardF3 As Long
Public COLevels(1 To 24) As Long
Public Shock As Boolean

Public Sub Kdraw()
StartSet
    picTmp.ForeColor = co3
    MoveToEx picTmp.hdc, 512, 1, pot
    DoEvents
    For x = 1 To 768
        LineTo picTmp.hdc, 512 - PrK(2, x) * Log(x) * Sin(x + LQT2), x + PrK(2, x) * Cos(PrK(2, x + LQT2)) * Sin(PrK(2, x + LQT2))
'        MoveToEx picTmp.hdc, x + 1, 512 - PrK(2, x), pot
    Next x
DrawAlpha
End Sub


Public Sub DrawP3()

On Error Resume Next
    With frmBase
    
    StartSet
    
    cTim = LQT2 / 1000000
    xCntr = .txtspm(22)
    yCntr = .txtspm(23)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For b = yst To yen Step SI5
        If Shock Then Shock = False: Exit Sub
        
        BStp = PrK(3, b)
        NStp = PrK(1, b)
        Stp = PrK(2, b)
         
        m = Stp / Ch1 * (SI2)
        c = BStp / Ch1 * (SI3)
        cc = c * c * (SI4)
        E = (m * cc) * (SI1)
        

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        SetCols
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If .chkPant(0) Then
            x = Sin(E * cTim) * Cos(cc * cTim) * cc / SI61 + xCntr
            y = Cos(E * cTim) * Cos(cc * cTim) * cc / SI61 + yCntr
        ElseIf .chkPant(1) Then
            x = Sin(E * cTim + b) * Cos(cc * cTim - b) * cc / SI61 + xCntr
            y = Cos(E * cTim - b) * Cos(cc * cTim + b) * cc / SI61 + yCntr
        Else
            x = Sin(E * cTim) * Cos(cc * cTim) * cc / SI61 + xCntr
            y = Cos(E * cTim) * Sin(cc * cTim) * cc / SI61 + yCntr
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If .chkTimeEnable(1) Then
            picTmp.ForeColor = co3 Xor co2
            picTmp.FillColor = co Xor co3
            xTm = x:   yTm = y
            
            Ellipse picTmp.hdc, xTm - m / 2 / SI62, yTm - m / 2 / SI62, xTm + m / 2 / SI62, yTm + m / 2 / SI62
        End If
    
    ''''''''''''''''''''''''''''''''' Strings '''''''''''''''''''''''''''''''''''''''
        
        If .chkTimeEnable(0) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co2 Xor co3
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(4) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co5 Xor co4
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(6) Then
            SetPixel picTmp.hdc, x, y, co3 Xor co5
        End If
        
        DoEvents
        PXY1(Sz1, 0) = x
        PXY1(Sz1, 1) = y
        Sz1 = Sz1 + 1

    '''''''''''''''''''''''''''''''' Points ''''''''''''''''''''''''''''''''''''''''
        
        If .chkTimeEnable(3) Then
            SetPixel picTmp.hdc, x, y, co4 Xor co5
            For Ti2 = b To b + PrK(1, b) '/ 3
                m = PrK(2, Ti3 + Ti2) / Ch1 * SI2
                c = PrK(3, Ti2) / Ch1 * SI3
                cc = c * c * SI4
                E = m * cc * SI1
                xTm = Sin(Ti2 + cTim) * Cos(b - Ti2 * cTim) * PrK(2, b) / SI63 + x
                yTm = Cos(Ti2 - cTim) * Sin(b + Ti2 * cTim) * PrK(2, b) / SI63 + y
                PXY2(Sz2, 0) = xTm
                PXY2(Sz2, 1) = yTm
                Sz2 = Sz2 + 1

                co5 = RGB(Stp * Tmp1, BStp * Stp / tmp2, Stp * NStp / tmp3) 'Xor co4
                SetPixel picTmp.hdc, xTm, yTm, co4 Xor co3

                If .chkTimeEnable(2) Then
                    If Ti2 = b Then MoveToEx picTmp.hdc, x, y, pot
                    picTmp.ForeColor = co5 Xor co3
                    LineTo picTmp.hdc, xTm, yTm
                End If
                
                If .chkTimeEnable(5) Then
                    For Ti3 = Ti2 To Ti2 + PrK(1, Ti2) '/ .txtspm(12)
                        m = PrK(2, Ti2) / Ch1 * SI2
                        c = PrK(3, Ti2) / Ch1 * SI3
                        cc = c * c * SI4
                        E = m * cc * SI1
                        xTm1 = Cos(cc + Ti3) * Sin(Ti2 * cTim) * PrK(2, Ti3) / SI63 + xTm
                        yTm1 = Sin(cc + Ti2) * Cos(Ti3 * cTim) * PrK(2, Ti3) / SI63 + yTm
                        PXY3(Sz3, 0) = xTm1
                        PXY3(Sz3, 1) = yTm1
                        Sz3 = Sz3 + 1
                        
                        co4 = RGB(PrK(2, Ti3) + PrK(3, Ti3), PrK(2, Ti3) * PrK(3, Ti3) - PrK(3, Ti3), E - PrK(2, Ti3) - PrK(3, Ti3)) Xor co5
                        SetPixel picTmp.hdc, xTm1, yTm1, co4 Xor co3
                        
                        If .chkTimeEnable(2) Then
                            If Ti3 = Ti2 Then MoveToEx picTmp.hdc, xTm, yTm, pot
                            picTmp.ForeColor = co4 Xor co3
                            LineTo picTmp.hdc, xTm1, yTm1
                        End If
                     Next Ti3
                 End If
            
            Next Ti2
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        
    
    Next b
    
    If .chkAutoFix Then .txtspm(33) = BStp \ 8
    If .chkALog Then Loger
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    CycleST
    
        DrawAlpha
        
        DrawTele
     
    CycleED
    Process(5, 1) = Round(tFa, 2)
    
    
    End With
    
Exit Sub
err:
  b = 1
    
End Sub


Public Sub DrawP2()
Dim TM1 As Integer, TM2 As Integer
    On Error Resume Next
   
    With frmBase
    
    StartSet
    
    cTim = LQT2 / 1000000
    xCntr = .txtspm(22)
    yCntr = .txtspm(23)
    
    TM1 = 1: TM2 = 1
    If .chkBGQ(2) Then TM1 = SI5
    If .chkBGQ(3) Then TM2 = SI5
    
    co5 = RGB((LQT2 Mod 4096) / NStp, (LQT2 Mod 2048) / 8, (LQT2 Mod 8192) / Stp)
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    coFl = False
    For b = yst To yen Step SI5
        If Shock Then Shock = False: Exit Sub
        
        BStp = PrK(3, b)
        NStp = PrK(1, b)
        Stp = PrK(2, b)
        
        m = Stp * (SI2) / Ch1
        c = BStp * (SI3) / Ch1
        cc = c * c * (SI4)
        E = (m * cc) * (SI1)
        

        
'        If coFl Then SetCols: coFl = False
        SetCols
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If .chkPant(0) Then
            x = Sin(E * cTim) * Cos(cc * cTim) * cc / SI61 + xCntr
            y = Cos(E * cTim) * Cos(cc * cTim) * cc / SI61 + yCntr
        ElseIf .chkPant(1) Then
            x = Sin(E * cTim + b) * Cos(cc * cTim - b) * cc / SI61 + xCntr
            y = Cos(E * cTim - b) * Cos(cc * cTim + b) * cc / SI61 + yCntr
        Else
            x = Sin(E * cTim) * Cos(cc * cTim) * cc / SI61 + xCntr
            y = Cos(E * cTim) * Sin(cc * cTim) * cc / SI61 + yCntr
        End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If .chkTimeEnable(1) Then
            picTmp.ForeColor = co4 Xor co2
            picTmp.FillColor = co5 And co5
            xTm = x:   yTm = y
            
            Ellipse picTmp.hdc, xTm - m / 2 / SI62, yTm - m / 2 / SI62, xTm + m / 2 / SI62, yTm + m / 2 / SI62
        End If
    ''''''''''''''''''''''''''''''''' Strings '''''''''''''''''''''''''''''''''''''''
        
        If .chkTimeEnable(0) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co Xor co5
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(2) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co2 Xor co5
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(4) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 Xor co5
            LineTo picTmp.hdc, x, y
        End If
            SetPixel picTmp.hdc, x, y, co3 Xor co
        
        If b Mod 29 < 2 Then DoEvents
        
        PXY1(Sz1, 0) = x
        PXY1(Sz1, 1) = y
        Sz1 = Sz1 + 1
        
    '''''''''''''''''''''''''''''''' Points ''''''''''''''''''''''''''''''''''''''''
        
        If .chkTimeEnable(3) Then
            SetPixel picTmp.hdc, x, y, co3 Xor co
            
            For Ti2 = b + NStp To b Step -TM1
                m = PrK(2, Ti2) / Ch1 * SI2
                c = PrK(3, Ti2) / Ch1 * SI3
                cc = c * c * SI4
                E = m * cc * SI1
                xTm = Sin(Ti2 + cTim) * Cos(b - Ti3 * cTim) * PrK(2, b) / SI62 + x
                yTm = Cos(Ti2 - cTim) * Sin(b + Ti2 * cTim) * PrK(2, b) / SI62 + y
                PXY2(Sz2, 0) = xTm
                PXY2(Sz2, 1) = yTm
                Sz2 = Sz2 + 1

                co5 = RGB(BStp, BStp, BStp) Xor co3
                SetPixel picTmp.hdc, xTm, yTm, co5 Xor co1
                
                If .chkTimeEnable(5) Then
                    For Ti3 = Ti2 + PrK(1, Ti2) To Ti2 Step -TM2
                        m = PrK(2, Ti3) / Ch1 * SI2
                        c = PrK(3, Ti3) / Ch1 * SI3
                        cc = c * c * SI4
                        E = m * cc * SI1
                        x = Sin(cc + Ti3) * Cos(b - Ti2 * cTim) * PrK(2, Ti2) / SI63 + x
                        y = Cos(cc + Ti3) * Sin(b + Ti2 * cTim) * PrK(2, Ti2) / SI63 + y
                        PXY3(Sz3, 0) = x
                        PXY3(Sz3, 1) = y
                        Sz3 = Sz3 + 1
                        co4 = RGB(PrK(2, Ti3) + PrK(3, Ti3), PrK(2, Ti3) * PrK(3, Ti3) - PrK(3, Ti3), E - PrK(2, Ti3) - PrK(3, Ti3)) 'Xor co3
                        SetPixel picTmp.hdc, x, y, co4 Xor co3
                     Next Ti3
                 End If
            
            Next Ti2
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
        
    Next b
    
    If .chkAutoFix Then .txtspm(33) = BStp \ 8
    If .chkALog Then Loger
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    CycleST
    
    DrawAlpha
    
    DrawTele
     
    CycleED
    Process(5, 1) = Round(tFa, 2)
    
    
    End With

End Sub

Public Sub DrawP4()
    
    With frmBase
    If .chkLogP Then DistancesCALC
    On Error Resume Next

    LQT = LQT + 1
    If LQT > 3001132 Then LQT = 1
    If LQT2 > 3001132 Then LQT2 = 1
    .txtspm(13) = LQT
    pd = (.txtspm(2))
    If .chkAvalue(0) Then LQT2 = LQT2 + pd
    If .chkAvalue(1) Then LQT2 = LQT2 - pd
    
    .txtspm(11) = LQT2: .txtspm(11).Refresh
    .txtLQT2 = Format$(LQT2, "###,###0") & vbCrLf & _
                Format$(Primes(LQT2), "###,###,###0") & vbCrLf & _
                Format$(PrK(3, LQT2), "####") & vbCrLf & _
                Format$(PrK(2, LQT2), "####") & vbCrLf & _
                Format$(PrK(1, LQT2), "####") & vbCrLf & _
                Format$(Sz1, "###,###,###0")
    .txtLQT2.Refresh
    
    .lblFullscr(1) = Round(LQT2, 2): .lblFullscr(1).Refresh
    .lblFullscr(2) = Primes(LQT2): .lblFullscr(2).Refresh
    .lblFullscr(3) = PrK(2, LQT2): .lblFullscr(3).Refresh
    .lblFullscr(4) = PrK(3, LQT2): .lblFullscr(4).Refresh
    .lblFullscr(5) = PrK(1, LQT2): .lblFullscr(5).Refresh
    
    '''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''
    
    If .ChkDraw(4) Then BitBlt picTmp.hdc, 0, 0, ResX, ResY, picTmp.hdc, 0, 0, 0
    
    '''''''''''''''
   
    If .chkAutoMax Then
          yen = .txtspm(11)
        Else
          yen = .txtspm(21)
    End If
    '''''''''''''''''''''''''''''''''''''''
    If .chkLastP Then
         yst = yen - .txtspm(32)
      Else
         yst = .txtspm(28)
    End If

    cTim = LQT2 / 1000000
    SI1 = 1: SI2 = 1: SI3 = 1: SI4 = 1: SI5 = 1: SI6 = 1
    SI61 = 1: SI62 = 1: SI63 = 1
    If IsNumeric(.txtspm(16)) Then SI1 = .txtspm(16)
    If IsNumeric(.txtspm(17)) Then SI2 = .txtspm(17)
    If IsNumeric(.txtspm(18)) Then SI3 = .txtspm(18)
    If IsNumeric(.txtspm(19)) Then SI4 = .txtspm(19)
    If IsNumeric(.txtspm(20)) Then SI5 = .txtspm(20)
    If IsNumeric(.txtspm(33)) Then SI6 = .txtspm(33)

    SI61 = SI6: SI62 = SI6: SI63 = SI6
    If 0 <> .chkZx(0).Value Then SI61 = SI6 * 2
    If 0 <> .chkZx(1).Value Then SI62 = SI6 * 2
    If 0 <> .chkZx(2).Value Then SI63 = SI6 * 2
    bBol(1) = .chkCol(1)
    bBol(2) = .chkCol(2)
    bBol(3) = .chkCol(3)
    bBol(4) = .chkCol(4)
    bBol(5) = .chkCol(5)
    bBol(0) = .chkCol(0)
    
    bBol(6) = .chkBox

    xCntr = .txtspm(22)
    yCntr = .txtspm(23)
    DoEvents
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For b = yst To yen Step SI5
        If Shock Then Shock = False: Exit Sub
       
        BStp = PrK(3, b)
        NStp = PrK(1, b)
        Stp = PrK(2, b)
         
        m = Stp * (SI2)
        c = BStp * (SI3)
        cc = c * c * (SI4)
        E = (m * cc) * (SI1)
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        SetCols
        co4 = RGB(Stp, Stp, Stp)
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        If .chkPant(0) Then
            x = Sin(E * cTim) * Cos(cc * cTim) * xCntr / SI61 + xCntr
            y = Cos(E * cTim) * Cos(cc * cTim) * yCntr / SI61 + yCntr
        ElseIf .chkPant(1) Then
            x = Sin(E * cTim + b) * Cos(cc * cTim - b) * xCntr / SI61 + xCntr
            y = Cos(E * cTim - b) * Cos(cc * cTim + b) * yCntr / SI61 + yCntr
        Else
            x = Sin(E * cTim) * Cos(cc * cTim) * xCntr / SI61 + xCntr
            y = Cos(E * cTim) * Sin(cc * cTim) * yCntr / SI61 + yCntr
        End If
        PXY1(Sz1, 0) = x
        PXY1(Sz1, 1) = y
        Sz1 = Sz1 + 1
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(1) Then
            picTmp.ForeColor = co3 Xor co4
            picTmp.FillColor = co3 Xor co2
            xTm1 = x
            yTm1 = y
            Ellipse picTmp.hdc, xTm1 - m / 2 / SI62, yTm1 - m / 2 / SI62, xTm1 + m / 2 / SI62, yTm1 + m / 2 / SI62
        End If
        ''''''''''''''''''''''''''''''''' Strings '''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(0) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 Xor co4
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(2) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 And co2
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(4) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 Xor coB
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(6) Then
'            B = B + 1
        End If
        DoEvents
        '''''''''''''''''''''''''''''''' Strings ''''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(3) Then
            For Ti2 = b To b + NStp
                m = PrK(2, Ti2) * SI2
                c = PrK(3, Ti2) * SI3
                cc = c * c * SI4 '/ prk(1, B)
                E = m * cc * SI1
                xTm1 = Cos(cc * Ti2 + cTim / Stp) * Cos(E * Ti2) * x / SI62 + x
                yTm1 = Sin(cc * Ti2) * Cos(E * Ti2 + cTim / Stp) * y / SI62 + y

                co5 = RGB(Stp * Tmp1, BStp * Stp / tmp2, Stp * NStp / tmp3) ' xor coB
                SetPixel picTmp.hdc, xTm1, yTm1, co5 Xor co3
                
                PXY2(Sz2, 0) = xTm1
                PXY2(Sz2, 1) = yTm1
                Sz2 = Sz2 + 1
                If .chkTimeEnable(5) Then
                    For Ti3 = 2 To NStp - Stp + 1 '+ PrK(2, Ti2) / 2
                        m = PrK(2, Ti3) * SI2
                        c = PrK(3, Ti3) * SI3
                        cc = c * c * SI4
                        E = m * cc * SI1
                        x = Cos(cc * Ti2 * cTim) * Cos(E * Ti2) * cc / SI63 + xTm1
                        y = Sin(cc * Ti2) * Cos(E * Ti2 * cTim) * cc / SI63 + yTm1
                        SetPixel picTmp.hdc, x, y, RGB(c - m, cc / m, c + m * 2) Xor coB
                        PXY3(Sz3, 0) = x
                        PXY3(Sz3, 1) = y
                        Sz3 = Sz3 + 1
                     Next Ti3
                 End If
            
            Next Ti2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
enx:

    Next b
    
    If .chkALog Then Loger
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    CycleST

    DrawAlpha
    
    DrawTele

    CycleED
    Process(5, 1) = Round(tFa, 2)
    
    End With
    
End Sub

Public Sub DrawP1()
    
    With frmBase
    
    On Error Resume Next
    If .chkLogP Then DistancesCALC

    LQT = LQT + 1
    If LQT > 3001132 Then LQT = 1
    If LQT2 > 3001132 Then LQT2 = 1
    .txtspm(13) = LQT
    pd = (.txtspm(2))
    If .chkAvalue(0) Then LQT2 = LQT2 + pd
    If .chkAvalue(1) Then LQT2 = LQT2 - pd
    .txtspm(11) = LQT2: .txtspm(11).Refresh
    .txtLQT2 = Format$(LQT2, "###,###0") & vbCrLf & _
        Format$(Primes(LQT2), "###,###,###0") & vbCrLf & _
        Format$(PrK(3, LQT2), "####") & vbCrLf & _
        Format$(PrK(2, LQT2), "####") & vbCrLf & _
        Format$(Sz1, "###,###,###0")
    .txtLQT2.Refresh
    
    .lblFullscr(1) = Round(LQT2, 2): .lblFullscr(1).Refresh
    .lblFullscr(2) = Primes(LQT2): .lblFullscr(2).Refresh
    .lblFullscr(3) = PrK(2, LQT2): .lblFullscr(3).Refresh
    .lblFullscr(4) = PrK(3, LQT2): .lblFullscr(4).Refresh
    
    '''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''
    
    If .ChkDraw(4) Then BitBlt picTmp.hdc, 0, 0, ResX, ResY, picTmp.hdc, 0, 0, 0
    
    
    If .chkAutoMax Then
          yen = .txtspm(11)
        Else
          yen = .txtspm(21)
    End If
    '''''''''''''''''''''''''''''''''''''''
    If .chkLastP Then
         yst = yen - .txtspm(32)
      Else
         yst = .txtspm(28)
    End If

    cTim = LQT2 / 1000000
    SI1 = 1: SI2 = 1: SI3 = 1: SI4 = 1: SI5 = 1: SI6 = 1
    SI1 = .txtspm(16) '* 2
    SI2 = .txtspm(17) '* 2
    SI3 = .txtspm(18) '* 2
    SI4 = .txtspm(19) '* 2
    SI5 = .txtspm(20) '* 2
    SI6 = .txtspm(33) '* 2

    bBol(1) = .chkCol(1)
    bBol(2) = .chkCol(2)
    bBol(3) = .chkCol(3)
    bBol(4) = .chkCol(4)
    bBol(5) = .chkCol(5)
    bBol(0) = .chkCol(0)
'
    bBol(6) = .chkBox

    xCntr = .txtspm(22)
    yCntr = .txtspm(23)
                                                                                                                
    co5 = RGB((LQT2 Mod 4096) / 16, (LQT2 Mod 2048) / 8, (LQT2 Mod 8192) / 32)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    For b = yst + Stp To yen Step SI5
        If Shock Then Shock = False: Exit Sub
        BStp = PrK(3, b)
        Stp = PrK(2, b)
        
        m = Stp * (SI2)
        cc = BStp * (SI3)
        If .chkCM Then cc = Stp * (SI3)
        cc = cc * cc * (SI4)
        E = (m * cc) * (SI1)
        
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        SetCols
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .chkPant(0) Then
            x = Sin(E * cTim) * Cos(cc * cTim) * cc / SI61 + xCntr
            y = Cos(E * cTim) * Cos(cc * cTim) * cc / SI61 + yCntr
        ElseIf .chkPant(1) Then
            x = Sin(E * cTim + b) * Cos(cc * cTim - b) * cc / SI61 + xCntr
            y = Cos(E * cTim - b) * Cos(cc * cTim + b) * cc / SI61 + yCntr
        Else
            x = Sin(E * cTim) * Cos(cc * cTim) * cc / SI61 + xCntr
            y = Cos(E * cTim) * Sin(cc * cTim) * cc / SI61 + yCntr
        End If
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(1) Then
            picTmp.ForeColor = co2 And co4
            picTmp.FillColor = co3 And co5
            xTm1 = x
            yTm1 = y
            Ellipse picTmp.hdc, xTm1 - m / SI6, yTm1 - m / SI6, xTm1 + m / SI6, yTm1 + m / SI6
        End If
        ''''''''''''''''''''''''''''''''' Strings '''''''''''''''''''''''''''''''''''''''
        If .chkTimeEnable(0) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 Xor co5
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(2) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co2 Or co3 Xor co5
            LineTo picTmp.hdc, x, y
        End If
        If .chkTimeEnable(4) Then
            If b = yst Then MoveToEx picTmp.hdc, x, y, pot
            picTmp.ForeColor = co3 Xor co4 Xor co3
            LineTo picTmp.hdc, x, y
        End If
        
        PXY1(Sz1, 0) = x
        PXY1(Sz1, 1) = y
        Sz1 = Sz1 + 1
        DoEvents
        '''''''''''''''''''''''''''''''' Strings ''''''''''''''''''''''''''''''''''''''''
        SetPixel picTmp.hdc, x, y, co4 Xor co5
        If .chkTimeEnable(3) Or .chkTimeEnable(5) Then
            
            For Ti2 = b To b + PrK(1, b) \ 3
                
                m = PrK(2, b) * SI2
                cc = PrK(3, b) * SI3
                cc = cc * cc * SI4
                E = (m * cc) * SI1
                xTm1 = Cos(cc * Ti2 + cTim / Stp) * Sin(E * Ti2) * PrK(2, b) / SI6 + xTm1
                yTm1 = Sin(cc * Ti2) * Cos(E * Ti2 + cTim / Stp) * PrK(2, b) / SI6 + yTm1
                SetPixel picTmp.hdc, xTm1, yTm1, co Xor co3
                PXY2(Sz2, 0) = xTm1
                PXY2(Sz2, 1) = yTm1
                Sz2 = Sz2 + 1
                        
                If .chkTimeEnable(5) Then
                    For Ti3 = Ti2 To Ti2 + PrK(1, Ti2) \ 3 '17
                        m = PrK(2, Ti2) * SI2
                        cc = PrK(3, Ti2) * SI3
                        cc = cc * cc * SI4
                        E = (m * cc) * SI1
                        x = Cos(cc * Ti3) * Sin(E + m * cTim) * PrK(2, b) / SI6 + x
                        y = Sin(cc * Ti3) * Cos(E - cc * cTim) * PrK(2, b) / SI6 + y
                        SetPixel picTmp.hdc, x, y, co3 And co2
                        PXY3(Sz3, 0) = x
                        PXY3(Sz3, 1) = y
                        Sz3 = Sz3 + 1
                     Next Ti3
                 End If
            
            Next Ti2
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        End If
enx:
       
    Next b
    
    If .chkAutoFix Then .txtspm(33) = BStp \ 8 + 1
    If .chkALog Then Loger
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    CycleST

       DrawAlpha
       DrawTele

    CycleED
    Process(5, 1) = Round(tFa, 2)
    
    End With

End Sub


Public Sub SetCols()
    
    With frmBase
        co4 = RGB(NStp, NStp, NStp)
        Tmp1 = .slrCol(0).Value: tmp2 = .slrCol(1).Value: tmp3 = .slrCol(2).Value
        co1 = RGB(Tmp1 * tmp2, tmp2 * tmp3, tmp3 * Tmp1)
        
        If bBol(1) Then
            co = RGB(256 - Stp * 5, 256 / Stp * 3, Stp * 5) Xor coB
            co2 = RGB(256 - Stp * 3, 256 - Stp * 5, Stp * 7)
            co3 = RGB(256 - Stp * 7, 256 / Stp * 7, Stp * 3)
        ElseIf bBol(2) Then
            co = RGB(256 - Stp * 7, 256 / Stp * 7, Stp * 5) Xor coB
            co2 = RGB(256 - Stp * 7, 256 / Stp * 7, Stp * 5)
            co3 = RGB((BStp - Stp + NStp) * Tmp1, (NStp - Stp) * tmp2, (BStp - NStp) * tmp3)
        ElseIf bBol(3) Then
            co = RGB(256 / Stp * 7, 256 / Stp * 3, 256 / Stp * 5) Xor coB
            co2 = RGB(256 / Stp * 3, 256 / Stp * 5, 256 / Stp * 7)
            co3 = RGB(256 / Stp * Tmp1, 256 / Stp * tmp2, 256 / Stp * tmp3)
        ElseIf bBol(4) Then
            co = RGB(256 - Stp * 3, 256 / Stp * 4, Stp * 2) Xor coB
            co2 = RGB(256 - Stp * 2, 256 / Stp * 3, Stp * 4)
            co3 = RGB(256 - PrK(1, Stp) * Tmp1, 256 - PrK(1, NStp) * tmp2, Stp - NStp + PrK(1, NStp) * tmp3)
        ElseIf bBol(5) Then
            co = RGB(256 - cc / Stp, cc / Stp, 256 - cc / Stp)
            co2 = RGB(cc / Stp, 256 - E / cc, cc / Stp) Xor coB
            co3 = RGB(256 - Stp * Tmp1, 256 - NStp * tmp2, (NStp + Stp) * tmp3)
        ElseIf bBol(0) Then
            co = RGB(256 / Stp * Tmp1, 256 / Stp * tmp2, 256 / Stp * tmp3) Xor coB
            co2 = RGB(256 / Stp * 3, 256 / Stp * 5, 256 / Stp * 7)
            co3 = RGB(BStp / Stp * Tmp1, BStp / Stp * tmp2, BStp / Stp * tmp3)
        Else
            co = RGB(cc / Stp, cc / Stp, cc / Stp) Xor coB
            co2 = RGB(cc / Stp, 256 - cc / Stp * 2, cc / Stp)
            co3 = RGB(E / Stp, 256 - E / Stp * 2, 256 - cc / Stp)
        End If
'        co6 = RGBBlue(co3)
        
   End With
End Sub



Public Sub DistancesCALC()
Dim zX As Currency, d As Currency, dist As Currency, s As String, sum1 As Currency, sum2 As Currency
Dim z As Double, z2 As Double, z3 As Double
Dim bx1 As Long, bx2 As Long, bx3 As Long
Dim by1 As Long, by2 As Long, by3 As Long
Dim la1 As Integer, la2 As Integer
Dim sngDx As Single
Dim sngDy As Single
'Exit Sub
z3 = 1 / Val(frmBase.txtspm(19))
z2 = frmBase.txtspm(20)
z = frmBase.txtspm(33)
        
        Sz1_b = Sz1_b - 1
        For zX = 2 To Sz1_b
            sngDx = PXY1(zX, 0) - PXY1(zX - 1, 0)
            sngDy = PXY1(zX, 1) - PXY1(zX - 1, 1)
            d = Sqr(sngDx * sngDx + sngDy * sngDy)
            dist = dist + d * z
        Next zX
        bx1 = PXY1(zX - 1, 0): by1 = PXY1(zX - 1, 1)
        d = GetDistance(xCntr, yCntr, bx1, by1)
        s = "  - All Points in Orbt 1 : " & Format$(Sz1_b * z2, "###,###,###,###,###,###,###0") & "    " & " Strings Lenght : " & Format$(dist * z2 * z3, "###,###,###,###,###,###0") & " Pixel's " & " ( Step " & frmBase.txtspm(20) & " )"
        Sz1_b = Sz1
        sum1 = Sz1 - 1
        Sz1 = 1
        sum2 = dist
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
        dist = 0
        Sz2_b = Sz2_b - 1
        For zX = 2 To Sz2_b
            sngDx = PXY2(zX, 0) - PXY2(zX - 1, 0)
            sngDy = PXY2(zX, 1) - PXY2(zX - 1, 1)
            d = Sqr(sngDx * sngDx + sngDy * sngDy)
            dist = dist + d * z
        Next zX
        bx2 = PXY2(zX - 1, 0): by2 = PXY2(zX - 1, 1)
        d = GetDistance(bx2, by2, bx1, by1)
        s = s & vbCrLf & "  -- All Points in Orbt 2 : " & Format$(Sz2_b, "###,###,###,###,###,###,###,###,###0") & "    " & " Strings Lenght : " & Format$(dist * z2 * z3, "###,###,###,###,###,###0") & " Pixel's"
        Sz2_b = Sz2
        sum1 = sum1 + Sz2 - 1
        sum2 = sum2 + dist
        Sz2 = 1
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        dist = 0
        Sz3_b = Sz3_b - 1
        For zX = 2 To Sz3_b
            sngDx = PXY3(zX, 0) - PXY3(zX - 1, 0)
            sngDy = PXY3(zX, 1) - PXY3(zX - 1, 1)
            d = Sqr(sngDx * sngDx + sngDy * sngDy)
            dist = dist + d * z
        Next zX
        bx3 = PXY3(zX - 1, 0): by3 = PXY3(zX - 1, 1)
        d = GetDistance(bx2, by2, bx3, by3)
        s = s & vbCrLf & "  --- All Points in Orbt 3 : " & Format$(Sz3_b, "###,###,###,###,###,###,###,###,###,###0") & "    " & " Strings Lenght : " & Format$(dist * z2 * z3, "###,###,###,###,###,###0") & " Pixel's"
        Sz3_b = Sz3
        sum1 = sum1 + Sz3
        sum2 = sum2 + dist
        Sz3 = 1
    
    With frmBase
    
    la1 = 0: la2 = 0
    For zX = 1 To 24
        
        If Val(.txtspm(28)) < Val(.List1.List(zX)) And Val(.txtspm(28)) >= Val(.List1.List(zX - 1)) Then
            la1 = zX
        End If
        If Val(.txtspm(21)) >= Val(.List1.List(zX - 1)) And Val(.txtspm(21)) < Val(.List1.List(zX)) Then
            la2 = zX + 1
        End If
        frmBaseSUP.chkN(zX).Value = 0
    Next zX
    frmBaseSUP.chkN(19).Value = 0
    
    For zX = la1 To la2
       frmBaseSUP.chkN(zX).Value = 1
    Next zX
    
    .Text4.Text = ""
        s = s & vbCrLf & vbCrLf & " Orbt 1 + Orbt 2 + Orbt 3 Total Strings Lenght : " & Format$(sum2 * z2 * z3, "###,###,###,###,###,###0") & " Pixel's "
        
        s = s & vbCrLf & vbCrLf & " Start on Primes ( " & Format$(.txtspm(28), "###,###,###0") & " ) " & " in Ring ( " & la1 & " ) = " & Format$(Primes(.txtspm(28)), "###,###,###,###,###,###0") & vbCrLf _
            & " Finish on Pimes ( " & Format$(.txtspm(21), "###,###,###0") & " )" & " in Ring ( " & la2 & " ) = " & Format$(Primes(.txtspm(21)), "###,###,###,###,###,###0")
        s = s & vbCrLf & vbCrLf & " E * " & .txtspm(16) & vbTab & " M * " & .txtspm(17) & vbTab & " C * " & .txtspm(18) & vbTab & " CC * " & .txtspm(19)
        
'        If .chkPant(0) + .chkPant(1) + .chkPant(2) = 0 Then .chkPant(1) = 1: .chkPant(0) = 0: .chkPant(2) = 0: DoEvents
        s = s & vbCrLf & "  " & 1 + .chkPant(0) * 1 + .chkPant(1) * 3 + .chkPant(2) * 2 & " Dimension View"
        .Text4 = s
    
   .picText(1).Height = 3000
    End With
    
End Sub

Public Sub StartSet()
Dim x As Long, d As Long, dist As Currency
On Error Resume Next
    
    With frmBase
        Ch1 = .txtR
'        If .chkShotAll.Value And .chkAutoShot.Value Then .cmdSF_Click
        
        If .ChkDraw(4) Then
           If .chkBW Then
                BitBlt picTmp.hdc, 0, 0, ResX, ResY, picTmp.hdc, 0, 0, vbWhite
                .Combo1.ListIndex = 12
           Else
                BitBlt picTmp.hdc, 0, 0, ResX, ResY, picTmp.hdc, 0, 0, 0
           End If
         End If
        
        
        
        If .chkAutoMax Then
             yen = .txtspm(11)
            Else
             yen = .txtspm(21)
        End If
        '''''''''''''''''''''''''''''''''''''''
        If .chkLastP Then
             yst = yen - .txtspm(32)
          Else
             yst = .txtspm(28)
        End If
    
        If .chkRGB_mu Then
            .slrCol(0).Value = (.slrCol(0).Value * 3 + ABass / 5) / 4
            .slrCol(1).Value = (.slrCol(1).Value * 3 + AMidl / 4) / 4
            .slrCol(2).Value = (.slrCol(1).Value * 3 + ATreb / 5) / 4
            co1 = RGB(.slrCol(0).Value, .slrCol(1).Value, .slrCol(2).Value)
        End If
    
        LQT = LQT + 1
        If LQT > 3001132 Then LQT = 1
        If LQT2 > 3001132 Then LQT2 = 1
        .txtspm(13) = LQT
        pd = (.txtspm(2))
        If .Check1 Then .txtspm(28) = .txtspm(28) + 1
        If .chkAvalue(0) Then LQT2 = LQT2 + pd
        If .chkAvalue(1) Then LQT2 = LQT2 - pd
        
        .txtspm(11) = LQT2: .txtspm(11).Refresh
        .txtLQT2 = Format$(LQT2, "###,###0") & vbCrLf & _
            Format$(Primes(LQT2), "###,###,###0") & vbCrLf & _
            Format$(PrK(3, LQT2), "####") & vbCrLf & _
            Format$(PrK(2, LQT2), "####") & vbCrLf & _
            Format$(PrK(1, LQT2), "####") & vbCrLf & _
            Format$(Sz1 - 1, "###,###,###0")
        .txtLQT2.Refresh
  
        
    If .chkAutoFix Then .txtspm(33) = BStp / 5 + 0.1
        .lblFullscr(1) = Round(LQT2, 2): .lblFullscr(1).Refresh
        .lblFullscr(2) = Primes(LQT2): .lblFullscr(2).Refresh
        .lblFullscr(3) = PrK(2, LQT2): .lblFullscr(3).Refresh
        .lblFullscr(4) = PrK(3, LQT2): .lblFullscr(4).Refresh
        .lblFullscr(5) = PrK(1, LQT2): .lblFullscr(5).Refresh
        SI1 = 1: SI2 = 1: SI3 = 1: SI4 = 1: SI5 = 1: SI6 = 1
        SI61 = 1: SI62 = 1: SI63 = 1
        If IsNumeric(.txtspm(16)) Then SI1 = .txtspm(16)
        If IsNumeric(.txtspm(17)) Then SI2 = .txtspm(17)
        If IsNumeric(.txtspm(18)) Then SI3 = .txtspm(18)
        If IsNumeric(.txtspm(19)) Then SI4 = .txtspm(19)
        If IsNumeric(.txtspm(20)) Then SI5 = .txtspm(20)
        If IsNumeric(.txtspm(33)) Then SI6 = .txtspm(33)
        
    
        SI61 = SI6: SI62 = SI6: SI63 = SI6
        If 0 <> .chkZx(0).Value Then SI61 = SI6 * 2
        If 0 <> .chkZx(1).Value Then SI62 = SI6 * 2
        If 0 <> .chkZx(2).Value Then SI63 = SI6 * 2
        bBol(1) = .chkCol(1)
        bBol(2) = .chkCol(2)
        bBol(3) = .chkCol(3)
        bBol(4) = .chkCol(4)
        bBol(5) = .chkCol(5)
        bBol(0) = .chkCol(0)
        bBol(6) = .chkBox
 
    If .chkLogP Then DistancesCALC
End With


End Sub


Public Sub DrawAlpha()
Dim z As Integer, x As Integer, y As Integer, ttm As String, s As String
Dim z2 As Integer, x2 As Integer, y2 As Integer, Tp As Integer

With frmBase
    SetStretchBltMode picBuff.hdc, 4
    ttm = 0
    
    If .fraProcess.Visible = False Then ttm = 1
        z2 = 1: x2 = 30: y2 = 15
        .picText(1).Cls
    
    If .chkLogP Then
        ttm = (.fraProcess.Height / Screen.TwipsPerPixelY)
        If .fraProcess.Visible = False Then ttm = 1
        .picText(1).ForeColor = &HEED0EE
        .picText(1).Print .Text4.Text
        StretchBlt .picBuffEE.hdc, 0, ttm, 1000, .picText(1).Height / 15, .picText(1).hdc, 0, 0, 1000, .picText(1).Height / 15, vbSrcPaint
    End If
    
    
    If .chkText2Pic Then
            
            If .chkTxt2Pic(4) Then
                Tp = 300 - 10 * Val(.txtTextSize(4).Text)
                .pic2Text(4).Cls
                TextBanner .txtPad(4).Text, .pic2Text(4), vbCrLf, .txtTextSize(4).BackColor, vbBlack
                StretchBlt .picBuffEE.hdc, 0, Tp, .pic2Text(4).Width \ 15, .pic2Text(4).Height \ 15, .pic2Text(4).hdc, 0, 0, .pic2Text(4).Width \ 15, .pic2Text(4).Height \ 15, vbSrcPaint
            End If
                        
            If .chkTxt2Pic(0) Then
                Tp = 300 - 10 * Val(.txtTextSize(0).Text)
                .pic2Text(0).Cls
                TextBanner "  " & .txtPad(0).Text & "   ", .pic2Text(0), vbCrLf, .txtTextSize(0).BackColor, vbBlack
                StretchBlt .picBuffEE.hdc, 0, Tp, .pic2Text(0).Width \ 15, .pic2Text(0).Height \ 15, .pic2Text(0).hdc, 0, 0, .pic2Text(0).Width \ 15, .pic2Text(0).Height \ 15, vbSrcPaint
            End If
            
            If .chkTxt2Pic(1) Then
                Tp = Tp + 300 - 10 * Val(.txtTextSize(1).Text)
                .pic2Text(1).Cls
                TextBanner "   " & .txtPad(1).Text & "  ", .pic2Text(1), vbCrLf, .txtTextSize(1).BackColor, vbBlack
                StretchBlt .picBuffEE.hdc, 0, Tp, .pic2Text(1).Width \ 15, .pic2Text(1).Height \ 15, .pic2Text(1).hdc, 0, 0, .pic2Text(1).Width \ 15, .pic2Text(1).Height \ 15, vbSrcPaint
            End If
            
            If .chkTxt2Pic(2) Then
                Tp = Tp + 300 - 10 * Val(.txtTextSize(2).Text)
                .pic2Text(2).Cls
                TextBanner "  " & .txtPad(2).Text & "  ", .pic2Text(2), vbCrLf, .txtTextSize(2).BackColor, vbBlack
                StretchBlt .picBuffEE.hdc, 0, Tp, .pic2Text(2).Width \ 15, .pic2Text(2).Height \ 15, .pic2Text(2).hdc, 0, 0, .pic2Text(2).Width \ 15, .pic2Text(2).Height \ 15, vbSrcPaint
            End If
            
            If .chkTxt2Pic(3) Then
                Tp = Tp + 300 - 10 * Val(.txtTextSize(3).Text)
                .pic2Text(3).Cls
                TextBanner "  " & .txtPad(3).Text & "  ", .pic2Text(3), vbCrLf, .txtTextSize(3).BackColor, vbBlack
                StretchBlt .picBuffEE.hdc, 0, Tp, .pic2Text(3).Width \ 15, .pic2Text(3).Height \ 15, .pic2Text(3).hdc, 0, 0, .pic2Text(3).Width \ 15, .pic2Text(3).Height \ 15, vbSrcPaint
            End If
                        
            If .chkTxt2Pic(5) Then
                Tp = Tp + 300 - 10 * Val(.txtTextSize(5).Text)
                .pic2Text(5).Cls
                TextBanner .pic2Text(5), .pic2Text(3), vbCrLf, .txtTextSize(5).BackColor, vbBlack
                StretchBlt .picBuffEE.hdc, 0, Tp, .pic2Text(5).Width \ 15, .pic2Text(5).Height \ 15, .pic2Text(5).hdc, 0, 0, .pic2Text(5).Width \ 15, .pic2Text(5).Height \ 15, vbSrcPaint
            End If
                                
            If .chkTxt2PicAN And .txtNumber < 31 Then          '  Number Ayeh & 7 Prime Levels
                .pic2Text(7).Cls
                .pic2Text(7).Width = 9000
                .pic2Text(7).Left = frmBase.Width \ 15 - .pic2Text(7).Width \ 15 - 30
                .pic2Text(7).Print Primes(Primes(Primes(Primes(Primes(Primes(Primes(.txtNumber))))))) & " : " & Primes(Primes(Primes(Primes(Primes(Primes(.txtNumber)))))) & " : " & Primes(Primes(Primes(Primes(Primes(.txtNumber))))) & " : " & Primes(Primes(Primes(Primes(.txtNumber)))) & " : " _
                    & Primes(Primes(Primes(.txtNumber))) & " : " & Primes(Primes(.txtNumber)) & " : " & Primes(.txtNumber) & "  ( " & .txtNumber & " )ÇáÈÞÑå " & "     "
                .pic2Text(7).ForeColor = vbWhite
                .pic2Text(7).FontSize = 14
                StretchBlt .picBuffEE.hdc, frmBase.Width \ 15 - .pic2Text(7).Width \ 15 - 10, 260 - 10 * Val(.txtTextSize(0).Text), .pic2Text(7).Width \ 15, .pic2Text(7).Height \ 15, _
                           .pic2Text(7).hdc, 0, 0, .pic2Text(7).Width \ 15, .pic2Text(7).Height \ 15, vbSrcPaint
            ElseIf .chkTxt2PicAN Then
                .pic2Text(7).Cls
                .pic2Text(7).Width = 1800
                .pic2Text(7).FontSize = 14
                .pic2Text(7).Left = frmBase.Width \ 15 - .pic2Text(7).Width \ 15 - 30
                .pic2Text(7).Print " ( " & .txtNumber & " ) ÇáÈÞÑå " & "     "
                .pic2Text(7).ForeColor = vbWhite

                StretchBlt .picBuffEE.hdc, frmBase.Width \ 15 - .pic2Text(7).Width \ 15 - 10, 260 - 10 * Val(.txtTextSize(0).Text), .pic2Text(7).Width \ 15, .pic2Text(7).Height \ 15, _
                           .pic2Text(7).hdc, 0, 0, .pic2Text(7).Width \ 15, .pic2Text(7).Height \ 15, vbSrcPaint
            End If
            
            ''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''
            ''''''''''''''''''''''''''''''''''
            
            .pic2Text(7).Cls
            .pic2Text(7).Width = 4400
            .pic2Text(7).Height = 4400
            .pic2Text(7).Left = frmBase.Width \ 15 - .pic2Text(7).Width \ 15 - 30
            .pic2Text(7).FontSize = 14: .pic2Text(7).FontBold = False
            .pic2Text(7).Print " " & Trim(.txtMain)
            .pic2Text(7).ForeColor = vbWhite
            StretchBlt .picBuffEE.hdc, frmBase.Width \ 15 - .pic2Text(7).Width \ 15 - .txtTextSize(5) * 30, frmBase.Height \ 15 - .pic2Text(7).Height \ 15, .pic2Text(7).Width \ 15, .pic2Text(7).Height \ 15, _
                       .pic2Text(7).hdc, 0, 0, .pic2Text(7).Width \ 15, .pic2Text(7).Height \ 15, vbSrcPaint
'        DoEvents
    End If
        
        DoEvents
        .pic2Text(6).Width = 3840
        .pic2Text(6).Cls:    s = .lblFullscr(0).Caption & "  " & .lblFullscr(9).Caption
        .pic2Text(6).Print s
        StretchBlt .picBuffEE.hdc, .picBuffEE.Width \ 15 - .pic2Text(6).Width \ 15, .picBuffEE.Height \ 15 - .pic2Text(6).Height \ 15, .pic2Text(6).Width \ 15, .pic2Text(6).Height \ 15, .pic2Text(6).hdc, 0, 0, .pic2Text(6).Width \ 15, .pic2Text(6).Height \ 15, vbSrcPaint
        
        .pic2Text(6).Width = .picBuffEE.Width - 3840
        .pic2Text(6).Cls:    s = "Time index: " & .lblFullscr(1).Caption & vbTab & "Time Value: " & .lblFullscr(2).Caption & vbTab & vbTab & .lblFullscr(3).Caption & vbTab & .lblFullscr(4).Caption & vbTab & .lblFullscr(5).Caption & vbTab & .lblFullscr(7).Caption & "   " & .lblFullscr(8).Caption
        s = s & vbTab & vbTab & .txtFrm & " fps" & vbTab & vbTab & .lVideoInfo.Caption
        .pic2Text(6).Print s
        StretchBlt .picBuffEE.hdc, 10, .picBuffEE.Height \ 15 - .pic2Text(6).Height \ 15, .pic2Text(6).Width \ 15, .pic2Text(6).Height \ 15, .pic2Text(6).hdc, 0, 0, .pic2Text(6).Width \ 15, .pic2Text(6).Height \ 15, vbSrcPaint
    
        If .chkImg2Pic Then
             StretchBlt .picBuffEE.hdc, (ResX - .txtspm(34)) \ 2 + .txtspm(36), (ResY - .txtspm(35)) \ 2 + .txtspm(14), .txtspm(34), .txtspm(35), _
                .picStore.hdc, 0, 0, .picStore.Width \ 15, .picStore.Height \ 15, vbSrcPaint
        End If
        If .chkImg2Pic2 Then
             StretchBlt .picBuffEE.hdc, (ResX - .picStore2.Width) \ 15 \ 2, (ResX - .picStore2.Height) \ 15 \ 2, .picStore2.Width \ 15, .picStore2.Height \ 15, _
                .picStore2.hdc, 0, 0, .picStore2.Width \ 15, .picStore2.Height \ 15, vbSrcPaint
        End If
        
    AO.SourceConstantAlpha = CByte(.txtspm(7))
    If .chkAlpha Then AO.SourceConstantAlpha = CByte(.txtspm(7)) / 3
    If .chkAlphaEnable Then AO.SourceConstantAlpha = AO.SourceConstantAlpha Xor (CByte(.txtspm(7)) / 3)
    RtlMoveMemory newAO, AO, 4
    
'    StretchBlt picView.hdc, 0, 0, ResX - 1, ResY - 1, .picBuffEE.hdc, 0, 0, ResX-1, ResY-1, vbSrcCopy
'    AlphaBlend picBuffEE2.hdc, 0, 0, ResX - 1, ResY - 1, .picBuffEE.hdc, 0, 0, ResX - 1, ResY - 1, newAO
    AlphaBlend picView.hdc, 0, 0, ResX - 1, ResY - 1, .picBuffEE.hdc, 0, 0, ResX - 1, ResY - 1, newAO
    
    DoEvents
    If .chkShotAll.Value And .chkAutoShot.Value And TiS Mod Val(Int(.txtspm(8)) + 1) <= 1 Then .cmdSF_Click
    
    
End With

End Sub

Public Sub DrawTele()
With frmBase
     
     Nvg(1) = .txtspm(36): Nvg(2) = .txtspm(34): Nvg(3) = .txtspm(35)

'     If .fraTelo.Visible = True And .chkImg2Pic Then
'         StretchBlt .picBuffEE.hdc, (ResX - 480) \ 2, (ResY - 480) \ 2, 480, 480, _
'          .picTele.hdc, 0, 0, 480, 480, vbSrcPaint
'     End If

     If bBol(6) Then
        picView.ForeColor = vbYellow
        MoveToEx picView.hdc, Nvg(2), Nvg(3), pot '0

        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3)
        LineTo picView.hdc, Nvg(2) + Nvg(1), Nvg(3) + Nvg(1)

        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)

        LineTo picView.hdc, Nvg(2), Nvg(3)
        LineTo picView.hdc, Nvg(2), Nvg(3) + Nvg(1)
     End If

End With
End Sub

Public Function GetDistance(ByVal lX1 As Long, ByVal lY1 As Long, ByVal lX2 As Long, ByVal lY2 As Long) As Long

    Dim sngDx As Single
    Dim sngDy As Single

    sngDx = lX2 - lX1
    sngDy = lY2 - lY1
    
    GetDistance = Sqr(sngDx * sngDx + sngDy * sngDy)

End Function

'''''''''                xtm1 = Cos(Cc * Ti2 * cTim) * Cos(E + B / cTim) * m \ SI6 + x
'''''''''                ytm1 = Sin(Cc * Ti2 * cTim) * Cos(E - B / cTim) * m \ SI6 + y
'''''''''                xtm1 = Sin(E * cTim) * Cos(Cc * cTim) * Cc \ SI6 + x
'''''''''                ytm1 = Cos(E * cTim) * Sin(Cc * cTim) * Cc \ SI6 + y
'''''''''                xtm1 = Sin(Ti2 + cTim) * Cos(B - Ti3 * cTim) * m \ SI62 + x
'''''''''                ytm1 = Cos(Ti2 - cTim) * Cos(B + Ti2 * cTim) * m \ SI62 + y



