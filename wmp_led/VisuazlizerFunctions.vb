Public Class VisualizerFunctions
    Const black = "000000"

    Public Sub RainbowDots(ByVal led As CLed, ByVal wave As CWave)
        'For c = 0 To 58 Step 2
        'spled.Write("S" & Microsoft.VisualBasic.Right("00" & Hex(c), 2) & Microsoft.VisualBasic.Right("000000" & Hex(Int(256 * 256 * 256 * 10 * CDbl(oMediaDescriptor.harmonic(c + 1)))), 6) & "X")
        'spled.Write("S" & Microsoft.VisualBasic.Right("00" & Hex(c + 1), 2) & Microsoft.VisualBasic.Right("000000" & Hex(Int(256 * 256 * 256 * 10 * CDbl(oMediaDescriptor.harmonic(c + 2)))), 6) & "X")
        'Next
        'spled.Write("Z")
        'ElseIf iSegment = 4 Then
        'spled.Write("S" & IIf(CDbl(oMediaDescriptor.harmonic(3)) / 0.02, "00", "FF") & IIf(CDbl(oMediaDescriptor.harmonic(5)) < 0.02, "00", "FF") & IIf(CDbl(oMediaDescriptor.harmonic(6)) < 0.02, "00", "FF") & "WZ")
        'For c = 0 To nbLed / 2 - 1
        '   WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(c), 2) & Microsoft.VisualBasic.Right("000000" & Hex(Int(256 * 256 * 256 * 10 * CDbl(oMediaDescriptor.harmonic(c + 1)))), 6) & "X")
        '    WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(nbLed - 1 - c), 2) & Microsoft.VisualBasic.Right("000000" & Hex(Int(256 * 256 * 256 * 10 * CDbl(oMediaDescriptor.harmonic(c + 1)))), 6) & "X")
        'Next
        led.WriteSPLED("S000000Y")
        For c = 0 To led.nbLed / 2 - 1
            led.WriteSPLED("S" & Microsoft.VisualBasic.Right("000000" & Hex(ColorFromLong(360 / 24 * c, CDbl(wave.harmonic(c + 1)))), 6) & "L")
            led.WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(c), 2) & "l")
            led.WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(led.nbLed - 1 - c), 2) & "l")
        Next
        'led.WriteSPLED("Z")
    End Sub
    Dim fourPc() As Integer = {0, 0, 3, 3}
    Dim fourPr() As Integer = {0, 4, 0, 4}
    Dim duck0 As Bitmap = file2bmp("C:\Users\Mathieu\Documents\visual studio 2010\Projects\wmp_led\wmp_led\duck0.bmp")
    Dim duck1 As Bitmap = file2bmp("C:\Users\Mathieu\Documents\visual studio 2010\Projects\wmp_led\wmp_led\duck1.bmp")

    Public Sub TestChar(ByVal led As CLed, ByVal wave As CWave)
        Threading.Thread.Sleep(250)
        If wave.sum(8) > 0.01 Then
            led.WriteSPLED(bmp2char(duck1, led))
        Else
            led.WriteSPLED(bmp2char(duck0, led))
        End If
    End Sub

    Public Function file2bmp(ByVal path As String) As Bitmap
        Return New Bitmap(Image.FromFile(path))
    End Function

    Public Function bmp2char(ByVal b As Bitmap, ByVal led As CLed) As String
        'Image.FromFile(path).Size
        'Dim i As Image = Image.FromFile(path)
        'Dim b As Bitmap = New Bitmap(i)
        Dim s As String = "S000000Y"
        For x = 0 To Min(6, b.Width) - 1
            For y = 0 To Min(8, b.Height) - 1
                Dim c = b.GetPixel(x, y)
                If c.R = 255 AndAlso c.G = 255 AndAlso c.B = 255 Then
                Else
                    s = s & "S" & led.grid(x, y) & strColor(b.GetPixel(x, y)) & "X"
                End If
            Next
        Next
        's = s & "Z"
        Return s
    End Function

    Public Sub FourRect(ByVal led As CLed, ByVal wave As CWave)
        Dim p(4) As Double
        p(0) = wave.sum(1, 2, 3, 4) / 4
        p(1) = wave.sum(5, 6, 7, 8, 9, 10) / 6
        p(2) = wave.sum(11, 12, 13, 14, 15, 16) / 6
        p(3) = wave.sum(17, 18, 19, 20, 21, 22, 23) / 6

        led.WriteSPLED("S000000Y")
        'Const cr = 4
        'Const cc = 3
        Dim n As Integer
        For n = 0 To 3
            Dim bc = fourPc(n)
            Dim br = fourPr(n)
            led.WriteSPLED("S" & strColor(ColorFromHSB(wave.precolors(n), 1.0, Min(1.0, p(n)))) & "L")
            For r = 0 To 3
                For c = 0 To 2
                    led.WriteSPLED("S" & led.grid(bc + c, br + r) & "l")
                Next
            Next
        Next
        'led.WriteSPLED("Z")
    End Sub
    Public Sub Equalizer(ByVal led As CLed, ByVal wave As CWave)
        led.WriteSPLED("S000000Y")
        For c = 0 To 7
            Dim p = 16 * wave.sum(c * 4 + 1, c * 4 + 2, c * 4 + 3, c * 4 + 4)
            For i = 0 To 5
                If p * 5 > i Then led.WriteSPLED("S" & led.grid(i, c) & strColor(ColorFromLong(360 / 8 * c, (1.0 / 10.0) * CDbl(i + 4), True)) & "X")
            Next
        Next
        'led.WriteSPLED("Z")
    End Sub

    Public Sub SEqualizer(ByVal led As CLed, ByVal wave As CWave)
        led.WriteSPLED("S000000Y")
        For c = 0 To 7
            Dim p = 16 * wave.sum(c * 4 + 1, c * 4 + 2, c * 4 + 3, c * 4 + 4)
            For i = 0 To 2
                If p * 2 > i Then
                    Dim col = strColor(ColorFromHSB(360 / 8 * c, 1.0, (1.0 / 10.0) * CDbl(i + 4)))
                    led.WriteSPLED("S" & col & "L")
                    led.WriteSPLED("S" & led.grid(i + 3, c) & "l")
                    led.WriteSPLED("S" & led.grid(2 - i, c) & "l")
                End If
            Next
        Next
        'led.WriteSPLED("Z")
    End Sub

    Public Sub DisplayTitle(ByVal led As CLed, ByVal t As Double, ByVal factorspeed As Double, ByVal label As String)
        Dim cp = Int(t * factorspeed)
        Dim charA = getPicChar(Mid(label, cp + 1, 1))
        Dim charB = getPicChar(Mid(label, cp + 2, 1))
        Dim charF = mixChar(charA, charB, Int(t * factorspeed * 6) Mod 6)
        led.WriteSPLED(convPicCharToSerial(charF, "444444"))
        'WriteSPLED(convPicCharToSerial(charF, Replace("   ", " ", Microsoft.VisualBasic.Right("00" & Hex(Int(255 * t / Len(label) * factorspeed)), 2))))
        'Dim v = (t - Int(t))
        'If v >= 0.1 And v <= 0.4 Then
        '            WriteSPLED(convPicCharToSerial(getPicChar("AU"), "FF2000"))
        'ElseIf v >= 0.6 And v <= 0.9 Then
        'WriteSPLED(convPicCharToSerial(getPicChar("AD"), "FF3000"))
        'Else
        'If t < 1 Then
        '            ' Affiche 3
        'WriteSPLED(convPicCharToSerial(getPicChar("3"), "FF0000"))
        'ElseIf t < 2 Then
        '' Affiche 2
        'WriteSPLED(convPicCharToSerial(getPicChar("2"), "FF0000"))
        'ElseIf t < 3 Then
        '' Affiche 1
        'WriteSPLED(convPicCharToSerial(getPicChar("1"), "FF0000"))
        'Else
        '' Affiche 0
        'WriteSPLED(convPicCharToSerial(getPicChar("0"), "FF0000"))
        'End If
        'End If
    End Sub

    Public Shared Function strColor(ByVal col As Long) As String
        Return Microsoft.VisualBasic.Right("000000" & Hex(col), 6)
    End Function

    Public Shared Function strColor(ByVal col As Color) As String
        Return strColor(ColorFromHSB(col.GetHue(), col.GetSaturation(), col.GetBrightness() / 5.0))
        'Return strColor(col.R * 256 * 256 + col.G * 256 + col.B)
    End Function

    Public Sub MultipleShape(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal duration As Double, ByVal iSegment As Integer)
        Dim c1 = IIf(CDbl(wave.harmonic(2) + wave.harmonic(3)) < 0.01, 0, 1)
        Dim c2 = IIf(CDbl(wave.harmonic(4) + wave.harmonic(5)) < 0.01, 0, 1)
        Dim c3 = IIf(CDbl(wave.harmonic(6) + wave.harmonic(7)) < 0.01, 0, 1)
        Dim cLst, imode
        If True OrElse iSegment >= 8 Then
            imode = IIf(iSegment >= 8, "SQ", "CI")
            led.WriteSPLED("S000000Y")
            cLst = VisualizerFunctions.strColor(ColorFromLong(t / duration * 360.0, 1, False))
            'cLst = oLEDParameters.SEG1(1)
            led.WriteSPLED(VisualizerFunctions.convPicCharToSerial(getPicChar(imode & "1"), IIf(c1 = 1, cLst, black), False))
            'cLst = oLEDParameters.SEG1(2)
            led.WriteSPLED(VisualizerFunctions.convPicCharToSerial(getPicChar(imode & "2"), IIf(c2 = 1, cLst, black), False))
            'cLst = oLEDParameters.SEG1(3)
            led.WriteSPLED(VisualizerFunctions.convPicCharToSerial(getPicChar(imode & "3"), IIf(c3 = 1, cLst, black), False))
            'led.WriteSPLED("Z")
        Else
            imode = IIf(iSegment >= 8, "1", "0")
            cLst = wave.SEG1(1)
            led.WriteSPLED("S000000Y")
            led.WriteSPLED("S" & imode & "0" & IIf(c1 = 1, cLst, black) & "V")
            led.WriteSPLED("S" & imode & "5" & IIf(c1 = 1, cLst, black) & "V")
            cLst = wave.SEG1(2)
            led.WriteSPLED("S" & imode & "1" & IIf(c2 = 1, cLst, black) & "V")
            led.WriteSPLED("S" & imode & "4" & IIf(c2 = 1, cLst, black) & "V")
            cLst = wave.SEG1(3)
            led.WriteSPLED("S" & imode & "2" & IIf(c3 = 1, cLst, black) & "V")
            led.WriteSPLED("S" & imode & "3" & IIf(c3 = 1, cLst, black) & "V")
            'led.WriteSPLED("Z")
        End If
    End Sub
    'Dim iCornerPlus As Integer
    Dim cntCP = countCharAnim(anim_CornerPlus)
    Dim cntR = countCharAnim(anim_Rect)
    Public Sub Rectangle(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal d As Double, Optional ByVal speed As Integer = 8, Optional ByVal col As String = "000000")
        'Dim col = strColor(ColorFromLong(Int(wave.mean(1) * 1000.0) Mod 360, 1))
        If Len(col) = 0 OrElse col = "000000" Then col = strColor(ColorFromLong(Int(CDbl(wave.mean(1)) ^ 0.5 * 3000) Mod 360, 1))
        led.WriteSPLED("S000000Y")
        led.WriteSPLED(convPicCharToSerial(getAnimChar(anim_Rect, cntR, Math.Abs(Int(speed * t) Mod ((cntR - 1) * 2) - cntR + 1)), col, False))
        'led.WriteSPLED("Z")
    End Sub
    'Dim TurnAroundPath As Integer() = {0, 1, 2, 3, 4, 5, 11, 17, 23, 29, 35, 41, 47, 46, 45, 44, 43, 42, 36, 30, 24, 18, 12, 6}
    'Const lengthTurn = 6
    Dim lengthShapes As New Hashtable
    Public Sub TurnAround(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal d As Double, ByVal shape As Integer(), ByVal lengthTurn As Integer, ByVal color As String, ByVal speed As Double)
        Dim startAt = Int(t * speed)
        Dim p, r, c
        led.WriteSPLED("S000000Y")
        'Dim lengths = lengthShapes.Item(shape)
        'If lengthTurn Is Nothing Then
        '        lengthShapes.Add(shape, shape.Length)
        'End If
        led.WriteSPLED("S" & color & "L")
        For i = startAt To startAt + lengthTurn
            p = shape(i Mod shape.Length)
            led.WriteSPLED("S" & led.grid(p Mod 6, Int(p / 6)) & "l")
        Next
    End Sub
    Public Sub LineFlash(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal d As Double, ByVal color As String, ByVal speed As Double)
        Dim startAt = Int(t * speed)
        led.WriteSPLED("S000000Y")
        'Dim lengths = lengthShapes.Item(shape)
        'If lengthTurn Is Nothing Then
        '        lengthShapes.Add(shape, shape.Length)
        'End If
        Select Case startAt Mod 4
            Case 0
                led.WriteSPLED("S00" & color & "V")
                led.WriteSPLED("S05" & color & "V")
            Case 1, 3
                led.WriteSPLED("S01" & color & "V")
                led.WriteSPLED("S04" & color & "V")
            Case 2
                led.WriteSPLED("S02" & color & "V")
                led.WriteSPLED("S03" & color & "V")
        End Select
    End Sub
    Public Sub LineFlash2(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal d As Double, ByVal color As String, ByVal speed As Double)
        Dim startAt = Int(t * speed)
        led.WriteSPLED("S000000Y")
        'Dim lengths = lengthShapes.Item(shape)
        'If lengthTurn Is Nothing Then
        '        lengthShapes.Add(shape, shape.Length)
        'End If
        led.WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(Math.Abs((startAt Mod 10) - 4)), 2) & color & "V")
    End Sub
    Dim RainCols As Integer() = {0, 2, 3, 5, 1, 3, 4, 0, 1, 3, 5, 1, 2, 4, 5, 2, 4, 3, 0, 1, 4, 2, 3, 1, 5, 0}
    Public Sub RainDots(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal d As Double, ByVal color As String, ByVal speed As Double)
        Dim startAt = Int(t * speed)
        'Dim p, r, c
        led.WriteSPLED("S000000Y")
        'Dim lengths = lengthShapes.Item(shape)
        'If lengthTurn Is Nothing Then
        '        lengthShapes.Add(shape, shape.Length)
        'End If
        'For i = startAt To startAt + lengthTurn
        '        p = shape(i Mod shape.Length)
        'led.WriteSPLED("S" & led.grid(p Mod 6, Int(p / 6)) & color & "X")
        'Next
        Dim nbCols = 5
        Dim v As Integer = t * speed
        'Dim colStart As Integer = Int(v / 8)
        Dim posStart As Integer = v
        led.WriteSPLED("S" & color & "L")
        For i = 0 To nbCols * 8 - 1 Step 7
            Dim p = (posStart + i)
            Dim r = p Mod 8
            Dim c = ((p - r) / 8)
            led.WriteSPLED("S" & led.grid(RainCols(c Mod RainCols.Length), r) & "l")
            'led.WriteSPLED("S" & led.grid(RainCols(((posStart + i) / 8) Mod RainCols.Length), (posStart + i) Mod 8) & color & "X")
        Next
    End Sub
    Public Sub CornerToPlus(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal d As Double)
        'oMediaDescriptor.mean(1)) ^ 0.5 * 300)))))
        Dim col = strColor(ColorFromLong(Int(CDbl(wave.mean(1)) ^ 0.5 * 3000) Mod 360, 1))
        led.WriteSPLED("S000000Y")
        led.WriteSPLED(convPicCharToSerial(getAnimChar(anim_CornerPlus, cntCP, Math.Abs(Int(2 * t) Mod ((cntCP - 1) * 2) - cntCP + 1)), col, False))
        'led.WriteSPLED("Z")
    End Sub
    Private Function f(ByVal v As Double) As Double
        Return v ^ 2 / 256
        If v < 5 Then
            Return 0
        ElseIf v > 250 Then
            Return 255
        ElseIf v < 60 Then
            Return v * 2
        Else
            Return v ^ 2 / 256
        End If
    End Function
    Public Sub Spirale(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal duration As Double)
        'For 
    End Sub
    Public Sub MovingSquare(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal duration As Double)
        Dim percbass As Double = 4000 * wave.sum(1, 2, 3, 4) / 2
        'If percbass > 100 Then
        'led.WriteSPLED("S000000YSFFFFFF00R")
        'ElseIf percbass > 10 Then
        If percbass > 10 Then
            'led.WriteSPLED("S000000YS000000A0R")
        Else
            led.WriteSPLED("S000000Y")
        End If
        'Dim percbass As Double = 4000 * wave.sum(1, 2, 3, 4) / 2
        Dim voice As Double = 4000 * wave.sum(12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23) / 12
        Dim medium As Double = 4000 * wave.sum(5, 6, 7, 8, 9, 10, 11, 12) / 8
        Dim coef1 = f(percbass)
        Dim coef2 = f(medium)
        Dim coef3 = f(voice)
        'Dim cr As Integer = Min(255, coef1 * oLEDParameters.coef(1, 1) + coef2 * oLEDParameters.coef(2, 1) + coef3 * oLEDParameters.coef(3, 1))
        'Dim cg As Integer = Min(255, coef1 * oLEDParameters.coef(1, 2) + coef2 * oLEDParameters.coef(2, 2) + coef3 * oLEDParameters.coef(3, 2))
        'Dim cb As Integer = Min(255, coef1 * oLEDParameters.coef(1, 3) + coef2 * oLEDParameters.coef(2, 3) + coef3 * oLEDParameters.coef(3, 3))
        'led.WriteSPLED("S000000Y")
        Dim c, r, bc, br
        'bc = Min(3, percbass / 10)
        'br = Min(5, voice / 10)
        Dim tm As Double = t * 3
        bc = Int(tm / 10) Mod 4

        br = (Int(tm) Mod 10) - 2
        If bc Mod 2 = 1 Then br = 6 - br
        'br = 3 + Int(2 * Math.Sin(tm))
        'bc = 2 + Int(2 * Math.Cos(tm))
        Dim col = strColor(ColorFromLong((t * 360.0 / 60) Mod 360, coef1, False))
        led.WriteSPLED("S" & col & "L")
        For c = 0 To 2
            If bc + c < 0 Then Continue For
            If bc + c >= 6 Then Continue For
            For r = 0 To 2
                If br + r < 0 Then Continue For
                If br + r >= 8 Then Continue For
                led.WriteSPLED("S" & led.grid(bc + c, br + r) & "l")
            Next
        Next
        'led.WriteSPLED("Z")
    End Sub


    Public Sub GrowingSquareColor(ByVal led As CLed, ByVal wave As CWave, ByVal t As Double, ByVal duration As Double)
        'Dim c1 As Integer = IIf(CDbl(oMediaDescriptor.harmonic(2) + oMediaDescriptor.harmonic(3)) < 0.01, 0, 1)
        'Dim c2 As Integer = IIf(CDbl(oMediaDescriptor.harmonic(4) + oMediaDescriptor.harmonic(5)) < 0.01, 0, 1)
        'Dim c3 As Integer = IIf(CDbl(oMediaDescriptor.harmonic(6) + oMediaDescriptor.harmonic(7)) < 0.01, 0, 1)
        'Dim c4 As Integer = IIf(CDbl(oMediaDescriptor.harmonic(8) + oMediaDescriptor.harmonic(9)) < 0.01, 0, 1)
        Dim voice As Double = 4000 * wave.sum(12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23) / 12
        Dim percbass As Double = 4000 * wave.sum(1, 2, 3, 4) / 2
        Dim medium As Double = 4000 * wave.sum(5, 6, 7, 8, 9, 10, 11, 12) / 8
        Dim total = voice + percbass + medium
        'Dim white = IIf(oLEDParameters.SEG1(1) = "FFFFFF", "FFFFFF", "FFFFA0")
        If percbass > 10 Then
            led.WriteSPLED("S000000Y")
            Dim cLst = "000000"
            cLst = VisualizerFunctions.strColor(ColorFromHSB(t / duration * 360.0, 1, percbass / 100))
            If percbass > 80 * total / 100 Then led.WriteSPLED(convPicCharToSerial(getPicChar("SQ3"), cLst, False))
            If percbass > 70 * total / 100 Then led.WriteSPLED(convPicCharToSerial(getPicChar("SQ2"), cLst, False))
            cLst = VisualizerFunctions.strColor(ColorFromHSB(t / duration * 360.0, 1, 1))
            If percbass > 50 * total / 100 Then led.WriteSPLED(convPicCharToSerial(getPicChar("SQ1"), cLst, False))
            'WriteSPLED(convPicCharToSerial(getPicChar("SQ2"), IIf(c2 = 1, cLst, black), False))
            'led.WriteSPLED("Z")
        End If
        'WriteSPLED("S" & IIf(percbass > 10 AndAlso percbass > 80 * total / 100, "FFFFFF", IIf(medium > 2 AndAlso medium > 50 * total / 100, "FFFF00", black)) & "YZ")
        'WriteSPLED("S" & IIf(percbass > 10 AndAlso percbass > 80 * total / 100, "FFFFFF", IIf(medium > 2 AndAlso medium > 50 * total / 100, "FFFF00", black)) & "YZ")
        'Dim cLst
        'cLst = oLEDParameters.SEG1(1)
        'WriteSPLED("S00" & IIf(c1 = 1, white, black) & "V")
        'WriteSPLED("S05" & IIf(c1 = 1, white, black) & "V")
        'cLst = oLEDParameters.SEG1(2)
        'WriteSPLED("S01" & IIf(c2 = 1, white, black) & "V")
        'WriteSPLED("S04" & IIf(c2 = 1, white, black) & "V")
        'cLst = oLEDParameters.SEG1(3)
        'WriteSPLED("S02" & IIf(c3 = 1, white, black) & "V")
        'WriteSPLED("S03" & IIf(c3 = 1, white, black) & "V")
        'WriteSPLED("Z")
    End Sub


    Public Interface CLed
        Sub WriteSPLED(ByVal val As String)
        Function grid(ByVal col As Integer, ByVal row As Integer) As String
        Function nbLed() As Integer
    End Interface

    Public Interface CWave
        Function sum(ByVal ParamArray lst() As Integer) As Double
        Function harmonic(ByVal iField As Integer) As String
        Function mean(ByVal p1 As Integer) As String
        Function segmenter(ByVal iField As Integer) As String
        Function SEG1(ByVal led As Integer) As String
        Function precolors(ByVal i As Integer) As Integer
    End Interface

    Public Shared Function convPicCharToSerial(ByVal pChar As String, ByVal pColor As String, Optional ByVal totalSeq As Boolean = True) As String
        Dim s As String = ""
        If totalSeq Then s = "S000000Y"
        Dim i, j, k
        k = 0
        s = s & "S" & pColor & "L"
        For j = 0 To 5
            For i = 0 To 7
                If Mid(pChar, i * 6 + j + 1, 1) = "1" Then s = s & "S" & Microsoft.VisualBasic.Right("00" & Hex(IIf(j Mod 2 = 0, i, 7 - i) + j * 8), 2) & "l"
            Next
            'If i Mod 2 = 1 Then
            '            For j = 5 To 0 Step -1
            '            If Mid(p1, i * 6 + j + 1, 1) = "1" Then s = s & "S" & Microsoft.VisualBasic.Right("00" & Hex(k), 2) & p2 & "X"
            'k = k + 1
            'Next
            'Else
            'For j = 0 To 7
            '            If Mid(p1, i * 6 + j + 1, 1) = "1" Then s = s & "S" & Microsoft.VisualBasic.Right("00" & Hex(k), 2) & p2 & "X"
            'k = k + 1
            'Next
            'End If
        Next
        'If totalSeq Then s = s & "Z"
        Return s
    End Function
    Public Shared Function mixChar(ByVal p1 As String, ByVal p2 As String, ByVal perc As Integer) As String
        Dim i
        Dim s = ""
        For i = 0 To 8
            s = s & Mid(p1, perc + i * 6 + 1, 6 - perc) & Mid(p2, i * 6 + 1, perc)
        Next
        Return s
    End Function

    Public Shared Function invChar(ByVal p1 As String) As String
        Return Replace(Replace(Replace(p1, "1", "Z"), "0", "1"), "Z", "0")
    End Function
End Class
