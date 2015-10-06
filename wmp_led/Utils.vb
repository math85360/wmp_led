Module Utils
    Public Function ColorFromHSB(ByVal col As Integer, ByVal sat As Double, ByVal bri As Double) As Long
        Dim x As Color = FromHSB(col, sat, bri)
        Return x.R * 256 * 256 + x.G * 256 + x.B
    End Function
    Public Function ColorFromLong(ByVal col As Integer, ByVal l As Double, Optional ByVal full As Boolean = False) As Long
        'Dim r As Double, g As Double, b As Double
        If l <> 0 And Not full Then
            l = 0.01 + 0.49 * Max(0.0, Min(l, 1.0))
        End If

        'If l < 0.5 Then
        '    l = 0
        'Else
        '    l = l - 0.5
        'End If

        Dim x As Color = FromHSB(col Mod 360, 1.0, l)
        'x = Color.Blue

        Return x.R * 256 * 256 + x.G * 256 + x.B
        'HlsToRgb(, , r, g, b)
        'Return Min(255, Max(r, 0)) * 256 * 256 + Min(255, Max(0, g)) * 256 + Min(255, Max(0, b))
    End Function

    Public Function Min(ByVal p1 As Object, ByVal p2 As Object) As Object
        If (p1 < p2) Then
            Return p1
        Else
            Return p2
        End If
    End Function
    Public Function Max(ByVal p1 As Object, ByVal p2 As Object) As Object
        If (p1 > p2) Then
            Return p1
        Else
            Return p2
        End If
    End Function

    Public Function FromHSB(ByVal Hue As Single, ByVal Saturation As Single, ByVal Brightness As Single) As Color
        Hue = Hue Mod 360
        Saturation = Min(1.0, Max(Saturation, 0.0))
        Brightness = Min(1.0, Max(Brightness, 0.0))
        Hue /= 360
        Dim r As Double = 0
        Dim g As Double = 0
        Dim b As Double = 0
        Dim temp1, temp2 As Double
        If Brightness = 0 Then
            r = 0
            g = 0
            b = 0
        Else
            If Saturation = 0 Then
                r = Brightness
                g = Brightness
                b = Brightness
            Else
                If Brightness <= 0.5 Then
                    temp2 = Brightness * (1.0 + Saturation)
                Else
                    temp2 = Brightness + Saturation - Brightness * Saturation
                End If
                temp1 = 2.0 * Brightness - temp2
                Dim t3() As Double = {Hue + 1.0 / 3.0, Hue, Hue - 1.0 / 3.0}
                Dim clr() As Double = {0, 0, 0}
                Dim i As Integer
                For i = 0 To 2
                    If t3(i) < 0 Then t3(i) += 1.0
                    If t3(i) > 1 Then t3(i) -= 1.0
                    If 6.0 * t3(i) < 1.0 Then
                        clr(i) = temp1 + (temp2 - temp1) * t3(i) * 6.0
                    ElseIf 2.0 * t3(i) < 1.0 Then
                        clr(i) = temp2
                    ElseIf 3.0 * t3(i) < 2.0 Then
                        clr(i) = temp1 + (temp2 - temp1) * (2.0 / 3.0 - t3(i)) * 6.0
                    Else
                        clr(i) = temp1
                    End If
                Next i
                r = clr(0)
                g = clr(1)
                b = clr(2)
            End If
        End If
        Return Color.FromArgb(CInt(255 * r), CInt(255 * g), CInt(255 * b))
    End Function
    Private Sub HlsToRgb(ByVal H As Double, ByVal L As Double, _
    ByVal S As Double, ByRef R As Double, ByRef G As _
    Double, ByRef B As Double)
        Dim p1 As Double
        Dim p2 As Double

        If L <= 0.5 Then
            p2 = L * (1 + S)
        Else
            p2 = L + S - L * S
        End If
        p1 = 2 * L - p2
        If S = 0 Then
            R = L
            G = L
            B = L
        Else
            R = QqhToRgb(p1, p2, H + 120)
            G = QqhToRgb(p1, p2, H)
            B = QqhToRgb(p1, p2, H - 120)
        End If
    End Sub

    Private Function QqhToRgb(ByVal q1 As Double, ByVal q2 As _
        Double, ByVal hue As Double) As Double
        If hue > 360 Then
            hue = hue - 360
        ElseIf hue < 0 Then
            hue = hue + 360
        End If
        If hue < 60 Then
            QqhToRgb = q1 + (q2 - q1) * hue / 60
        ElseIf hue < 180 Then
            QqhToRgb = q2
        ElseIf hue < 240 Then
            QqhToRgb = q1 + (q2 - q1) * (240 - hue) / 60
        Else
            QqhToRgb = q1
        End If
    End Function

End Module
