Public Class Form1
    Implements VisualizerFunctions.CLed
    Implements VisualizerFunctions.CWave
    Dim vf As New VisualizerFunctions
    Function colorsn(ByVal i As Integer) As Integer Implements VisualizerFunctions.CWave.precolors
        Return oLEDParameters.colors(i)
    End Function
    Function SEG1(ByVal i As Integer) As String Implements VisualizerFunctions.CWave.SEG1
        Return oLEDParameters.SEG1(i)
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        lightOperation.RunWorkerAsync()
    End Sub

    Private Sub lightOperation_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles lightOperation.DoWork
        'lightOperation.ReportProgress 
        ' Enumération des .mp3
        'Dim obj As WMPLib.IWMPMediaCollection2
        'obj = mediaPlayer.mediaCollection
        Dim oEncode As System.Text.Encoding = System.Text.Encoding.GetEncoding(1252)

        For j = 0 To oPlaylist.count - 1
            Dim omedia As WMPLib.IWMPMedia
            omedia = oPlaylist.Item(j)
            Dim sFile = omedia.sourceURL
            Dim sPrefix = Mid(sFile, 1, Len(sFile) - 4)
            If Not checkFile(sPrefix) Then
                'If Not My.Computer.FileSystem.FileExists(sPrefix & "_vamp_vamp-libxtract_harmonic_spectrum_amplitudes.csv") Then
                Debug.Print(omedia.name & " | " & omedia.sourceURL)
                lightOperation.ReportProgress(99, sFile)
                'sFile = System.Text.Encoding.Convert(System.Text.Encoding.Default,oEncoding,System.Text.Encoding.Default.GetBytes(sFile)

                Shell("C:\Program Files (x86)\Vamp Plugins\sonic-annotator.exe -d vamp:qm-vamp-plugins:qm-segmenter:segmentation -d vamp:vamp-libxtract:harmonic_spectrum:amplitudes -d vamp:vamp-libxtract:mean:mean """ & (New String(oEncode.GetChars(System.Text.Encoding.Default.GetBytes(sFile)))) & """ -w csv", AppWinStyle.Hide, True, -1)
                'Shell("e:\sonic\sonic-annotator.exe -d vamp:qm-vamp-plugins:qm-segmenter:segmentation """ & sFile & """ -w csv", AppWinStyle.Hide, True, -1)
                'Dim tp = My.Computer.FileSystem.OpenTextFieldParser(sPrefix & "_vamp_vamp-libxtract_harmonic_spectrum_amplitudes.csv", ",")
                'Dim k = -1
                'While Not tp.EndOfData
                'Dim fields() As String = tp.ReadFields()
                'fields(0)
                'If k = -1 Then
                'For k = 1 To fields.Count - 1
                'Debug.Print(k & ":" & fields(k))
                'Next
                'End If
                'End While
                'tp.Close()
                'ListBox1.Items.Add(omedia.sourceURL)
                'System.err.println(omedia.sourceURL)
            End If
            If e.Cancel Then Exit For
        Next
        lightOperation.ReportProgress(100, "Terminé")
        '   oPlaylist.
        'End If
        'mediaPlayer.currentPlaylist.
        'obj.getPlaylistByQuery()
    End Sub

    Private Sub lightOperation_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles lightOperation.ProgressChanged
        Me.Text = e.UserState
    End Sub

    Private Sub lightOperation_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles lightOperation.RunWorkerCompleted

    End Sub

    Private Sub mediaPlayer_CurrentItemChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_CurrentItemChangeEvent) Handles mediaPlayer.CurrentItemChange
        ' Nouvelle lecture
        Dim omedia As WMPLib.IWMPMedia = e.pdispMedia
        oMediaDescriptor.closeFile()
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(omedia.sourceURL) Then
            oMediaDescriptor.setFile(omedia.sourceURL)
            oMediaDescriptor.openFile()
        End If
        'Debug.Print("media : " & omedia.name)
        '
        'mediaPlayer.Ctlcontrols.
        oLEDParameters.reset(LCase(omedia.name))
        'x.NextDouble()
    End Sub
    Dim oLEDParameters As New LDescriptor

    Dim oMediaDescriptor As New FDescriptor
    Private Class LDescriptor
        Public coef(4, 4) As Double
        Dim sLst As String
        Public chars() As String
        Const xSEG1_COLORS = "FFFFFF" & "FF0000" & "00FF00" & "0000FF" & "FFFF00" & "FF00FF" & "00FFFF" & "00FF00" & "FF8800" & "8800FF"
        Public colors(10) As Integer
        Public label As String
        'Const chars_LIST = "HEART*HEART2|:)*:D*XD"
        Public Sub reset(ByVal name As String)
            label = UCase(name)
            If InStr(label, "(") > 0 Then label = Mid(label, 1, InStr(label, "(") - 1)
            label = " " & label & " "
            Dim x As New Random()
            Dim s = "123"
            x.Next()
            For i = 1 To 3
                For j = 1 To 3
                    coef(j, i) = 0
                Next
                Dim c = Int(x.NextDouble * Len(s)) + 1
                coef(CInt(Mid(s, c, 1)), i) = 1
                s = Replace(s, CStr(c), "")
            Next
            For i = 0 To 9
                x.Next()
                colors(i) = Int(360.0 * x.NextDouble)
            Next
            sLst = ""
            Dim cnt = Len(xSEG1_COLORS) / 6
            For i = 1 To 3
                x.Next()
                sLst = sLst & Mid(xSEG1_COLORS, Int(cnt * x.NextDouble()) * 6 + 1, 6)
            Next
            If InStr(name, "love") > 0 OrElse InStr(name, "life") > 0 OrElse InStr(name, "vie") > 0 OrElse InStr(name, "amour") > 0 Then
                chars = New String() {"HEART", "HEART2"}
            Else
                chars = New String() {":)", ":D"}
            End If
        End Sub

        Public Function SEG1(ByVal led As Integer) As String
            Return Mid(sLst, (led - 1) * 6 + 1, 6)
        End Function


    End Class
    Private Class FDescriptor
        Private Class FSubDescriptor
            Dim currentFP As Microsoft.VisualBasic.FileIO.TextFieldParser
            Dim currentLine As String()
            Dim lastLine As String()
            Public Class FNullSD
                Inherits FSubDescriptor
                Public Overloads Sub openFile(ByVal url As String)

                End Sub
                Private Overloads Sub nextLine()

                End Sub
                Public Overloads Sub closeFile()

                End Sub
                Overloads Function readyForTime(ByVal p1 As Double) As Boolean
                    Return False
                End Function
                Overloads Function field(ByVal iField As Integer) As String
                    Return "0"
                End Function
            End Class
            Public Sub openFile(ByVal url As String)
                currentFP = My.Computer.FileSystem.OpenTextFieldParser(url, ",")
                nextLine()
            End Sub

            Private Sub nextLine()
                If currentFP.EndOfData Then
                    If currentLine Is Nothing Then
                        lastLine = Nothing
                    Else
                        currentLine = lastLine
                        lastLine = Nothing
                    End If
                    'lastLine = Nothing
                    'currentLine = Nothing
                Else
                    lastLine = currentLine
                    currentLine = currentFP.ReadFields()
                End If
            End Sub

            Public Sub closeFile()
                lastLine = Nothing
                currentLine = Nothing
                If currentFP Is Nothing Then Exit Sub
                currentFP.Close()
                currentFP = Nothing
            End Sub

            Function readyForTime(ByVal p1 As Double) As Boolean
                If currentFP Is Nothing Then Return False
                If CDbl(field(0)) < p1 Then
                    While Not (currentLine Is Nothing) AndAlso CDbl(currentLine(0)) < p1
                        nextLine()
                    End While
                End If
                Return True
            End Function

            Function field(ByVal iField As Integer) As String
                If lastLine Is Nothing Then
                    If currentLine Is Nothing Then
                        Return "0"
                    Else
                        Return currentLine(iField)
                    End If
                End If
                Return lastLine(iField)
            End Function

        End Class

        Dim oHarmonic As New FSubDescriptor
        Dim oSegmenter As New FSubDescriptor
        Dim oMean As New FSubDescriptor
        Dim oMidi As New FSubDescriptor
        Dim sMidi As String
        Dim currentFName As String
        Dim strongMidiContent As New ArrayList
        'Dim currentFP As Microsoft.VisualBasic.FileIO.TextFieldParser
        Public Sub setFile(ByVal url As String)
            currentFName = url
        End Sub
        Public Sub closeFile()
            oMidi.closeFile()
            If strongMidiContent.Count > 0 Then
                If Len(sMidi) > 0 Then
                    'Dim r = MsgBoxResult.Cancel
                    'If My.Computer.FileSystem.FileExists(sMidi) Then
                    '    While r = MsgBoxResult.Cancel
                    '        r = MsgBox("Ecraser le fichier ?", MsgBoxStyle.YesNoCancel Or MsgBoxStyle.DefaultButton3 Or MsgBoxStyle.Exclamation)
                    '    End While
                    'Else
                    '    r = MsgBoxResult.Yes
                    'End If
                    'If r = MsgBoxResult.Yes Then
                    Dim j = 0
                    While My.Computer.FileSystem.FileExists(sMidi & IIf(j = 0, "", "_" & j))
                        j = j + 1
                    End While
                    Dim sw = My.Computer.FileSystem.OpenTextFileWriter(sMidi & IIf(j = 0, "", "_" & j), False)
                    For i = 0 To strongMidiContent.Count - 1
                        sw.WriteLine(strongMidiContent(i))
                    Next
                    sw.Flush()
                    sw.Close()
                    'End If
                    strongMidiContent = New ArrayList()
                End If
            End If
            sMidi = ""
            oHarmonic.closeFile()
            oSegmenter.closeFile()
            oMean.closeFile()
        End Sub
        Public Sub writeMidi(ByVal s As String)
            strongMidiContent.Add(s)
        End Sub
        Public Sub openFile()
            'currentFName = url
            Dim sPrefix = Mid(currentFName, 1, Len(currentFName) - 4)
            If Form1.checkFile(sPrefix) Then
                Dim sFileH = sPrefix & "_vamp_vamp-libxtract_harmonic_spectrum_amplitudes.csv"
                Dim sFileS = sPrefix & "_vamp_qm-vamp-plugins_qm-segmenter_segmentation.csv"
                Dim sFileM = sPrefix & "_vamp_vamp-libxtract_mean_mean.csv"
                sMidi = sPrefix & "_midi.csv"
                oHarmonic.openFile(sFileH)
                oSegmenter.openFile(sFileS)
                oMean.openFile(sFileM)
                If My.Computer.FileSystem.FileExists(sMidi) Then
                    oMidi = New FSubDescriptor
                    oMidi.openFile(sMidi)
                Else
                    oMidi = New FDescriptor.FSubDescriptor.FNullSD
                End If
            End If
        End Sub

        Sub resetFile()
            closeFile()
            openFile()
        End Sub

        Function readyForTime(ByVal p1 As Double) As Boolean
            Dim x As Boolean = (oHarmonic.readyForTime(p1) AndAlso oSegmenter.readyForTime(p1) AndAlso oMean.readyForTime(p1))
            Dim y As Boolean = oMidi.readyForTime(p1)
            Return y Or x
        End Function
        Function segmenter(ByVal iField As Integer) As String
            Return oSegmenter.field(iField)
        End Function

        Function harmonic(ByVal iField As Integer) As String
            Return oHarmonic.field(iField)
        End Function

        Function mean(ByVal p1 As Integer) As String
            Return oMean.field(p1)
        End Function

        Function midi(ByVal p1 As Integer) As Integer
            Return oMidi.field(p1)
        End Function

    End Class

    Public Function harmonic(ByVal iField As Integer) As String Implements VisualizerFunctions.CWave.harmonic
        Return oMediaDescriptor.harmonic(iField)
    End Function
    Function segmenter(ByVal iField As Integer) As String Implements VisualizerFunctions.CWave.segmenter
        Return oMediaDescriptor.segmenter(iField)
    End Function

    Function cntLed() As Integer Implements VisualizerFunctions.CLed.nbLed
        Return nbLed
    End Function
    Function mean(ByVal p1 As Integer) As String Implements VisualizerFunctions.CWave.mean
        Return oMediaDescriptor.mean(p1)
    End Function

    Private Sub mediaPlayer_PlayStateChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_PlayStateChangeEvent) Handles mediaPlayer.PlayStateChange
        ' Statut lecteur
        If e.newState = WMPLib.WMPPlayState.wmppsStopped Then
            'Debug.Print("state " & e.newState)
            ''//closeFile()
            ''//openfile(currentFName)
            oMediaDescriptor.resetFile()
            'nextLine()
        End If
    End Sub

    Private Sub mediaPlayer_PositionChange(ByVal sender As Object, ByVal e As AxWMPLib._WMPOCXEvents_PositionChangeEvent) Handles mediaPlayer.PositionChange
        ' Position changée
        'resetFile()
    End Sub

    Private Sub ComboBox1_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.GotFocus
        '  ComboBox1.Items = My.Computer.Ports.SerialPortNames()
        Dim s
        ComboBox1.Items.Clear()
        For Each s In My.Computer.Ports.SerialPortNames()
            ComboBox1.Items.Add(s)
        Next
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If spled.IsOpen Then spled.Close()
        spled.PortName = ComboBox1.Items(ComboBox1.SelectedIndex)
        spled.Open()
        'spled.Write("S000000YZ")
    End Sub
    Dim lastTick As Double
    Dim black = "000000"

    Dim nbLed = 60

    Private Sub WriteSPLED(ByVal val As String) Implements VisualizerFunctions.CLed.WriteSPLED
        Try
            spled.Write(val)
        Catch ex As InvalidOperationException
        Catch ex As Exception
            'Debug.Write(".")
            'System.Out.Print(".")
        End Try
    End Sub
    Dim binloop As New Object
    Dim TurnAroundPathRect As Integer() = {0, 1, 2, 3, 4, 5, 11, 17, 23, 29, 35, 41, 47, 46, 45, 44, 43, 42, 36, 30, 24, 18, 12, 6}
    Dim TurnAroundPathSpiraleRect As Integer() = {0, 1, 2, 3, 4, 5, 11, 17, 23, 29, 35, 41, 47, 46, 45, 44, 43, 42, 36, 30, 24, 18, 12, 6, 7, 8, 9, 10, 16, 22, 28, 34, 40, 39, 38, 37, 31, 25, 19, 13, 14, 15, 21, 27, 33, 32, 26, 20, 21}
    Dim TurnAroundPathDoubleRect As Integer() = {6, 7, 1, 2, 8, 9, 3, 4, 10, 11, 17, 16, 22, 23, 29, 28, 34, 35, 41, 40, 46, 45, 39, 38, 44, 43, 37, 36, 30, 31, 25, 24, 18, 19, 13, 12}
    Dim TurnAroundPathHeart As Integer() = {7, 14, 20, 21, 15, 10, 17, 23, 29, 34, 39, 45, 44, 38, 31, 24, 18, 12}
    Dim TurnAroundPathCircle As Integer() = {9, 16, 23, 29, 34, 39, 38, 31, 24, 18, 13, 8}
    Private Sub doWork()
        SyncLock binloop
            If Not spled.IsOpen Then Exit Sub
            'If Threading.Interlocked.Read(binloop) > 0 Then Exit Sub
            'Threading.Interlocked.Increment(binloop)
            Try
                Dim t As Double = mediaPlayer.Ctlcontrols.currentPosition
                lastTick = t
                Dim duration = mediaPlayer.currentMedia.duration
                Dim ready As Boolean = oMediaDescriptor.readyForTime(t)
                If editMode Then
                    'Debug.Print(noteValues(t))
                    oMediaDescriptor.writeMidi(noteValues(t))
                    refreshNotes(ready, t, duration)
                    Exit Try
                End If
                'Me.Text = FormatNumber(t - lastTick, 6)
                'mediaPlayer.settings.
                If Not ready Then Exit Sub
                'Me.Text = CDbl(lastLine(0)) & "|" & IIf(CDbl(lastLine(1)) < 0.02, "0", "1") & IIf(CDbl(lastLine(2)) < 0.02, "0", "1") & IIf(CDbl(lastLine(3)) < 0.02, "0", "1") & IIf(CDbl(lastLine(4)) < 0.02, "0", "1") & IIf(CDbl(lastLine(5)) < 0.02, "0", "1") & IIf(CDbl(lastLine(6)) < 0.02, "0", "1") & IIf(CDbl(lastLine(7)) < 0.02, "0", "1")
                If t = 0 Then
                    WriteSPLED("S000000YZ")
                    'spled.Write("SFFFFFFYZ")
                    Exit Sub
                End If
                'spled.Write(IIf((CDbl(lastLine(4)) + CDbl(lastLine(3)) + CDbl(lastLine(2)) + CDbl(lastLine(1))) > 0.05, "SFFFFFFW", "S000000W"))

                Dim iSegment As Integer = CInt(oMediaDescriptor.segmenter(2))
                Me.Text = iSegment & mediaPlayer.currentMedia.name

                'Me()
                Const factorspeed = 3
                'TODO
                'pluie sur 6 cols (vitesse de défilement basée sur 9)
                'flèche <> qui se croisent
                'For c = 0 To nbLed / 2 - 1
                '    WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(c), 2) & Microsoft.VisualBasic.Right("000000" & Hex(ColorFromLong(iSegment, CDbl(oMediaDescriptor.harmonic(c + 1)))), 6) & "X")
                '    WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(nbLed - 1 - c), 2) & Microsoft.VisualBasic.Right("000000" & Hex(ColorFromLong(iSegment, CDbl(oMediaDescriptor.harmonic(c + 1)))), 6) & "X")
                'Next
                'WriteSPLED("Z")
                'Exit Try
                Dim percbass As Double = sum(1, 2, 3)

                If False AndAlso t < Len(oLEDParameters.label) / factorspeed Then
                    vf.DisplayTitle(Me, t, factorspeed, oLEDParameters.label)
                    'Exit tr
                    'ElseIf percbass > 100 Then
                    'WriteSPLED("S000000YS0000FF00RZ")
                    'ElseIf cntNotes > 0 Then
                    '   refreshNotes(ready, t, duration)
                    '  Exit Try
                ElseIf oMediaDescriptor.midi(0) <> 0 Then
                    For i = Midi.Pitch.C2 To Midi.Pitch.C7
                        notes(i) = oMediaDescriptor.midi(i - Midi.Pitch.C2 + 1)
                        'If notes(i) <> 0 Then
                        'Debug.Print(t & " | " & i & "|" & notes(i))
                        'MsgBox(notes(i))
                        '                        End If
                    Next
                    refreshNotes(ready, t, duration)
                    Exit Try
                ElseIf iSegment = 10 AndAlso InStr(oLEDParameters.label, "RATTLE") > 0 Then
                    vf.TestChar(Me, Me)
                ElseIf iSegment = 10 AndAlso InStr(oLEDParameters.label, "RIDDLE") > 0 Then
                    'vf.TurnAround(Me, Me, t, duration, TurnAroundPath2, 10, "0000FF", 25)
                    vf.TurnAround(Me, Me, t, duration, TurnAroundPathCircle, 3, "FF00FF", 20)
                ElseIf iSegment = 10 Then
                    vf.GrowingSquareColor(Me, Me, t, duration)
                ElseIf iSegment = 6 OrElse iSegment = 9 Then
                    vf.SEqualizer(Me, Me)
                    WriteSPLED("S000000FFR")
                ElseIf (Not iSegment = 9) AndAlso (iSegment <= 1 OrElse iSegment >= 8) Then
                    vf.MultipleShape(Me, Me, t, duration, iSegment)
                ElseIf iSegment = 2 Then
                    vf.RainbowDots(Me, Me)
                ElseIf iSegment = 4 Then
                    vf.CornerToPlus(Me, Me, t, duration)
                ElseIf iSegment = 5 Then
                    vf.Rectangle(Me, Me, t, duration)
                ElseIf iSegment = 7 Then
                    'Dim kkk = Int(t * 20) Mod 8
                    'If kkk = 1 Then
                    '    WriteSPLED("SFFFFFF00R")
                    'ElseIf kkk = 3 Then
                    '    WriteSPLED("S000000FFR")
                    'Else
                    '    WriteSPLED("S00000000R")
                    'End If
                    vf.MovingSquare(Me, Me, t, duration)
                ElseIf iSegment = 3 Then
                    vf.FourRect(Me, Me)
                Else
                    Dim voice As Double = 4000 * sum(12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23) / 12
                    Dim medium As Double = 4000 * sum(5, 6, 7, 8, 9, 10, 11, 12) / 8
                    'Label1.Text = Microsoft.VisualBasic.Right("00" & Hex(Int(percbass)), 2)
                    'Label2.Text = Microsoft.VisualBasic.Right("00" & Hex(Int(medium)), 2)
                    'Label3.Text = Microsoft.VisualBasic.Right("00" & Hex(Int(voice)), 2)
                    Me.Text = oMediaDescriptor.segmenter(2)
                    'Dim envoilepate As Boolean = False
                    'If percbass > 64 AndAlso percbass > medium AndAlso percbass > voice Then
                    '        envoilepate = True
                    'End If
                    'If voice > percbass AndAlso voice > medium Then
                    'percbass = 0
                    'medium = 0
                    'ElseIf medium > percbass AndAlso medium > voice Then
                    'voice = 0
                    'percbass = 0
                    'ElseIf percbass > medium AndAlso percbass > voice Then
                    'medium = 0
                    '                voice = 0

                    'End If
                    'medium = 0
                    ' RGBCMYOPVW
                    ' RG=Y, RB=M, GB=C
                    ' RRG=O,RGG=?,RRB=P,RBB=V,GGB=?,GBB=Azur
                    'voice = (voice + medium + percbass)
                    'medium = voice
                    'percbass = voice
                    Dim coef1 = f(percbass)
                    Dim coef2 = f(medium)
                    Dim coef3 = f(voice)
                    Dim r As Integer = Min(255, coef1 * oLEDParameters.coef(1, 1) + coef2 * oLEDParameters.coef(2, 1) + coef3 * oLEDParameters.coef(3, 1))
                    Dim g As Integer = Min(255, coef1 * oLEDParameters.coef(1, 2) + coef2 * oLEDParameters.coef(2, 2) + coef3 * oLEDParameters.coef(3, 2))
                    Dim b As Integer = Min(255, coef1 * oLEDParameters.coef(1, 3) + coef2 * oLEDParameters.coef(2, 3) + coef3 * oLEDParameters.coef(3, 3))
                    'b = b / 2
                    'g = 0
                    'g = g / 2

                    'Me.Text = b
                    'Me.Text = voice & " | " & percbass
                    Dim v As String = ""
                    Try

                        'Me.Text = Microsoft.VisualBasic.Left(v, 20)
                        'spled.Write("Z")
                        If iSegment = 3 Then
                            WriteSPLED("S000000YS" & Microsoft.VisualBasic.Right("00" & Hex((Int(t * 6) Mod 6)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "V" & "S" & Microsoft.VisualBasic.Right("00" & Hex(5 - (Int(t * 6) Mod 6)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "V")
                        ElseIf iSegment = 4 Then
                            'spled.Write("S" & Microsoft.VisualBasic.Right("00" & Hex((Int(t * 1000) Mod 60)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "XZ")
                            'Me.Text = CDbl(oMediaDescriptor.mean(1)) * 1000000
                            WriteSPLED("S1" & Microsoft.VisualBasic.Right("0" & Hex((Max(0, Min(14, Int(CDbl(oMediaDescriptor.mean(1)) ^ 0.5 * 300))))), 1) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "U")
                            'spled.Write("S0" & Microsoft.VisualBasic.Right("0" & Hex((Int(t * 1) Mod 16)), 1) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "UZ")
                        ElseIf iSegment = 5 Then
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex((Int(t * 250) Mod 6)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "V")
                            'If envoilepate Then
                            'Microsoft.VisualBasic.StrDup(3,
                        ElseIf iSegment = 9 Then
                            'Dim z As String = "S000000YS" & Microsoft.VisualBasic.Right("00" & Hex((Int(t * 30) Mod 6)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "V" & "S" & Microsoft.VisualBasic.Right("00" & Hex(5 - (Int(t * 6) Mod 6)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X"
                            'MsgBox(z)
                            Dim x = 3
                            Dim y = (nbLed)
                            Dim z = -(nbLed * 3 / 4)
                            WriteSPLED("S000000Y")
                            'For c = 0 To 15 Step x
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(((Int(z + t * x)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(((Int(z + t * x + 1)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex((nbLed / 4 + (Int(z + t * x)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex((nbLed / 4 + (Int(z + t * x + 1)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(nbLed - 1 - ((Int(z + t * x)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(nbLed - 1 - ((Int(z + t * x + 1)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(nbLed - 1 - (nbLed / 4 + (Int(z + t * x)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(nbLed - 1 - (nbLed / 4 + (Int(z + t * x + 1)) Mod y) Mod (nbLed / 2)), 2) & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "X")
                            'Next
                            'WriteSPLED("Z")
                        ElseIf iSegment = 7 Then
                            Dim sH = getPicChar(oLEDParameters.chars(IIf(voice > 40, 1, 0)))
                            'If Int(t * 2) Mod 2 = 1 Then
                            'sH = invChar(sH)
                            'End If
                            WriteSPLED(VisualizerFunctions.convPicCharToSerial(sH, Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2)))
                        ElseIf iSegment >= 6 Then
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "Y")
                        ElseIf iSegment = 6 Then
                            WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "Y")
                            'Else
                            'ElseIf iSegment = 7 Then
                            'WriteSPLED("S" & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "WZ")
                            'ElseIf iSegment = 7 Then
                            'spled.Write("S" & Microsoft.VisualBasic.Right("00" & Hex(r), 2) & Microsoft.VisualBasic.Right("00" & Hex(g), 2) & Microsoft.VisualBasic.Right("00" & Hex(b), 2) & "WZ")
                        End If
                        'End If
                    Catch ex As Exception
                        Debug.Print(ex.ToString())
                        Debug.Print(ex.StackTrace.ToString())
                    End Try
                End If
                'Label1.Text = FormatNumber(percbass, 3, TriState.True)
                'If percbass > 0.15 Then WriteSPLED("S000000FFR")
                'If percbass > 0.15 Then WriteSPLED("SFFFFFF00R")
                'If percbass > 0.1 Then WriteSPLED("SFFFF0000R")
                'If percbass > 0.1 Then WriteSPLED("SFFFFFFYSFFFFFF00R")
            Catch ex2 As Exception
                Debug.Print(ex2.Message)
                Debug.Print(ex2.StackTrace.ToString())
            Finally
                WriteSPLED("Z")
            End Try
            System.Threading.Thread.Sleep(19)
            'System.Threading.Thread.Sleep(49)
            'Threading.Interlocked.Decrement(binloop)
        End SyncLock
    End Sub

    Private Sub tmrLed_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrLed.Tick
        doWork()
    End Sub

    Function sum(ByVal ParamArray lst() As Integer) As Double Implements VisualizerFunctions.CWave.sum
        Dim i As Integer
        Dim r As Double
        For i = 0 To lst.Count - 1
            r = r + CDbl(oMediaDescriptor.harmonic(lst(i)))
        Next
        Return r
    End Function
    Const s1Bass = 1
    Const s1Medium = 2
    Const s1Voice = 3
    Const s2R = 1
    Const s2G = 2
    Const s2B = 3
    Dim oPlaylist As WMPLib.IWMPPlaylist

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        oMediaDescriptor.closeFile()
        If Not openedMidi Is Nothing Then
            openedMidi.StopReceiving()
            openedMidi.Close()
            openedMidi = Nothing
        End If
    End Sub
    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'coef(s1Bass, s2R) = 1
        'coef(s1Bass, s2G) = 0.5
        'coef(s1Bass, s2B) = 0.5
        'coef(s1Medium, s2G) = 0.5
        'coef(s1Voice, s2B) = 0.5
        'coef(s1Bass, s2R) = 1
        'coef(s1Medium, s2G) = 1
        'coef(s1Voice, s2B) = 1
        Dim oMPArray As WMPLib.IWMPPlaylistArray
        Dim i, j As Integer
        oMPArray = mediaPlayer.playlistCollection.getAll()
        mediaPlayer.settings.autoStart = False
        For i = 0 To oMPArray.count - 1
            oPlaylist = oMPArray.Item(i)
            If InStr(LCase(oPlaylist.name), "light_") > 0 Then
                mediaPlayer.currentPlaylist = oPlaylist
                Exit For
            End If
        Next
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
    Public Function checkFile(ByVal sPrefix As String) As Boolean
        Dim sFileH = sPrefix & "_vamp_vamp-libxtract_harmonic_spectrum_amplitudes.csv"
        Dim sFileS = sPrefix & "_vamp_qm-vamp-plugins_qm-segmenter_segmentation.csv"
        Dim sFileM = sPrefix & "_vamp_vamp-libxtract_mean_mean.csv"
        Return System.IO.File.Exists(sFileH) AndAlso System.IO.File.Exists(sFileS) AndAlso System.IO.File.Exists(sFileM)
    End Function

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        nbLed = CInt(TextBox1.Text)
    End Sub


    Private Function grid(ByVal col As Integer, ByVal row As Integer) As String Implements VisualizerFunctions.CLed.grid
        Return Microsoft.VisualBasic.Right("00" & Hex(col * 8 + IIf(col Mod 2 = 1, 7 - row, row)), 2)
    End Function
    Dim WithEvents openedMidi As Midi.InputDevice

    'Dim clock As new Midi.Clock (
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'tmrLed.Stop()
        If Not openedMidi Is Nothing Then
            openedMidi.StopReceiving()
            openedMidi.Close()
            openedMidi = Nothing
        End If
        For Each id In Midi.InputDevice.InstalledDevices
            'If InStr(id.Name, "MIDI") > 0 Then Continue For
            openedMidi = id
            Debug.Print(id.Name)
            openedMidi.Open()
            openedMidi.StartReceiving(Nothing)
            Exit For
        Next
    End Sub
    Dim notes(200) As Integer
    Dim cntNotes As Integer
    Dim sustain As Boolean
    Dim collectNotes As New ArrayList()
    Private Class Note
        Public device As String
        Public control As Boolean
        Public channel As Integer
        Public entry As Integer
        Public value As Integer

        Public Sub New(ByVal vdevice As String, ByVal vcontrol As Boolean, ByVal vchannel As Integer, ByVal ventry As Integer, ByVal vvalue As Integer)
            device = vdevice
            control = vcontrol
            channel = vchannel
            entry = ventry
            value = vvalue
        End Sub
    End Class
    'Private Sub proceedNote(ByVal device As String, ByVal control As Boolean, ByVal channel As Integer, ByVal entry As Integer, ByVal value As Integer)
    Private Sub proceedNote(ByVal note As Note)
        If note.control AndAlso note.entry = 64 Then
            If note.value <> 0 Then
                sustain = Not sustain
                If Not sustain Then
                    ' Empty collect
                    While collectNotes.Count > 0
                        proceedNote(collectNotes(0))
                        collectNotes.RemoveAt(0)
                    End While
                End If
            End If
            Exit Sub
        End If
        If sustain Then
            collectNotes.Add(note)
            Exit Sub
        End If
        If False Then
            notes(note.entry) = note.value
        ElseIf note.control Then

            Select Case note.entry
                Case 1 To 3
                    notes(Midi.Pitch.G3 + (note.entry - 1) * 2) = note.value
                Case 6
                    notes(Midi.Pitch.B2) = note.value
                Case 7
                    notes(Midi.Pitch.ASharp2) = note.value
                Case 8
                    notes(Midi.Pitch.A2) = note.value
            End Select
        Else
            Select Case note.entry
                'Case 36 To 38
                'notes(Midi.Pitch.C3 + (entry - 36) * 2) = value
                'Case 39
                'notes(Midi.Pitch.F3) = value
                Case 42
                    notes(Midi.Pitch.A5) = note.value
                Case 43
                    notes(Midi.Pitch.ASharp5) = note.value
                Case 36 To 47
                    notes(Midi.Pitch.C5 + (note.entry - 36)) = note.value
                Case Else
                    notes(note.entry) = note.value
            End Select
        End If
    End Sub

    Private Sub openedMidi_ControlChange(ByVal msg As Midi.ControlChangeMessage) Handles openedMidi.ControlChange
        Debug.Print("CtrlChg " & msg.Device.Name & " / " & msg.Channel & " / " & msg.Control & " / " & msg.Value)
        proceedNote(New Note(msg.Device.Name, True, msg.Channel, msg.Control, msg.Value))
    End Sub

    Private Sub openedMidi_NoteOff(ByVal msg As Midi.NoteOffMessage) Handles openedMidi.NoteOff
        'Debug.Print("NoteOff " & msg.ToString)
        Debug.Print("NoteOff " & msg.Device.Name & " / " & msg.Channel & " / " & msg.Pitch & " / " & msg.Velocity)
    End Sub
    Private Sub openedMidi_NoteOn(ByVal msg As Midi.NoteOnMessage) Handles openedMidi.NoteOn
        'If msg.Velocity = 0 Then
        'WriteSPLED("S00000000RZ")
        'Else
        'WriteSPLED("S000000FFRZ")
        'End If
        '
        'cntNotes = cntNotes + IIf(msg.Velocity = 0, -1, 1)
        'notes(msg.Pitch) = msg.Velocity
        'refreshNotes()
        Debug.Print("NoteOn " & msg.Device.Name & " / " & msg.Channel & " / " & msg.Pitch & " / " & msg.Velocity)
        proceedNote(New Note(msg.Device.Name, False, msg.Channel, msg.Pitch, msg.Velocity))
    End Sub
    Public Sub refreshNotes(ByVal ready As Boolean, ByVal t As Double, ByVal duration As Double)
        WriteSPLED("S000000Y")
        Dim done As Boolean = False
        If notes(Midi.Pitch.A2) > 0 AndAlso (Math.Cos(Math.PI * t / 2.0 * notes(Midi.Pitch.A2))) < (CDbl(notes(Midi.Pitch.ASharp2) - 63.0) / 64.0) Then
            'If notes(Midi.Pitch.A2) > 0 AndAlso (Math.Cos(Math.PI * t / 2.0 * notes(Midi.Pitch.A2))) > IIf(notes(Midi.Pitch.ASharp2) = 0, 0.0, (CDbl(notes(Midi.Pitch.ASharp2) - 63.0) / 64.0)) Then
            WriteSPLED("Z")
            Exit Sub
        End If
        If notes(Midi.Pitch.ASharp3) = 0 Then
            Dim col As String
            Dim colb = IIf(notes(Midi.Pitch.GSharp3) = 0, notes(Midi.Pitch.G3) * 2 * 256 * 256 + notes(Midi.Pitch.A3) * 2 * 256 + notes(Midi.Pitch.B3) * 2, &HFFFFFF)
            If notes(Midi.Pitch.B2) = 0 Then
                col = VisualizerFunctions.strColor(colb)
            Else
                If colb = 0 Then
                    col = VisualizerFunctions.strColor(ColorFromLong((t * notes(Midi.Pitch.B2) * 2.0) Mod 360, 0.5, True))
                Else
                    Dim c2 As Color = Color.FromArgb(colb)
                    col = VisualizerFunctions.strColor(ColorFromHSB(c2.GetHue(), c2.GetSaturation(), Math.Abs(CDbl(Int(100.0 * t * (CDbl(notes(Midi.Pitch.B2)) / 8.0) Mod 200) - 100) / 100) / 2))
                End If
            End If

            'Dim line As Integer = notes(Midi.Pitch.C4) <> 0
            'If notes(Midi.Pitch.C4) <> 0 Then
            'WriteSPLED("S" & col & "Y")
            'Else
            'Dim r = checkNotes(notes, Midi.Pitch.D4, Midi.Pitch.E4, Midi.Pitch.F4)
            done = True
            If ready AndAlso notes(Midi.Pitch.C5) <> 0 Then
                vf.SEqualizer(Me, Me)
            ElseIf ready AndAlso notes(Midi.Pitch.CSharp5) <> 0 Then
                vf.RainbowDots(Me, Me)
            ElseIf ready AndAlso notes(Midi.Pitch.D5) <> 0 Then
                vf.FourRect(Me, Me)
            ElseIf ready AndAlso notes(Midi.Pitch.DSharp5) <> 0 Then
                vf.RainDots(Me, Me, t, duration, col, notes(Midi.Pitch.DSharp5) / 8)
            ElseIf ready AndAlso notes(Midi.Pitch.E5) <> 0 Then
                vf.Rectangle(Me, Me, t, duration, notes(Midi.Pitch.E5) / 6, col)
            ElseIf ready AndAlso notes(Midi.Pitch.F5) <> 0 Then
                vf.TurnAround(Me, Me, t, duration, TurnAroundPathRect, 6, col, notes(Midi.Pitch.F5) / 4)
            ElseIf ready AndAlso notes(Midi.Pitch.FSharp5) <> 0 Then
                vf.TurnAround(Me, Me, t, duration, TurnAroundPathDoubleRect, 7, col, notes(Midi.Pitch.FSharp5) / 4)
            ElseIf ready AndAlso notes(Midi.Pitch.G5) <> 0 Then
                vf.TurnAround(Me, Me, t, duration, TurnAroundPathSpiraleRect, 9, col, notes(Midi.Pitch.G5) / 4)
            ElseIf ready AndAlso notes(Midi.Pitch.GSharp5) <> 0 Then
                vf.TurnAround(Me, Me, t, duration, TurnAroundPathCircle, 5, col, notes(Midi.Pitch.GSharp5) / 4)
            ElseIf ready AndAlso notes(Midi.Pitch.A5) <> 0 Then
                vf.LineFlash(Me, Me, t, duration, col, notes(Midi.Pitch.A5) / 4)
            ElseIf ready AndAlso notes(Midi.Pitch.ASharp5) <> 0 Then
                vf.LineFlash2(Me, Me, t, duration, col, notes(Midi.Pitch.ASharp5) / 4)
            Else
                done = False
                Dim right As Boolean = notes(Midi.Pitch.CSharp4) = 0
                Dim left As Boolean = notes(Midi.Pitch.DSharp4) = 0
                If notes(Midi.Pitch.C4) <> 0 Then
                    If notes(Midi.Pitch.B4) <> 0 Then
                        If left Then WriteSPLED("S10" & col & "V")
                        If right Then WriteSPLED("S15" & col & "V")
                    Else
                        If left Then WriteSPLED("S00" & col & "V")
                        If right Then WriteSPLED("S05" & col & "V")
                    End If
                End If
                If notes(Midi.Pitch.D4) <> 0 Then
                    If notes(Midi.Pitch.B4) <> 0 Then
                        If left Then WriteSPLED("S11" & col & "V")
                        If right Then WriteSPLED("S14" & col & "V")
                    Else
                        If left Then WriteSPLED("S01" & col & "V")
                        If right Then WriteSPLED("S04" & col & "V")
                    End If
                End If
                If notes(Midi.Pitch.E4) <> 0 Then
                    If notes(Midi.Pitch.B4) <> 0 Then
                        If left Then WriteSPLED("S12" & col & "V")
                        If right Then WriteSPLED("S13" & col & "V")
                    Else
                        If left Then WriteSPLED("S02" & col & "V")
                        If right Then WriteSPLED("S03" & col & "V")
                    End If
                End If
                If notes(Midi.Pitch.F4) <> 0 Then WriteSPLED(VisualizerFunctions.convPicCharToSerial(CCharViz.getPicChar("CI1"), col, False))
                If notes(Midi.Pitch.FSharp4) <> 0 Then WriteSPLED(VisualizerFunctions.convPicCharToSerial(CCharViz.getPicChar("SQ1"), col, False))
                If notes(Midi.Pitch.G4) <> 0 Then WriteSPLED(VisualizerFunctions.convPicCharToSerial(CCharViz.getPicChar("CI2"), col, False))
                If notes(Midi.Pitch.GSharp4) <> 0 Then WriteSPLED(VisualizerFunctions.convPicCharToSerial(CCharViz.getPicChar("SQ2"), col, False))
                If notes(Midi.Pitch.A4) <> 0 Then WriteSPLED(VisualizerFunctions.convPicCharToSerial(CCharViz.getPicChar("CI3"), col, False))
                If notes(Midi.Pitch.ASharp4) <> 0 Then WriteSPLED(VisualizerFunctions.convPicCharToSerial(CCharViz.getPicChar("SQ3"), col, False))
            End If
        End If
        'If notes(Midi.Pitch.G4) <> 0 Then VisualizerFunctions.convPicCharToSerial("CI1", col, False)
        'If notes(Midi.Pitch.GSharp4) <> 0 Then VisualizerFunctions.convPicCharToSerial("CI2", col, False)
        'If notes(Midi.Pitch.A4) <> 0 Then VisualizerFunctions.convPicCharToSerial("CI3", col, False)
        'End If
        Dim sFirst = IIf(notes(Midi.Pitch.FSharp3) = 0,
                   IIf(notes(Midi.Pitch.C3) = 0, "00", "FF") &
                   IIf(notes(Midi.Pitch.D3) = 0, "00", "FF") &
                   IIf(notes(Midi.Pitch.E3) = 0, "00", "FF"), "FFFFFF")
        WriteSPLED("S" & sFirst & IIf(notes(Midi.Pitch.F3) = 0, "00", "FF") & "R")
        If (Not done) AndAlso (sFirst = "000000") Then WriteSPLED("SFFQ")
        WriteSPLED("Z")
    End Sub
    Public Function checkNotes(ByRef notes As Integer(), ByVal ParamArray vals() As Integer) As Integer
        Dim i
        For i = 0 To vals.Length - 1
            If notes(vals(i)) <> 0 Then
                Return vals(i)
            End If
        Next
        Return -1
    End Function
    Dim editMode As Boolean
    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        editMode = CheckBox1.Checked
    End Sub

    Public Function noteValues(ByVal l As Double) As String
        Dim s As String = l
        For i = Midi.Pitch.C2 To Midi.Pitch.C7
            s = s & "," & notes(i)
        Next
        Return s
    End Function

    Protected Overrides Sub Finalize()

        MyBase.Finalize()
    End Sub

    Private Sub openedMidi_PitchBend(ByVal msg As Midi.PitchBendMessage) Handles openedMidi.PitchBend
        'Debug.Print("PitchBend " & msg.ToString)
        Debug.Print("PitchBend " & msg.Device.Name & " / " & msg.Channel & " / " & msg.Value)
    End Sub

    Private Sub openedMidi_ProgramChange(ByVal msg As Midi.ProgramChangeMessage) Handles openedMidi.ProgramChange
        Debug.Print("PrgChg " & msg.ToString)
    End Sub
End Class
