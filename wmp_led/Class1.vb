﻿Module CCharViz
    Private char_EXCL = "" &
        "000000" &
        "000100" &
        "000100" &
        "000100" &
        "000100" &
        "000100" &
        "000000" &
        "000100"
    Private char_SMILEY = "" &
        "000000" &
        "110011" &
        "110011" &
        "000000" &
        "100001" &
        "011110" &
        "000000" &
        "000000"
    Private char_SMILEYD = "" &
        "000000" &
        "110011" &
        "110011" &
        "000000" &
        "011110" &
        "100001" &
        "011110" &
        "000000"
    Private char_DOT = "" &
        "000000" &
        "000000" &
        "000000" &
        "000000" &
        "000000" &
        "000000" &
        "000000" &
        "000100"
    Private char_APOS = "" &
        "000000" &
        "000100" &
        "000100" &
        "001000" &
        "000000" &
        "000000" &
        "000000" &
        "000000"
    Private char_COMA = "" &
        "000000" &
        "000000" &
        "000000" &
        "000000" &
        "000000" &
        "000100" &
        "000100" &
        "001000"
    Private char_MINUS = "" &
        "000000" &
        "000000" &
        "000000" &
        "000000" &
        "011111" &
        "000000" &
        "000000" &
        "000000"
    Private char_HEART = "" &
        "000000" &
        "010010" &
        "101101" &
        "101101" &
        "100001" &
        "010010" &
        "001100" &
        "001100"
    Private char_HEART2 = "" &
        "000000" &
        "010010" &
        "111111" &
        "111111" &
        "111111" &
        "011110" &
        "001100" &
        "001100"
    Private char_0 = "" &
            "000000" &
            "001110" &
            "010001" &
            "010001" &
            "010001" &
            "010001" &
            "010001" &
            "001110"
    Private char_1 = "" &
        "000000" &
        "000100" &
        "001100" &
        "010100" &
        "000100" &
        "000100" &
        "000100" &
        "011111"
    Private char_2 = "" &
        "000000" &
        "001110" &
        "010001" &
        "000010" &
        "000100" &
        "001000" &
        "010000" &
        "011111"
    Private char_3 = "" &
        "000000" &
        "001110" &
        "010001" &
        "000010" &
        "001100" &
        "000010" &
        "010001" &
        "001110"
    Private char_AMP = "" &
        "000000" &
        "000110" &
        "001001" &
        "000101" &
        "000010" &
        "001101" &
        "010010" &
        "001100"
    Private char_BO = "" &
        "000000" &
        "000100" &
        "001000" &
        "010000" &
        "010000" &
        "010000" &
        "001000" &
        "000100"
    Private char_BC = "" &
        "000000" &
        "000100" &
        "000010" &
        "000001" &
        "000001" &
        "000001" &
        "000010" &
        "000100"
    Private char_ARRLD = "" &
        "000000" &
        "000000" &
        "000001" &
        "000010" &
        "100100" &
        "101000" &
        "110000" &
        "111100"
    Private char_ARRU = "" &
        "000000" &
        "000100" &
        "001110" &
        "010101" &
        "000100" &
        "000100" &
        "000100" &
        "000100"
    Private char_ARRD = "" &
        "000000" &
        "000100" &
        "000100" &
        "000100" &
        "000100" &
        "010101" &
        "001110" &
        "000100"
    Private char_A = "" &
        "000000" &
        "000100" &
        "001010" &
        "001010" &
        "011111" &
        "010001" &
        "010001" &
        "010001"
    Private char_B = "" &
        "000000" &
        "011110" &
        "010001" &
        "010001" &
        "011110" &
        "010001" &
        "010001" &
        "011110"
    Private char_C = "" &
        "000000" &
        "001110" &
        "010001" &
        "010000" &
        "010000" &
        "010000" &
        "010001" &
        "001110"
    Private char_D = "" &
        "000000" &
        "011110" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "011110"
    Private char_E = "" &
        "000000" &
        "011111" &
        "010000" &
        "010000" &
        "011110" &
        "010000" &
        "010000" &
        "011111"
    Private char_F = "" &
        "000000" &
        "011111" &
        "010000" &
        "010000" &
        "011100" &
        "010000" &
        "010000" &
        "010000"
    Private char_G = "" &
        "000000" &
        "001110" &
        "010001" &
        "010000" &
        "010000" &
        "010111" &
        "010001" &
        "001111"
    Private char_H = "" &
        "000000" &
        "010001" &
        "010001" &
        "010001" &
        "011111" &
        "010001" &
        "010001" &
        "010001"
    Private char_I = "" &
        "000000" &
        "001110" &
        "000100" &
        "000100" &
        "000100" &
        "000100" &
        "000100" &
        "001110"
    Private char_J = "" &
        "000000" &
        "000111" &
        "000010" &
        "000010" &
        "000010" &
        "000010" &
        "010010" &
        "001110"
    Private char_K = "" &
        "000000" &
        "010001" &
        "010010" &
        "010100" &
        "011000" &
        "010100" &
        "010010" &
        "010001"
    Private char_L = "" &
        "000000" &
        "010000" &
        "010000" &
        "010000" &
        "010000" &
        "010000" &
        "010000" &
        "011111"
    Private char_M = "" &
        "000000" &
        "010001" &
        "011011" &
        "010101" &
        "010101" &
        "010001" &
        "010001" &
        "010001"
    Private char_N = "" &
        "000000" &
        "010001" &
        "011001" &
        "010101" &
        "010101" &
        "010101" &
        "010011" &
        "010001"
    Private char_O = "" &
        "000000" &
        "000100" &
        "001010" &
        "010001" &
        "010001" &
        "010001" &
        "001010" &
        "000100"
    Private char_P = "" &
        "000000" &
        "011110" &
        "010001" &
        "010001" &
        "011110" &
        "010000" &
        "010000" &
        "010000"
    Private char_Q = "" &
        "000000" &
        "001110" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "010011" &
        "001111"
    Private char_R = "" &
        "000000" &
        "011110" &
        "010001" &
        "010001" &
        "011110" &
        "010100" &
        "010010" &
        "010001"
    Private char_S = "" &
        "000000" &
        "001111" &
        "010000" &
        "010000" &
        "001110" &
        "000001" &
        "000001" &
        "011110"
    Private char_T = "" &
        "000000" &
        "011111" &
        "000100" &
        "000100" &
        "000100" &
        "000100" &
        "000100" &
        "000100"
    Private char_U = "" &
        "000000" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "001110"
    Private char_V = "" &
        "000000" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "001010" &
        "001010" &
        "000100"
    Private char_W = "" &
        "000000" &
        "010001" &
        "010001" &
        "010001" &
        "010001" &
        "010101" &
        "010101" &
        "001010"
    Private char_X = "" &
        "000000" &
        "010001" &
        "010001" &
        "001010" &
        "000100" &
        "001010" &
        "010001" &
        "010001"
    Private char_Y = "" &
        "000000" &
        "010001" &
        "010001" &
        "001010" &
        "001010" &
        "000100" &
        "000100" &
        "000100"
    Private char_Z = "" &
        "000000" &
        "011111" &
        "000001" &
        "000010" &
        "000100" &
        "001000" &
        "010010" &
        "011111"
    Private char_SQ1 = "" &
    "000000" &
        "000000" &
        "000000" &
        "001100" &
        "001100" &
        "000000" &
        "000000" &
        "000000"
    Private char_CI1 = "" &
    "000000" &
        "000000" &
        "000000" &
        "001100" &
        "001100" &
        "000000" &
        "000000" &
        "000000"
    Private char_SQ2 = "" &
    "000000" &
        "000000" &
        "011110" &
        "010010" &
        "010010" &
        "011110" &
        "000000" &
        "000000"
    Private char_CI2 = "" &
    "000000" &
        "000000" &
        "001100" &
        "010010" &
        "010010" &
        "001100" &
        "000000" &
        "000000"
    Private char_SQ3 = "" &
    "000000" &
        "111111" &
        "100001" &
        "100001" &
        "100001" &
        "100001" &
        "111111" &
        "000000"
    Private char_CI3 = "" &
        "000000" &
        "001100" &
        "010010" &
        "100001" &
        "100001" &
        "010010" &
        "001100" &
        "000000"
    Private char_DEATH = "" &
        "000000" &
        "010100" &
        "101010" &
        "101010" &
        "111110" &
        "010100" &
        "010100" &
        "001000"
    Public anim_CornerPlus = "" &
        "000000" & "000000" & "000000" & "000000" & "000000" & "000000" &
        "100001" & "010010" & "001100" & "000000" & "000000" & "111111" &
        "000000" & "110011" & "001100" & "001100" & "011110" & "100001" &
        "000000" & "000000" & "111111" & "011110" & "010010" & "100001" &
        "000000" & "000000" & "111111" & "011110" & "010010" & "100001" &
        "000000" & "110011" & "001100" & "001100" & "011110" & "100001" &
        "100001" & "010010" & "001100" & "000000" & "000000" & "111111" &
        "000000" & "000000" & "000000" & "000000" & "000000" & "000000" &
        ""
    Public anim_Rect = "" &
       "000000" & "000000" & "000000" & "111111" & "111111" & "111111" & "111111" & "111111" &
       "000000" & "000000" & "011110" & "100001" & "111111" & "100001" & "111111" & "111111" &
       "000000" & "001100" & "010010" & "100001" & "110011" & "101101" & "110011" & "111111" &
       "000000" & "001100" & "010010" & "100001" & "110011" & "101101" & "110011" & "111111" &
       "000000" & "001100" & "010010" & "100001" & "110011" & "101101" & "110011" & "111111" &
       "000000" & "001100" & "010010" & "100001" & "110011" & "101101" & "110011" & "111111" &
       "000000" & "000000" & "011110" & "100001" & "111111" & "100001" & "111111" & "111111" &
       "000000" & "000000" & "000000" & "111111" & "111111" & "111111" & "111111" & "111111" &
       ""
    Private anim_PlusCross = "" &
        "000000" & "000000" & "000000" & "000000" & "000000" & "000000" &
        "100001" & "010010" & "001100" & "000000" & "000000" & "111111" &
        "000000" & "110011" & "001100" & "001100" & "011110" & "100001" &
        "000000" & "000000" & "111111" & "011110" & "010010" & "100001" &
        "000000" & "000000" & "111111" & "011110" & "010010" & "100001" &
        "000000" & "110011" & "001100" & "001100" & "011110" & "100001" &
        "100001" & "010010" & "001100" & "000000" & "000000" & "111111" &
        "000000" & "000000" & "000000" & "000000" & "000000" & "000000" &
        ""
    

    Public Function countCharAnim(ByVal anim As String) As Integer
        Return Len(anim) / 48
    End Function
    Public Function getAnimChar(ByVal anim As String, ByVal cnt As Integer, ByVal pos As String) As String
        If pos < Len(anim) / 48 Then
            Dim s As String = ""
            For j = 0 To 7
                s = s & Mid(anim, 6 * (pos + cnt * j) + 1, 6)
            Next
            Return s
        End If
        Return Space(48)
    End Function
    Public Function getPicChar(ByVal p1 As String) As String
        If p1 = "0" Then
            Return char_0
        ElseIf p1 = "1" Then
            Return char_1
        ElseIf p1 = "2" Then
            Return char_2
        ElseIf p1 = "3" Then
            Return char_3
        ElseIf p1 = "&" Then
            Return char_AMP
        ElseIf p1 = "(" Then
            Return char_BO
        ElseIf p1 = ")" Then
            Return char_BC
        ElseIf p1 = "AU" Then
            Return char_ARRU
        ElseIf p1 = "AD" Then
            Return char_ARRD
        ElseIf p1 = "A" Then
            Return char_A
        ElseIf p1 = "B" Then
            Return char_B
        ElseIf p1 = "C" Then
            Return char_C
        ElseIf p1 = "D" Then
            Return char_D
        ElseIf p1 = "E" Then
            Return char_E
        ElseIf p1 = "F" Then
            Return char_F
        ElseIf p1 = "G" Then
            Return char_G
        ElseIf p1 = "H" Then
            Return char_H
        ElseIf p1 = "I" Then
            Return char_I
        ElseIf p1 = "J" Then
            Return char_J
        ElseIf p1 = "K" Then
            Return char_K
        ElseIf p1 = "L" Then
            Return char_L
        ElseIf p1 = "M" Then
            Return char_M
        ElseIf p1 = "N" Then
            Return char_N
        ElseIf p1 = "O" Then
            Return char_O
        ElseIf p1 = "P" Then
            Return char_P
        ElseIf p1 = "Q" Then
            Return char_Q
        ElseIf p1 = "R" Then
            Return char_R
        ElseIf p1 = "S" Then
            Return char_S
        ElseIf p1 = "T" Then
            Return char_T
        ElseIf p1 = "U" Then
            Return char_U
        ElseIf p1 = "V" Then
            Return char_V
        ElseIf p1 = "W" Then
            Return char_W
        ElseIf p1 = "X" Then
            Return char_X
        ElseIf p1 = "Y" Then
            Return char_Y
        ElseIf p1 = "Z" Then
            Return char_Z
        ElseIf p1 = "HEART" Then
            Return char_HEART
        ElseIf p1 = "HEART2" Then
            Return char_HEART2
        ElseIf p1 = "-" Then
            Return char_MINUS
        ElseIf p1 = "," Then
            Return char_COMA
        ElseIf p1 = "." Then
            Return char_DOT
        ElseIf p1 = "!" Then
            Return char_EXCL
        ElseIf p1 = "'" Then
            Return char_APOS
        ElseIf p1 = ":)" Then
            Return char_SMILEY
        ElseIf p1 = ":D" Then
            Return char_SMILEYD
        ElseIf p1 = "SQ1" Then
            Return char_SQ1
        ElseIf p1 = "SQ2" Then
            Return char_SQ2
        ElseIf p1 = "SQ3" Then
            Return char_SQ3
        ElseIf p1 = "CI1" Then
            Return char_CI1
        ElseIf p1 = "CI2" Then
            Return char_CI2
        ElseIf p1 = "CI3" Then
            Return char_CI3

        End If
        Return Space(48)
    End Function
End Module
