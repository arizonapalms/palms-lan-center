Attribute VB_Name = "modCups"
' Palms LAN Center - Tournaments: keeps the brackets the server sends and draws them.
' Also runs as VBScript, for testing without Visual Basic.
Option Explicit

Public Const CUP_MAX As Long = 8           ' tournaments at once (the server's limit too)
Public Const CUP_SLOTS As Long = 64        ' players in one bracket
Private Const CUP_M As Long = 64           ' room for one tournament's matches (63 at most)

' Bracket sizes, in pixels
Private Const CB_M As Long = 10            ' margin
Private Const CB_HEAD As Long = 24         ' round titles along the top
Private Const CB_W As Long = 150           ' a match
Private Const CB_H As Long = 36
Private Const CB_GAP As Long = 30          ' between rounds, for the lines
Private Const CB_PITCH As Long = 44        ' first-round matches, top to top
Private Const CB_CHAMP_W As Long = 140
Private Const CB_LINE As Long = &H808080
Private Const CB_ROWLINE As Long = &HC8D0D4

' ---- the tournaments, in the order the server lists them
Public gCupCount As Long
Public gCupId(0 To 7) As String
Public gCupState(0 To 7) As String         ' signup, drawn, running or done
Public gCupMax(0 To 7) As Long             ' most players, 0 = no limit
Public gCupTime(0 To 7) As String          ' when it starts, as the host typed it
Public gCupGame(0 To 7) As String
Public gCupChamp(0 To 7) As String         ' the winner, once there is one
Public gCupName(0 To 7) As String
Public gCupNotes(0 To 7) As String         ' the host's details, ~ between lines
Public gCupPlayers(0 To 7) As String       ' sign-ups (by name), or the draw (- for a bye), separated by commas
Public gCupRounds(0 To 7) As Long          ' 0 until the bracket is drawn

' ---- their matches: tournament k's are from k * 64, round by round
Private mA(0 To 511) As String             ' a player, - for a bye, blank while not known yet
Private mB(0 To 511) As String
Private mWin(0 To 511) As String           ' a, b or blank
Private mScore(0 To 511) As String
Private mRepBy(0 To 511) As String         ' a result a player reported, waiting for the host
Private mRepWin(0 To 511) As String
Private mRepScore(0 To 511) As String

' ================================================================ from the server

Public Sub CupClear()
    gCupCount = 0
End Sub

' Takes one CUP... line from the server. Returns the tournament it was about, or -1.
' CUPCLEAR|n   CUP|id|state|max|time|game|champion|name   CUPNOTE|id|notes   CUPP|id|players
' CUPM|id|round|match|a|b|win|score|reported by|reported win|reported score   CUPEND|id|rounds   CUPDEL|id
Public Function CupLine(ByVal lineText As String) As Long
    Dim f As Variant, cmd As String, k As Long, n As Long, i As Long, mi As Long
    CupLine = -1
    f = Split(lineText, "|")
    cmd = UCase(f(0))
    If cmd = "CUPCLEAR" Then
        gCupCount = 0
        Exit Function
    End If
    If UBound(f) < 1 Then Exit Function
    k = CupIndex(f(1))
    Select Case cmd
    Case "CUP"
        If k < 0 Then
            If gCupCount >= CUP_MAX Then Exit Function
            k = gCupCount
            gCupCount = gCupCount + 1
            gCupId(k) = f(1)
        End If
        gCupState(k) = Part(f, 2)
        gCupMax(k) = Val(Part(f, 3))
        gCupTime(k) = Part(f, 4)
        gCupGame(k) = Part(f, 5)
        gCupChamp(k) = Part(f, 6)
        gCupName(k) = Rest(f, 7)
        gCupNotes(k) = ""
        gCupPlayers(k) = ""
        gCupRounds(k) = 0
        For i = k * CUP_M To k * CUP_M + CUP_M - 1
            mA(i) = ""
            mB(i) = ""
            mWin(i) = ""
            mScore(i) = ""
            mRepBy(i) = ""
            mRepWin(i) = ""
            mRepScore(i) = ""
        Next
    Case "CUPNOTE"
        If k >= 0 Then gCupNotes(k) = Rest(f, 2)
    Case "CUPP"
        If k >= 0 Then gCupPlayers(k) = Part(f, 2)
    Case "CUPM"
        If k >= 0 Then
            mi = MatchIdx(k, Val(Part(f, 2)), Val(Part(f, 3)), SlotCount(k))
            If mi >= 0 Then
                mA(mi) = Part(f, 4)
                mB(mi) = Part(f, 5)
                mWin(mi) = Part(f, 6)
                mScore(mi) = Part(f, 7)
                mRepBy(mi) = Part(f, 8)
                mRepWin(mi) = Part(f, 9)
                mRepScore(mi) = Part(f, 10)
            End If
        End If
    Case "CUPEND"
        If k >= 0 Then
            n = Val(Part(f, 2))
            If n < 0 Or n > 6 Then n = 0
            gCupRounds(k) = n
        End If
    Case "CUPDEL"
        If k >= 0 Then RemoveCup k
        k = -1
    Case Else
        k = -1
    End Select
    CupLine = k
End Function

Private Sub RemoveCup(ByVal k As Long)
    Dim j As Long, i As Long
    For j = k To gCupCount - 2
        gCupId(j) = gCupId(j + 1)
        gCupState(j) = gCupState(j + 1)
        gCupMax(j) = gCupMax(j + 1)
        gCupTime(j) = gCupTime(j + 1)
        gCupGame(j) = gCupGame(j + 1)
        gCupChamp(j) = gCupChamp(j + 1)
        gCupName(j) = gCupName(j + 1)
        gCupNotes(j) = gCupNotes(j + 1)
        gCupPlayers(j) = gCupPlayers(j + 1)
        gCupRounds(j) = gCupRounds(j + 1)
        For i = 0 To CUP_M - 1
            mA(j * CUP_M + i) = mA((j + 1) * CUP_M + i)
            mB(j * CUP_M + i) = mB((j + 1) * CUP_M + i)
            mWin(j * CUP_M + i) = mWin((j + 1) * CUP_M + i)
            mScore(j * CUP_M + i) = mScore((j + 1) * CUP_M + i)
            mRepBy(j * CUP_M + i) = mRepBy((j + 1) * CUP_M + i)
            mRepWin(j * CUP_M + i) = mRepWin((j + 1) * CUP_M + i)
            mRepScore(j * CUP_M + i) = mRepScore((j + 1) * CUP_M + i)
        Next
    Next
    gCupCount = gCupCount - 1
End Sub

Private Function Part(f As Variant, ByVal i As Long) As String
    Part = ""
    If i <= UBound(f) Then Part = f(i)
End Function

' Field i and everything after it, so text can contain "|"
Private Function Rest(f As Variant, ByVal i As Long) As String
    Dim k As Long, s As String
    Rest = ""
    If i > UBound(f) Then Exit Function
    s = f(i)
    For k = i + 1 To UBound(f)
        s = s & "|" & f(k)
    Next
    Rest = s
End Function

' ================================================================ questions about them

Public Function CupIndex(ByVal id As String) As Long
    Dim k As Long
    CupIndex = -1
    For k = 0 To gCupCount - 1
        If gCupId(k) = id Then
            CupIndex = k
            Exit Function
        End If
    Next
End Function

' How many places in the bracket (a power of 2), from the draw
Private Function SlotCount(ByVal k As Long) As Long
    SlotCount = 0
    If gCupState(k) = "signup" Or Len(gCupPlayers(k)) = 0 Then Exit Function
    SlotCount = UBound(Split(gCupPlayers(k), ",")) + 1
End Function

' Where match i of round r is kept: the rounds one after another
Private Function MatchIdx(ByVal k As Long, ByVal r As Long, ByVal i As Long, ByVal size As Long) As Long
    Dim q As Long, start As Long, n As Long
    MatchIdx = -1
    If size < 2 Then Exit Function
    n = size \ 2
    start = 0
    For q = 0 To r - 1
        start = start + n
        n = n \ 2
    Next
    If n < 1 Or i < 0 Or i >= n Or r < 0 Then Exit Function
    MatchIdx = k * CUP_M + start + i
End Function

Private Function MI(ByVal k As Long, ByVal r As Long, ByVal i As Long) As Long
    MI = MatchIdx(k, r, i, SlotCount(k))
End Function

' Matches in round r
Public Function CupMatches(ByVal k As Long, ByVal r As Long) As Long
    Dim n As Long, q As Long
    n = SlotCount(k) \ 2
    For q = 1 To r
        n = n \ 2
    Next
    CupMatches = n
End Function

' A match, for frmMain: what is a, b, win, score, repby, repwin or repscore
Public Function CupM(ByVal k As Long, ByVal r As Long, ByVal i As Long, ByVal what As String) As String
    Dim x As Long
    CupM = ""
    x = MI(k, r, i)
    If x < 0 Then Exit Function
    Select Case what
    Case "a": CupM = mA(x)
    Case "b": CupM = mB(x)
    Case "win": CupM = mWin(x)
    Case "score": CupM = mScore(x)
    Case "repby": CupM = mRepBy(x)
    Case "repwin": CupM = mRepWin(x)
    Case "repscore": CupM = mRepScore(x)
    End Select
End Function

' "Final", "Semi-final", "Quarter-final" or "Round 1"
Public Function CupRoundWord(ByVal k As Long, ByVal r As Long) As String
    Select Case gCupRounds(k) - r
    Case 1: CupRoundWord = "Final"
    Case 2: CupRoundWord = "Semi-final"
    Case 3: CupRoundWord = "Quarter-final"
    Case Else: CupRoundWord = "Round " & (r + 1)
    End Select
End Function

' The heading over a round: "Final", "Semi-finals"...
Public Function CupRoundTitle(ByVal k As Long, ByVal r As Long) As String
    Select Case gCupRounds(k) - r
    Case 1: CupRoundTitle = "Final"
    Case 2: CupRoundTitle = "Semi-finals"
    Case 3: CupRoundTitle = "Quarter-finals"
    Case Else: CupRoundTitle = "Round " & (r + 1)
    End Select
End Function

' "the Semi-final", or "Round 1"
Private Function InRound(ByVal k As Long, ByVal r As Long) As String
    If gCupRounds(k) - r <= 3 Then
        InRound = "the " & CupRoundWord(k, r)
    Else
        InRound = CupRoundWord(k, r)
    End If
End Function

' A real player (not a bye, and not still to be decided)
Private Function IsPlayer(ByVal s As String) As Boolean
    IsPlayer = (Len(s) > 0 And s <> "-")
End Function

Private Function Same(ByVal a As String, ByVal b As String) As Boolean
    Same = (StrComp(a, b, vbTextCompare) = 0)
End Function

Public Function CupPlayerCount(ByVal k As Long) As Long
    Dim p As Variant, i As Long, n As Long
    n = 0
    If Len(gCupPlayers(k)) > 0 Then
        p = Split(gCupPlayers(k), ",")
        For i = 0 To UBound(p)
            If IsPlayer(p(i)) Then n = n + 1
        Next
    End If
    CupPlayerCount = n
End Function

Public Function CupIsIn(ByVal k As Long, ByVal nick As String) As Boolean
    CupIsIn = (Len(nick) > 0 And InStr(1, "," & gCupPlayers(k) & ",", "," & nick & ",", vbTextCompare) > 0)
End Function

' Both players known and no result yet
Public Function CupReady(ByVal k As Long, ByVal r As Long, ByVal i As Long) As Boolean
    Dim x As Long
    CupReady = False
    x = MI(k, r, i)
    If x < 0 Then Exit Function
    CupReady = IsPlayer(mA(x)) And IsPlayer(mB(x)) And Len(mWin(x)) = 0
End Function

' Results reported by players, waiting for the host
Public Function CupPending(ByVal k As Long) As Long
    Dim r As Long, i As Long, x As Long, n As Long
    n = 0
    For r = 0 To gCupRounds(k) - 1
        For i = 0 To CupMatches(k, r) - 1
            x = MI(k, r, i)
            If Len(mRepBy(x)) > 0 And Len(mWin(x)) = 0 Then n = n + 1
        Next
    Next
    CupPending = n
End Function

' The first match waiting for the host, as round * 100 + match, or -1
Public Function CupFirstPending(ByVal k As Long) As Long
    Dim r As Long, i As Long, x As Long
    CupFirstPending = -1
    For r = 0 To gCupRounds(k) - 1
        For i = 0 To CupMatches(k, r) - 1
            x = MI(k, r, i)
            If Len(mRepBy(x)) > 0 And Len(mWin(x)) = 0 Then
                CupFirstPending = r * 100 + i
                Exit Function
            End If
        Next
    Next
End Function

' The match a player is still to play (ready, or waiting for an opponent), as round * 100 + match, or -1
Public Function CupMyMatch(ByVal k As Long, ByVal nick As String) As Long
    Dim r As Long, i As Long, x As Long
    CupMyMatch = -1
    If Len(nick) = 0 Or gCupState(k) = "signup" Then Exit Function
    For r = 0 To gCupRounds(k) - 1
        For i = 0 To CupMatches(k, r) - 1
            x = MI(k, r, i)
            If Len(mWin(x)) = 0 And (Same(mA(x), nick) Or Same(mB(x), nick)) Then
                CupMyMatch = r * 100 + i
                Exit Function
            End If
        Next
    Next
End Function

' A match ready for this player to play and report (and not reported yet)
Public Function CupMyTurn(ByVal k As Long, ByVal nick As String) As Boolean
    Dim m As Long, x As Long
    CupMyTurn = False
    If gCupState(k) <> "running" Then Exit Function
    m = CupMyMatch(k, nick)
    If m < 0 Then Exit Function
    If Not CupReady(k, m \ 100, m Mod 100) Then Exit Function
    x = MI(k, m \ 100, m Mod 100)
    CupMyTurn = (Len(mRepBy(x)) = 0)
End Function

' The other player in a match
Public Function CupOpponent(ByVal k As Long, ByVal r As Long, ByVal i As Long, ByVal nick As String) As String
    Dim x As Long
    CupOpponent = ""
    x = MI(k, r, i)
    If x < 0 Then Exit Function
    If Same(mA(x), nick) Then
        CupOpponent = mB(x)
    ElseIf Same(mB(x), nick) Then
        CupOpponent = mA(x)
    End If
End Function

Private Function SideName(ByVal x As Long, ByVal side As String) As String
    If side = "a" Then
        SideName = mA(x)
    ElseIf side = "b" Then
        SideName = mB(x)
    Else
        SideName = ""
    End If
End Function

Private Function OtherName(ByVal x As Long, ByVal side As String) As String
    If side = "a" Then
        OtherName = mB(x)
    ElseIf side = "b" Then
        OtherName = mA(x)
    Else
        OtherName = ""
    End If
End Function

' "Bob won 16-12."
Private Function WonText(ByVal who As String, ByVal score As String) As String
    If Len(score) > 0 Then
        WonText = who & " won " & score & "."
    Else
        WonText = who & " won."
    End If
End Function

' Two lines about a match for the bar under the bracket: what it is, and how it stands
Public Function CupMatchText(ByVal k As Long, ByVal r As Long, ByVal i As Long, ByVal nick As String, ByVal isHost As Boolean) As String
    Dim x As Long, l1 As String, l2 As String, mine As Boolean
    CupMatchText = ""
    x = MI(k, r, i)
    If x < 0 Then Exit Function
    mine = (Same(mA(x), nick) Or Same(mB(x), nick)) And Len(nick) > 0
    l1 = CupRoundWord(k, r) & ", match " & (i + 1) & ": " & WhoText(mA(x)) & " vs " & WhoText(mB(x))
    If mA(x) = "-" Or mB(x) = "-" Then
        If IsPlayer(mA(x)) Then
            l2 = mA(x) & " has a bye into the next round."
        ElseIf IsPlayer(mB(x)) Then
            l2 = mB(x) & " has a bye into the next round."
        Else
            l2 = "Waiting for the previous round."
        End If
    ElseIf Len(mWin(x)) > 0 Then
        l2 = WonText(SideName(x, mWin(x)), mScore(x))
        If isHost Then l2 = l2 & " To change the result, click Enter Result."
    ElseIf Not IsPlayer(mA(x)) Or Not IsPlayer(mB(x)) Then
        l2 = "Waiting for the previous round."
    ElseIf gCupState(k) <> "running" Then
        l2 = "Not started. The host starts the tournament."
    ElseIf Len(mRepBy(x)) > 0 Then
        l2 = "Reported by " & mRepBy(x) & ": " & WonText(SideName(x, mRepWin(x)), mRepScore(x))
        If isHost Then
            l2 = l2 & " Click Confirm Result."
        Else
            l2 = l2 & " Waiting for the host to confirm."
        End If
    ElseIf mine Then
        l2 = "Your match. When it has been played, click Report Result."
    ElseIf isHost Then
        l2 = "Not played yet. The players report the result, or click Enter Result."
    Else
        l2 = "Not played yet."
    End If
    CupMatchText = l1 & vbLf & l2
End Function

Private Function WhoText(ByVal s As String) As String
    If s = "-" Then
        WhoText = "(bye)"
    ElseIf Len(s) = 0 Then
        WhoText = "?"
    Else
        WhoText = s
    End If
End Function

' Where a player stands, for the bar under the bracket when no match is selected
Public Function CupStatusText(ByVal k As Long, ByVal nick As String) As String
    Dim m As Long, r As Long, i As Long, x As Long, s As String, n As Long
    s = ""
    Select Case gCupState(k)
    Case "signup"
        If CupIsIn(k, nick) Then
            s = "You have signed up. The host draws the bracket when sign-ups close."
        ElseIf gCupMax(k) > 0 And CupPlayerCount(k) >= gCupMax(k) Then
            s = "Sign-ups are full (" & gCupMax(k) & " players)."
        Else
            s = "Sign-ups are open. To enter, click Sign Up."
        End If
    Case "drawn"
        If CupIsIn(k, nick) Then
            s = "The bracket has been drawn. It starts when the host clicks Start."
        Else
            s = "The bracket has been drawn. The tournament has not started yet."
        End If
    Case "running", "done"
        If Len(gCupChamp(k)) > 0 Then
            If Same(gCupChamp(k), nick) Then
                s = "You won the tournament."
            Else
                s = gCupChamp(k) & " won the tournament."
            End If
        ElseIf CupIsIn(k, nick) Then
            m = CupMyMatch(k, nick)
            If m >= 0 Then
                r = m \ 100
                i = m Mod 100
                x = MI(k, r, i)
                If Not CupReady(k, r, i) Then
                    s = "You are through to " & InRound(k, r) & ". Waiting for your opponent."
                ElseIf Len(mRepBy(x)) > 0 Then
                    s = "Your " & CupRoundWord(k, r) & " result has been reported. Waiting for the host to confirm."
                Else
                    s = "Your next match: " & CupRoundWord(k, r) & " against " & CupOpponent(k, r, i, nick) & "."
                End If
            Else
                s = KnockedOut(k, nick)
            End If
        Else
            n = 0
            For r = 0 To gCupRounds(k) - 1
                For i = 0 To CupMatches(k, r) - 1
                    x = MI(k, r, i)
                    If Len(mWin(x)) = 0 Then n = n + 1
                Next
            Next
            If n = 1 Then
                s = "In progress. 1 match to play."
            Else
                s = "In progress. " & n & " matches to play."
            End If
        End If
    End Select
    CupStatusText = s
End Function

Private Function KnockedOut(ByVal k As Long, ByVal nick As String) As String
    Dim r As Long, i As Long, x As Long
    KnockedOut = "You are out of the tournament."
    For r = 0 To gCupRounds(k) - 1
        For i = 0 To CupMatches(k, r) - 1
            x = MI(k, r, i)
            If Len(mWin(x)) > 0 Then
                If Same(OtherName(x, mWin(x)), nick) Then
                    KnockedOut = "You were knocked out in " & InRound(k, r) & " by " & SideName(x, mWin(x)) & "."
                    Exit Function
                End If
            End If
        Next
    Next
End Function

' The game, the start time and the first line of the host's details
Public Function CupInfoText(ByVal k As Long) As String
    Dim s As String, n As String, first As Variant
    s = ""
    If Len(gCupGame(k)) > 0 Then s = gCupGame(k)
    If Len(gCupTime(k)) > 0 Then
        If Len(s) > 0 Then s = s & ", "
        s = s & "starts " & gCupTime(k)
    End If
    first = Split(gCupNotes(k) & "~", "~")
    n = Trim(first(0))
    If Len(n) > 0 Then
        If Len(s) > 0 Then s = s & ". "
        s = s & n
    End If
    If Len(s) > 0 And Right(s, 1) <> "." Then s = s & "."
    CupInfoText = s
End Function

' ================================================================ the bracket's size and layout

Private Function ColX(ByVal r As Long) As Long
    ColX = CB_M + r * (CB_W + CB_GAP)
End Function

' The top of match i in round r: centred on the two matches that feed it
Private Function MatchY(ByVal r As Long, ByVal i As Long) As Long
    Dim span As Long, q As Long
    span = 1
    For q = 1 To r
        span = span * 2
    Next
    MatchY = CB_M + CB_HEAD + CB_PITCH * span * i + (CB_PITCH * (span - 1)) \ 2
End Function

' How big the whole view is, for the scroll bars. The sign-up view fits the width it's given.
Public Function CupViewW(ByVal k As Long, ByVal viewW As Long) As Long
    If k < 0 Or k >= gCupCount Then
        CupViewW = 0
    ElseIf gCupState(k) = "signup" Or gCupRounds(k) = 0 Then
        CupViewW = viewW
    Else
        CupViewW = ColX(gCupRounds(k)) + CB_CHAMP_W + CB_M
    End If
End Function

Public Function CupViewH(p As PictureBox, ByVal k As Long, ByVal viewW As Long) As Long
    If k < 0 Or k >= gCupCount Then
        CupViewH = 0
    ElseIf gCupState(k) = "signup" Or gCupRounds(k) = 0 Then
        CupViewH = SignUpView(p, k, viewW, 0, False, "")
    Else
        CupViewH = MatchY(0, CupMatches(k, 0) - 1) + CB_H + CB_M
    End If
End Function

' The match at a point in the view (scrolled position added), as round * 100 + match, or -1
Public Function CupHit(ByVal k As Long, ByVal x As Long, ByVal y As Long) As Long
    Dim r As Long, i As Long, y0 As Long
    CupHit = -1
    If k < 0 Or k >= gCupCount Then Exit Function
    If gCupState(k) = "signup" Or gCupRounds(k) = 0 Then Exit Function
    For r = 0 To gCupRounds(k) - 1
        If x >= ColX(r) And x < ColX(r) + CB_W Then
            For i = 0 To CupMatches(k, r) - 1
                y0 = MatchY(r, i)
                If y >= y0 And y < y0 + CB_H Then
                    CupHit = r * 100 + i
                    Exit Function
                End If
            Next
        End If
    Next
End Function

' Where a match is in the view, to scroll it into sight
Public Sub CupMatchSpot(ByVal r As Long, ByVal i As Long, x As Long, y As Long, w As Long, h As Long)
    x = ColX(r)
    y = MatchY(r, i)
    w = CB_W
    h = CB_H
End Sub

' ================================================================ drawing

' The whole view: the sign-up list, or the bracket. sx, sy: how far it's scrolled.
' sel: the match picked (round * 100 + match), or -1. nick: who's looking.
Public Sub DrawCup(p As PictureBox, ByVal k As Long, ByVal sx As Long, ByVal sy As Long, ByVal sel As Long, ByVal nick As String)
    Dim h As Long
    p.Cls
    If k < 0 Or k >= gCupCount Then Exit Sub
    If gCupState(k) = "signup" Or gCupRounds(k) = 0 Then
        h = SignUpView(p, k, p.ScaleWidth, sy, True, nick)
    Else
        DrawBracket p, k, sx, sy, sel, nick
    End If
End Sub

' Sign-ups: what it is, the host's details, and who has signed up. Returns how tall it is.
Private Function SignUpView(p As PictureBox, ByVal k As Long, ByVal w As Long, ByVal sy As Long, ByVal draw As Boolean, ByVal nick As String) As Long
    Dim y As Long, s As String, lines As Variant, i As Long, j As Long, wrapped As Variant
    Dim names As Variant, n As Long, cols As Long, colW As Long, x As Long, yy As Long, mine As Boolean
    y = 10 - sy
    If draw Then
        DrawIcon32 p, IC_TROPHY, 10, y
        p.FontSize = 12
        TextAt p, 50, y - 1, FitText(p, gCupName(k), w - 60), C_TITLE, True
        p.FontSize = 8
        s = ""
        If Len(gCupGame(k)) > 0 Then s = gCupGame(k) & ". "
        If Len(gCupTime(k)) > 0 Then s = s & "Starts " & gCupTime(k) & ". "
        If gCupState(k) = "signup" Then
            s = s & "Sign-ups open."
        Else
            s = s & "Sign-ups closed."
        End If
        TextAt p, 51, y + 20, FitText(p, s, w - 60), C_SUB, False
    End If
    y = y + 44
    ' the host's details, a line at a time
    If Len(gCupNotes(k)) > 0 Then
        lines = Split(gCupNotes(k), "~")
        For i = 0 To UBound(lines)
            wrapped = Split(WrapText(p, Trim(lines(i)), w - 30), vbLf)
            For j = 0 To UBound(wrapped)
                If draw Then TextAt p, 12, y, wrapped(j), vbBlack, False
                y = y + 15
            Next
        Next
        y = y + 8
    End If
    ' who has signed up
    n = CupPlayerCount(k)
    If gCupMax(k) > 0 Then
        s = "Players (" & n & " of " & gCupMax(k) & ")"
    Else
        s = "Players (" & n & ")"
    End If
    If draw Then
        TextAt p, 12, y, s, C_TITLE, True
        Box p, 12, y + 16, w - 24, 1, CB_ROWLINE
    End If
    y = y + 22
    If n = 0 Then
        If draw Then TextAt p, 12, y, "No players have signed up yet.", C_GREY, False
        y = y + 18
    Else
        colW = 130
        cols = (w - 24) \ colW
        If cols < 1 Then cols = 1
        names = Split(gCupPlayers(k), ",")
        j = 0
        For i = 0 To UBound(names)
            If IsPlayer(names(i)) Then
                x = 12 + (j Mod cols) * colW
                yy = y + (j \ cols) * 18
                If draw And yy > -20 And yy < p.ScaleHeight Then
                    mine = Same(names(i), nick)
                    If mine Then
                        DrawIcon16 p, IC_OK, x, yy
                        TextAt p, x + 20, yy + 2, FitText(p, names(i), colW - 26), C_BLUE, True
                    Else
                        DrawIcon16 p, IC_USER, x, yy
                        TextAt p, x + 20, yy + 2, FitText(p, names(i), colW - 26), vbBlack, False
                    End If
                End If
                j = j + 1
            End If
        Next
        y = y + ((j + cols - 1) \ cols) * 18
    End If
    SignUpView = y + 10 + sy
End Function

' Words over lines no wider than w, with vbLf between
Private Function WrapText(p As PictureBox, ByVal s As String, ByVal w As Long) As String
    Dim words As Variant, i As Long, cur As String, out As String, try As String
    out = ""
    cur = ""
    words = Split(s, " ")
    For i = 0 To UBound(words)
        If Len(cur) = 0 Then
            try = words(i)
        Else
            try = cur & " " & words(i)
        End If
        If p.TextWidth(try) > w And Len(cur) > 0 Then
            If Len(out) > 0 Then out = out & vbLf
            out = out & cur
            cur = words(i)
        Else
            cur = try
        End If
    Next
    If Len(out) > 0 Then out = out & vbLf
    WrapText = out & cur
End Function

' The bracket: a column per round, joined by lines, and the champion at the end.
' Bold text goes on in a second pass, as switching fonts is slow on old PCs.
Private Sub DrawBracket(p As PictureBox, ByVal k As Long, ByVal sx As Long, ByVal sy As Long, ByVal sel As Long, ByVal nick As String)
    Dim r As Long, i As Long, x As Long, y As Long, n As Long, vw As Long, vh As Long, pass As Long
    Dim x2 As Long, ya As Long, yb As Long, rounds As Long, s As String, fx As Long, fy As Long
    vw = p.ScaleWidth
    vh = p.ScaleHeight
    rounds = gCupRounds(k)
    ' the lines between rounds
    For r = 0 To rounds - 2
        x = ColX(r) - sx + CB_W
        x2 = x + CB_GAP \ 2
        If x < vw And x + CB_GAP > 0 Then
            n = CupMatches(k, r)
            For i = 0 To n - 1 Step 2
                ya = MatchY(r, i) - sy + CB_H \ 2
                yb = MatchY(r, i + 1) - sy + CB_H \ 2
                If yb >= 0 And ya < vh Then
                    Box p, x, ya, CB_GAP \ 2, 1, CB_LINE
                    Box p, x, yb, CB_GAP \ 2, 1, CB_LINE
                    Box p, x2, ya, 1, yb - ya + 1, CB_LINE
                    Box p, x2, (ya + yb) \ 2, CB_GAP - CB_GAP \ 2, 1, CB_LINE
                End If
            Next
        End If
    Next
    ' the final to the champion
    fx = ColX(rounds - 1) - sx + CB_W
    fy = MatchY(rounds - 1, 0) - sy + CB_H \ 2
    Box p, fx, fy, CB_GAP, 1, CB_LINE
    ' the matches: boxes and ordinary text first, then the round titles and the bold text
    For pass = 0 To 1
        If pass = 1 Then
            p.FontBold = True
            For r = 0 To rounds - 1
                x = ColX(r) - sx
                If x + CB_W > 0 And x < vw Then TextKeep p, x + 2, CB_M - sy, CupRoundTitle(k, r), C_TITLE
            Next
            x = ColX(rounds) - sx
            If x < vw Then TextKeep p, x + 2, CB_M - sy, "Champion", C_TITLE
        End If
        For r = 0 To rounds - 1
            x = ColX(r) - sx
            If x + CB_W > 0 And x < vw Then
                n = CupMatches(k, r)
                For i = 0 To n - 1
                    y = MatchY(r, i) - sy
                    If y + CB_H > 0 And y < vh Then DrawMatch p, k, r, i, x, y, (sel = r * 100 + i), nick, pass
                Next
            End If
        Next
        DrawChampion p, k, ColX(rounds) - sx, fy - 22, nick, pass
    Next
    p.FontBold = False
End Sub

Private Sub DrawMatch(p As PictureBox, ByVal k As Long, ByVal r As Long, ByVal i As Long, ByVal x As Long, ByVal y As Long, ByVal picked As Boolean, ByVal nick As String, ByVal pass As Long)
    Dim m As Long, fill As Long, edgeCol As Long, ready As Boolean, bye As Boolean
    m = MI(k, r, i)
    If m < 0 Then Exit Sub
    bye = (mA(m) = "-" Or mB(m) = "-")
    ready = CupReady(k, r, i) And gCupState(k) = "running"
    If pass = 0 Then
        fill = vbWhite
        If ready Then fill = C_TIP
        If bye Then fill = vbButtonFace
        If picked Then
            edgeCol = C_SEL
        ElseIf bye Then
            edgeCol = vb3DShadow
        Else
            edgeCol = CB_LINE
        End If
        Box p, x, y, CB_W, CB_H, fill
        Box p, x + 1, y + CB_H \ 2, CB_W - 2, 1, CB_ROWLINE
        Edge p, x, y, CB_W, CB_H, edgeCol, edgeCol
        If picked Then Edge p, x + 1, y + 1, CB_W - 2, CB_H - 2, edgeCol, edgeCol
    End If
    DrawMatchSide p, m, "a", x, y + 2, nick, pass
    DrawMatchSide p, m, "b", x, y + CB_H \ 2 + 2, nick, pass
End Sub

' One player's row in a match: the name (bold for the winner), and the score or a reported result
Private Sub DrawMatchSide(p As PictureBox, ByVal m As Long, ByVal side As String, ByVal x As Long, ByVal y As Long, ByVal nick As String, ByVal pass As Long)
    Dim nm As String, col As Long, bold As Boolean, s As String, tw As Long, rightX As Long
    nm = SideName(m, side)
    bold = (mWin(m) = side)
    rightX = x + CB_W - 6
    If pass = 0 Then
        ' the score goes next to the winner; a result waiting for the host gets a warning sign
        If bold And Len(mScore(m)) > 0 Then
            tw = p.TextWidth(mScore(m))
            TextAt p, rightX - tw, y + 1, mScore(m), C_SUB, False
            rightX = rightX - tw - 6
        ElseIf Len(mWin(m)) = 0 And Len(mRepBy(m)) > 0 And mRepWin(m) = side Then
            DrawIcon16 p, IC_WARN, rightX - 16, y - 1
            rightX = rightX - 20
            If Len(mRepScore(m)) > 0 Then
                tw = p.TextWidth(mRepScore(m))
                TextAt p, rightX - tw, y + 1, mRepScore(m), C_ORANGE, False
                rightX = rightX - tw - 6
            End If
        End If
    End If
    If bold <> (pass = 1) Then Exit Sub
    If nm = "-" Then
        TextAt p, x + 6, y + 1, "(bye)", C_GREY, False
        Exit Sub
    End If
    If Len(nm) = 0 Then Exit Sub
    If Same(nm, nick) Then
        col = C_BLUE
    ElseIf Len(mWin(m)) > 0 And Not bold Then
        col = C_GREY
    Else
        col = vbBlack
    End If
    If bold Then
        ' (the font is already bold in this pass)
        TextKeep p, x + 6, y + 1, FitText(p, nm, rightX - x - 8), col
    Else
        TextAt p, x + 6, y + 1, FitText(p, nm, rightX - x - 8), col, False
    End If
End Sub

Private Sub DrawChampion(p As PictureBox, ByVal k As Long, ByVal x As Long, ByVal y As Long, ByVal nick As String, ByVal pass As Long)
    Dim col As Long
    If x > p.ScaleWidth Or y > p.ScaleHeight Or y + 44 < 0 Then Exit Sub
    If pass = 0 Then
        If Len(gCupChamp(k)) > 0 Then
            Box p, x, y, CB_CHAMP_W, 44, C_TIP
        Else
            Box p, x, y, CB_CHAMP_W, 44, vbWhite
        End If
        Edge p, x, y, CB_CHAMP_W, 44, CB_LINE, CB_LINE
        DrawIcon32 p, IC_TROPHY, x + 6, y + 6
        If Len(gCupChamp(k)) = 0 Then TextAt p, x + 44, y + 15, "To be decided", C_GREY, False
    ElseIf Len(gCupChamp(k)) > 0 Then
        col = vbBlack
        If Same(gCupChamp(k), nick) Then col = C_BLUE
        TextKeep p, x + 44, y + 15, FitText(p, gCupChamp(k), CB_CHAMP_W - 50), col
    End If
End Sub
