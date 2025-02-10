' Dla ThisWorkBook

Private Sub Workbook_Open()
    Rows("1:8").RowHeight = 66
    Columns("A:H").ColumnWidth = 11
    Columns("I").ColumnWidth = 3
    Columns("J").ColumnWidth = 20
    Range("A1:H8").Borders.LineStyle = xlContinuous
    
    button_maker ActiveSheet.Range("J2"), "Init", "Nowa Gra"
    button_maker ActiveSheet.Range("J4"), "DoMove", "Wykonaj Ruch"
    Init
End Sub


Sub button_maker(r As Range, Action As String, Text As String)

        ActiveSheet.Buttons.Add(r.Left, r.Top, r.Width, r.Height).Select
        With Selection
            .OnAction = Action
            .Characters.Text = Text
        End With
    r.Select
End Sub

' Koniec ThisWorkBook


Option Explicit

Dim Board(1 To 8, 1 To 8) As Byte
' 1 - pion biały
' 2 - dama biały
' 3 - pion czarny
' 4 - dama czarna

Dim lastRound As Byte
' 0 - biale
' 1 - czarne
Dim LastCapture As Boolean
Dim NextCaptures As Boolean
Dim NextCapturesX As Integer
Dim NextCapturesY As Integer

Sub Init()
    Dim i As Integer, j As Integer

    lastRound = 0
    NextCaptures = False
    LastCapture = False
    
    ' Wyczyść planszę
    For i = 1 To 8
        For j = 1 To 8
            If (i + j) Mod 2 = 0 Then
                Board(i, j) = 0
            ElseIf i <= 3 Then
                Board(i, j) = 1 ' Białe pionki
            ElseIf i >= 6 Then
                Board(i, j) = 3 ' Czarne pionki
            Else:
                Board(i, j) = 0
            End If
        Next j
    Next i
    
    ' Wyświetl planszę
    DrawMap
End Sub

Sub DrawMap()
    Dim i As Integer, j As Integer
    
    Application.ScreenUpdating = False
    ' Tworzenie wizualnej planszy w arkuszu
    For i = 1 To 8
        For j = 1 To 8
            With Cells(i, j)
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
                .Font.Size = 58
                
                ' Kolorowanie pól
                If (i + j) Mod 2 = 1 Then ' Czarne pola
                    .Interior.Color = RGB(0, 0, 0)
                    .Value = PickEmoji(Board(i, j))
                    
                    If Board(i, j) = 0 Then
                        .Font.Color = RGB(0, 0, 0)
                    ElseIf Board(i, j) >= 3 Then
                        .Font.Color = RGB(255, 255, 255)
                    Else
                        .Font.Color = RGB(150, 75, 0)
                        
                    End If
                Else ' Białe pola
                    .Value = ""
                    .Interior.Color = RGB(255, 255, 255) ' Białe pola
                    .Font.Color = RGB(255, 255, 255)
                End If
            End With
        Next j
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

Sub DoMove()
    Dim startX As Integer, startY As Integer, endX As Integer, endY As Integer
    Dim res As Byte
    Dim Sel_Start As Range
    Dim Sel_End As Range
    ' Pobranie ruchu od użytkownika
    If NextCaptures = True Then
        startX = NextCapturesX
        startY = NextCapturesY
        NextCaptures = False
        
    Else
        Set Sel_Start = Application.InputBox(":", "Wybierz figure do przesunięcia", Type:=8)
        startX = Sel_Start.Row
        startY = Sel_Start.Column
    End If
    
    Set Sel_End = Application.InputBox(":", "Wybierz pole, na które chcesz przesunąć figure", Type:=8)
    endX = Sel_End.Row
    endY = Sel_End.Column
    
    ' Sprawdzenie poprawności ruchu
    res = IsValidMove(startX, startY, endX, endY)
    
    If res = 0 Or res = 4 Then
        Kill startX, startY, endX, endY
        Board(endX, endY) = Board(startX, startY) + IIf(res = 4, 1, 0)
        Board(startX, startY) = 0
        DrawMap
        
        ' Sprawdzenie, czy istnieje możliwość kolejnego bicia
        If (LastCapture = True And CheckForFurtherCaptures(endX, endY)) Then
            If MsgBox("Czy chcesz wykonać kolejne bicie?", vbYesNo) = vbYes Then
                NextCaptures = True
                NextCapturesX = endX
                NextCapturesY = endY
                DoMove
            Else
                lastRound = 1 - lastRound
            End If

        Else
            lastRound = 1 - lastRound
        End If
       
    ElseIf res = 1 Then
        MsgBox "Ruch przeciwnika!", vbExclamation
    ElseIf res = 2 Then
        MsgBox "Nieprawidłowy ruch!", vbExclamation
    ElseIf res = 3 Then
        MsgBox "Ruch wykracza poza plansze!", vbExclamation
    End If
    
    LastCapture = False
End Sub

Function IsValidMove(startX As Integer, startY As Integer, endX As Integer, endY As Integer) As Byte
    ' 0 - Poprawny ruch
    ' 1 - Ruch Przeciwnika
    ' 2 - Nieprawidłowy ruch
    ' 3 - Ruch wykracza poza planszę
    ' 4 - Awans pionka na dame

    ' Sprawdzenie, czy współrzędne są w zakresie 1-8
    If startX < 1 Or startX > 8 Or startY < 1 Or startY > 8 Or _
       endX < 1 Or endX > 8 Or endY < 1 Or endY > 8 Then
        IsValidMove = 3
        Exit Function
    End If

    Dim piece As Byte: piece = Board(startX, startY)
    Dim Target As Byte: Target = Board(endX, endY)
    
    ' Sprawdzenie, czy na polu startowym znajduje się pionek i czy pole docelowe jest puste
    If piece = 0 Or Target <> 0 Then
        IsValidMove = 2
        Debug.Print "1) " & piece & ", " & Target
        Exit Function
    End If

    ' Ruchy muszą być ukośne (pionki) (o 1 w pionie i 1 w poziomie)
    If (piece = 1 Or piece = 3) Then
        If Abs(endX - startX) = 2 And Abs(endY - startY) = 2 Then
        
            ' Możliwość bicia
            
            Dim midX As Integer, midY As Integer
            midX = (startX + endX) \ 2
            midY = (startY + endY) \ 2

            If Not ((piece <= 2 And (Board(midX, midY)) >= 3) Or (piece >= 3 And (Board(midX, midY)) <= 2)) Then
                IsValidMove = 2
                Exit Function
            ElseIf Not (Board(midX, midY) <> 0 And (Board(endX, endY)) = 0) Then
                IsValidMove = 2
                Exit Function
            End If
            
            Debug.Print "13) " & piece & ", " & startX & ", " & startY & ", " & endX & ", " & endY
            
        ElseIf Not (Abs(endX - startX) = 1 And Abs(endY - startY) = 1) Then

            IsValidMove = 2
            Exit Function
        End If
    End If

    ' Ruchy muszą być ukośne (damki) (o X w pionie i X w poziomie)
    If (piece = 2 Or piece = 4) Then
        ' Sprawdzenie, czy ruch jest ukośny
        If Abs(endX - startX) <> Abs(endY - startY) Or Abs(endX - startX) < 1 Then
            IsValidMove = 2
            Debug.Print "3) " & piece & ", " & startX & ", " & startY & ", " & endX & ", " & endY
            Exit Function
        End If
    
        ' Sprawdzenie, czy na drodze nie ma innych figur
        Dim stepX As Integer, stepY As Integer
        stepX = Sgn(endX - startX)
        stepY = Sgn(endY - startY)
    
        Dim x As Integer, y As Integer
        x = startX + stepX
        y = startY + stepY
    
        Do While x <> endX And y <> endY
            If Board(x, y) <> 0 And Board(x + stepX, y + stepY) <> 0 Then ' Jeśli na ścieżce znajduje się figura, ruch jest nieprawidłowy
                IsValidMove = 2
                Exit Function
            End If
            x = x + stepX
            y = y + stepY
        Loop
    End If

    ' Sprawdzenie kierunku ruchu (pionki mogą poruszać się tylko do przodu)
    If (piece = 1 And endX <= startX) Or (piece = 3 And endX >= startX) Then
        IsValidMove = 2
        Debug.Print "4) " & piece & ", " & startX & ", " & startY & ", " & endX & ", " & endY
        Exit Function
    End If
    
    ' Sprawdzenie tury gracza
    If Not ((piece <= 2 And lastRound = 1) Or (piece >= 3 And lastRound = 0)) Then
        IsValidMove = 1 ' Ruch przeciwnika
        Exit Function
    End If
    
    
    ' Sprawdzenie, czy pionek dotarł na koniec planszy i staje się damką
    If (piece = 1 And endX = 8) Or (piece = 3 And endX = 1) Then
        IsValidMove = 4
        Exit Function
    End If
    
End Function

Function Kill(startX As Integer, startY As Integer, endX As Integer, endY As Integer)
    Dim stepX As Integer, stepY As Integer, x As Integer, y As Integer
    stepX = Sgn(endX - startX)
    stepY = Sgn(endY - startY)
    
    x = startX + stepX
    y = startY + stepY
    
    Do While x <> endX Or y <> endY
        If Board(x, y) <> 0 Then
            LastCapture = True
        End If
        Board(x, y) = 0
        x = x + stepX
        y = y + stepY
    Loop
End Function

Function PickEmoji(Number As Byte) As String
    Select Case Number
        Case 1
            PickEmoji = ChrW(&H26C0)
        Case 2
            PickEmoji = ChrW(&H26C1)
        Case 3
            PickEmoji = ChrW(&H26C2)
        Case 4
            PickEmoji = ChrW(&H26C3)
        Case Else
            PickEmoji = ""
    End Select
End Function

Function CheckForFurtherCaptures(x As Integer, y As Integer) As Boolean
    Dim piece As Byte: piece = Board(x, y)
    Dim i As Integer, j As Integer
    Dim dx As Integer, dy As Integer
    Dim stepX As Integer, stepY As Integer
    Dim midX As Integer, midY As Integer
    Dim hasCapture As Boolean: hasCapture = False
    
    ' Sprawdzenie możliwych kierunków ruchu
    For i = -1 To 1 Step 2
        For j = -1 To 1 Step 2
            Debug.Print "Sprawdzanie_ruchu) " & x & ", " & y & ", " & dx & ", " & dy
            ' Dla pionków sprawdzamy tylko o 2 pola
            If piece = 1 Or piece = 3 Then
            Debug.Print "Sprawdzanie_ruchu2) " & x & ", " & y & ", " & dx & ", " & dy
                dx = x + 2 * i
                dy = y + 2 * j
                
                ' Sprawdzenie, czy ruch jest w zakresie planszy
                If dx >= 1 And dx <= 8 And dy >= 1 And dy <= 8 Then
                    Debug.Print "Sprawdzanie_ruchu3) " & x & ", " & y & ", " & dx & ", " & dy
                    ' Sprawdzenie, czy ruch jest poprawny (czy jest bicie)
                    If IsValidMove(x, y, dx, dy) = 0 Or IsValidMove(x, y, dx, dy) = 4 Then
                        Debug.Print "Sprawdzanie_ruchu4) " & x & ", " & y & ", " & dx & ", " & dy
                        hasCapture = True
                        Exit For
                    End If
                End If
            ' Dla damki sprawdzamy wszystkie możliwe ruchy po przekątnej
            ElseIf piece = 2 Or piece = 4 Then
                stepX = i
                stepY = j
                dx = x + stepX
                dy = y + stepY
                
                ' Przeszukaj całą przekątną
                Do While dx >= 1 And dx <= 8 And dy >= 1 And dy <= 8
                    ' Jeśli napotkamy pionek przeciwnika, sprawdź, czy za nim jest puste pole
                    If Board(dx, dy) <> 0 And Board(dx, dy) <> piece Then
                        midX = dx
                        midY = dy
                        dx = dx + stepX
                        dy = dy + stepY
                        
                        ' Sprawdź, czy pole za pionkiem przeciwnika jest puste
                        If dx >= 1 And dx <= 8 And dy >= 1 And dy <= 8 Then
                            If Board(dx, dy) = 0 Then
                                If IsValidMove(x, y, dx, dy) = 0 Then
                                    hasCapture = True
                                    Exit For
                                End If
                            End If
                        End If
                        Exit Do
                    ElseIf Board(dx, dy) = piece Then
                        ' Nie można przeskoczyć własnego pionka
                        Exit Do
                    End If
                    
                    dx = dx + stepX
                    dy = dy + stepY
                Loop
            End If
        Next j
        If hasCapture Then Exit For
    Next i
    
    CheckForFurtherCaptures = hasCapture
End Function
