VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim TicksRowA As Integer


'How to create a timer
Sub Sampletimer()
    Application.OnTime Now() + TimeValue("00:00:10"), "TimerMssg"
End Sub




Private Sub worksheet_calculate()
    'This runs every time excel calculates or every time a cell changes which
    'has a formula referenced to it (it will run every time a price changes)

    If RecordOn = 1 Then

    CheckTradesCCCC

    End If

End Sub



Sub CalculateGain()
'This looks for a completed trade (WIN, Timeout, or stopout) and that the order was "Filled", then calculates the gain for the trade
Dim cgxx As Integer
Dim cgyy As Integer

On Error GoTo Errhandler

If MicroWarn <> 10 Then

cgxx = 3
cgyy = 10

marker:

If (Cells(24, cgxx) = "WIN" Or Cells(24, cgxx) = "Timeout" Or Cells(24, cgxx) = "stopout") And Sheets("Conditional Orders").Cells(cgyy, 20) = "Filled" And Cells(28, cgxx) <> "" Then
    If Cells(28, cgxx) = "BUY" Then
    Cells(17, cgxx) = Cells(17, cgxx) + Cells(27, cgxx) * (Sheets("Conditional Orders").Cells(cgyy, 23) - Sheets("Conditional Orders").Cells(cgyy, 86)) / Range("AQ2")
    End If
    If Cells(28, cgxx) = "SELL" Then
    Cells(17, cgxx) = Cells(17, cgxx) + Cells(27, cgxx) * (Sheets("Conditional Orders").Cells(cgyy, 86) - Sheets("Conditional Orders").Cells(cgyy, 23)) / Range("AQ2")
    End If
        With ThisWorkbook.Worksheets("Control Panel")
    .Range(.Cells(25, cgxx), .Cells(29, cgxx)).ClearContents
    End With
End If

cgyy = cgyy + 25
If (Cells(56, cgxx) = "WIN" Or Cells(56, cgxx) = "Timeout" Or Cells(56, cgxx) = "stopout") And Sheets("Conditional Orders").Cells(cgyy, 20) = "Filled" And Cells(60, cgxx) <> "" Then
    If Cells(60, cgxx) = "BUY" Then
    Cells(49, cgxx) = Cells(49, cgxx) + Cells(59, cgxx) * (Sheets("Conditional Orders").Cells(cgyy, 23) - Sheets("Conditional Orders").Cells(cgyy, 86)) / Range("AQ2")
    End If
    If Cells(60, cgxx) = "SELL" Then
    Cells(49, cgxx) = Cells(49, cgxx) + Cells(59, cgxx) * (Sheets("Conditional Orders").Cells(cgyy, 86) - Sheets("Conditional Orders").Cells(cgyy, 23)) / Range("AQ2")
    End If
        With ThisWorkbook.Worksheets("Control Panel")
    .Range(.Cells(57, cgxx), .Cells(61, cgxx)).ClearContents
    End With
End If
cgyy = cgyy - 25




cgxx = cgxx + 2
cgyy = cgyy + 1
If cgxx < 52 Then
GoTo marker:
End If


End If

Errhandler:

End Sub


Sub WinCheck()
'This looks for a trade that exited automatically by hitting your price target.
'It removes the used trade data from the cells and sets up the conditional orders page to be ready for a new trade on that symbol.
'timeout exit and stopout exits are not reset automatically
Dim wcxx As Integer
Dim wcyy As Integer


On Error GoTo Errhandler

If MicroWarn <> 10 Then
wcxx = 3
wcyy = 10

marker:

If Cells(24, wcxx) = "EXIT" And Sheets("Conditional Orders").Cells(wcyy, 20) = "Filled" Then
Cells(24, wcxx) = "WIN"
CalculateGain
Cells(15, wcxx) = "x"
    With ThisWorkbook.Worksheets("Conditional Orders")
    .Range(.Cells(wcyy, 13), .Cells(wcyy, 25)).ClearContents
    End With
End If

wcyy = wcyy + 25
If Cells(56, wcxx) = "EXIT" And Sheets("Conditional Orders").Cells(wcyy, 20) = "Filled" Then
Cells(56, wcxx) = "WIN"
CalculateGain
Cells(47, wcxx) = "x"
    With ThisWorkbook.Worksheets("Conditional Orders")
    .Range(.Cells(wcyy, 13), .Cells(wcyy, 25)).ClearContents
    End With
End If
wcyy = wcyy - 25


wcxx = wcxx + 2
wcyy = wcyy + 1
If wcxx < 52 Then
GoTo marker:
End If



End If

Errhandler:

End Sub


Sub TakeProfits()
'This sends your order to exit your position.
Dim tpxx As Integer
Dim tpyy As Integer

On Error GoTo Errhandler

tpxx = 3
tpyy = 10


MicroWarn = 10

marker:

If Cells(24, tpxx) = "EXITset" Then
    If Cells(28, tpxx) = "BUY" Then
    Sheets("Conditional Orders").Cells(tpyy, 16) = Cells(11, tpxx)
    Else
    Sheets("Conditional Orders").Cells(tpyy, 16) = Cells(14, tpxx)
    End If
    Sheets("Conditional Orders").Select
    Sheets("Conditional Orders").Cells(tpyy, 1).Select
    Sheets("Conditional Orders").placeConditionalOrder
    Sheets("Control Panel").Select
    y = Timer + 2
    Do While Timer < y
    DoEvents
    Loop
    Cells(24, tpxx) = "EXIT"
End If

tpyy = tpyy + 25
If Cells(56, tpxx) = "EXITset" Then
    If Cells(60, tpxx) = "BUY" Then
    Sheets("Conditional Orders").Cells(tpyy, 16) = Cells(43, tpxx)
    Else
    Sheets("Conditional Orders").Cells(tpyy, 16) = Cells(46, tpxx)
    End If
    Sheets("Conditional Orders").Select
    Sheets("Conditional Orders").Cells(tpyy, 1).Select
    Sheets("Conditional Orders").placeConditionalOrder
    Sheets("Control Panel").Select
    y = Timer + 2
    Do While Timer < y
    DoEvents
    Loop
    Cells(56, tpxx) = "EXIT"
End If
tpyy = tpyy - 25


tpxx = tpxx + 2
tpyy = tpyy + 1
If tpxx < 52 Then
GoTo marker:
End If

Errhandler:

MicroWarn = 0

End Sub



Sub SendExitOrder()
'This checks to see if your entry order was "Filled"
'If your entry order was only partially filled, it will cancel that order
'then it sets up the first part of your exit order

On Error GoTo Errhandler

If MicroWarn <> 10 And Range("W2") = True Then

exx = 3
eyy = 10

marker:

'this section looks for a partial order fill and a price that has reached your price target
'it cancels the rest of the unfilled order and allows the process to move to the next step
'if the price has reached your price target without any shares filled, then it cancels the order and resets the conditional orders page
If Cells(12, exx) <> "" And Cells(11, exx) <> "" Then
    If Cells(24, exx) = "BUY" And Cells(20, exx) >= Cells(11, exx) Then
        MicroWarn = 10
        Sheets("Conditional Orders").Select
        Sheets("Conditional Orders").Cells(eyy, 1).Select
        Sheets("Conditional Orders").cancelConditionalOrder
        Sheets("Control Panel").Select
        y = Timer + 2
        Do While Timer < y
        DoEvents
        Loop
            If Sheets("Conditional Orders").Cells(eyy, 21) = 0 Then
                With ThisWorkbook.Worksheets("Conditional Orders")
                .Range(.Cells(eyy, 13), .Cells(eyy, 25)).ClearContents
                End With
                With ThisWorkbook.Worksheets("Control Panel")
                .Range(.Cells(24, exx), .Cells(29, exx)).ClearContents
                End With
                Cells(15, exx) = "x"
            ResetConditionalOrders
            Else
            Sheets("Conditional Orders").Cells(eyy, 20) = "Filled"
            End If
        MicroWarn = 0
    End If
End If
If Cells(13, exx) <> "" And Cells(14, exx) <> "" Then
    If Cells(24, exx) = "SELL" And Cells(20, exx) <= Cells(14, exx) Then
        MicroWarn = 10
        Sheets("Conditional Orders").Select
        Sheets("Conditional Orders").Cells(eyy, 1).Select
        Sheets("Conditional Orders").cancelConditionalOrder
        Sheets("Control Panel").Select
        y = Timer + 2
        Do While Timer < y
        DoEvents
        Loop
            If Sheets("Conditional Orders").Cells(eyy, 21) = 0 Then
                With ThisWorkbook.Worksheets("Conditional Orders")
                .Range(.Cells(eyy, 13), .Cells(eyy, 25)).ClearContents
                End With
                With ThisWorkbook.Worksheets("Control Panel")
                .Range(.Cells(24, exx), .Cells(29, exx)).ClearContents
                End With
                Cells(15, exx) = "x"
            ResetConditionalOrders
            Else
            Sheets("Conditional Orders").Cells(eyy, 20) = "Filled"
            End If
        MicroWarn = 0
    End If
End If

'this section checks for the entry order to be filled, then sets of the first part of the exit order
If (Cells(24, exx) = "BUY" Or Cells(24, exx) = "SELL") Then
    If Sheets("Conditional Orders").Cells(eyy, 20) = "Filled" Then
    MicroWarn = 10
    Cells(24, exx) = "OPENset"
    Sheets("Conditional Orders").Cells(eyy, 86) = Sheets("Conditional Orders").Cells(eyy, 23)
    If Cells(28, exx) = "BUY" Then
    Sheets("Conditional Orders").Cells(eyy, 13) = "SELL"
    Else
    Sheets("Conditional Orders").Cells(eyy, 13) = "BUY"
    End If
    Sheets("Conditional Orders").Cells(eyy, 14) = Sheets("Conditional Orders").Cells(eyy, 21)
    Sheets("Conditional Orders").Cells(eyy, 15) = "LMT"
    With ThisWorkbook.Worksheets("Conditional Orders")
    .Range(.Cells(eyy, 18), .Cells(eyy, 25)).ClearContents
    End With
    Sheets("Conditional Orders").Cells(eyy, 20) = "Submitted"
    MicroWarn = 0
    Cells(24, exx) = "OPEN"
    End If
End If


'---------expansion cells:


eyy = eyy + 25
If Cells(44, exx) <> "" And Cells(43, exx) <> "" Then
    If Cells(56, exx) = "BUY" And Cells(52, exx) >= Cells(43, exx) Then
        MicroWarn = 10
        Sheets("Conditional Orders").Select
        Sheets("Conditional Orders").Cells(eyy, 1).Select
        Sheets("Conditional Orders").cancelConditionalOrder
        Sheets("Control Panel").Select
        y = Timer + 2
        Do While Timer < y
        DoEvents
        Loop
            If Sheets("Conditional Orders").Cells(eyy, 21) = 0 Then
                With ThisWorkbook.Worksheets("Conditional Orders")
                .Range(.Cells(eyy, 13), .Cells(eyy, 25)).ClearContents
                End With
                With ThisWorkbook.Worksheets("Control Panel")
                .Range(.Cells(56, exx), .Cells(61, exx)).ClearContents
                End With
                Cells(47, exx) = "x"
            ResetConditionalOrders
            Else
            Sheets("Conditional Orders").Cells(eyy, 20) = "Filled"
            End If
        MicroWarn = 0
    End If
End If
If Cells(45, exx) <> "" And Cells(46, exx) <> "" Then
    If Cells(56, exx) = "SELL" And Cells(52, exx) <= Cells(46, exx) Then
        MicroWarn = 10
        Sheets("Conditional Orders").Select
        Sheets("Conditional Orders").Cells(eyy, 1).Select
        Sheets("Conditional Orders").cancelConditionalOrder
        Sheets("Control Panel").Select
        y = Timer + 2
        Do While Timer < y
        DoEvents
        Loop
            If Sheets("Conditional Orders").Cells(eyy, 21) = 0 Then
                With ThisWorkbook.Worksheets("Conditional Orders")
                .Range(.Cells(eyy, 13), .Cells(eyy, 25)).ClearContents
                End With
                With ThisWorkbook.Worksheets("Control Panel")
                .Range(.Cells(56, exx), .Cells(61, exx)).ClearContents
                End With
                Cells(47, exx) = "x"
            ResetConditionalOrders
            Else
            Sheets("Conditional Orders").Cells(eyy, 20) = "Filled"
            End If
        MicroWarn = 0
    End If
End If

If (Cells(56, exx) = "BUY" Or Cells(56, exx) = "SELL") Then
    If Sheets("Conditional Orders").Cells(eyy, 20) = "Filled" Then
    MicroWarn = 10
    Cells(56, exx) = "OPENset"
    Sheets("Conditional Orders").Select
    Sheets("Conditional Orders").Cells(eyy, 1).Select
    Sheets("Conditional Orders").cancelConditionalOrder
    Sheets("Control Panel").Select
    y = Timer + 2
    Do While Timer < y
    DoEvents
    Loop
    Sheets("Conditional Orders").Cells(eyy, 86) = Sheets("Conditional Orders").Cells(eyy, 23)
    If Cells(60, exx) = "BUY" Then
    Sheets("Conditional Orders").Cells(eyy, 13) = "SELL"
    Else
    Sheets("Conditional Orders").Cells(eyy, 13) = "BUY"
    End If
    Sheets("Conditional Orders").Cells(eyy, 14) = Sheets("Conditional Orders").Cells(eyy, 21)
    Sheets("Conditional Orders").Cells(eyy, 15) = "LMT"
    With ThisWorkbook.Worksheets("Conditional Orders")
    .Range(.Cells(eyy, 18), .Cells(eyy, 25)).ClearContents
    End With
    Sheets("Conditional Orders").Cells(eyy, 20) = "Submitted"
    MicroWarn = 0
    Cells(56, exx) = "OPEN"
    End If
End If
eyy = eyy - 25

exx = exx + 2
eyy = eyy + 1
If exx < 52 Then
GoTo marker:
End If

End If

Errhandler:


End Sub


Sub PTpaste()
'This sub inserts our default price targets
Dim EEEE As Integer

On Error GoTo Errhandler
If MicroWarn <> 10 And Range("C3") = "OPEN" Then
EEEE = 3
marker:

If Cells(12, EEEE) <> "" And Cells(13, EEEE) <> "" And Cells(11, EEEE) = "" Then
    Cells(11, EEEE) = (Cells(12, EEEE) - Cells(13, EEEE)) * 1.5 + Cells(12, EEEE)
    Cells(11, EEEE) = WorksheetFunction.Round(Cells(11, EEEE), 2)
End If
If Cells(12, EEEE) <> "" And Cells(13, EEEE) <> "" And Cells(14, EEEE) = "" Then
    Cells(14, EEEE) = Cells(13, EEEE) - (Cells(12, EEEE) - Cells(13, EEEE)) * 1.5
    Cells(14, EEEE) = WorksheetFunction.Round(Cells(14, EEEE), 2)
End If

If Cells(44, EEEE) <> "" And Cells(45, EEEE) <> "" And Cells(43, EEEE) = "" Then
    Cells(43, EEEE) = (Cells(44, EEEE) - Cells(45, EEEE)) * 1.5 + Cells(44, EEEE)
    Cells(43, EEEE) = WorksheetFunction.Round(Cells(43, EEEE), 2)
End If
If Cells(44, EEEE) <> "" And Cells(45, EEEE) <> "" And Cells(46, EEEE) = "" Then
    Cells(46, EEEE) = Cells(45, EEEE) - (Cells(44, EEEE) - Cells(45, EEEE)) * 1.5
    Cells(46, EEEE) = WorksheetFunction.Round(Cells(46, EEEE), 2)
End If


EEEE = EEEE + 2
If EEEE < 52 Then
GoTo marker:
End If

End If
Errhandler:

End Sub

Sub CheckTradesCCCC()
    'this module runs up to 5 times per second
    'it runs the buy and short sell code
    'it looks for stopout criteria
    'it runs part of the live data recording code
    'it watches to exit if your price hits your price target

    Dim CCCC As Integer


    On Error GoTo Errhandler
    If RecordOn = 1 And MicroWarn = 0 Then
    PTpaste
    MicroWarn = 10

    CCCC = 3
    marker:

    'this part of the module updates the live data recording cells (row 31 and 32)
    If Range("S2") = True And Cells(20, CCCC) <> "" And Cells(20, CCCC) > 0 Then
            If Cells(20, CCCC) > Cells(31, CCCC) Then
            Cells(31, CCCC) = Cells(20, CCCC)
            End If
            If Cells(20, CCCC) < Cells(32, CCCC) Or Cells(32, CCCC) = "" Then
            Cells(32, CCCC) = Cells(20, CCCC)
            End If
    End If

    'this part of the module watches for the price to hit your price target and triggers the "Take Profits" module to send the exit order
    If Cells(11, CCCC) <> "" And Range("W2") = True Then
        If Cells(20, CCCC) >= Cells(11, CCCC) And Cells(24, CCCC) = "OPEN" And Cells(28, CCCC) = "BUY" Then
        Cells(24, CCCC) = "EXITset"
        'it waits 2 seconds before sending the order. you can adjust it to whatever you want below
        Application.OnTime Now() + TimeValue("00:00:02"), "getback34"
        End If
    End If
    If Cells(14, CCCC) <> "" And Range("W2") = True Then
        If Cells(20, CCCC) <= Cells(14, CCCC) And Cells(24, CCCC) = "OPEN" And Cells(28, CCCC) = "SELL" Then
        Cells(24, CCCC) = "EXITset"
        Application.OnTime Now() + TimeValue("00:00:02"), "getback34"
        End If
    End If

    'the part of the module looks for a buy trade, sets up the order, and sends it (if "automatically send orders" is true)
    If Cells(5, CCCC) <> "" And Cells(13, CCCC) <> "" And Cells(12, CCCC) <> "" And Cells(20, CCCC) <> "" And Range("C3") <> "premarket" Then
        
        If Cells(15, CCCC) = "b" Then
            If Cells(20, CCCC) > Cells(12, CCCC) And Cells(24, CCCC) = "" Then
                ResetConditionalOrders
                Cells(24, CCCC) = "BUY"
                Cells(27, CCCC) = Range("AQ2").value / Cells(20, CCCC) * (Cells(20, CCCC) * 0.01 / ((Cells(12, CCCC) - Cells(13, CCCC)) * 0.85))
                    If Cells(27, CCCC) * Cells(20, CCCC) > 2 * Range("AQ2").value Then
                    Cells(27, CCCC) = 2 * Range("AQ2").value / Cells(20, CCCC)
                    End If
                Cells(27, CCCC) = WorksheetFunction.Round(Cells(27, CCCC), 0)
                Cells(28, CCCC) = "BUY"
                Cells(29, CCCC) = Cells(12, CCCC) + (Cells(12, CCCC) - Cells(13, CCCC)) * 0.18
                Cells(29, CCCC) = WorksheetFunction.Round(Cells(29, CCCC).value, 2)
                TestPlayWavFile
                    If Range("W2") = True Then
                    Cells(25, CCCC) = "y"
                    y = Timer + 2
                    Do While Timer < y
                    DoEvents
                    Loop
                    End If
            End If
            'if the price breaks your stop price before triggering a buy order, then the system replaces the buy setup with a 45 minute timer
                If Cells(20, CCCC) < Cells(13, CCCC) And Cells(28, CCCC) = "" And Cells(24, CCCC) = "" Then
                Cells(15, CCCC) = 45
                End If
                'this part of the module looks for a stopout if the price is 5% of the range width below your stop price
                If Cells(20, CCCC) < Cells(13, CCCC) - 0.05 * (Cells(12, CCCC) - Cells(13, CCCC)) And Cells(28, CCCC) = "BUY" And Range("AC2") = True And (Cells(24, CCCC) = "OPEN" Or Cells(24, CCCC) = "EXIT" Or Cells(24, CCCC) = "Timeout") Then
                TestPlayWavFile6
                Cells(24, CCCC) = "stopout"
                MicroWarn = 0
                Timeout
                End If
        End If
        
        'this part of the module is the same as above but is for a short selling order
        If Cells(15, CCCC) = "s" Then
            If Cells(20, CCCC) < Cells(13, CCCC) And Cells(24, CCCC) = "" Then
                ResetConditionalOrders
                Cells(24, CCCC) = "SELL"
                Cells(27, CCCC) = Range("AQ2").value / Cells(20, CCCC) * (Cells(20, CCCC) * 0.01 / ((Cells(12, CCCC) - Cells(13, CCCC)) * 0.85))
                    If Cells(27, CCCC) * Cells(20, CCCC) > 2 * Range("AQ2").value Then
                    Cells(27, CCCC) = 2 * Range("AQ2").value / Cells(20, CCCC)
                    End If
                Cells(27, CCCC) = WorksheetFunction.Round(Cells(27, CCCC), 0)
                Cells(28, CCCC) = "SELL"
                Cells(29, CCCC) = Cells(13, CCCC) - (Cells(12, CCCC) - Cells(13, CCCC)) * 0.18
                Cells(29, CCCC) = WorksheetFunction.Round(Cells(29, CCCC), 2)
                TestPlayWavFile
                    If Range("W2") = True Then
                    Cells(25, CCCC) = "y"
                    y = Timer + 2
                    Do While Timer < y
                    DoEvents
                    Loop
                    End If
            End If
                If Cells(20, CCCC) > Cells(12, CCCC) And Cells(28, CCCC) = "" And Cells(24, CCCC) = "" Then
                Cells(15, CCCC) = 45
                End If
                If Cells(20, CCCC) > Cells(12, CCCC) + 0.05 * (Cells(12, CCCC) - Cells(13, CCCC)) And Cells(28, CCCC) = "SELL" And Range("AC2") = True And (Cells(24, CCCC) = "OPEN" Or Cells(24, CCCC) = "EXIT" Or Cells(24, CCCC) = "Timeout") Then
                TestPlayWavFile6
                Cells(24, CCCC) = "stopout"
                MicroWarn = 0
                Timeout
                End If
        End If
    End If

    ' expansion cells --------------------------------------------------------------------------------------------------------------------------------

    If Range("S2") = True And Cells(52, CCCC) <> "" And Cells(52, CCCC) > 0 Then
            If Cells(52, CCCC) > Cells(63, CCCC) Then
            Cells(63, CCCC) = Cells(52, CCCC)
            End If
            If Cells(52, CCCC) < Cells(64, CCCC) Or Cells(64, CCCC) = "" Then
            Cells(64, CCCC) = Cells(52, CCCC)
            End If
    End If

    If Cells(43, CCCC) <> "" And Range("W2") = True Then
        If Cells(52, CCCC) >= Cells(43, CCCC) And Cells(56, CCCC) = "OPEN" And Cells(60, CCCC) = "BUY" Then
        Cells(56, CCCC) = "EXITset"
        Application.OnTime Now() + TimeValue("00:00:02"), "getback34"
        End If
    End If
    If Cells(46, CCCC) <> "" And Range("W2") = True Then
        If Cells(52, CCCC) <= Cells(46, CCCC) And Cells(56, CCCC) = "OPEN" And Cells(60, CCCC) = "SELL" Then
        Cells(56, CCCC) = "EXITset"
        Application.OnTime Now() + TimeValue("00:00:02"), "getback34"
        End If
    End If
    If Cells(37, CCCC) <> "" And Cells(45, CCCC) <> "" And Cells(44, CCCC) <> "" And Cells(52, CCCC) <> "" And Range("C3") <> "premarket" Then
        If Cells(47, CCCC) = "b" Then
        
            If Cells(52, CCCC) > Cells(44, CCCC) And Cells(56, CCCC) = "" Then
                ResetConditionalOrders
                Cells(56, CCCC) = "BUY"
                Cells(59, CCCC) = Range("AQ2").value / Cells(52, CCCC) * (Cells(52, CCCC) * 0.01 / ((Cells(44, CCCC) - Cells(45, CCCC)) * 0.85))
                    If Cells(59, CCCC) * Cells(52, CCCC) > 2 * Range("AQ2").value Then
                    Cells(59, CCCC) = 2 * Range("AQ2").value / Cells(52, CCCC)
                    End If
                Cells(59, CCCC) = WorksheetFunction.Round(Cells(59, CCCC), 0)
                Cells(60, CCCC) = "BUY"
                Cells(61, CCCC) = Cells(44, CCCC) + (Cells(44, CCCC) - Cells(45, CCCC)) * 0.18
                Cells(61, CCCC) = WorksheetFunction.Round(Cells(61, CCCC).value, 2)
                TestPlayWavFile
                    If Range("W2") = True Then
                    Cells(57, CCCC) = "y"
                    y = Timer + 2
                    Do While Timer < y
                    DoEvents
                    Loop
                    End If
            End If
                If Cells(52, CCCC) < Cells(45, CCCC) And Cells(60, CCCC) = "" And Cells(56, CCCC) = "" Then
                Cells(47, CCCC) = 45
                End If
                If Cells(52, CCCC) < Cells(45, CCCC) - 0.05 * (Cells(44, CCCC) - Cells(45, CCCC)) And Cells(60, CCCC) = "BUY" And Range("AC2") = True And (Cells(56, CCCC) = "OPEN" Or Cells(56, CCCC) = "EXIT" Or Cells(56, CCCC) = "Timeout") Then
                TestPlayWavFile6
                Cells(56, CCCC) = "stopout"
                MicroWarn = 0
                Timeout
                End If
        End If
        
        If Cells(47, CCCC) = "s" Then
            If Cells(52, CCCC) < Cells(45, CCCC) And Cells(56, CCCC) = "" Then
                ResetConditionalOrders
                Cells(56, CCCC) = "SELL"
                Cells(59, CCCC) = Range("AQ2").value / Cells(52, CCCC) * (Cells(52, CCCC) * 0.01 / ((Cells(44, CCCC) - Cells(45, CCCC)) * 0.85))
                    If Cells(59, CCCC) * Cells(52, CCCC) > 2 * Range("AQ2").value Then
                    Cells(59, CCCC) = 2 * Range("AQ2").value / Cells(52, CCCC)
                    End If
                Cells(59, CCCC) = WorksheetFunction.Round(Cells(59, CCCC), 0)
                Cells(60, CCCC) = "SELL"
                Cells(61, CCCC) = Cells(45, CCCC) - (Cells(44, CCCC) - Cells(45, CCCC)) * 0.18
                Cells(61, CCCC) = WorksheetFunction.Round(Cells(61, CCCC), 2)
                TestPlayWavFile
                    If Range("W2") = True Then
                    Cells(57, CCCC) = "y"
                    y = Timer + 2
                    Do While Timer < y
                    DoEvents
                    Loop
                    End If
            End If
                If Cells(52, CCCC) > Cells(44, CCCC) And Cells(60, CCCC) = "" And Cells(56, CCCC) = "" Then
                Cells(47, CCCC) = 45
                End If
                If Cells(52, CCCC) > Cells(44, CCCC) + 0.05 * (Cells(44, CCCC) - Cells(45, CCCC)) And Cells(60, CCCC) = "SELL" And Range("AC2") = True And (Cells(56, CCCC) = "OPEN" Or Cells(56, CCCC) = "EXIT" Or Cells(56, CCCC) = "Timeout") Then
                TestPlayWavFile6
                Cells(56, CCCC) = "stopout"
                MicroWarn = 0
                Timeout
                End If
        End If
    End If

    CCCC = CCCC + 2
    If CCCC < 52 Then
    GoTo marker:
    End If

    Errhandler:
    MicroWarn = 0

    End If

    If CCCC <> 53 And CCCC <> 0 Then
        Range("I3") = CCCC
    End If
    SendExitOrder
    WinCheck

End Sub




Sub HighTimer()
'This sub calls the timer every 15 seconds
'It also calls the sub that adjusts the timeouts or the stopouts every 15 seconds
On Error GoTo Errhandler:

If RecordOn = 1 Then
Application.OnTime Now() + TimeValue("00:00:15"), "getback18"
   
End If



Errhandler:

Timeout
Alarm50

End Sub


Sub CloseTime()
    'this sub sets any open positions to "stopout" mode near the end of the day
    Dim ctxx As Integer

    ctxx = 3
    marker:

    If Cells(24, ctxx) = "OPEN" Then
        Cells(24, ctxx) = "stopout"
    End If

    If Cells(56, ctxx) = "OPEN" Then
        Cells(56, ctxx) = "stopout"
    End If

    ctxx = ctxx + 2
    If ctxx < 52 Then
        GoTo marker:
    End If

End Sub


Sub Timeout()
'this sub sends our stopout orders
On Error GoTo Errhandler

SendExitOrder

Dim toxx As Integer
Dim toyy As Integer

toxx = 3
toyy = 10

marker:

If RecordOn = 1 And Range("AC2") = True And MicroWarn <> 10 Then

If (Cells(24, toxx) = "Timeout" Or Cells(24, toxx) = "stopout") And Sheets("Conditional Orders").Cells(toyy, 20) = "Submitted" Then
    If Cells(28, toxx) = "SELL" Then
    Cells(29, toxx) = Cells(19, toxx)
    Else
    Cells(29, toxx) = Cells(22, toxx)
    End If
    MicroWarn = 10
    Cells(29, toxx) = WorksheetFunction.Round(Cells(29, toxx).value, 2)
    Sheets("Conditional Orders").Cells(toyy, 16) = Cells(29, toxx)
    Sheets("Conditional Orders").Select
    Sheets("Conditional Orders").Cells(toyy, 1).Select
    Sheets("Conditional Orders").placeConditionalOrder
    y = Timer + 2
    Do While Timer < y
    DoEvents
    Loop
    Sheets("Control Panel").Select
    MicroWarn = 0
End If

toyy = toyy + 25
If (Cells(56, toxx) = "Timeout" Or Cells(56, toxx) = "stopout") And Sheets("Conditional Orders").Cells(toyy, 20) = "Submitted" Then
    If Cells(60, toxx) = "SELL" Then
    Cells(61, toxx) = Cells(51, toxx)
    Else
    Cells(61, toxx) = Cells(54, toxx)
    End If
    MicroWarn = 10
    Cells(61, toxx) = WorksheetFunction.Round(Cells(61, toxx).value, 2)
    Sheets("Conditional Orders").Cells(toyy, 16) = Cells(61, toxx)
    Sheets("Conditional Orders").Select
    Sheets("Conditional Orders").Cells(toyy, 1).Select
    Sheets("Conditional Orders").placeConditionalOrder
    y = Timer + 2
    Do While Timer < y
    DoEvents
    Loop
    Sheets("Control Panel").Select
    MicroWarn = 0
End If
toyy = toyy - 25

toxx = toxx + 2
toyy = toyy + 1
If toxx < 52 Then
GoTo marker:
End If


CalculateGain

End If

Errhandler:

End Sub



Sub RecordDataOn()
'the second time this module runs is at exactly 8:30:00 which is the market open in my time zone
'at that time it starts the live market data recording code (getback 30) and set the first row to 3


    If Range("C2") = "ON" And Range("C3") = "premarket" Then
        TicksRowA = 3
        getback30

        Range("C3").Interior.colorIndex = 4
        Range("C3") = "OPEN"
    End If

    If Range("C2") = "off" Then
    'the first time this module runs it sets the time to automatically exit all open positions ("14:52:10")
    'then is sets the time to run this module again exactly when the market opens ("8:30:00") in my time zone
    'then it starts the alarm code (HighTimer), the timer code (Getback13), and the clock code runClock (getback20)


        If Sheets("Ticks").Range("B3") <> "" And Range("S2") = True Then
            Dim Answer2 As String
            Answer2 = MsgBox("Clear Live Recorded Data?", vbQuestion + vbYesNo, "Clear Data?")

            If Answer2 = vbYes Then
                Sheets("Ticks").Range("B3:GS2220").ClearContents
            End If
        End If

        RecordOn = 1

        'enable this event to close all open positions at 2:52:10 pm
        'Application.OnTime TimeValue("14:52:10"), "getback21"

        'this event activates the system at the market open
        Application.OnTime TimeValue("8:30:00"), "recordDataOnWrapper"
        'Application.OnTime TimeValue("22:38:00"), "recordDataOnWrapper"

        HighTimer

        'getback13 runs the timers
        Getback13

        'getback20 runs the clock
        runClock


        Range("C2").Interior.colorIndex = 4
        Range("C2") = "ON"

        CheckSetup
    End If
End Sub

Sub CheckSetup()
'This module runs as part of the Set Up Workbook process
'it checks for the setup error "0" and reminds the user to add the text "ISLAND" to the primary exchange column on the tickers page

Dim CCC As Integer
CCC = 3
marker:
If Cells(5, CCC) <> "" And Cells(18, CCC) = 0 Then
MsgBox "Setup Error on " & Cells(5, CCC)
MsgBox "Set Tickers Primary Exchange column to ISLAND. Set Conditional Orders Primary Exchange column to ISLAND."
End If
If Cells(37, CCC) <> "" And Cells(50, CCC) = 0 Then
MsgBox "Setup Error on " & Cells(37, CCC)
MsgBox "Set Tickers Primary Exchange column to ISLAND. Set Conditional Orders Primary Exchange column to ISLAND."
End If
CCC = CCC + 2
If CCC < 52 Then
GoTo marker:
End If
End Sub

Sub RecordDataOff()

    If Range("C28") <> "" Or Range("E28") <> "" Or Range("G28") <> "" Or Range("I28") <> "" Or Range("K28") <> "" Or Range("M28") <> "" Or Range("O28") <> "" Or Range("Q28") <> "" Or Range("S28") <> "" Or Range("U28") <> "" Or Range("W28") <> "" Or Range("Y28") <> "" Or Range("AA28") <> "" Or Range("AC28") <> "" Or Range("AE28") <> "" Or Range("AG28") <> "" Or Range("AI28") <> "" Or Range("AK28") <> "" Or Range("AM28") <> "" Or Range("AO28") <> "" Or Range("AQ28") <> "" Or Range("AS28") <> "" Or Range("AU28") <> "" Or Range("AW28") <> "" Or Range("AY28") <> "" Then
        Dim Answer As String
        Answer = MsgBox("You still have open orders! Are you sure you want to deactive the system?", vbQuestion + vbYesNo, "Turn System Off?")
        If Answer = vbYes Then
            RecordOn = 0
            Range("C2").Interior.colorIndex = 3
            Range("C2") = "off"
            Range("C3").Interior.colorIndex = 3
            Range("C3") = "premarket"
        End If
    Else
        RecordOn = 0
        Range("C2").Interior.colorIndex = 3
        Range("C2") = "off"
        Range("C3").Interior.colorIndex = 3
        Range("C3") = "premarket"
    End If

    Dim Answer2 As String
    Answer2 = MsgBox("Log Data?", vbQuestion + vbYesNo, "Log?")

    If Answer2 = vbYes Then
        LogRow = 2
        Gains = 0
        Sheets("Log").Select
        Sheets("Log").Logworksheet
        Sheets("Control Panel").Select
    End If

End Sub




Sub SetUpWorksheets()
    'This module runs as part of the Set Up Workbook process
    'If the workbook is already live, then it will only add the new symbols to the tickers page and the conditional orders page.
    'the user must then hit the "request market"  button manually for each new symbol added

    'if the workbook is not live (normal setup) then this module clears all the old links on the
    'tickers page and waits for a few seconds for that code to fully process before taking the
    'next step

    If RecordOn = 1 Then
        MsgBox "System is currently live!"

        Dim tickersSymbolPos As Variant
        tickersSymbolPos = Array( _
            "A12", "A13", "A14", "A15", "A16", "A17", "A18", "A19", "A20", _
            "A21", "A22", "A23", "A24", "A25", "A26", "A27", "A28", "A29", _
            "A30", "A31", "A32", "A33", "A34", "A35", "A36", "A37", "A38", _
            "A39", "A40", "A41", "A42", "A43", "A44", "A45", "A46", "A47", _
            "A48", "A49", "A50", "A51", "A52", "A53", "A54", "A55", "A56", _
            "A57", "A58", "A59", "A60", "A61" _
        )
        
        Dim ctrPanelSymbolPos As Variant
        ctrPanelSymbolPos = Array( _
            "C5", "E5", "G5", "I5", "K5", "M5", "O5", "Q5", "S5", "U5", "W5", "Y5", _
            "AA5", "AC5", "AE5", "AG5", "AI5", "AK5", "AM5", "AO5", "AQ5", "AS5", "AU5", "AW5", "AY5", _
            "C37", "E37", "G37", "I37", "K37", "M37", "O37", "Q37", "S37", "U37", "W37", "Y37", _
            "AA37", "AC37", "AE37", "AG37", "AI37", "AK37", "AM37", "AO37", "AQ37", "AS37", "AU37", "AW37", "AY37" _
        )

        Dim idx As Integer
        For idx = 0 To 49
            Sheets("Tickers").Range(tickersSymbolPos(idx)) = Range(ctrPanelSymbolPos(idx))
        Next

        Sheets("Tickers").Select
        Sheets("Tickers").Range("A12:A61").Select
        Selection.Copy
        Sheets("Conditional Orders").Select
        Sheets("Conditional Orders").Range("A10").Select
        ActiveSheet.Paste
        Sheets("Tickers").Select
        Range("C1").ClearContents
        Exit Sub
    End If

    Range("E2") = Date
    'Range("C24:AY29").ClearContents
    'Range("C56:AY61").ClearContents

    Sheets("Tickers").Select
    Sheets("Tickers").clearLinks

    Application.OnTime Now() + TimeValue("00:00:05"), "getback9"

End Sub


Sub continue45()
'this module copies each symbol from the "Control Panel" sheet to the "Tickers" sheet
'then it copies those symbols and pastes them into the "Conditional Orders" sheet too

Sheets("Tickers").Range("A12") = Range("C5")
Sheets("Tickers").Range("A13") = Range("E5")
Sheets("Tickers").Range("A14") = Range("G5")
Sheets("Tickers").Range("A15") = Range("I5")
Sheets("Tickers").Range("A16") = Range("K5")
Sheets("Tickers").Range("A17") = Range("M5")
Sheets("Tickers").Range("A18") = Range("O5")
Sheets("Tickers").Range("A19") = Range("Q5")
Sheets("Tickers").Range("A20") = Range("S5")
Sheets("Tickers").Range("A21") = Range("U5")
Sheets("Tickers").Range("A22") = Range("W5")
Sheets("Tickers").Range("A23") = Range("Y5")
Sheets("Tickers").Range("A24") = Range("AA5")
Sheets("Tickers").Range("A25") = Range("AC5")
Sheets("Tickers").Range("A26") = Range("AE5")
Sheets("Tickers").Range("A27") = Range("AG5")
Sheets("Tickers").Range("A28") = Range("AI5")
Sheets("Tickers").Range("A29") = Range("AK5")
Sheets("Tickers").Range("A30") = Range("AM5")
Sheets("Tickers").Range("A31") = Range("AO5")
Sheets("Tickers").Range("A32") = Range("AQ5")
Sheets("Tickers").Range("A33") = Range("AS5")
Sheets("Tickers").Range("A34") = Range("AU5")
Sheets("Tickers").Range("A35") = Range("AW5")
Sheets("Tickers").Range("A36") = Range("AY5")
Sheets("Tickers").Range("A37") = Range("C37")
Sheets("Tickers").Range("A38") = Range("E37")
Sheets("Tickers").Range("A39") = Range("G37")
Sheets("Tickers").Range("A40") = Range("I37")
Sheets("Tickers").Range("A41") = Range("K37")
Sheets("Tickers").Range("A42") = Range("M37")
Sheets("Tickers").Range("A43") = Range("O37")
Sheets("Tickers").Range("A44") = Range("Q37")
Sheets("Tickers").Range("A45") = Range("S37")
Sheets("Tickers").Range("A46") = Range("U37")
Sheets("Tickers").Range("A47") = Range("W37")
Sheets("Tickers").Range("A48") = Range("Y37")
Sheets("Tickers").Range("A49") = Range("AA37")
Sheets("Tickers").Range("A50") = Range("AC37")
Sheets("Tickers").Range("A51") = Range("AE37")
Sheets("Tickers").Range("A52") = Range("AG37")
Sheets("Tickers").Range("A53") = Range("AI37")
Sheets("Tickers").Range("A54") = Range("AK37")
Sheets("Tickers").Range("A55") = Range("AM37")
Sheets("Tickers").Range("A56") = Range("AO37")
Sheets("Tickers").Range("A57") = Range("AQ37")
Sheets("Tickers").Range("A58") = Range("AS37")
Sheets("Tickers").Range("A59") = Range("AU37")
Sheets("Tickers").Range("A60") = Range("AW37")
Sheets("Tickers").Range("A61") = Range("AY37")

Sheets("Tickers").Select
Sheets("Tickers").Range("A12:A61").Select
Selection.Copy


Sheets("Conditional Orders").Select
Sheets("Conditional Orders").Range("A10").Select
ActiveSheet.Paste
Sheets("Conditional Orders").Range("L10:Y59").ClearContents
ResetConditionalOrders


Sheets("Tickers").Select
Sheets("Tickers").Range("A12").Select
Sheets("Tickers").GetMarketDataN



End Sub



Sub ResetConditionalOrders()
'this sheet sets up the "Conditional Orders" sheet to be ready to send trades via the manual "y" method

Dim rco As Integer
rco = 10

Application.Calculation = xlManual
Sheets("Control Panel").Select
Range("C25:AY25").ClearContents
Range("C57:AY57").ClearContents

marker:

Sheets("Conditional Orders").Cells(rco, 27) = "ADD"
Sheets("Conditional Orders").Cells(rco, 30) = "LMT"

rco = rco + 1
If rco < 60 Then
GoTo marker:
End If

Sheets("Conditional Orders").Range("Z10").Formula = "='Control Panel'!C25=""y"""
Sheets("Conditional Orders").Range("Z11").Formula = "='Control Panel'!E25=""y"""
Sheets("Conditional Orders").Range("Z12").Formula = "='Control Panel'!G25=""y"""
Sheets("Conditional Orders").Range("Z13").Formula = "='Control Panel'!I25=""y"""
Sheets("Conditional Orders").Range("Z14").Formula = "='Control Panel'!K25=""y"""
Sheets("Conditional Orders").Range("Z15").Formula = "='Control Panel'!M25=""y"""
Sheets("Conditional Orders").Range("Z16").Formula = "='Control Panel'!O25=""y"""
Sheets("Conditional Orders").Range("Z17").Formula = "='Control Panel'!Q25=""y"""
Sheets("Conditional Orders").Range("Z18").Formula = "='Control Panel'!S25=""y"""
Sheets("Conditional Orders").Range("Z19").Formula = "='Control Panel'!U25=""y"""
Sheets("Conditional Orders").Range("Z20").Formula = "='Control Panel'!W25=""y"""
Sheets("Conditional Orders").Range("Z21").Formula = "='Control Panel'!Y25=""y"""
Sheets("Conditional Orders").Range("Z22").Formula = "='Control Panel'!AA25=""y"""
Sheets("Conditional Orders").Range("Z23").Formula = "='Control Panel'!AC25=""y"""
Sheets("Conditional Orders").Range("Z24").Formula = "='Control Panel'!AE25=""y"""
Sheets("Conditional Orders").Range("Z25").Formula = "='Control Panel'!AG25=""y"""
Sheets("Conditional Orders").Range("Z26").Formula = "='Control Panel'!AI25=""y"""
Sheets("Conditional Orders").Range("Z27").Formula = "='Control Panel'!AK25=""y"""
Sheets("Conditional Orders").Range("Z28").Formula = "='Control Panel'!AM25=""y"""
Sheets("Conditional Orders").Range("Z29").Formula = "='Control Panel'!AO25=""y"""
Sheets("Conditional Orders").Range("Z30").Formula = "='Control Panel'!AQ25=""y"""
Sheets("Conditional Orders").Range("Z31").Formula = "='Control Panel'!AS25=""y"""
Sheets("Conditional Orders").Range("Z32").Formula = "='Control Panel'!AU25=""y"""
Sheets("Conditional Orders").Range("Z33").Formula = "='Control Panel'!AW25=""y"""
Sheets("Conditional Orders").Range("Z34").Formula = "='Control Panel'!AY25=""y"""
Sheets("Conditional Orders").Range("Z35").Formula = "='Control Panel'!C57=""y"""
Sheets("Conditional Orders").Range("Z36").Formula = "='Control Panel'!E57=""y"""
Sheets("Conditional Orders").Range("Z37").Formula = "='Control Panel'!G57=""y"""
Sheets("Conditional Orders").Range("Z38").Formula = "='Control Panel'!I57=""y"""
Sheets("Conditional Orders").Range("Z39").Formula = "='Control Panel'!K57=""y"""
Sheets("Conditional Orders").Range("Z40").Formula = "='Control Panel'!M57=""y"""
Sheets("Conditional Orders").Range("Z41").Formula = "='Control Panel'!O57=""y"""
Sheets("Conditional Orders").Range("Z42").Formula = "='Control Panel'!Q57=""y"""
Sheets("Conditional Orders").Range("Z43").Formula = "='Control Panel'!S57=""y"""
Sheets("Conditional Orders").Range("Z44").Formula = "='Control Panel'!U57=""y"""
Sheets("Conditional Orders").Range("Z45").Formula = "='Control Panel'!W57=""y"""
Sheets("Conditional Orders").Range("Z46").Formula = "='Control Panel'!Y57=""y"""
Sheets("Conditional Orders").Range("Z47").Formula = "='Control Panel'!AA57=""y"""
Sheets("Conditional Orders").Range("Z48").Formula = "='Control Panel'!AC57=""y"""
Sheets("Conditional Orders").Range("Z49").Formula = "='Control Panel'!AE57=""y"""
Sheets("Conditional Orders").Range("Z50").Formula = "='Control Panel'!AG57=""y"""
Sheets("Conditional Orders").Range("Z51").Formula = "='Control Panel'!AI57=""y"""
Sheets("Conditional Orders").Range("Z52").Formula = "='Control Panel'!AK57=""y"""
Sheets("Conditional Orders").Range("Z53").Formula = "='Control Panel'!AM57=""y"""
Sheets("Conditional Orders").Range("Z54").Formula = "='Control Panel'!AO57=""y"""
Sheets("Conditional Orders").Range("Z55").Formula = "='Control Panel'!AQ57=""y"""
Sheets("Conditional Orders").Range("Z56").Formula = "='Control Panel'!AS57=""y"""
Sheets("Conditional Orders").Range("Z57").Formula = "='Control Panel'!AU57=""y"""
Sheets("Conditional Orders").Range("Z58").Formula = "='Control Panel'!AW57=""y"""
Sheets("Conditional Orders").Range("Z59").Formula = "='Control Panel'!AY57=""y"""

Sheets("Conditional Orders").Range("AB10").Formula = "='Control Panel'!C28"
Sheets("Conditional Orders").Range("AB11").Formula = "='Control Panel'!E28"
Sheets("Conditional Orders").Range("AB12").Formula = "='Control Panel'!G28"
Sheets("Conditional Orders").Range("AB13").Formula = "='Control Panel'!I28"
Sheets("Conditional Orders").Range("AB14").Formula = "='Control Panel'!K28"
Sheets("Conditional Orders").Range("AB15").Formula = "='Control Panel'!M28"
Sheets("Conditional Orders").Range("AB16").Formula = "='Control Panel'!O28"
Sheets("Conditional Orders").Range("AB17").Formula = "='Control Panel'!Q28"
Sheets("Conditional Orders").Range("AB18").Formula = "='Control Panel'!S28"
Sheets("Conditional Orders").Range("AB19").Formula = "='Control Panel'!U28"
Sheets("Conditional Orders").Range("AB20").Formula = "='Control Panel'!W28"
Sheets("Conditional Orders").Range("AB21").Formula = "='Control Panel'!Y28"
Sheets("Conditional Orders").Range("AB22").Formula = "='Control Panel'!AA28"
Sheets("Conditional Orders").Range("AB23").Formula = "='Control Panel'!AC28"
Sheets("Conditional Orders").Range("AB24").Formula = "='Control Panel'!AE28"
Sheets("Conditional Orders").Range("AB25").Formula = "='Control Panel'!AG28"
Sheets("Conditional Orders").Range("AB26").Formula = "='Control Panel'!AI28"
Sheets("Conditional Orders").Range("AB27").Formula = "='Control Panel'!AK28"
Sheets("Conditional Orders").Range("AB28").Formula = "='Control Panel'!AM28"
Sheets("Conditional Orders").Range("AB29").Formula = "='Control Panel'!AO28"
Sheets("Conditional Orders").Range("AB30").Formula = "='Control Panel'!AQ28"
Sheets("Conditional Orders").Range("AB31").Formula = "='Control Panel'!AS28"
Sheets("Conditional Orders").Range("AB32").Formula = "='Control Panel'!AU28"
Sheets("Conditional Orders").Range("AB33").Formula = "='Control Panel'!AW28"
Sheets("Conditional Orders").Range("AB34").Formula = "='Control Panel'!AY28"
Sheets("Conditional Orders").Range("AB35").Formula = "='Control Panel'!C60"
Sheets("Conditional Orders").Range("AB36").Formula = "='Control Panel'!E60"
Sheets("Conditional Orders").Range("AB37").Formula = "='Control Panel'!G60"
Sheets("Conditional Orders").Range("AB38").Formula = "='Control Panel'!I60"
Sheets("Conditional Orders").Range("AB39").Formula = "='Control Panel'!K60"
Sheets("Conditional Orders").Range("AB40").Formula = "='Control Panel'!M60"
Sheets("Conditional Orders").Range("AB41").Formula = "='Control Panel'!O60"
Sheets("Conditional Orders").Range("AB42").Formula = "='Control Panel'!Q60"
Sheets("Conditional Orders").Range("AB43").Formula = "='Control Panel'!S60"
Sheets("Conditional Orders").Range("AB44").Formula = "='Control Panel'!U60"
Sheets("Conditional Orders").Range("AB45").Formula = "='Control Panel'!W60"
Sheets("Conditional Orders").Range("AB46").Formula = "='Control Panel'!Y60"
Sheets("Conditional Orders").Range("AB47").Formula = "='Control Panel'!AA60"
Sheets("Conditional Orders").Range("AB48").Formula = "='Control Panel'!AC60"
Sheets("Conditional Orders").Range("AB49").Formula = "='Control Panel'!AE60"
Sheets("Conditional Orders").Range("AB50").Formula = "='Control Panel'!AG60"
Sheets("Conditional Orders").Range("AB51").Formula = "='Control Panel'!AI60"
Sheets("Conditional Orders").Range("AB52").Formula = "='Control Panel'!AK60"
Sheets("Conditional Orders").Range("AB53").Formula = "='Control Panel'!AM60"
Sheets("Conditional Orders").Range("AB54").Formula = "='Control Panel'!AO60"
Sheets("Conditional Orders").Range("AB55").Formula = "='Control Panel'!AQ60"
Sheets("Conditional Orders").Range("AB56").Formula = "='Control Panel'!AS60"
Sheets("Conditional Orders").Range("AB57").Formula = "='Control Panel'!AU60"
Sheets("Conditional Orders").Range("AB58").Formula = "='Control Panel'!AW60"
Sheets("Conditional Orders").Range("AB59").Formula = "='Control Panel'!AY60"

Sheets("Conditional Orders").Range("AC10").Formula = "='Control Panel'!C27"
Sheets("Conditional Orders").Range("AC11").Formula = "='Control Panel'!E27"
Sheets("Conditional Orders").Range("AC12").Formula = "='Control Panel'!G27"
Sheets("Conditional Orders").Range("AC13").Formula = "='Control Panel'!I27"
Sheets("Conditional Orders").Range("AC14").Formula = "='Control Panel'!K27"
Sheets("Conditional Orders").Range("AC15").Formula = "='Control Panel'!M27"
Sheets("Conditional Orders").Range("AC16").Formula = "='Control Panel'!O27"
Sheets("Conditional Orders").Range("AC17").Formula = "='Control Panel'!Q27"
Sheets("Conditional Orders").Range("AC18").Formula = "='Control Panel'!S27"
Sheets("Conditional Orders").Range("AC19").Formula = "='Control Panel'!U27"
Sheets("Conditional Orders").Range("AC20").Formula = "='Control Panel'!W27"
Sheets("Conditional Orders").Range("AC21").Formula = "='Control Panel'!Y27"
Sheets("Conditional Orders").Range("AC22").Formula = "='Control Panel'!AA27"
Sheets("Conditional Orders").Range("AC23").Formula = "='Control Panel'!AC27"
Sheets("Conditional Orders").Range("AC24").Formula = "='Control Panel'!AE27"
Sheets("Conditional Orders").Range("AC25").Formula = "='Control Panel'!AG27"
Sheets("Conditional Orders").Range("AC26").Formula = "='Control Panel'!AI27"
Sheets("Conditional Orders").Range("AC27").Formula = "='Control Panel'!AK27"
Sheets("Conditional Orders").Range("AC28").Formula = "='Control Panel'!AM27"
Sheets("Conditional Orders").Range("AC29").Formula = "='Control Panel'!AO27"
Sheets("Conditional Orders").Range("AC30").Formula = "='Control Panel'!AQ27"
Sheets("Conditional Orders").Range("AC31").Formula = "='Control Panel'!AS27"
Sheets("Conditional Orders").Range("AC32").Formula = "='Control Panel'!AU27"
Sheets("Conditional Orders").Range("AC33").Formula = "='Control Panel'!AW27"
Sheets("Conditional Orders").Range("AC34").Formula = "='Control Panel'!AY27"
Sheets("Conditional Orders").Range("AC35").Formula = "='Control Panel'!C59"
Sheets("Conditional Orders").Range("AC36").Formula = "='Control Panel'!E59"
Sheets("Conditional Orders").Range("AC37").Formula = "='Control Panel'!G59"
Sheets("Conditional Orders").Range("AC38").Formula = "='Control Panel'!I59"
Sheets("Conditional Orders").Range("AC39").Formula = "='Control Panel'!K59"
Sheets("Conditional Orders").Range("AC40").Formula = "='Control Panel'!M59"
Sheets("Conditional Orders").Range("AC41").Formula = "='Control Panel'!O59"
Sheets("Conditional Orders").Range("AC42").Formula = "='Control Panel'!Q59"
Sheets("Conditional Orders").Range("AC43").Formula = "='Control Panel'!S59"
Sheets("Conditional Orders").Range("AC44").Formula = "='Control Panel'!U59"
Sheets("Conditional Orders").Range("AC45").Formula = "='Control Panel'!W59"
Sheets("Conditional Orders").Range("AC46").Formula = "='Control Panel'!Y59"
Sheets("Conditional Orders").Range("AC47").Formula = "='Control Panel'!AA59"
Sheets("Conditional Orders").Range("AC48").Formula = "='Control Panel'!AC59"
Sheets("Conditional Orders").Range("AC49").Formula = "='Control Panel'!AE59"
Sheets("Conditional Orders").Range("AC50").Formula = "='Control Panel'!AG59"
Sheets("Conditional Orders").Range("AC51").Formula = "='Control Panel'!AI59"
Sheets("Conditional Orders").Range("AC52").Formula = "='Control Panel'!AK59"
Sheets("Conditional Orders").Range("AC53").Formula = "='Control Panel'!AM59"
Sheets("Conditional Orders").Range("AC54").Formula = "='Control Panel'!AO59"
Sheets("Conditional Orders").Range("AC55").Formula = "='Control Panel'!AQ59"
Sheets("Conditional Orders").Range("AC56").Formula = "='Control Panel'!AS59"
Sheets("Conditional Orders").Range("AC57").Formula = "='Control Panel'!AU59"
Sheets("Conditional Orders").Range("AC58").Formula = "='Control Panel'!AW59"
Sheets("Conditional Orders").Range("AC59").Formula = "='Control Panel'!AY59"

Sheets("Conditional Orders").Range("AE10").Formula = "='Control Panel'!C29"
Sheets("Conditional Orders").Range("AE11").Formula = "='Control Panel'!E29"
Sheets("Conditional Orders").Range("AE12").Formula = "='Control Panel'!G29"
Sheets("Conditional Orders").Range("AE13").Formula = "='Control Panel'!I29"
Sheets("Conditional Orders").Range("AE14").Formula = "='Control Panel'!K29"
Sheets("Conditional Orders").Range("AE15").Formula = "='Control Panel'!M29"
Sheets("Conditional Orders").Range("AE16").Formula = "='Control Panel'!O29"
Sheets("Conditional Orders").Range("AE17").Formula = "='Control Panel'!Q29"
Sheets("Conditional Orders").Range("AE18").Formula = "='Control Panel'!S29"
Sheets("Conditional Orders").Range("AE19").Formula = "='Control Panel'!U29"
Sheets("Conditional Orders").Range("AE20").Formula = "='Control Panel'!W29"
Sheets("Conditional Orders").Range("AE21").Formula = "='Control Panel'!Y29"
Sheets("Conditional Orders").Range("AE22").Formula = "='Control Panel'!AA29"
Sheets("Conditional Orders").Range("AE23").Formula = "='Control Panel'!AC29"
Sheets("Conditional Orders").Range("AE24").Formula = "='Control Panel'!AE29"
Sheets("Conditional Orders").Range("AE25").Formula = "='Control Panel'!AG29"
Sheets("Conditional Orders").Range("AE26").Formula = "='Control Panel'!AI29"
Sheets("Conditional Orders").Range("AE27").Formula = "='Control Panel'!AK29"
Sheets("Conditional Orders").Range("AE28").Formula = "='Control Panel'!AM29"
Sheets("Conditional Orders").Range("AE29").Formula = "='Control Panel'!AO29"
Sheets("Conditional Orders").Range("AE30").Formula = "='Control Panel'!AQ29"
Sheets("Conditional Orders").Range("AE31").Formula = "='Control Panel'!AS29"
Sheets("Conditional Orders").Range("AE32").Formula = "='Control Panel'!AU29"
Sheets("Conditional Orders").Range("AE33").Formula = "='Control Panel'!AW29"
Sheets("Conditional Orders").Range("AE34").Formula = "='Control Panel'!AY29"
Sheets("Conditional Orders").Range("AE35").Formula = "='Control Panel'!C61"
Sheets("Conditional Orders").Range("AE36").Formula = "='Control Panel'!E61"
Sheets("Conditional Orders").Range("AE37").Formula = "='Control Panel'!G61"
Sheets("Conditional Orders").Range("AE38").Formula = "='Control Panel'!I61"
Sheets("Conditional Orders").Range("AE39").Formula = "='Control Panel'!K61"
Sheets("Conditional Orders").Range("AE40").Formula = "='Control Panel'!M61"
Sheets("Conditional Orders").Range("AE41").Formula = "='Control Panel'!O61"
Sheets("Conditional Orders").Range("AE42").Formula = "='Control Panel'!Q61"
Sheets("Conditional Orders").Range("AE43").Formula = "='Control Panel'!S61"
Sheets("Conditional Orders").Range("AE44").Formula = "='Control Panel'!U61"
Sheets("Conditional Orders").Range("AE45").Formula = "='Control Panel'!W61"
Sheets("Conditional Orders").Range("AE46").Formula = "='Control Panel'!Y61"
Sheets("Conditional Orders").Range("AE47").Formula = "='Control Panel'!AA61"
Sheets("Conditional Orders").Range("AE48").Formula = "='Control Panel'!AC61"
Sheets("Conditional Orders").Range("AE49").Formula = "='Control Panel'!AE61"
Sheets("Conditional Orders").Range("AE50").Formula = "='Control Panel'!AG61"
Sheets("Conditional Orders").Range("AE51").Formula = "='Control Panel'!AI61"
Sheets("Conditional Orders").Range("AE52").Formula = "='Control Panel'!AK61"
Sheets("Conditional Orders").Range("AE53").Formula = "='Control Panel'!AM61"
Sheets("Conditional Orders").Range("AE54").Formula = "='Control Panel'!AO61"
Sheets("Conditional Orders").Range("AE55").Formula = "='Control Panel'!AQ61"
Sheets("Conditional Orders").Range("AE56").Formula = "='Control Panel'!AS61"
Sheets("Conditional Orders").Range("AE57").Formula = "='Control Panel'!AU61"
Sheets("Conditional Orders").Range("AE58").Formula = "='Control Panel'!AW61"
Sheets("Conditional Orders").Range("AE59").Formula = "='Control Panel'!AY61"

Application.Calculation = xlAutomatic

End Sub



Sub ClearWorksheet()
'this module clears the "Control Panel" sheet of all user data and clears the links on the "Tickers" sheet and deletes all of the live recorded data on the "Ticks" sheet

If RecordOn = 1 Then
MsgBox "Please stop the system first!"
End If

If RecordOn = 0 Then

Dim Answer As String
Answer = MsgBox("Delete the data entries on this worksheet?", vbQuestion + vbYesNo, "Delete?")
If Answer = vbYes Then
Range("C5:Ay10").ClearContents
Range("C11:Ay17").ClearContents
Range("C24:Ay33").ClearContents
Range("C34:Ay42").ClearContents
Range("C43:Ay49").ClearContents
Range("C56:Ay65").ClearContents
Range("C66:Ay67").ClearContents
ResetConditionalOrders
Sheets("Conditional Orders").Range("M10:Y59").ClearContents
Sheets("Conditional Orders").Range("CH10:CH59").ClearContents
Sheets("Ticks").Range("B3:GS2220").ClearContents
Range("C15:Ay16").Interior.colorIndex = 2
Range("C47:Ay48").Interior.colorIndex = 2
Sheets("Tickers").Select
Sheets("Tickers").clearLinks
End If

End If

End Sub

Sub SilenceAlarms()
'this module runs when you hit the "Silence Alarms" button
AlarmOn = 0
MicroWarn = 0
Range("I3") = ""
Range("I3").Interior.colorIndex = 2
End Sub




'This module runs every 15 seconds
'It changes the colors of the cells in rows 15 and 47
'type b in those cells to look for a buy trade
'type s to look for a short trade
'type x to make the cell grayed out
'enter a number to start a countdown timer
Sub Alarm50()

    Dim DDDD As Integer
    On Error GoTo Errhandler:


    If MicroWarn <> 10 Then
    DDDD = 3
    marker:

    If IsNumeric(Cells(15, DDDD)) = True And Cells(15, DDDD) = 0 And Cells(15, DDDD) <> "" Then
        Cells(15, DDDD) = ""
        AlarmOn = 1
        Cells(15, DDDD).Interior.colorIndex = 3
    End If
    If (Cells(15, DDDD) = "b" Or Cells(15, DDDD) = "s" Or Cells(15, DDDD) = "e") And Cells(12, DDDD) <> "" And Cells(13, DDDD) <> "" And Cells(15, DDDD).Interior.colorIndex <> 7 Then
    Cells(15, DDDD).Interior.colorIndex = 4
    End If
    If Cells(15, DDDD) = "x" Then
    Cells(15, DDDD) = ""
    Cells(15, DDDD).Interior.colorIndex = 16
    End If



    If IsNumeric(Cells(47, DDDD)) = True And Cells(47, DDDD) = 0 And Cells(47, DDDD) <> "" Then
        Cells(47, DDDD) = ""
        AlarmOn = 1
        Cells(47, DDDD).Interior.colorIndex = 3
    End If
    If (Cells(47, DDDD) = "b" Or Cells(47, DDDD) = "s" Or Cells(47, DDDD) = "e") And Cells(44, DDDD) <> "" And Cells(45, DDDD) <> "" And Cells(47, DDDD).Interior.colorIndex <> 7 Then
    Cells(47, DDDD).Interior.colorIndex = 4
    End If
    If Cells(47, DDDD) = "x" Then
    Cells(47, DDDD) = ""
    Cells(47, DDDD).Interior.colorIndex = 16
    End If



    DDDD = DDDD + 2
    If DDDD < 52 Then
    GoTo marker:
    End If

    End If

    If Range("I3") <> "" Then
    AlarmOn = 1
    Range("I3").Interior.colorIndex = 3
    End If


    Errhandler:
End Sub


Sub Timers()
    Dim ZZZZ As Integer
    On Error GoTo Errhandler:

    If MicroWarn <> 10 Then
        ZZZZ = 3
        marker:

        If IsNumeric(Cells(15, ZZZZ)) = True And Cells(15, ZZZZ) <> "" Then
            Cells(15, ZZZZ) = Cells(15, ZZZZ) - 1
            Cells(15, ZZZZ).Interior.Color = 16760950
        End If

        If IsNumeric(Cells(47, ZZZZ)) = True And Cells(47, ZZZZ) <> "" Then
            Cells(47, ZZZZ) = Cells(47, ZZZZ) - 1
            Cells(47, ZZZZ).Interior.Color = 16760950
        End If


        ZZZZ = ZZZZ + 2
        If ZZZZ < 52 Then
            GoTo marker:
        End If

    End If

    Errhandler:
    Alarm50
End Sub


Sub Record2minTick()
'this module pastes the data recorded over the last 2 minutes into the sheet "Ticks"
'it enters the time, the high, the low, and the volume for each symbol
'volume is recorded in thousands of shares (1.5 is 1,500 shares)
'all premarket data is recorded as a single tick(the first one on the sheet)

Dim RRRR As Integer
Dim TTTT As Integer
Dim Ticx As Integer

On Error GoTo Errhandler

If Range("S2") = True Then
MicroWarn = 10

Ticx = 12
RRRR = 3
TTTT = 2
marker:


If Cells(5, RRRR) <> "" And Cells(20, RRRR) <> -1 Then
Sheets("Ticks").Cells(TicksRowA, TTTT) = Time()
Sheets("Ticks").Cells(TicksRowA, TTTT + 1) = Cells(31, RRRR)
Sheets("Ticks").Cells(TicksRowA, TTTT + 2) = Cells(32, RRRR)
'the volume equals the total volume so far minus the sum total volume of all previous ticks
Sheets("Ticks").Cells(TicksRowA, TTTT + 3) = Sheets("Tickers").Cells(Ticx, 32) / 10 - Cells(33, RRRR)
Sheets("Ticks").Cells(TicksRowA, TTTT + 3) = WorksheetFunction.Round(Sheets("Ticks").Cells(TicksRowA, TTTT + 3), 4)
    If EarlyTicks = 1 Then
    Cells(31, RRRR) = ""
    Cells(32, RRRR) = ""
    End If
Cells(33, RRRR) = Sheets("Tickers").Cells(Ticx, 32) / 10
End If


Ticx = Ticx + 25
If Cells(37, RRRR) <> "" And Cells(52, RRRR) <> -1 Then
Sheets("Ticks").Cells(TicksRowA, TTTT + 100) = Time()
Sheets("Ticks").Cells(TicksRowA, TTTT + 101) = Cells(63, RRRR)
Sheets("Ticks").Cells(TicksRowA, TTTT + 102) = Cells(64, RRRR)
Sheets("Ticks").Cells(TicksRowA, TTTT + 103) = Sheets("Tickers").Cells(Ticx, 32) / 10 - Cells(65, RRRR)
Sheets("Ticks").Cells(TicksRowA, TTTT + 103) = WorksheetFunction.Round(Sheets("Ticks").Cells(TicksRowA, TTTT + 103), 4)
    If EarlyTicks = 1 Then
    Cells(63, RRRR) = ""
    Cells(64, RRRR) = ""
    End If
Cells(65, RRRR) = Sheets("Tickers").Cells(Ticx, 32) / 10
End If
Ticx = Ticx - 25

Ticx = Ticx + 1
RRRR = RRRR + 2
TTTT = TTTT + 4
If RRRR < 52 Then
GoTo marker:
End If

If EarlyTicks = 1 Then
TicksRowA = TicksRowA + 1
End If

End If

Errhandler:
MicroWarn = 0



End Sub





Sub PositionExit()
'The button called "Force Exit" runs this code.
'It looks for the active cell you have selected and runs for the symbol closest to that cell
'If there is a partially filled order, it will tell the system to cancel it and give you a message box saying it was cancelled
'Sets the TRADE cell to "Timeout" which will tell the system to send an exit order on the bid or ask (see Timeout module)

Dim Actcolc As Integer
Actcolc = ActiveCell.Column
If ActiveCell.Row < 36 Then
Cells(24, ActiveCell.Column) = "Timeout"
Else
Cells(56, ActiveCell.Column) = "Timeout"
End If
Timeout


End Sub



