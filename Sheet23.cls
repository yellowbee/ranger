VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub testMinMax()
Dim BuyRow As Integer

BuyRow = 32
MsgBox Application.WorksheetFunction.Max(Range(Cells(BuyRow - 29, 16), Cells(BuyRow, 17)))

End Sub



Sub BUYbacktest()


Dim SMA As Double
Dim BuyRow As Integer
Dim TradeRow As Integer
Dim OpenPosition As Integer

OpenPosition = 0
BuyRow = 3
TradeRow = 7


marker:

If BuyRow > 33 Then
    'entry code
    SMA = Application.Average(Range(Cells(BuyRow - 29, 16), Cells(BuyRow, 17)))
    If Cells(BuyRow, 16) > SMA * 1.01 And Cells(BuyRow, 17) < SMA * 1.01 And OpenPosition = 0 Then
        Cells(TradeRow, 7) = Cells(BuyRow, 14)
        Cells(TradeRow, 8) = SMA * 1.011
        Cells(TradeRow, 8) = WorksheetFunction.Round(Cells(TradeRow, 8).value, 2)
        OpenPosition = 1
    End If
    'exit code
    If Cells(BuyRow, 16) > SMA * 1.04 And OpenPosition = 1 Then
        Cells(TradeRow, 9) = SMA * 1.04
        Cells(TradeRow, 9) = WorksheetFunction.Round(Cells(TradeRow, 9).value, 2)
        Cells(TradeRow, 11) = (Cells(TradeRow, 9) - Cells(TradeRow, 8)) / Cells(TradeRow, 8)
        Cells(TradeRow, 10) = Cells(BuyRow, 14)
        TradeRow = TradeRow + 1
        OpenPosition = 0
    End If
    'stopout code:
    If Cells(BuyRow, 17) < Cells(TradeRow, 8) * 0.96 And OpenPosition = 1 Then
        Cells(TradeRow, 9) = Cells(TradeRow, 8) * 0.959
        Cells(TradeRow, 9) = WorksheetFunction.Round(Cells(TradeRow, 9).value, 2)
        Cells(TradeRow, 10) = Cells(BuyRow, 14)
            'if the opening price gaps down less than the stop, then assume the trade stops out at the opening price
            If Cells(BuyRow, 15) < Cells(TradeRow, 8) * 0.96 Then
            Cells(TradeRow, 9) = Cells(BuyRow, 15)
            End If
        Cells(TradeRow, 11) = (Cells(TradeRow, 9) - Cells(TradeRow, 8)) / Cells(TradeRow, 8)
        TradeRow = TradeRow + 1
        OpenPosition = 0
    End If
    
End If

BuyRow = BuyRow + 1
If Cells(BuyRow, 14) <> "" Then
GoTo marker:
End If


End Sub





Sub SHORTbacktest()


Dim SMA As Double
Dim BuyRow As Integer
Dim TradeRow As Integer
Dim OpenPosition As Integer

OpenPosition = 0
BuyRow = 3
TradeRow = 7
marker:

If BuyRow > 33 Then
    'entry code
    SMA = Application.Average(Range(Cells(BuyRow - 29, 16), Cells(BuyRow, 17)))
    If Cells(BuyRow, 17) < SMA * 0.99 And Cells(BuyRow, 16) > SMA * 0.99 And OpenPosition = 0 Then
        Cells(TradeRow, 1) = Cells(BuyRow, 14)
        Cells(TradeRow, 2) = SMA * 0.989
        Cells(TradeRow, 2) = WorksheetFunction.Round(Cells(TradeRow, 2).value, 2)
        OpenPosition = 1
    End If
    'exit code
    If Cells(BuyRow, 17) < SMA * 0.96 And OpenPosition = 1 Then
        Cells(TradeRow, 3) = SMA * 0.96
        Cells(TradeRow, 3) = WorksheetFunction.Round(Cells(TradeRow, 3).value, 2)
        Cells(TradeRow, 5) = (Cells(TradeRow, 2) - Cells(TradeRow, 3)) / Cells(TradeRow, 2)
        Cells(TradeRow, 4) = Cells(BuyRow, 14)
        TradeRow = TradeRow + 1
        OpenPosition = 0
    End If
    'stopout code:
    If Cells(BuyRow, 17) > Cells(TradeRow, 2) * 1.04 And OpenPosition = 1 Then
        Cells(TradeRow, 3) = Cells(TradeRow, 2) * 1.041
        Cells(TradeRow, 3) = WorksheetFunction.Round(Cells(TradeRow, 3).value, 2)
        Cells(TradeRow, 4) = Cells(BuyRow, 14)
            'if the opening price gaps up more than the stop, then assume the trade stops out at the opening price
            If Cells(BuyRow, 15) > Cells(TradeRow, 2) * 1.05 Then
            Cells(TradeRow, 3) = Cells(BuyRow, 15)
            End If
        Cells(TradeRow, 5) = (Cells(TradeRow, 2) - Cells(TradeRow, 3)) / Cells(TradeRow, 2)
        TradeRow = TradeRow + 1
        OpenPosition = 0
    End If
    
End If

BuyRow = BuyRow + 1
If Cells(BuyRow, 14) <> "" Then
GoTo marker:
End If


End Sub
