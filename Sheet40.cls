VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet40"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True



Sub Logworksheet()

marker:
LogRow = LogRow + 8

If Sheets("Log").Cells(LogRow, 2) <> "" Then
Gains = Gains + WorksheetFunction.Sum(Sheets("Log").Range(Cells(LogRow + 6, 2), Cells(LogRow + 6, 50)))
GoTo marker
End If
Sheets("Control Panel").Range("C5:AY10").Copy Destination:=Sheets("Log").Cells(LogRow, 2)
Sheets("Control Panel").Range("C17:AY17").Copy Destination:=Sheets("Log").Cells(LogRow + 6, 2)
Sheets("Control Panel").Range("E2").Copy Destination:=Sheets("Log").Cells(LogRow, 1)
Gains = Gains + WorksheetFunction.Sum(Sheets("Log").Range(Cells(LogRow + 6, 2), Cells(LogRow + 6, 50)))




Range(Cells(LogRow + 5, 2), Cells(LogRow + 5, 50)).Select
Selection.Copy
'Range(LogRow + 5, 2).Select
Selection.PasteSpecial xlValues
If Cells(LogRow + 6, 2) > 0 Then
Cells(LogRow + 6, 2).Style = "good"
End If
If Cells(LogRow + 6, 2) < 0 Then
Cells(LogRow + 6, 2).Style = "bad"
End If
If Cells(LogRow + 6, 4) > 0 Then
Cells(LogRow + 6, 4).Style = "good"
End If
If Cells(LogRow + 6, 4) < 0 Then
Cells(LogRow + 6, 4).Style = "bad"
End If
If Cells(LogRow + 6, 6) > 0 Then
Cells(LogRow + 6, 6).Style = "good"
End If
If Cells(LogRow + 6, 6) < 0 Then
Cells(LogRow + 6, 6).Style = "bad"
End If
If Cells(LogRow + 6, 8) > 0 Then
Cells(LogRow + 6, 8).Style = "good"
End If
If Cells(LogRow + 6, 8) < 0 Then
Cells(LogRow + 6, 8).Style = "bad"
End If
If Cells(LogRow + 6, 10) > 0 Then
Cells(LogRow + 6, 10).Style = "good"
End If
If Cells(LogRow + 6, 10) < 0 Then
Cells(LogRow + 6, 10).Style = "bad"
End If
If Cells(LogRow + 6, 12) > 0 Then
Cells(LogRow + 6, 12).Style = "good"
End If
If Cells(LogRow + 6, 12) < 0 Then
Cells(LogRow + 6, 12).Style = "bad"
End If
If Cells(LogRow + 6, 14) > 0 Then
Cells(LogRow + 6, 14).Style = "good"
End If
If Cells(LogRow + 6, 14) < 0 Then
Cells(LogRow + 6, 14).Style = "bad"
End If
If Cells(LogRow + 6, 16) > 0 Then
Cells(LogRow + 6, 16).Style = "good"
End If
If Cells(LogRow + 6, 16) < 0 Then
Cells(LogRow + 6, 16).Style = "bad"
End If
If Cells(LogRow + 6, 18) > 0 Then
Cells(LogRow + 6, 18).Style = "good"
End If
If Cells(LogRow + 6, 18) < 0 Then
Cells(LogRow + 6, 18).Style = "bad"
End If
If Cells(LogRow + 6, 20) > 0 Then
Cells(LogRow + 6, 20).Style = "good"
End If
If Cells(LogRow + 6, 20) < 0 Then
Cells(LogRow + 6, 20).Style = "bad"
End If
If Cells(LogRow + 6, 22) > 0 Then
Cells(LogRow + 6, 22).Style = "good"
End If
If Cells(LogRow + 6, 22) < 0 Then
Cells(LogRow + 6, 22).Style = "bad"
End If
If Cells(LogRow + 6, 24) > 0 Then
Cells(LogRow + 6, 24).Style = "good"
End If
If Cells(LogRow + 6, 24) < 0 Then
Cells(LogRow + 6, 24).Style = "bad"
End If
If Cells(LogRow + 6, 26) > 0 Then
Cells(LogRow + 6, 26).Style = "good"
End If
If Cells(LogRow + 6, 26) < 0 Then
Cells(LogRow + 6, 26).Style = "bad"
End If
If Cells(LogRow + 6, 28) > 0 Then
Cells(LogRow + 6, 28).Style = "good"
End If
If Cells(LogRow + 6, 28) < 0 Then
Cells(LogRow + 6, 28).Style = "bad"
End If
If Cells(LogRow + 6, 30) > 0 Then
Cells(LogRow + 6, 30).Style = "good"
End If
If Cells(LogRow + 6, 30) < 0 Then
Cells(LogRow + 6, 30).Style = "bad"
End If
If Cells(LogRow + 6, 32) > 0 Then
Cells(LogRow + 6, 32).Style = "good"
End If
If Cells(LogRow + 6, 32) < 0 Then
Cells(LogRow + 6, 32).Style = "bad"
End If
If Cells(LogRow + 6, 34) > 0 Then
Cells(LogRow + 6, 34).Style = "good"
End If
If Cells(LogRow + 6, 34) < 0 Then
Cells(LogRow + 6, 34).Style = "bad"
End If
If Cells(LogRow + 6, 36) > 0 Then
Cells(LogRow + 6, 36).Style = "good"
End If
If Cells(LogRow + 6, 36) < 0 Then
Cells(LogRow + 6, 36).Style = "bad"
End If
If Cells(LogRow + 6, 38) > 0 Then
Cells(LogRow + 6, 38).Style = "good"
End If
If Cells(LogRow + 6, 38) < 0 Then
Cells(LogRow + 6, 38).Style = "bad"
End If
If Cells(LogRow + 6, 40) > 0 Then
Cells(LogRow + 6, 40).Style = "good"
End If
If Cells(LogRow + 6, 40) < 0 Then
Cells(LogRow + 6, 40).Style = "bad"
End If
If Cells(LogRow + 6, 42) > 0 Then
Cells(LogRow + 6, 42).Style = "good"
End If
If Cells(LogRow + 6, 42) < 0 Then
Cells(LogRow + 6, 42).Style = "bad"
End If
If Cells(LogRow + 6, 44) > 0 Then
Cells(LogRow + 6, 44).Style = "good"
End If
If Cells(LogRow + 6, 44) < 0 Then
Cells(LogRow + 6, 44).Style = "bad"
End If
If Cells(LogRow + 6, 46) > 0 Then
Cells(LogRow + 6, 46).Style = "good"
End If
If Cells(LogRow + 6, 46) < 0 Then
Cells(LogRow + 6, 46).Style = "bad"
End If
If Cells(LogRow + 6, 48) > 0 Then
Cells(LogRow + 6, 48).Style = "good"
End If
If Cells(LogRow + 6, 48) < 0 Then
Cells(LogRow + 6, 48).Style = "bad"
End If
If Cells(LogRow + 6, 50) > 0 Then
Cells(LogRow + 6, 50).Style = "good"
End If
If Cells(LogRow + 6, 50) < 0 Then
Cells(LogRow + 6, 50).Style = "bad"
End If

Range("A3").Select
Selection.ClearContents
Range("A3") = Gains

If Sheets("Control Panel").Range("C37") <> "" Then
RepeatLog
End If

End Sub

Sub RepeatLog()
marker:
LogRow = LogRow + 8

If Sheets("Log").Cells(LogRow, 2) <> "" Then
Gains = Gains + WorksheetFunction.Sum(Sheets("Log").Range(Cells(LogRow + 6, 2), Cells(LogRow + 6, 50)))
GoTo marker
End If
Sheets("Control Panel").Range("C37:AY42").Copy Destination:=Sheets("Log").Cells(LogRow, 2)
Sheets("Control Panel").Range("C49:AY49").Copy Destination:=Sheets("Log").Cells(LogRow + 6, 2)
Sheets("Control Panel").Range("E2").Copy Destination:=Sheets("Log").Cells(LogRow, 1)
Gains = Gains + WorksheetFunction.Sum(Sheets("Log").Range(Cells(LogRow + 6, 2), Cells(LogRow + 6, 50)))




Range(Cells(LogRow + 5, 2), Cells(LogRow + 5, 50)).Select
Selection.Copy
'Range(LogRow + 5, 2).Select
Selection.PasteSpecial xlValues
If Cells(LogRow + 6, 2) > 0 Then
Cells(LogRow + 6, 2).Style = "good"
End If
If Cells(LogRow + 6, 2) < 0 Then
Cells(LogRow + 6, 2).Style = "bad"
End If
If Cells(LogRow + 6, 4) > 0 Then
Cells(LogRow + 6, 4).Style = "good"
End If
If Cells(LogRow + 6, 4) < 0 Then
Cells(LogRow + 6, 4).Style = "bad"
End If
If Cells(LogRow + 6, 6) > 0 Then
Cells(LogRow + 6, 6).Style = "good"
End If
If Cells(LogRow + 6, 6) < 0 Then
Cells(LogRow + 6, 6).Style = "bad"
End If
If Cells(LogRow + 6, 8) > 0 Then
Cells(LogRow + 6, 8).Style = "good"
End If
If Cells(LogRow + 6, 8) < 0 Then
Cells(LogRow + 6, 8).Style = "bad"
End If
If Cells(LogRow + 6, 10) > 0 Then
Cells(LogRow + 6, 10).Style = "good"
End If
If Cells(LogRow + 6, 10) < 0 Then
Cells(LogRow + 6, 10).Style = "bad"
End If
If Cells(LogRow + 6, 12) > 0 Then
Cells(LogRow + 6, 12).Style = "good"
End If
If Cells(LogRow + 6, 12) < 0 Then
Cells(LogRow + 6, 12).Style = "bad"
End If
If Cells(LogRow + 6, 14) > 0 Then
Cells(LogRow + 6, 14).Style = "good"
End If
If Cells(LogRow + 6, 14) < 0 Then
Cells(LogRow + 6, 14).Style = "bad"
End If
If Cells(LogRow + 6, 16) > 0 Then
Cells(LogRow + 6, 16).Style = "good"
End If
If Cells(LogRow + 6, 16) < 0 Then
Cells(LogRow + 6, 16).Style = "bad"
End If
If Cells(LogRow + 6, 18) > 0 Then
Cells(LogRow + 6, 18).Style = "good"
End If
If Cells(LogRow + 6, 18) < 0 Then
Cells(LogRow + 6, 18).Style = "bad"
End If
If Cells(LogRow + 6, 20) > 0 Then
Cells(LogRow + 6, 20).Style = "good"
End If
If Cells(LogRow + 6, 20) < 0 Then
Cells(LogRow + 6, 20).Style = "bad"
End If
If Cells(LogRow + 6, 22) > 0 Then
Cells(LogRow + 6, 22).Style = "good"
End If
If Cells(LogRow + 6, 22) < 0 Then
Cells(LogRow + 6, 22).Style = "bad"
End If
If Cells(LogRow + 6, 24) > 0 Then
Cells(LogRow + 6, 24).Style = "good"
End If
If Cells(LogRow + 6, 24) < 0 Then
Cells(LogRow + 6, 24).Style = "bad"
End If
If Cells(LogRow + 6, 26) > 0 Then
Cells(LogRow + 6, 26).Style = "good"
End If
If Cells(LogRow + 6, 26) < 0 Then
Cells(LogRow + 6, 26).Style = "bad"
End If
If Cells(LogRow + 6, 28) > 0 Then
Cells(LogRow + 6, 28).Style = "good"
End If
If Cells(LogRow + 6, 28) < 0 Then
Cells(LogRow + 6, 28).Style = "bad"
End If
If Cells(LogRow + 6, 30) > 0 Then
Cells(LogRow + 6, 30).Style = "good"
End If
If Cells(LogRow + 6, 30) < 0 Then
Cells(LogRow + 6, 30).Style = "bad"
End If
If Cells(LogRow + 6, 32) > 0 Then
Cells(LogRow + 6, 32).Style = "good"
End If
If Cells(LogRow + 6, 32) < 0 Then
Cells(LogRow + 6, 32).Style = "bad"
End If
If Cells(LogRow + 6, 34) > 0 Then
Cells(LogRow + 6, 34).Style = "good"
End If
If Cells(LogRow + 6, 34) < 0 Then
Cells(LogRow + 6, 34).Style = "bad"
End If
If Cells(LogRow + 6, 36) > 0 Then
Cells(LogRow + 6, 36).Style = "good"
End If
If Cells(LogRow + 6, 36) < 0 Then
Cells(LogRow + 6, 36).Style = "bad"
End If
If Cells(LogRow + 6, 38) > 0 Then
Cells(LogRow + 6, 38).Style = "good"
End If
If Cells(LogRow + 6, 38) < 0 Then
Cells(LogRow + 6, 38).Style = "bad"
End If
If Cells(LogRow + 6, 40) > 0 Then
Cells(LogRow + 6, 40).Style = "good"
End If
If Cells(LogRow + 6, 40) < 0 Then
Cells(LogRow + 6, 40).Style = "bad"
End If
If Cells(LogRow + 6, 42) > 0 Then
Cells(LogRow + 6, 42).Style = "good"
End If
If Cells(LogRow + 6, 42) < 0 Then
Cells(LogRow + 6, 42).Style = "bad"
End If
If Cells(LogRow + 6, 44) > 0 Then
Cells(LogRow + 6, 44).Style = "good"
End If
If Cells(LogRow + 6, 44) < 0 Then
Cells(LogRow + 6, 44).Style = "bad"
End If
If Cells(LogRow + 6, 46) > 0 Then
Cells(LogRow + 6, 46).Style = "good"
End If
If Cells(LogRow + 6, 46) < 0 Then
Cells(LogRow + 6, 46).Style = "bad"
End If
If Cells(LogRow + 6, 48) > 0 Then
Cells(LogRow + 6, 48).Style = "good"
End If
If Cells(LogRow + 6, 48) < 0 Then
Cells(LogRow + 6, 48).Style = "bad"
End If
If Cells(LogRow + 6, 50) > 0 Then
Cells(LogRow + 6, 50).Style = "good"
End If
If Cells(LogRow + 6, 50) < 0 Then
Cells(LogRow + 6, 50).Style = "bad"
End If

Range("A3").Select
Selection.ClearContents
Range("A3") = Gains

End Sub

