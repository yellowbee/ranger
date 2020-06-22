Attribute VB_Name = "Module1"
Public Gains As Double
Public RecordOn As Integer
Public MicroWarn As Integer
Public MicroWarn2 As Integer
Public MicroWarn3 As Integer
Public LogRow As Integer
Public AlarmOn As Integer
Public EarlyTicks As Integer

'the code below is needed to play sounds
Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long
Sub PlayWavFile(WavFileName As String, Wait As Boolean)
'this module plays the sounds
If Sheets("Control Panel").Range("AI2") <> True Then
    If Dir(WavFileName) = "" Then Exit Sub ' no file to play
        sndPlaySound WavFileName, 1
End If
End Sub

Sub Auto_Open()
'this module runs every time the workbook opens
'it selects the "Control Panel" sheet and clears the "Send? >" cells containing any "y" values to make sure duplicate trades are not sent accidentally
    Sheets("Control Panel").Select

    Sheets("Control Panel").Range("C25:AY25").ClearContents
    Sheets("Control Panel").Range("C57:AY57").ClearContents

End Sub

Sub TestPlayWavFile()
PlayWavFile "Z:\repos\ranger1.0\trade.wav", False
End Sub


Sub TestPlayWavFile4()
PlayWavFile "Z:\repos\ranger1.0\cash.wav", False
End Sub


Sub TestPlayWavFile6()
PlayWavFile "Z:\repos\ranger1.0\stopout.wav", False
End Sub


Sub TestPlayWavFile7()
PlayWavFile "Z:\repos\ranger1.0\exit.wav", False
End Sub



Sub TestPlayWavFile10()
PlayWavFile "C:\fart.wav", False
End Sub



Sub getback5()
'this module is part of the Set Up Workbook process
Sheets("Tickers").GetMarketDataN
End Sub


Sub getback9()
'this module is part of the Set Up Workbook process
Sheets("Control Panel").continue45
End Sub


'this module runs the timers
Sub Getback13()
    If RecordOn = 1 Then
    time789 = 60 - Second(Sheets("Control Panel").Range("I2"))
    Application.OnTime Format(DateAdd("s", time789, Now()), "hh:mm:ss"), "getback13"
    Sheets("Control Panel").Timers
    End If
End Sub


Sub getback18()
'this module plays the alarm sound every 15 seconds until you hit the "silence alarms" button
'it also calls the "highTimer" code on sheet19 (which also is designed to run every 15 seconds)
Sheets("Control Panel").HighTimer
If AlarmOn = 1 Then
TestPlayWavFile7
End If
End Sub


Sub getback19()
'this module sends the "activate system?" dialogue box at the end of the Set Up Workbook process
Dim Answer As String
Answer = MsgBox("Activate System?", vbQuestion + vbYesNo, "Start?")
If Answer = vbYes Then
Sheets("Control Panel").RecordDataOn
'Sheets("Control Panel").Select
MsgBox "System on!"
End If
End Sub

'a recursive sub for running a clock
'old name: getback20()
Sub runClock()
    'this module updates the clock every second
    If RecordOn = 1 Then
        Application.OnTime Now() + TimeValue("00:00:01"), "runClock"
        If MicroWarn = 0 Then
            Sheets("Control Panel").Range("I2") = Time()
        End If
    End If
End Sub

Sub getback21()
'this module automatically closes out any open postions before the market closes
Sheets("Control Panel").CloseTime
End Sub


'this module records the high and the low of each 2 minute tick after the market opens
'(starting at 8:30 in my time zone)
'if microwarn = 0, then there isn't any higher priority code running and recording
'takes place immediately
'if microwarn is not 0, that means order sending code is running at the same time that
'this function is being called and this module will wait 2 seconds until the order sending
'code finishes
Sub getback30()
    If RecordOn = 1 Then
        Application.OnTime Now() + TimeValue("00:02:00"), "getback30"

        If MicroWarn = 0 Then
             EarlyTicks = 1
             Sheets("Control Panel").Record2minTick
             EarlyTicks = 0
        Else
             getback31
        End If
    End If
End Sub

Sub getback31()
'this module is the second half of the getback30 code above it, which waits for 2 seconds,
'then runs as soon as microwarn = 0 (no longer any order sending code running)
If RecordOn = 1 Then
    If MicroWarn = 0 Then
    EarlyTicks = 1
    Sheets("Control Panel").Record2minTick
    EarlyTicks = 0
    Else
    Application.OnTime Now() + TimeValue("00:00:02"), "getback31"
    End If
End If
End Sub

'/*
' * This module turns the workbook on automatically at exactly 8:30:00 AM,
' * which is when the market opens in my time zone
' */
'old name: getback33()
Sub recordDataOnWrapper()
    If RecordOn = 1 Then
        Sheets("Control Panel").RecordDataOn
    End If
End Sub

Sub getBack34()
'When the price reaches the Price Target, this sends the exit order to take profits
'if there is order routing code already runnning when this module is triggered (microwarn = 10), it will wait for 2 seconds and try running again
If RecordOn = 1 Then
    If MicroWarn = 0 Then
    Sheets("Control Panel").TakeProfits
    Else
    Application.OnTime Now() + TimeValue("00:00:02"), "getback34"
    End If
End If
End Sub


'TimerMssg sub to be called by Sampletimer in sheet19
Sub TimerMssg
    MsgBox "Hello!"
End Sub






