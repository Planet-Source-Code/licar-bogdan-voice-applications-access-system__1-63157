Attribute VB_Name = "VoiceAccess"
'Module recycled and modified from my Voice Recognition program:
'http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=62860&lngWId=1
'
'Author:      Licar Bogdan (copyright).
'
'Description: Functions for comparing sounds. It bases on heuristic-statistic
'             methods. Perhaps there are other more efficient methods to to this.
'             I've included 2 other comparing functions, without using them in
'             this project.

Option Explicit

'Variables used in more routines
Public Waves() As String, CommandsFld As String

Sub LoadWave(ByVal Path As String, values() As Double, Y() As Double)
Dim j As Long, Buff As Long, Yrate As Double

Dim Min As Double, Max As Double


'Not using other variables just for lack of will:
'values(0)=the startpoint of the important part
'values(1)=the endpoint        "    "
'values(3)=number of low peaks
'values(4)=number of high peaks

On Error Resume Next
        
        'Reads it and registers all values in an array
        j = 44 'Set i To 44, since the wave sample begins at Byte 44.
        Open Path For Random As #1
        Do
            Get #1, j, Buff
            j = j + 1: ReDim Preserve values(j)
            values(j) = Buff
            If Buff > Max Then Max = Buff
            If Buff < Min Then Min = Buff
        Loop Until EOF(1)
        Close #1

        'Change the values of the *.wav in cartesian coordinates
        Yrate = (Max - Min) / (500)
        For j = 44 To UBound(values)
            ReDim Preserve Y(j - 43)
            Y(j - 43) = (values(j) / Yrate)
            
            'Count peaks above/below a certain constant
            If Y(j - 43) > (500 / 3.5) Then values(4) = values(4) + 1
            If Y(j - 43) < (-500 / 3.5) Then values(3) = values(3) + 1
        Next j

            'Loops for isolating the important part of the wave file
            For j = 1 To UBound(Y) - 1          'The beginning of the "talking" part
                If (Abs(Y(j) - Y(j + 1))) > 100 Then values(0) = j: Exit For
            Next j
            For j = UBound(Y) To 1 Step -1      'The end of the "talking"
                If (Abs(Y(j) - Y(j - 1))) > 100 Then values(1) = j: Exit For
            Next j

End Sub

Function SearchCommand(ByVal Path As String) As String
Dim values1() As Double, values2() As Double, Y1() As Double, Y2() As Double
Dim i As Long, MatchLevel() As Integer, MaxMatchLevel As Integer
    
    GetWaves CommandsFld, Waves
    LoadWave Path, values2, Y2
    ReDim MatchLevel(UBound(Waves))

    For i = 1 To UBound(Waves)
        LoadWave CommandsFld & Waves(i), values1, Y1
        
        If AccessGranted(Y1, Y2, values1(0), values1(1), values2(0), _
        values2(1), values1(4), values1(3), values2(4), values2(3), MatchLevel(i)) = True Then SearchCommand = CommandsFld & Waves(i)
        
        'In case more sounds match, it searches for the sound with the greatest
        'MatchLevel. If you find other solutions to this possible problem, let me know.
        If MatchLevel(i) > MaxMatchLevel Then MaxMatchLevel = i
    Next i

    If Waves(MaxMatchLevel) = "" Then Exit Function
SearchCommand = CommandsFld & Waves(MaxMatchLevel)
End Function

Function AccessGranted(Y1() As Double, Y2() As Double, ByVal StartPoint1 As Long, _
ByVal Endpoint1 As Long, ByVal StartPoint2 As Long, ByVal EndPoint2 As Long, ByVal HighPeaks1, _
ByVal LowPeaks1, ByVal HighPeaks2, ByVal Lowpeaks2, MatchLevel As Integer) As Boolean

Dim i As Integer

'Result based on the number of high/low peaks and on statistics of the whole "important part".
'If you find a better way to verify waves, please let me know.

'On Error Resume Next

'It considers match levels (i.e. a more/less compatible sound than another). More
'predefined commands could match with the recorded sound, so it verifies which command
'is more similar to the recorded sound.
AccessGranted = True
Do While AccessGranted = True
    i = i + 1

    If Abs(UBound(Y1) - UBound(Y2)) > 250 Then AccessGranted = False: Exit Function
    If (Abs(HighPeaks1 - HighPeaks2) <= 20 - i) And (Abs(LowPeaks1 - Lowpeaks2) <= 20 - i) And _
    (Abs(ArithmeticMean(Y1, StartPoint1, Endpoint1) - ArithmeticMean(Y2, StartPoint2, EndPoint2)) < 8 - i) And _
    (Abs(StandardDeviation(Y1, StartPoint1, Endpoint1) - StandardDeviation(Y2, StartPoint2, EndPoint2)) < 20 - i) Then

        MatchLevel = i
        AccessGranted = True
    Else

    AccessGranted = False
    End If
Loop

End Function

Function ArithmeticMean(vals() As Double, ByVal StartPoint As Long, ByVal EndPoint As Long) As Single
Dim i As Long, result As Single
On Error Resume Next
    For i = StartPoint To EndPoint
        result = result + vals(i)
    Next i
ArithmeticMean = result / (EndPoint - StartPoint)
End Function

Function StandardDeviation(vals() As Double, ByVal StartPoint As Long, ByVal EndPoint As Long) As Double
Dim i As Long, Am As Single, result As Double
On Error Resume Next
Am = ArithmeticMean(vals(), StartPoint, EndPoint)
    For i = StartPoint To EndPoint
        result = result + ((vals(i) - Am) ^ 2)
    Next i
StandardDeviation = Format(Sqr(result / (EndPoint - StartPoint)), "#.##")
End Function

'-------------------------------------------------------------------------------
'Others

Sub GetWaves(ByVal Path As String, Files() As String)
Dim result() As String, filename As String, count As Long
        
        'Returns an array with all the wave files in a folder
        filename = Dir$(Path)
        Do While Len(filename)
            If LCase(Right$(filename, 3)) = "wav" Then
            count = count + 1
            ReDim Preserve result(count)
            result(count) = filename
            End If
            filename = Dir$
        Loop
Files = result
End Sub

Function GetFileName(ByVal Path As String, Optional Extension As Boolean = True) As String
Dim i As Integer, str As String, iStart As Integer
    
    If Extension = True Then
        iStart = 1
    ElseIf Extension = False Then
        iStart = Len(Path) - InStr(1, Path, ".") + 2
    End If
    
    For i = iStart To Len(Path)
        str = str & Mid$(StrReverse(Path), i, 1)
        If Mid$(StrReverse(Path), i + 1, 1) = "\" Then str = StrReverse(str): Exit For
    Next i

GetFileName = str
End Function

'-----------------------------------------------------------------------------------

'If you're not satisfied of the result with AccessGranted, try these 2 functions.
'I suggest the Statistic_Comparison, since is much more reliable.

Function PointToPoint_Comparison(vals1() As Double, vals2() As Double, _
ByVal StartPoint1 As Long, ByVal Endpoint1 As Long, ByVal StartPoint2 _
As Long, ByVal EndPoint2 As Long) As Double

Dim j As Long, i As Long, Same As Long, ErrRange As Integer
On Error Resume Next
ErrRange = 10
'Compares each value of the default sound with the near values of the sound used
'for the matching process. It leaves a small range of error. With a greater number
'than 10, there are more chances the sounds to match, but they could also be different,
'giving a high percentage of matching.
'This is a not so highly efficient method.

For j = 1 To (Endpoint1 - StartPoint1)
    If j = EndPoint2 Then Exit For

    For i = -2 To 2
        If (Abs(vals1(StartPoint1 + j) - vals2(StartPoint2 + j + i))) < ErrRange Then Same = Same + 1: Exit For
    Next i

Next j

PointToPoint_Comparison = Format((Same * 100) / (Endpoint1 - StartPoint1), "#.##")
End Function

Function Statistic_Comparison(vals1() As Double, vals2() As Double, ByVal StartPoint1 _
As Long, ByVal Endpoint1 As Long, ByVal StartPoint2 As Long, ByVal EndPoint2 As Long) As Double

Dim i As Long, j As Long, Same2 As Long, ErrRange As Integer, v1() As Double, v2() As Double, ArrSize As Integer
ArrSize = 20
ReDim v1(ArrSize): ReDim v2(ArrSize)
On Error Resume Next

'This could be done also by dividing the wave in totally separate parts and analyze them
For j = 1 To (Endpoint1 - StartPoint1)
    If (j + StartPoint2) > EndPoint2 Then Exit For
    
    For i = 1 To ArrSize
        v1(i) = vals1(StartPoint1 + j + i): v2(i) = vals2(StartPoint2 + j + i)
    Next i
    
    If (Abs(ArithmeticMean(v1, LBound(v1), UBound(v1)) - _
        ArithmeticMean(v2, LBound(v2), UBound(v2)))) < 5 And _
        (Abs(StandardDeviation(v1, LBound(v1), UBound(v1)) - _
        StandardDeviation(v2, LBound(v2), UBound(v2)))) < 15 Then Same2 = Same2 + 1
    
Next j
'I think it's better than the point-to-point technique.
Statistic_Comparison = Format((Same2 * 100) / Round(Endpoint1 - StartPoint1), "#.##")
End Function
'---------------------------------------------------------------------------------
'School sucks.
