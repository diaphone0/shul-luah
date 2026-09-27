Attribute VB_Name = "mod_TMStruct"
Public Type TMStruct
    tm_sec As Long
    tm_min As Long
    tm_hour As Long
    tm_mday As Long
    tm_mon As Long
    tm_year As Long
    tm_wday As Long
    tm_yday As Long
    tm_isdst As Long
End Type

Private refDate As Date

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wday As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Public Enum TIME_ZONE
    TIME_ZONE_ID_INVALID = 0
    TIME_ZONE_STANDARD = 1
    TIME_ZONE_DAYLIGHT = 2
End Enum
#If VBA7 Then
Private Declare PtrSafe Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare PtrSafe Function GetTimeZoneInformationForYear Lib "kernel32" ( _
    ByVal wYear As Long, _
    ByVal pdtzi As Any, _
    ByRef ptzi As TIME_ZONE_INFORMATION) As Long
#Else
Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Function GetTimeZoneInformationForYear Lib "kernel32" ( _
    ByVal wYear As Long, _
    ByVal pdtzi As Any, _
    ByRef ptzi As TIME_ZONE_INFORMATION) As Long
#End If
Function IsDaylightSavingTime(dt As Date) As Long
    Dim tzi As TIME_ZONE_INFORMATION
    Dim result As Long
    On Error GoTo NoYearFunc
    
    result = GetTimeZoneInformationForYear(year(dt), ByVal 0&, tzi)
    GoTo proceed
NoYearFunc:
    result = GetTimeZoneInformation(tzi)
proceed:
    If result <> 0 Then ' TIME_ZONE_ID_DAYLIGHT
        ' Determine if the given date is within daylight saving time
        Dim dstStartDate As Date
        Dim dstEndDate As Date
        
        dstStartDate = DateSerial(year(dt), tzi.DaylightDate.wMonth, tzi.DaylightDate.wday)
        dstEndDate = DateSerial(year(dt), tzi.StandardDate.wMonth, tzi.StandardDate.wday)
        
        If dt >= dstStartDate And dt < dstEndDate Then
            IsDaylightSavingTime = 1
        End If
    End If
End Function

Function IsDST(date_in As Date) As Long
    IsDST = IsDaylightSavingTime(date_in)
End Function

Public Function mktm(ByVal date_in As Date) As TMStruct
    Dim result As TMStruct
    
    With result
    .tm_sec = second(date_in)
    .tm_min = minute(date_in)
    .tm_hour = hour(date_in)
    .tm_mday = day(date_in)
    .tm_mon = month(date_in) - 1
    .tm_year = year(date_in) - 1900
    .tm_wday = Weekday(date_in)
    .tm_yday = DateDiff("d", jan1, DateSerial(result.tm_year, result.tm_mon, result.tm_mday)) + 1
    .tm_isdst = IsDST(date_in)
    End With
    mktm = result
End Function

Public Function mktime(ByRef tm_in As TMStruct) As Date
    mktime = mkdate(tm_in, True)
End Function

Public Function mkdate(ByRef tm_in As TMStruct, Optional time_only As Boolean = False) As Date
    ' Trimming values and calculating remainders
    Dim offsetMonth As Double
    Dim offsetDay As Double
    Dim offsetHour As Double
    Dim offsetMinute As Double
    Dim offsetSecond As Double
    If time_only = False Then
        tm_in.tm_mon = tm_in.tm_mon + 1
        tm_in.tm_year = tm_in.tm_year + 1900
        ' Ensure the month is within 1 to 12 range
        If tm_in.tm_mon < 1 Or tm_in.tm_mon > 12 Then offsetMonth = tm_in.tm_mon - 1
        tm_in.tm_mon = IIf(tm_in.tm_mon < 1, 1, IIf(tm_in.tm_mon > 12, 12, tm_in.tm_mon))
    
        ' Ensure the day of the month is within valid range
        Dim lastDayOfMonth As Date
        lastDayOfMonth = DateAdd("d", -1, DateSerial(tm_in.tm_year, tm_in.tm_mon + 1, 1))
        If tm_in.tm_mday < 1 Or tm_in.tm_mday > day(lastDayOfMonth) Then offsetDay = tm_in.tm_mday - 1
        tm_in.tm_mday = IIf(tm_in.tm_mday < 1, 1, IIf(tm_in.tm_mday > day(lastDayOfMonth), day(lastDayOfMonth), tm_in.tm_mday))
    End If
    
    ' Normalize time elements
    If tm_in.tm_hour < 0 Or tm_in.tm_hour > 23 Then offsetHour = tm_in.tm_hour
    tm_in.tm_hour = IIf(tm_in.tm_hour < 0, 0, IIf(tm_in.tm_hour > 23, 23, tm_in.tm_hour))
    If tm_in.tm_min < 0 Or tm_in.tm_min > 59 Then offsetMinute = tm_in.tm_min
    tm_in.tm_min = IIf(tm_in.tm_min < 0, 0, IIf(tm_in.tm_min > 59, 59, tm_in.tm_min))
    If tm_in.tm_sec < 0 Or tm_in.tm_sec > 59 Then offsetSecond = tm_in.tm_sec
    tm_in.tm_sec = IIf(tm_in.tm_sec < 0, 0, IIf(tm_in.tm_sec > 59, 59, tm_in.tm_sec))
    
    refDate = TimeSerial(tm_in.tm_hour, tm_in.tm_min, tm_in.tm_sec)
    If time_only = False Then refDate = refDate + DateSerial(tm_in.tm_year, tm_in.tm_mon, tm_in.tm_mday)
    
    ' Applying offset
    Call ApplyOffset(0, offsetMonth, offsetDay, offsetHour, offsetMinute, offsetSecond, refDate)
    
    ' Calculate the day of the week (0 = Sunday, 1 = Monday, etc.)
    If time_only = False Then tm_in.tm_wday = Weekday(DateSerial(tm_in.tm_year, tm_in.tm_mon, tm_in.tm_mday))
    
    ' Calculate the day of the year
    Dim jan1 As Date
    If time_only = False Then
        jan1 = DateSerial(tm_in.tm_year, 1, 1)
        tm_in.tm_yday = DateDiff("d", jan1, DateSerial(tm_in.tm_year, tm_in.tm_mon, tm_in.tm_mday)) + 1
    
        tm_in.tm_isdst = IsDST(refDate)
        tm_in.tm_mon = tm_in.tm_mon - 1
        tm_in.tm_year = tm_in.tm_year - 1900
    End If
    mkdate = refDate
End Function

Private Sub ApplyOffset(yearOffset As Double, monthOffset As Double, dayOffset As Double, _
                       hourOffset As Double, minuteOffset As Double, secondOffset As Double, _
                       ByRef dateValue As Date)
    dateValue = DateAdd("yyyy", yearOffset, dateValue)
    dateValue = DateAdd("m", monthOffset, dateValue)
    dateValue = DateAdd("d", dayOffset, dateValue)
    dateValue = DateAdd("h", hourOffset, dateValue)
    dateValue = DateAdd("n", minuteOffset, dateValue)
    dateValue = DateAdd("s", secondOffset, dateValue)
End Sub

Sub test_tm()
Dim tm As TMStruct
Dim dt As Date

With tm
    .tm_hour = 11 '#6/15/2023 3:30:00 'Fri Apr 22 11:53:36 2016
    .tm_min = 53
    .tm_sec = 36
    .tm_year = 2016 - 1900
    .tm_mon = 4 - 1
    .tm_mday = 22
    End With
dt = mkdate(tm)

Debug.Print dt
Debug.Print "is dst: " + CStr(tm.tm_isdst) + "wday: " + CStr(tm.tm_wday) + "yday: " + CStr(tm.tm_yday)

tm.tm_mon = tm.tm_mon - 100
tm.tm_mday = tm.tm_mday - 22 - 31
dt = mkdate(tm)

Debug.Print dt
Debug.Print "is dst: " + CStr(tm.tm_isdst) + "wday: " + CStr(tm.tm_wday) + "yday: " + CStr(tm.tm_yday)



End Sub


Sub TestIsDST()
    Dim date1 As Date
    Dim date2 As Date
    Dim date3 As Date
    Dim date4 As Date
    
    ' Case 1: Date is in the daylight saving period
    date1 = #6/15/2023 3:30:00 AM#
    Debug.Print "Case 1 (DST): " & IsDST(date1) ' Output: True
    
    ' Case 2: Date is not in the daylight saving period
    date2 = #12/15/2023 2:30:00 AM#
    Debug.Print "Case 2 (Standard Time): " & IsDST(date2) ' Output: False
    
    ' Case 3: Date is on the transition day from DST to standard time
    date3 = #10/30/2023 2:30:00 AM#
    Debug.Print "Case 3 (Transition DST to Standard): " & IsDST(date3) ' Output: False
    
    ' Case 4: Date is on the transition day from standard time to DST
    date4 = #3/26/2023 2:30:00 AM#
    Debug.Print "Case 4 (Transition Standard to DST): " & IsDST(date4) ' Output: True
End Sub


