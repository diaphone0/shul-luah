Attribute VB_Name = "mod_hebrewcalendar"
' Define data structure for Hebrew date
Public Type hdate
    year As Long '0 - infinity
    month As Long 'starts from nissan 1 - 13
    day As Long '1-30
    dayOfYear As Long 'starts from tishrei
    wday As Long 'weekdays 1 - 6, 0
    leap As Long 'if leap year
    hour As Long '0-23
    min As Long '0-59
    sec As Long '0-59
    msec As Long
    offset As Long 'timezone offset in seconds
    EY As Boolean 'if Eretz Yisroel (for yomtov & parshah)
End Type

' Define enumeration for yomtov
Public Enum yomtov
    CHOL
    PESACH_DAY1
    PESACH_DAY2
    SHVEI_SHEL_PESACH
    ACHRON_SHEL_PESACH
    SHAVOUS_DAY1
    SHAVOUS_DAY2
    ROSH_HASHANAH_DAY1
    ROSH_HASHANAH_DAY2
    YOM_KIPPUR
    SUKKOS_DAY1
    SUKKOS_DAY2
    SHMEINI_ATZERES
    SIMCHAS_TORAH
    CHOL_HAMOED_PESACH_DAY1
    CHOL_HAMOED_PESACH_DAY2
    CHOL_HAMOED_PESACH_DAY3
    CHOL_HAMOED_PESACH_DAY4
    CHOL_HAMOED_PESACH_DAY5
    CHOL_HAMOED_SUKKOS_DAY1
    CHOL_HAMOED_SUKKOS_DAY2
    CHOL_HAMOED_SUKKOS_DAY3
    CHOL_HAMOED_SUKKOS_DAY4
    CHOL_HAMOED_SUKKOS_DAY5
    HOSHANA_RABBAH
    PESACH_SHEINI
    LAG_BAOMER
    TU_BAV
    CHANUKAH_DAY1
    CHANUKAH_DAY2
    CHANUKAH_DAY3
    CHANUKAH_DAY4
    CHANUKAH_DAY5
    CHANUKAH_DAY6
    CHANUKAH_DAY7
    CHANUKAH_DAY8
    TU_BISHVAT
    PURIM_KATAN
    SHUSHAN_PURIM_KATAN
    PURIM
    SHUSHAN_PURIM
    SHIVA_ASAR_BTAAMUZ
    TISHA_BAV
    TZOM_GEDALIA
    ASARAH_BTEVES
    TAANIS_ESTER
    EREV_PESACH
    EREV_SHAVOUS
    EREV_ROSH_HASHANAH
    EREV_YOM_KIPPUR
    EREV_SUKKOS
    SHKALIM
    ZACHOR
    PARAH
    HACHODESH
    ROSH_CHODESH
    MACHAR_CHODESH
    SHABBOS_MEVORCHIM
    HAGADOL
    CHAZON
    NACHAMU
    SHUVA
    SHIRA
    SHABBOS_CHOL_HAMOED
End Enum


' Define enumeration for parshah
Public Enum parshah
    NOPARSHAH
    BERESHIT
    NOACH
    LECH_LECHA
    VAYEIRA
    CHAYEI_SARAH
    TOLEDOT
    VAYETZE
    VAYISHLACH
    VAYESHEV
    MIKETZ
    VAYIGASH
    VAYECHI
    SHEMOT
    VAEIRA
    BO
    BESHALACH
    YITRO
    MISHPATIM
    TERUMAH
    TETZAVEH
    KI_TISA
    VAYAKHEL
    PEKUDEI
    VAYIKRA
    TZAV
    SHEMINI
    TAZRIA
    METZORA
    ACHAREI_MOT
    KEDOSHIM
    EMOR
    BEHAR
    BECHUKOTAI
    BAMIDBAR
    NASO
    BEHAALOTECHA
    SHLACH
    KORACH
    CHUKAT
    BALAK
    PINCHAS
    MATOT
    MASEI
    DEVARIM
    VAETCHANAN
    EIKEV
    REEH
    SHOFTIM
    KI_TEITZEI
    KI_TAVO
    NITZAVIM
    VAYELECH
    HAAZINU
    VZOT_HABERACHAH
    VAYAKHEL_PEKUDEI
    TAZRIA_METZORA
    ACHAREI_MOT_KEDOSHIM
    BEHAR_BECHUKOTAI
    CHUKAT_BALAK
    MATOT_MASEI
    NITZAVIM_VAYELECH
End Enum

'is the year a leap year?
Public Function HebrewLeapYear(year As Long) As Long
    If (((7 * year) + 1) Mod 19) < 7 Then
        HebrewLeapYear = 1
    Else
        HebrewLeapYear = 0
    End If
End Function

'day from molad tohu until Rosh Hashana
Public Function HebrewCalendarElapsedDays(year As Long) As Long
    Dim MonthsElapsed As Long
    MonthsElapsed = (235 * ((year - 1) \ 19)) + (12 * ((year - 1) Mod 19)) + ((7 * ((year - 1) Mod 19) + 1) \ 19)
    Dim PartsElapsed As Long
    PartsElapsed = 204 + 793 * (MonthsElapsed Mod 1080)
    Dim HoursElapsed As Long
    HoursElapsed = 5 + 12 * MonthsElapsed + 793 * (MonthsElapsed \ 1080) + (PartsElapsed \ 1080)
    Dim ConjunctionDay As Long
    ConjunctionDay = 1 + 29 * MonthsElapsed + (HoursElapsed \ 24)
    Dim ConjunctionParts As Long
    ConjunctionParts = 1080 * (HoursElapsed Mod 24) + PartsElapsed Mod 1080
    Dim AlternativeDay As Long
    Dim cdw As Long
    cdw = (ConjunctionDay Mod 7)
    
    If (ConjunctionParts >= 19440) Or _
       ((cdw = 2) And (ConjunctionParts >= 9924) And Not (HebrewLeapYear(year) = 1)) Or _
       ((cdw = 1) And (ConjunctionParts >= 16789) And (HebrewLeapYear(year - 1) = 1)) Then
        AlternativeDay = ConjunctionDay + 1
    Else
        AlternativeDay = ConjunctionDay
    End If
    
    Dim adw As Long
    adw = (AlternativeDay Mod 7)
    
    If (adw = 0) Or (adw = 3) Or (adw = 5) Then
        HebrewCalendarElapsedDays = 1 + AlternativeDay
    Else
        HebrewCalendarElapsedDays = AlternativeDay
    End If
End Function

Public Function DaysInHebrewYear(year As Long) As Long
    DaysInHebrewYear = (HebrewCalendarElapsedDays(year + 1)) - (HebrewCalendarElapsedDays(year))
End Function

Public Function LongHeshvan(year As Long) As Long
    If (DaysInHebrewYear(year) Mod 10) = 5 Then
        LongHeshvan = 1
    Else
        LongHeshvan = 0
    End If
End Function

Public Function ShortKislev(year As Long) As Long
    If (DaysInHebrewYear(year) Mod 10) = 3 Then
        ShortKislev = 1
    Else
        ShortKislev = 0
    End If
End Function

Public Function LastDayOfHebrewMonth(month As Long, year As Long) As Long

    If (month = 2) Or (month = 4) Or (month = 6) Or _
       ((month = 8) And Not (LongHeshvan(year) = 1)) Or _
       ((month = 9) And ShortKislev(year) = 1) Or _
       (month = 10) Or _
       ((month = 12) And Not (HebrewLeapYear(year) = 1)) Or _
       (month = 13) Then
        LastDayOfHebrewMonth = 29
    Else
        LastDayOfHebrewMonth = 30
    End If
End Function

'day of year for alef nissan
Public Function NissanCount(year As Long) As Long
    Select Case DaysInHebrewYear(year)
        Case 353
            NissanCount = 176
        Case 354
            NissanCount = 177
        Case 355
            NissanCount = 178
        Case 383
            NissanCount = 206
        Case 384
            NissanCount = 207
        Case 385
            NissanCount = 208
    End Select
End Function

Public Function HDateNew(year As Long, month As Long, day As Long, _
                  hour As Long, min As Long, sec As Long, _
                  msec As Long, offset As Long) As hdate
    Dim result As hdate
    result.year = year
    result.month = month
    result.day = day
    result.hour = hour
    result.min = min
    result.sec = sec
    result.msec = msec
    result.offset = offset
'    HDateSetDOY result
    HDateNew = result
End Function

Sub SetEY(ByRef date_in As hdate, EY As Boolean)
    date_in.EY = EY
End Sub


'Public Function ConvertDate(date_in As TMStruct) As hdate
Public Function ConvertDate(date_in As Date) As hdate
    Dim result As hdate
    Dim julianDay As Double
    Dim d As Long
    Dim m As Double
    Dim year As Long
    Dim month As Long
    Dim daycount As Long
    Dim dayOfYear As Long
    Dim nissanStart As Long
    Dim tm_date_in As TMStruct
    tm_date_in = mktm(date_in)
    
    julianDay = GregorianJulian(tm_date_in)
    d = Fix(julianDay) - 347996
    m = (d * 25920#) / 765433#
    year = Fix((19 * m) / 235)
    
    While d >= HebrewCalendarElapsedDays(year + 1)
        year = year + 1
    Wend
    
    Dim ys As Long
    ys = HebrewCalendarElapsedDays(year)
    dayOfYear = (d - ys) + 1
    nissanStart = NissanCount(year)
    
    If dayOfYear <= nissanStart Then
        month = 7 ' Start at Tishri
        daycount = 0
    Else
        month = 1 ' Start at Nisan
        daycount = nissanStart
    End If
    
    While dayOfYear > (daycount + LastDayOfHebrewMonth(month, year))
        daycount = daycount + LastDayOfHebrewMonth(month, year)
        month = month + 1
    Wend
    
    Dim day As Long
    day = dayOfYear - daycount
    
    result.year = year
    result.month = month
    result.day = day
    result.wday = (HebrewCalendarElapsedDays(year) + dayOfYear) Mod 7
    result.dayOfYear = dayOfYear
    result.leap = HebrewLeapYear(year)
    result.hour = tm_date_in.tm_hour
    result.min = tm_date_in.tm_min
    result.sec = tm_date_in.tm_sec
'    HDateSetDoy result
    ConvertDate = result
End Function

'convert a hdate to gregorian date
Public Function HDateGregorian(date_in As hdate) As Date
    Dim result As TMStruct
    Dim JD As Double
    Dim a As Double
    Dim b As Double
    Dim C As Double
    Dim d As Double
    Dim E As Double
    Dim m As Long
    Dim y As Long
    
    JD = HDateJulian(date_in) + 0.5
    a = Int((JD - 1867216.25) / 36524.25)
    b = (JD + 1525 + a - Int(a / 4))
    C = Int((b - 122.1) / 365.25)
    d = Int(C * 365.25)
    E = Int((b - d) / 30.6001)
    
    If E > 13 Then
        m = E - 13
    Else
        m = E - 1
    End If
    
    If m > 2 Then
        y = C - 4716
    Else
        y = C - 4715
    End If
    
    result.tm_year = y - 1900
    result.tm_mon = m - 1
    result.tm_mday = (b - d - Int(E * 30.6001))
    result.tm_hour = date_in.hour
    result.tm_min = date_in.min
    result.tm_sec = date_in.sec
    result.tm_isdst = -1
    
    HDateGregorian = mkdate(result)
End Function

'convert a gregorian date to julian day
Public Function GregorianJulian(date_in As TMStruct) As Double
    Dim year As Long
    Dim month As Long
    Dim day As Long
    year = date_in.tm_year + 1900
    month = date_in.tm_mon + 1
    day = date_in.tm_mday
    If month <= 2 Then
        year = year - 1
        month = month + 12
    End If
    Dim a As Long
    Dim b As Long
    a = Fix(year / 100)
    b = 2 - a + Fix(a / 4)
    Dim JD As Double
    JD = Fix(365.25 * (year + 4716)) + Fix(30.6001 * (month + 1)) + day + b - 1524.5
    GregorianJulian = JD
End Function

'convert a hdate to julian day
Public Function HDateJulian(date_in As hdate) As Double
    Dim diff As Double
    diff = 347996.5
    Dim yearstart As Long
    yearstart = HebrewCalendarElapsedDays(date_in.year)
    HDateJulian = (date_in.dayOfYear - 1) + yearstart + diff
End Function

Public Function HDateTime_t(date_in As hdate) As Double
    Dim result As Double
    result = (HebrewCalendarElapsedDays(date_in.year) + (date_in.dayOfYear - 1)) - 2092591
    result = ((((((result * 24) + date_in.hour) * 60) + date_in.min) * 60) + date_in.sec)
    result = result - date_in.offset
    HDateTime_t = result
End Function

Public Function Time_THDate(time As Double, offset As Long) As hdate
    Dim temp As Double
    temp = time + offset
    Dim result As hdate
    result.sec = temp Mod 60
    temp = temp / 60
    result.min = temp Mod 60
    temp = temp / 60
    result.hour = temp Mod 24
    Dim d As Long
    d = (temp / 24) + 2092591
    Dim m As Double
    m = ((d * 25920#) / 765433#)
    Dim year As Long
    year = Int((19# * m) / 235#)
    Dim month As Long
    Dim daycount As Long
    Do While d >= HebrewCalendarElapsedDays(year + 1)
        year = year + 1
    Loop
    Dim ys As Long
    ys = HebrewCalendarElapsedDays(year)
    Dim dayOfYear As Long
    dayOfYear = (d - ys) + 1
    Dim nissanStart As Long
    nissanStart = NissanCount(year)
    If dayOfYear <= nissanStart Then
        month = 7 ' Start at Tishri
        daycount = 0
    Else
        month = 1 ' Start at Nisan
        daycount = nissanStart
    End If
    Do While dayOfYear > (daycount + LastDayOfHebrewMonth(month, year))
        daycount = daycount + LastDayOfHebrewMonth(month, year)
        month = month + 1
    Loop
    Dim day As Long
    day = dayOfYear - daycount
    result.year = year
    result.month = month
    result.day = day
    result.offset = offset
    HDateSetDoy result
    Time_THDate = result
End Function

'compare 2 hdate:
' returns 0 if they are the same
' 1 if date1 < date2
' and -1 if date1 > date2
Public Function HDateCompare(date1 As hdate, date2 As hdate) As Long
    If date1.year < date2.year Then
        HDateCompare = 1
    ElseIf date1.year > date2.year Then
        HDateCompare = -1
    ElseIf date1.dayOfYear < date2.dayOfYear Then
        HDateCompare = 1
    ElseIf date1.dayOfYear > date2.dayOfYear Then
        HDateCompare = -1
    ElseIf date1.hour < date2.hour Then
        HDateCompare = 1
    ElseIf date1.hour > date2.hour Then
        HDateCompare = -1
    ElseIf date1.min < date2.min Then
        HDateCompare = 1
    ElseIf date1.min > date2.min Then
        HDateCompare = -1
    ElseIf date1.sec < date2.sec Then
        HDateCompare = 1
    ElseIf date1.sec > date2.sec Then
        HDateCompare = -1
    ElseIf date1.msec < date2.msec Then
        HDateCompare = 1
    ElseIf date1.msec > date2.msec Then
        HDateCompare = -1
    Else
        HDateCompare = 0
    End If
End Function

' normalize a hdate and set the wday, dayofyear, and leap
' converts 13 (Adar II) in a non leap year to 12
' converts 30 cheshvan, kislev & adar to 29 in a year that only has 29 days
Sub HDateSetDoy(ByRef date_in As hdate)
    Dim year As Long
    Dim month As Long
    Dim day As Long
    Dim monthcount As Long
    Dim dayOfYear As Long

    'dayOfYear = date_in.day
    'For month = 1 To (date_in.month - 6) - 1
    '    dayOfYear = dayOfYear + LastDayOfHebrewMonth(month, date_in.year)
    'Next month
    
    'date_in.dayOfYear = dayOfYear
'       if (month < TISHREI) {
'           // this year before and after Nisan.
'           for (int m = TISHREI; m <= getLastMonthOfJewishYear(year); m++) {
'               elapsedDays += getDaysInJewishMonth(m, year);
'           }
'           for (int m = NISSAN; m < month; m++) {
'               elapsedDays += getDaysInJewishMonth(m, year);
'           }
'       } else { // Add days in prior months this year
'           for (int m = TISHREI; m < month; m++) {
'               elapsedDays += getDaysInJewishMonth(m, year);
'           }
'       }
'       return elapsedDays;
'    Exit Sub
    year = date_in.year
    month = date_in.month
    If month = 13 And Not (HebrewLeapYear(year) = 1) Then
        month = 12
    End If
    If date_in.day = 30 And LastDayOfHebrewMonth(month, year) = 29 Then
        date_in.day = 29
    End If
    day = date_in.day
    If month < 7 Then
        monthcount = 1
        dayOfYear = NissanCount(year)
    Else
        monthcount = 7
        dayOfYear = 0
    End If
    Do While monthcount < month
        dayOfYear = dayOfYear + LastDayOfHebrewMonth(monthcount, year)
        monthcount = monthcount + 1
    Loop
    dayOfYear = dayOfYear + day
    date_in.dayOfYear = dayOfYear
    date_in.wday = (HebrewCalendarElapsedDays(year) + dayOfYear) Mod 7
    date_in.leap = HebrewLeapYear(year)
End Sub

' functions to add or subtract from a hdate field and then normalize the result
Sub HDateAddYear(ByRef date_in As hdate, years As Long)
    Dim year As Long
    year = date_in.year
    Dim month As Long
    month = date_in.month
    Dim leap1 As Long
    leap1 = date_in.leap
    year = year + years
    Dim leap2 As Long
    leap2 = IIf(HebrewLeapYear(year) = 1, 1, 0)
    If leap1 <> leap2 Then
        If leap1 And month = 13 Then
            month = 12
        ElseIf Not (leap1 = 1) And month = 12 Then
            month = 13
        End If
    End If
    date_in.year = year
    date_in.month = month
    HDateSetDoy date_in
End Sub

Sub HDateAddMonth(ByRef date_in As hdate, months As Long)
    Dim last As Boolean
    last = False
    If date_in.day = 30 Then
        last = True
    End If
    Dim monthcount As Long
    monthcount = months
    Do While monthcount > 0
        Select Case date_in.month
            Case 12
                If date_in.leap = 1 Then
                    date_in.month = date_in.month + 1
                Else
                    date_in.month = 1
                End If
                monthcount = monthcount - 1
            Case 13
                date_in.month = 1
                monthcount = monthcount - 1
            Case 6
                date_in.month = date_in.month + 1
                HDateAddYear date_in, 1
                monthcount = monthcount - 1
            Case Else
                date_in.month = date_in.month + 1
                monthcount = monthcount - 1
        End Select
    Loop
    Do While monthcount < 0
        Select Case date_in.month
            Case 1
                If date_in.leap = 1 Then
                    date_in.month = 13
                Else
                    date_in.month = 12
                End If
                monthcount = monthcount + 1
            Case 7
                date_in.month = date_in.month - 1
                HDateAddYear date_in, -1
                monthcount = monthcount + 1
            Case Else
                date_in.month = date_in.month - 1
                monthcount = monthcount + 1
        End Select
    Loop
    If last And LastDayOfHebrewMonth(date_in.month, date_in.year) = 30 Then
        date_in.day = 30
    End If
    HDateSetDoy date_in
End Sub

Public Sub HDateAddDay(ByRef date_in As hdate, days As Long)
    Dim daycount As Long
    daycount = days
    Do While daycount > 0
        Select Case date_in.day
            Case 30
                date_in.day = 1
                HDateAddMonth date_in, 1
                daycount = daycount - 1
            Case 29
                If LastDayOfHebrewMonth(date_in.month, date_in.year) = 29 Then
                    date_in.day = 1
                    HDateAddMonth date_in, 1
                Else
                    date_in.day = date_in.day + 1
                End If
                daycount = daycount - 1
            Case Else
                date_in.day = date_in.day + 1
                daycount = daycount - 1
        End Select
    Loop
    Do While daycount < 0
        Select Case date_in.day
            Case 1
                HDateAddMonth date_in, -1
                If LastDayOfHebrewMonth(date_in.month, date_in.year) = 30 Then
                    date_in.day = 30
                Else
                    date_in.day = 29
                End If
                daycount = daycount + 1
            Case Else
                date_in.day = date_in.day - 1
                daycount = daycount + 1
        End Select
    Loop
    HDateSetDoy date_in
End Sub

Sub DivideAndCarry(start As Long, ByRef finish As Long, ByRef carry As Long, divisor As Long)
    finish = start Mod divisor
    carry = start \ divisor
    If finish < 0 Then
        finish = finish + divisor
        carry = carry - 1
    End If
End Sub
Sub HDateAddHour(ByRef date_in As hdate, hours As Long)
    Dim hour As Long
    hour = date_in.hour + hours
    Dim carry As Long
    DivideAndCarry hour, date_in.hour, carry, 24
    If carry Then
        HDateAddDay date_in, carry
    Else
        HDateSetDoy date_in
    End If
End Sub

Sub HDateAddMinute(ByRef date_in As hdate, minutes As Long)
    Dim minute As Long
    minute = date_in.min + minutes
    Dim carry As Long
    DivideAndCarry minute, date_in.min, carry, 60
    If carry Then
        HDateAddHour date_in, carry
    Else
        HDateSetDoy date_in
    End If
End Sub

Sub HDateAddSecond(ByRef date_in As hdate, seconds As Long)
    Dim second As Long
    second = date_in.sec + seconds
    Dim carry As Long
    DivideAndCarry second, date_in.sec, carry, 60
    If carry Then
        HDateAddMinute date_in, carry
    Else
        HDateSetDoy date_in
    End If
End Sub

Sub HDateAddMSecond(ByRef date_in As hdate, mseconds As Long)
    Dim msecond As Long
    msecond = date_in.msec + mseconds
    Dim carry As Long
    DivideAndCarry msecond, date_in.msec, carry, 1000
    If carry Then
        HDateAddSecond date_in, carry
    Else
        HDateSetDoy date_in
    End If
End Sub

Sub HDateAdd(ByRef date_in As hdate, years As Long, months As Long, days As Long, hours As Long, minutes As Long, seconds As Long, mseconds As Long)
    If years Then HDateAddYear date_in, years
    If months Then HDateAddMonth date_in, months
    If days Then HDateAddDay date_in, days
    If hours Then HDateAddHour date_in, hours
    If minutes Then HDateAddMinute date_in, minutes
    If seconds Then HDateAddSecond date_in, seconds
    If mseconds Then HDateAddMSecond date_in, mseconds
End Sub

' Returns the molad of the given month remember month 1 is Nissan
' beware that the seconds are actually chalakim 1/1080 of an hour
' and that the hebrew date is actually the next one when the hour > 18
' also the molad is calculated in Yerushalayim Mean Time so cannot be easily converted.
Public Function GetMolad(year As Long, month As Long) As hdate
    Dim result As hdate
    result = HDateNew(0, 0, 0, 0, 0, 0, 0, 0)
    Dim MonthsElapsed As Long
    MonthsElapsed = _
        (235 * ((year - 1) \ 19)) + _
        (12 * ((year - 1) Mod 19)) + _
        (7 * ((year - 1) Mod 19) + 1) \ 19
    
    If month > 6 Then
        MonthsElapsed = MonthsElapsed + (month - 7)
    Else
        MonthsElapsed = MonthsElapsed + (month + 5)
        If HebrewLeapYear(year) = 1 Then
            MonthsElapsed = MonthsElapsed + 1
        End If
    End If
    
    Dim PartsElapsed As Long
    PartsElapsed = 204 + 793 * (MonthsElapsed Mod 1080)
    Dim HoursElapsed As Long
    HoursElapsed = _
        5 + 12 * MonthsElapsed + _
        793 * (MonthsElapsed \ 1080) + _
        PartsElapsed \ 1080
    
    Dim ConjunctionDay As Long
    ConjunctionDay = 29 * MonthsElapsed + (HoursElapsed) \ 24
    Dim ConjunctionHour As Long
    ConjunctionHour = (HoursElapsed Mod 24)
    Dim ConjunctionMinute As Long
    ConjunctionMinute = (PartsElapsed Mod 1080) \ 18
    Dim ConjunctionParts As Long
    ConjunctionParts = (PartsElapsed Mod 1080) Mod 18
    
    If ConjunctionDay > 2085362 Then
        'snapshot for shorter calculation - alef nissan 5710
        result.year = 5710
        result.month = 1
        result.day = 1
        result.wday = 1
        
        HDateAddDay result, ConjunctionDay - 2085362
    Else
        result.year = 1
        result.month = 7
        result.day = 1
        HDateAddDay result, ConjunctionDay
    End If
    result.hour = ConjunctionHour
    result.min = ConjunctionMinute
    result.sec = ConjunctionParts
    result.offset = 8456
    HDateAddHour result, -6
    
    GetMolad = result
End Function

'get the relevant type of parsha list for the year of the specified hdate (there are 16 different types in total)
Public Function GetYearType(date_in As hdate) As Long
    Dim yearWday As Long
    GetYearType = -1
    yearWday = (HebrewCalendarElapsedDays(date_in.year) + 1) Mod 7
    If yearWday = 0 Then
        yearWday = 7
    End If
    If date_in.leap = 1 Then
        Select Case yearWday
            Case 2
                If ShortKislev(date_in.year) Then
                    If date_in.EY Then
                        GetYearType = 14
                    Else
                        GetYearType = 6
                    End If
                ElseIf LongHeshvan(date_in.year) Then
                    If date_in.EY Then
                        GetYearType = 15
                    Else
                        GetYearType = 7
                    End If
                End If
            Case 3
                If date_in.EY Then
                    GetYearType = 15
                Else
                    GetYearType = 7
                End If
            Case 5
                If ShortKislev(date_in.year) Then
                    GetYearType = 8
                ElseIf LongHeshvan(date_in.year) Then
                    GetYearType = 9
                End If
            Case 7
                If ShortKislev(date_in.year) Then
                    GetYearType = 10
                ElseIf LongHeshvan(date_in.year) Then
                    If date_in.EY Then
                        GetYearType = 16
                    Else
                        GetYearType = 11
                    End If
                End If
        End Select
    Else
        Select Case yearWday
            Case 2
                If ShortKislev(date_in.year) Then
                    GetYearType = 0
                ElseIf LongHeshvan(date_in.year) Then
                    If date_in.EY Then
                        GetYearType = 12
                    Else
                        GetYearType = 1
                    End If
                End If
            Case 3
                If date_in.EY Then
                    GetYearType = 12
                Else
                    GetYearType = 1
                End If
            Case 5
                If LongHeshvan(date_in.year) Then
                    GetYearType = 3
                ElseIf Not (ShortKislev(date_in.year) = 1) Then
                    If date_in.EY Then
                        GetYearType = 13
                    Else
                        GetYearType = 2
                    End If
                End If
            Case 7
                If ShortKislev(date_in.year) Then
                    GetYearType = 4
                ElseIf LongHeshvan(date_in.year) Then
                    GetYearType = 5
                End If
        End Select
    End If
    
End Function

'if Shabbos get the current parshah otherwise return 0
Public Function GetParshah(date_in As hdate) As parshah
    Dim yearType As Long
    Dim yearWday As Long
    Dim day As Long
    
    If parashaList_ready = False Then Init_parashalist
    
    yearType = GetYearType(date_in)
    yearWday = HebrewCalendarElapsedDays(date_in.year) Mod 7
    Call HDateSetDoy(date_in)
    day = yearWday + date_in.dayOfYear
    
    If date_in.wday <> 0 Then
        GetParshah = NOPARSHAH
        Exit Function
    End If
    
    If yearType >= 0 Then
        GetParshah = parashaList(yearType, day \ 7)
    Else
        GetParshah = NOPARSHAH
    End If
End Function

'if yomtov get the current yomtov otherwise return 0
Public Function GetYomTov(date_in As hdate) As yomtov
    GetYomTov = CHOL
    Select Case date_in.month
        Case 1
            Select Case date_in.day
                Case 14: GetYomTov = EREV_PESACH
                Case 15: GetYomTov = PESACH_DAY1
                Case 16: If date_in.EY Then GetYomTov = CHOL_HAMOED_PESACH_DAY1 Else GetYomTov = PESACH_DAY2
                Case 17: If date_in.EY Then GetYomTov = CHOL_HAMOED_PESACH_DAY2 Else GetYomTov = CHOL_HAMOED_PESACH_DAY1
                Case 18: If date_in.EY Then GetYomTov = CHOL_HAMOED_PESACH_DAY3 Else GetYomTov = CHOL_HAMOED_PESACH_DAY2
                Case 19: If date_in.EY Then GetYomTov = CHOL_HAMOED_PESACH_DAY4 Else GetYomTov = CHOL_HAMOED_PESACH_DAY3
                Case 20: If date_in.EY Then GetYomTov = CHOL_HAMOED_PESACH_DAY5 Else GetYomTov = CHOL_HAMOED_PESACH_DAY4
                Case 21:  GetYomTov = SHVEI_SHEL_PESACH
                Case 22: If Not date_in.EY Then GetYomTov = ACHRON_SHEL_PESACH
            End Select
        Case 2
            Select Case date_in.day
                Case 14: GetYomTov = PESACH_SHEINI: Exit Function
                Case 18: GetYomTov = LAG_BAOMER: Exit Function
            End Select
        Case 3
            Select Case date_in.day
                Case 5: GetYomTov = EREV_SHAVOUS: Exit Function
                Case 6: GetYomTov = SHAVOUS_DAY1: Exit Function
                Case 7
                    If Not date_in.EY Then
                        GetYomTov = SHAVOUS_DAY2: Exit Function
                    End If
            End Select
        Case 4
            Select Case date_in.day
                Case 17, 18
                    If (date_in.day = 17 And date_in.wday <> 0) Or (date_in.day = 10 And date_in.wday = 1) Then
                        GetYomTov = SHIVA_ASAR_BTAAMUZ: Exit Function
                    End If
            End Select
        Case 5
            Select Case date_in.day
                Case 9, 10
                    If (date_in.day = 9 And date_in.wday <> 0) Or (date_in.day = 10 And date_in.wday = 1) Then
                        GetYomTov = TISHA_BAV: Exit Function
                    End If
                Case 15: GetYomTov = TU_BAV: Exit Function
            End Select
        Case 6
            If date_in.day = 29 Then
                GetYomTov = EREV_ROSH_HASHANAH: Exit Function
            End If
        Case 7
            Select Case date_in.day
                Case 1: GetYomTov = ROSH_HASHANAH_DAY1
                Case 2: GetYomTov = ROSH_HASHANAH_DAY2
                Case 3: If date_in.wday <> 0 Then GetYomTov = TZOM_GEDALIA
                Case 4: If date_in.wday = 1 Then GetYomTov = TZOM_GEDALIA
                Case 9: GetYomTov = EREV_YOM_KIPPUR
                Case 10: GetYomTov = YOM_KIPPUR
                Case 14: GetYomTov = EREV_SUKKOS
                Case 15: GetYomTov = SUKKOS_DAY1
                Case 16: If date_in.EY Then GetYomTov = CHOL_HAMOED_SUKKOS_DAY1 Else GetYomTov = SUKKOS_DAY2
                Case 17: If date_in.EY Then GetYomTov = CHOL_HAMOED_SUKKOS_DAY2 Else GetYomTov = CHOL_HAMOED_SUKKOS_DAY1
                Case 18: If date_in.EY Then GetYomTov = CHOL_HAMOED_SUKKOS_DAY3 Else GetYomTov = CHOL_HAMOED_SUKKOS_DAY2
                Case 19: If date_in.EY Then GetYomTov = CHOL_HAMOED_SUKKOS_DAY4 Else GetYomTov = CHOL_HAMOED_SUKKOS_DAY3
                Case 20: If date_in.EY Then GetYomTov = CHOL_HAMOED_SUKKOS_DAY5 Else GetYomTov = CHOL_HAMOED_SUKKOS_DAY4
                Case 21: GetYomTov = HOSHANA_RABBAH
                Case 22: GetYomTov = SHMEINI_ATZERES
                Case 23: If Not date_in.EY Then GetYomTov = SIMCHAS_TORAH
            End Select
        Case 9
            Select Case date_in.day
                Case 25: GetYomTov = CHANUKAH_DAY1
                Case 26: GetYomTov = CHANUKAH_DAY2
                Case 27: GetYomTov = CHANUKAH_DAY3
                Case 28: GetYomTov = CHANUKAH_DAY4
                Case 29: GetYomTov = CHANUKAH_DAY5
                Case 30: GetYomTov = CHANUKAH_DAY6
            End Select
        Case 10
            Select Case date_in.day
                Case 1
                    If ShortKislev(date_in.year) Then
                        GetYomTov = CHANUKAH_DAY6: Exit Function
                    Else
                        GetYomTov = CHANUKAH_DAY7: Exit Function
                    End If
                Case 2
                    If ShortKislev(date_in.year) Then
                        GetYomTov = CHANUKAH_DAY7: Exit Function
                    Else
                        GetYomTov = CHANUKAH_DAY8: Exit Function
                    End If
                Case 3
                    If ShortKislev(date_in.year) Then
                        GetYomTov = CHANUKAH_DAY8: Exit Function
                    End If
                Case 10: GetYomTov = ASARAH_BTEVES: Exit Function
            End Select
        Case 11
            If date_in.day = 15 Then
                GetYomTov = TU_BISHVAT: Exit Function
            End If
        Case 12
            Select Case date_in.day
                Case 11
                    If date_in.leap = 0 And date_in.wday = 5 Then
                        GetYomTov = TAANIS_ESTER: Exit Function
                    End If
                Case 13
                    If date_in.leap = 0 And date_in.wday <> 0 Then
                        GetYomTov = TAANIS_ESTER: Exit Function
                    End If
                Case 14
                    If date_in.leap = 1 Then
                        GetYomTov = PURIM_KATAN: Exit Function
                    Else
                        GetYomTov = PURIM: Exit Function
                    End If
                Case 15
                    If date_in.leap = 1 Then
                        GetYomTov = SHUSHAN_PURIM_KATAN: Exit Function
                    Else
                        GetYomTov = SHUSHAN_PURIM: Exit Function
                    End If
            End Select
        Case 13
            Select Case date_in.day
                Case 11
                    If date_in.wday = 5 Then
                        GetYomTov = TAANIS_ESTER: Exit Function
                    End If
                Case 13
                    If date_in.wday <> 0 Then
                        GetYomTov = TAANIS_ESTER: Exit Function
                    End If
                Case 14: GetYomTov = PURIM: Exit Function
                Case 15: GetYomTov = SHUSHAN_PURIM: Exit Function
            End Select
    End Select
    Select Case GetYomTov
        Case CHOL_HAMOED_PESACH_DAY1 To CHOL_HAMOED_PESACH_DAY5, CHOL_HAMOED_SUKKOS_DAY1 To CHOL_HAMOED_SUKKOS_DAY5
            If date_in.wday = 0 Then GetYomTov = SHABBOS_CHOL_HAMOED
    End Select
End Function

'if Shabbos get the current special parshah otherwise return 0
Public Function GetSpecialShabbos(date_in As hdate) As yomtov
    GetSpecialShabbos = CHOL
    If date_in.wday = 0 Then
        If (date_in.month = 11 And Not (date_in.leap = 1)) Or (date_in.month = 12 And (date_in.leap = 1)) Then
            Select Case date_in.day
                Case 25, 27, 29
                    GetSpecialShabbos = SHKALIM
            End Select
        End If
        If (date_in.month = 12 And Not (date_in.leap = 1)) Or date_in.month = 13 Then
            Select Case date_in.day
                Case 1
                    GetSpecialShabbos = SHKALIM
                Case 8, 9, 11, 13
                    GetSpecialShabbos = ZACHOR
                Case 18, 20, 22, 23
                    GetSpecialShabbos = PARAH
                Case 25, 27, 29
                    GetSpecialShabbos = HACHODESH
            End Select
        End If
        If date_in.month = 1 Then
            If date_in.day = 1 Then GetSpecialShabbos = HACHODESH
            If date_in.day >= 8 And date_in.day <= 14 Then GetSpecialShabbos = HAGADOL
        End If
        If date_in.month = 5 Then
            If date_in.day >= 4 And date_in.day <= 9 Then GetSpecialShabbos = CHAZON
            If date_in.day >= 10 And date_in.day <= 16 Then GetSpecialShabbos = NACHAMU
        End If
        If date_in.month = 7 Then
            If date_in.day >= 3 And date_in.day <= 8 Then GetSpecialShabbos = SHUVA
        End If
        If GetParshah(date_in) = BESHALACH Then GetSpecialShabbos = SHIRA
    End If
End Function

'if rosh chodesh return rosh chodesh otherwise return 0
Public Function GetRoshChodesh(date_in As hdate) As yomtov
    If date_in.day = 30 Or (date_in.day = 1 And date_in.month <> 7) Then
        GetRoshChodesh = ROSH_CHODESH
    Else
        GetRoshChodesh = CHOL
    End If
End Function

'if machar chodesh return machar chodesh otherwise return 0
Public Function GetMacharChodesh(date_in As hdate) As yomtov
    If date_in.wday Then
        GetMacharChodesh = CHOL
    ElseIf date_in.day = 30 Or date_in.day = 29 Then
        GetMacharChodesh = MACHAR_CHODESH
    Else
        GetMacharChodesh = CHOL
    End If
End Function

'if shabbos mevorchim return shabbos mevorchim otherwise return 0
Public Function GetShabbosMevorchim(date_in As hdate) As yomtov
    If date_in.wday Then
        GetShabbosMevorchim = CHOL
    ElseIf date_in.day >= 23 And date_in.day <= 29 Then
        GetShabbosMevorchim = SHABBOS_MEVORCHIM
    Else
        GetShabbosMevorchim = CHOL
    End If
End Function

'The omer count 1 - 49 or 0 if none
Public Function GetOmer(date_in As hdate) As Long
    Dim omer As Long
    If date_in.month = 1 And date_in.day >= 16 Then
        omer = date_in.day - 15
    ElseIf date_in.month = 2 Then
        omer = date_in.day + 15
    ElseIf date_in.month = 3 And date_in.day <= 5 Then
        omer = date_in.day + 44
    End If
    GetOmer = omer
End Function

' if Shabbos get the current chapter of avos otherwise return 0
' returns 1 - 6 or 12 or 34 or 56 for double chapter
Public Function GetAvos(date_in As hdate) As Long
    If date_in.wday Then
        GetAvos = 0 ' Shabbos
    Else
        Dim avos_start As Long
        avos_start = NissanCount(date_in.year) + 23 ' 23 Nissan
        Dim avos_day As Long
        avos_day = date_in.dayOfYear - avos_start ' days from start of avos
        If avos_day <= 0 Then
            GetAvos = 0
            Exit Function
        End If
        Dim chapter As Long
        chapter = avos_day \ 7 ' current chapter

        ' corrections
        Dim avos_day_of_week As Long
        avos_day_of_week = avos_day Mod 7
        If avos_day_of_week = 6 Then ' tisha bav is Shabbos
            If avos_day = 104 Then ' tisha bav
                GetAvos = 0
                Exit Function
            ElseIf avos_day > 104 Then ' after tisha bav
                chapter = chapter - 1
            End If
        ElseIf avos_day_of_week = 5 Then ' tisha bav is Sunday
            If avos_day = 103 Then ' erev tisha bav
                GetAvos = 0
                Exit Function
            ElseIf avos_day > 103 Then ' after tisha bav
                chapter = chapter - 1
            End If
        ElseIf avos_day_of_week = 1 And Not date_in.EY Then ' 2nd day of Shavous is Shabbos in Chutz Laaretz
            If avos_day = 43 Then ' Shavous
                GetAvos = 0
                Exit Function
            ElseIf avos_day > 43 Then ' after Shavous
                chapter = chapter - 1
            End If
        End If

        ' normalize
        chapter = chapter Mod 6
        chapter = chapter + 1

        ' Elul double chapters
        If date_in.month = 6 Then
            If date_in.day > 22 Then
                chapter = 56 ' Last week Chapter 5 - 6
            ElseIf date_in.day > 15 Then
                chapter = 34 ' Chapter 3 - 4
            ElseIf date_in.day > 8 And chapter = 1 Then
                chapter = 12 ' Chapter 1 - 2
            End If
        End If

        GetAvos = chapter
    End If
End Function

' return true if a ta'anis day
Public Function IsTaAnis(date_in As hdate) As Boolean
    Dim current As yomtov
    current = GetYomTov(date_in)
    If current = YOM_KIPPUR Or (current >= SHIVA_ASAR_BTAAMUZ And current <= TAANIS_ESTER) Then
        IsTaAnis = True
    Else
        IsTaAnis = False
    End If
End Function

' return true if assur be-melachah (shabbos, yom tov etc..)
Public Function IsAssurBeMelachah(date_in As hdate) As Boolean
    Dim current As yomtov
    current = GetYomTov(date_in)
    If (date_in.wday = 0) Or (current >= PESACH_DAY1 And current <= SIMCHAS_TORAH) Then
        IsAssurBeMelachah = True
    Else
        IsAssurBeMelachah = False
    End If
End Function

' return 1 if cadlelighting regular, 2 if at nightfall, 3 if chanukah, or 0 if none
Public Function IsCandleLighting(date_in As hdate) As Long
    If date_in.wday = 6 Then
        IsCandleLighting = 1
        Exit Function
    End If

    Dim current As yomtov
    current = GetYomTov(date_in)

    If (current >= EREV_PESACH And current <= EREV_SUKKOS) _
    Or (current = CHOL_HAMOED_PESACH_DAY4 And Not date_in.EY) _
    Or (current = CHOL_HAMOED_PESACH_DAY5 And date_in.EY) _
    Or current = HOSHANA_RABBAH Then
        If date_in.wday = 0 Then
            IsCandleLighting = 2
        Else
            IsCandleLighting = 1
        End If
        Exit Function
    End If

    If current = ROSH_HASHANAH_DAY1 Then
        IsCandleLighting = 2
        Exit Function
    End If

    If current = PESACH_DAY1 _
    Or current = SHVEI_SHEL_PESACH _
    Or current = SHAVOUS_DAY1 _
    Or current = SUKKOS_DAY1 _
    Or current = SHMEINI_ATZERES Then
        If Not date_in.EY Then
            IsCandleLighting = 2
        End If
        Exit Function
    End If

    'TODO
    'If (date_in.month = 9 And date_in.day = 24) _
    'Or (current >= CHANUKAH_DAY1 And current <= CHANUKAH_DAY7) Then
    '    If date_in.wday = 0 Then
    '        IsCandleLighting = 2
    '    Else
    '        IsCandleLighting = 3
    '    End If
    '    Exit Function
    'End If

    IsCandleLighting = 0
End Function

' return true if birchas hachama on date
Public Function IsBirchasHaChama(date_in As hdate) As Boolean
    Dim yearstart As Long
    yearstart = HebrewCalendarElapsedDays(date_in.year)
    Dim day As Long
    day = yearstart + date_in.dayOfYear
    If day Mod 10227 = 172 Then
        IsBirchasHaChama = True
    Else
        IsBirchasHaChama = False
    End If
End Function

' Returns True if the day is Shabbos Mevorchim.
Public Function IsShabbosMevorchim(date_in As hdate) As Boolean
    If date_in.wday <> 0 Then
        IsShabbosMevorchim = False
        Exit Function
    End If
    If date_in.month = 6 Then
        'Elul - no shabbos mevorchim
        IsShabbosMevorchim = False
        Exit Function
    End If
    If date_in.day >= 23 And date_in.day <= 29 Then
        IsShabbosMevorchim = True
    Else
        IsShabbosMevorchim = False
    End If
End Function

Public Function TekufasTishreiElapsedDays(date_in As hdate) As Long
    ' days since Rosh Hashana year 1
    ' add 1/2 day as the first tekufas tishrei was 9 hours into the day
    ' this allows all 4 years of the secular leap year cycle to share 47 days
    ' make from 47D,9H to 47D for simplicity
    Dim days As Double
    days = HebrewCalendarElapsedDays(date_in.year) + (date_in.dayOfYear - 1) + 0.5
    ' days of completed solar years
    Dim solar As Double
    solar = (date_in.year - 1) * 365.25
    TekufasTishreiElapsedDays = CInt(days - solar)
End Function

' return true if birchas hashanim is switched on date (winter)
Public Function IsBirchasHaShanim(date_in As hdate) As Boolean
    If date_in.EY Then
        If date_in.month = 7 And date_in.day = 7 Then
            IsBirchasHaShanim = True
        End If
    ElseIf TekufasTishreiElapsedDays(date_in) = 47 Then
        IsBirchasHaShanim = True
    Else
        IsBirchasHaShanim = False
    End If
End Function

' return true if tal umatar livrachah is said in birchas hashanim
Public Function GetBirchasHaShanim(date_in As hdate) As Boolean
    If date_in.month = 1 And date_in.day < 15 Then
        GetBirchasHaShanim = True
        Exit Function
    End If

    If date_in.month < 7 Then
        GetBirchasHaShanim = False
        Exit Function
    End If

    If date_in.EY Then
        If date_in.month = 7 And date_in.day < 7 Then
            GetBirchasHaShanim = False
        Else
            GetBirchasHaShanim = True
        End If
    Else
        If TekufasTishreiElapsedDays(date_in) < 47 Then
            GetBirchasHaShanim = False
        Else
            GetBirchasHaShanim = True
        End If
    End If
End Function


