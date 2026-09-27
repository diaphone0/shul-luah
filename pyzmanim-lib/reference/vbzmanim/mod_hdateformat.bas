Attribute VB_Name = "mod_hdateformat"
Public hchar(24) As String
Public hmonth(13) As String
Public hwday(7) As String
Public parshahchar(61) As String
Public hdateformat_ready As Boolean

Public Sub init_hdateformat()
    hchar(0) = "?"
    hchar(1) = "א"
    hchar(2) = "ב"
    hchar(3) = "ג"
    hchar(4) = "ד"
    hchar(5) = "ה"
    hchar(6) = "ו"
    hchar(7) = "ז"
    hchar(8) = "ח"
    hchar(9) = "ט"
    hchar(10) = "י"
    hchar(11) = "כ"
    hchar(12) = "ל"
    hchar(13) = "מ"
    hchar(14) = "נ"
    hchar(15) = "ס"
    hchar(16) = "ע"
    hchar(17) = "פ"
    hchar(18) = "צ"
    hchar(19) = "ק"
    hchar(20) = "ר"
    hchar(21) = "ש"
    hchar(22) = "ת"
    hchar(23) = "״"
    hchar(24) = "׳"
    hmonth(0) = "אדר א׳"
    hmonth(1) = "ניסן"
    hmonth(2) = "אייר"
    hmonth(3) = "סיון"
    hmonth(4) = "תמוז"
    hmonth(5) = "אב"
    hmonth(6) = "אלול"
    hmonth(7) = "תשרי"
    hmonth(8) = "חשון"
    hmonth(9) = "כסלו"
    hmonth(10) = "טבת"
    hmonth(11) = "שבט"
    hmonth(12) = "אדר"
    hmonth(13) = "אדר ב׳"
    hwday(0) = "שביעי"
    hwday(1) = "ראשון"
    hwday(2) = "שני"
    hwday(3) = "שלישי"
    hwday(4) = "רביעי"
    hwday(5) = "חמישי"
    hwday(6) = "שישי"
    hwday(7) = "שבת"
    parshahchar(0) = ""
    parshahchar(1) = "בראשית"
    parshahchar(2) = "נח"
    parshahchar(3) = "לך לך"
    parshahchar(4) = "וירא"
    parshahchar(5) = "חיי שרה"
    parshahchar(6) = "תולדות"
    parshahchar(7) = "ויצא"
    parshahchar(8) = "וישלח"
    parshahchar(9) = "וישב"
    parshahchar(10) = "מקץ"
    parshahchar(11) = "ויגש"
    parshahchar(12) = "ויחי"
    parshahchar(13) = "שמות"
    parshahchar(14) = "וארא"
    parshahchar(15) = "בא"
    parshahchar(16) = "בשלח"
    parshahchar(17) = "יתרו"
    parshahchar(18) = "משפטים"
    parshahchar(19) = "תרומה"
    parshahchar(20) = "תצוה"
    parshahchar(21) = "כי תשא"
    parshahchar(22) = "ויקהל"
    parshahchar(23) = "פקודי"
    parshahchar(24) = "ויקרא"
    parshahchar(25) = "צו"
    parshahchar(26) = "שמיני"
    parshahchar(27) = "תזריע"
    parshahchar(28) = "מצורע"
    parshahchar(29) = "אחרי מות"
    parshahchar(30) = "קדושים"
    parshahchar(31) = "אמור"
    parshahchar(32) = "בהר"
    parshahchar(33) = "בחוקותי"
    parshahchar(34) = "במדבר"
    parshahchar(35) = "נשא"
    parshahchar(36) = "בהעלותך"
    parshahchar(37) = "שלח"
    parshahchar(38) = "קרח"
    parshahchar(39) = "חקת"
    parshahchar(40) = "בלק"
    parshahchar(41) = "פינחס"
    parshahchar(42) = "מטות"
    parshahchar(43) = "מסעי"
    parshahchar(44) = "דברים"
    parshahchar(45) = "ואתחנן"
    parshahchar(46) = "עקב"
    parshahchar(47) = "ראה"
    parshahchar(48) = "שופטים"
    parshahchar(49) = "כי תצא"
    parshahchar(50) = "כי תבוא"
    parshahchar(51) = "נצבים"
    parshahchar(52) = "וילך"
    parshahchar(53) = "האזינו"
    parshahchar(54) = "וזאת הברכה"
    parshahchar(55) = "ויקהל - פקודי"
    parshahchar(56) = "תזריע - מצורע"
    parshahchar(57) = "אחרי מות - קדושים"
    parshahchar(58) = "בהר - בחוקותי"
    parshahchar(59) = "חקת - בלק"
    parshahchar(60) = "מטות - מסעי"
    parshahchar(61) = "נצבים - וילך"
    
    
    
    hdateformat_ready = True
End Sub

Public Function ParshahFormat(current As parshah) As String
    If hdateformat_ready = False Then init_hdateformat
    
    ParshahFormat = parshahchar(current)
End Function

Function GetHChar(ByVal num As Integer) As String
    Dim hchar1 As String
    If hdateformat_ready = False Then init_hdateformat
    Select Case num
        Case 1
            hchar1 = hchar(1)
        Case 2
            hchar1 = hchar(2)
        Case 3
            hchar1 = hchar(3)
        Case 4
            hchar1 = hchar(4)
        Case 5
            hchar1 = hchar(5)
        Case 6
            hchar1 = hchar(6)
        Case 7
            hchar1 = hchar(7)
        Case 8
            hchar1 = hchar(8)
        Case 9
            hchar1 = hchar(9)
        Case 10
            hchar1 = hchar(10)
        Case 20
            hchar1 = hchar(11)
        Case 30
            hchar1 = hchar(12)
        Case 40
            hchar1 = hchar(13)
        Case 50
            hchar1 = hchar(14)
        Case 60
            hchar1 = hchar(15)
        Case 70
            hchar1 = hchar(16)
        Case 80
            hchar1 = hchar(17)
        Case 90
            hchar1 = hchar(18)
        Case 100
            hchar1 = hchar(19)
        Case 200
            hchar1 = hchar(20)
        Case 300
            hchar1 = hchar(21)
        Case 400
            hchar1 = hchar(22)
        Case 99
            hchar1 = hchar(23)
        Case 999
            hchar1 = hchar(24)
        Case Else
            hchar1 = hchar(0)
    End Select
    GetHChar = hchar1
End Function

Function AddChar(ByVal year As String, ByVal charnum As Integer, ByRef num As Integer, ByRef counter As Long, ByVal limit As Long) As String
    Dim charvalue As Integer
    If hdateformat_ready = False Then init_hdateformat
    charvalue = charnum
    If charvalue = 99 Or charvalue = 999 Then charvalue = 0
    Dim Len1 As Long
    Len1 = limit - counter
    Dim endPos As Long
    endPos = counter + IIf(Len1 > 2, 2, Len1)
    Dim endStr As String
    endStr = Left$(year, counter) & GetHChar(charnum) & Mid$(year, endPos + 1)
    counter = counter + IIf(Len1 > 1, 1, Len1)
    num = num - charvalue
    AddChar = endStr
End Function

'convert an int to a Hebrew char based representation. 5779 becomes תשע"ט
Public Function NumToHChar(ByVal innum As Integer) As String
    Dim num As Integer
    num = innum
    Dim year As String
    Dim counter As Long
    If hdateformat_ready = False Then init_hdateformat
    year = String(13, vbNullChar)
    counter = 0
    If num >= 1000 And num <= 10000 And num Mod 1000 = 0 Then
        year = AddChar(year, num \ 1000, num, counter, 13)
        year = AddChar(year, 999, num, counter, 13)
        NumToHChar = Left$(year, counter)
        Exit Function
    End If
    If num >= 1000 And num <= 10000 Then num = num Mod 1000
    Do While num > 0 And counter < 13
        If num = 15 Or num = 16 Then
            year = AddChar(year, 9, num, counter, 13)
            year = AddChar(year, 99, num, counter, 13)
            year = AddChar(year, num, num, counter, 13)
            NumToHChar = Left$(year, counter)
            Exit Function
        ElseIf num < 10 Or (num < 100 And num Mod 10 = 0) Or (num < 500 And num Mod 100 = 0) Then
            If counter <> 0 Then year = AddChar(year, 99, num, counter, 13)
            year = AddChar(year, num, num, counter, 13)
            NumToHChar = Left$(year, counter)
            If innum < 11 Then NumToHChar = NumToHChar + "'"
            Exit Function
        ElseIf num > 400 Then
            year = AddChar(year, 400, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num > 300 Then
            year = AddChar(year, 300, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num > 200 Then
            year = AddChar(year, 200, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num > 100 Then
            year = AddChar(year, 100, num, counter, 13)
            GoTo ContinueLoop
        ElseIf num \ 10 > 0 Then
            year = AddChar(year, num - (num Mod 10), num, counter, 13)
            GoTo ContinueLoop
        End If
ContinueLoop:
    Loop
    NumToHChar = Left$(year, counter)
    If innum < 11 Then NumToHChar = NumToHChar + "'"
End Function

'convert int based Hebrew weekday (hdate.wday) to char based representation.
'second argument is a booean if to use שבת (true) or שביעי (false)
Function NumToWDay(date_in As hdate, ByVal shabbos As Boolean) As String
    If hdateformat_ready = False Then init_hdateformat
    If shabbos And date_in.wday = 0 Then
        NumToWDay = hwday(7)
    Else
        NumToWDay = hwday(date_in.wday)
    End If
End Function

'convert int based Hebrew month (hdate.month) to char based representation.
Function NumToHMonth(ByVal month As Integer, ByVal leap As Integer) As String
    If hdateformat_ready = False Then init_hdateformat
    If leap <> 0 Then
        If month = 12 Then
            NumToHMonth = hmonth(0)
            Exit Function
        ElseIf month = 13 Then
            NumToHMonth = hmonth(month)
            Exit Function
        End If
    End If
    If month > 0 And month < 13 Then
        NumToHMonth = hmonth(month)
    Else
        NumToHMonth = vbNullString
    End If
End Function

'convert hdate to string based representation
Function HDateFormat(date_in As hdate) As String
    Dim day As String
    Dim year As String
    If hdateformat_ready = False Then init_hdateformat
    day = NumToHChar(date_in.day)
    Dim month As String
    month = NumToHMonth(date_in.month, date_in.leap)
    year = NumToHChar(date_in.year)
    HDateFormat = day & " " & month & " " & year
End Function

'convert hdate to string based representation, with evening consideration;
'if the time is between zais to sunrise of next day then add "אור ל-" for next day
Function HDateOrFormat(date_in As hdate, here As location) As String
    Dim day As String
    Dim year As String
    Dim date_next As hdate
    Dim date_before As hdate
    Dim is_or As Boolean
    Dim sunset_today As Date
    Dim sunset_yesterday As Date
    Dim current_date As Date
    Dim sunrise_tomorrow As Date
    Dim date_result As hdate
    date_next = date_in
    date_before = date_in
    current_date = (HDateGregorian(date_in))
    Call HDateAddDay(date_next, 1)
    Call HDateAddDay(date_before, -1)
    sunset_today = (HDateGregorian(gettzais8p5(date_in, here)))
    sunrise_today = (HDateGregorian(getsunrise(date_in, here)))
    sunrise_tomorrow = (HDateGregorian(getsunrise(date_next, here)))
    sunset_yesterday = (HDateGregorian(gettzais8p5(date_before, here)))
    If current_date >= sunset_today Then
        is_or = True
        date_result = date_next
    Else
        date_result = date_in
    End If
    If current_date < sunrise_today Then is_or = True
    If hdateformat_ready = False Then init_hdateformat
    day = NumToHChar(date_result.day)
    Dim month As String
    month = NumToHMonth(date_result.month, date_result.leap)
    year = NumToHChar(date_result.year)
    HDateOrFormat = day & " " & month & " " & year
    If is_or Then HDateOrFormat = "אור ל-" & HDateOrFormat
End Function

'convert hdate holding molad info to string representation, suitable for molad announcement
Function MoladFormat(molad As hdate, Optional full_date As Boolean = True) As String
    If full_date Then
        MoladFormat = "יום " & NumToWDay(molad, True) & ", " & NumToHChar(molad.day) & " " & NumToHMonth(molad.month, molad.leap) & ", שעה " & Format(TimeSerial(molad.hour, molad.min, 0), "hh:mm") & " ו-" & molad.sec & " חלקים"
    Else
        MoladFormat = "יום " & NumToWDay(molad, True) & " שעה " & Format(TimeSerial(molad.hour, molad.min, 0), "hh:mm") & " ו-" & molad.sec & " חלקים"
    End If
End Function

'convert yomtov enum to char based representation
Function YomTovFormat(ByVal current As yomtov) As String
    If hdateformat_ready = False Then init_hdateformat
    Select Case current
        Case CHOL
            YomTovFormat = ""
        Case PESACH_DAY1, PESACH_DAY2
            YomTovFormat = "פסח"
        Case SHVEI_SHEL_PESACH
            YomTovFormat = "שביעי של פסח"
        Case ACHRON_SHEL_PESACH
            YomTovFormat = "אחרון של פסח"
        Case SHAVOUS_DAY1, SHAVOUS_DAY2
            YomTovFormat = "שבועות"
        Case ROSH_HASHANAH_DAY1, ROSH_HASHANAH_DAY2
            YomTovFormat = "ראש השנה"
        Case YOM_KIPPUR
            YomTovFormat = "יום כיפור"
        Case SUKKOS_DAY1, SUKKOS_DAY2
            YomTovFormat = "סוכות"
        Case SHMEINI_ATZERES
            YomTovFormat = "שמיני עצרת"
        Case SIMCHAS_TORAH
            YomTovFormat = "שמחת תורה"
        Case CHOL_HAMOED_PESACH_DAY1 To CHOL_HAMOED_PESACH_DAY5
            YomTovFormat = "חול המועד פסח"
        Case CHOL_HAMOED_SUKKOS_DAY1 To CHOL_HAMOED_SUKKOS_DAY5
            YomTovFormat = "חול המועד סוכות"
        Case HOSHANA_RABBAH
            YomTovFormat = "הושענא רבה"
        Case PESACH_SHEINI
            YomTovFormat = "פסח שני"
        Case LAG_BAOMER
            YomTovFormat = "ל״ג בעומר"
        Case TU_BAV
            YomTovFormat = "ט״ו באב"
        Case CHANUKAH_DAY1 To CHANUKAH_DAY8
            YomTovFormat = "חנוכה"
        Case TU_BISHVAT
            YomTovFormat = "ט״ו בשבט"
        Case PURIM_KATAN
            YomTovFormat = "פורים קטן"
        Case SHUSHAN_PURIM_KATAN
            YomTovFormat = "שושן פורים קטן"
        Case PURIM
            YomTovFormat = "פורים"
        Case SHUSHAN_PURIM
            YomTovFormat = "שושן פורים"
        Case SHIVA_ASAR_BTAAMUZ
            YomTovFormat = "שבעה עשר בתמוז"
        Case TISHA_BAV
            YomTovFormat = "ט״ב"
        Case TZOM_GEDALIA
            YomTovFormat = "צום גדליה"
        Case ASARAH_BTEVES
            YomTovFormat = "עשרה בטבת"
        Case TAANIS_ESTER
            YomTovFormat = "תענית אסתר"
        Case EREV_PESACH
            YomTovFormat = "ערב פסח"
        Case EREV_SHAVOUS
            YomTovFormat = "ערב שבועות"
        Case EREV_ROSH_HASHANAH
            YomTovFormat = "ערב ראש השנה"
        Case EREV_YOM_KIPPUR
            YomTovFormat = "ערב יום כיפור"
        Case EREV_SUKKOS
            YomTovFormat = "ערב סוכות"
        Case SHKALIM
            YomTovFormat = "פרשת שקלים"
        Case ZACHOR
            YomTovFormat = "פרשת זכור"
        Case PARAH:
            YomTovFormat = "פרשת פרה"
        Case HACHODESH:
            YomTovFormat = "פרשת החודש"
        Case ROSH_CHODESH:
            YomTovFormat = "ראש חודש"
        Case MACHAR_CHODESH:
            YomTovFormat = "מחר חודש"
        Case SHABBOS_MEVORCHIM:
            YomTovFormat = "שבת מברכים"
        Case HAGADOL:
            YomTovFormat = "שבת הגדול"
        Case CHAZON:
            YomTovFormat = "שבת חזון"
        Case NACHAMU:
            YomTovFormat = "שבת נחמו"
        Case SHUVA:
            YomTovFormat = "שבת שובה"
        Case SHIRA:
            YomTovFormat = "שבת שירה"
        Case SHABBOS_CHOL_HAMOED:
            YomTovFormat = "שבת חול המועד"
        Case Else
            YomTovFormat = ""
    End Select
End Function

'convert avos int to char based representation
Function AvosFormat(ByVal avos As Integer) As String
    If hdateformat_ready = False Then init_hdateformat
    Select Case avos
        Case 1
            AvosFormat = "א"
        Case 2
            AvosFormat = "ב"
        Case 3
            AvosFormat = "ג"
        Case 4
            AvosFormat = "ד"
        Case 5
            AvosFormat = "ה"
        Case 6
            AvosFormat = "ו"
        Case 12
            AvosFormat = "א-ב"
        Case 34
            AvosFormat = "ג-ד"
        Case 56
            AvosFormat = "ה-ו"
        Case Else
            AvosFormat = ""
    End Select
End Function

'rounds seconds of a time to an added or substracted minute
Function tround(t As Date) As Date
    tround = IIf(second(t) > 29, DateAdd("n", 1, t), t)
End Function

