Attribute VB_Name = "calendar_utils"
Option Explicit
Public Const NUM_WEEKDAYS As Integer = 7
Public Const NUM_ROWS As Integer = 6
Public Const NUM_COLS As Integer = NUM_WEEKDAYS

'returns multi-dimentional array representing a whole month display with remainders from next & prev months (42 days in total)
'each day item is a string containing the following format:
'<day_date>;<is_in_selected_month>;<first_date_for_selected_month>
Public Function calendar_utils_get_month_matrix(ByVal sel_month As Integer, ByVal sel_year As Integer, is_hebrew As Boolean) As Variant
    Dim currentDay As Date
    Dim startDate As Date
    Dim hebrewDate As hdate
    
    Dim checkdate As hdate
    Dim currentDate As Date
    Dim firstDayOfMonth As Date
    Dim lastDayOfMonth As Date
    
    Dim I As Integer, J As Integer
    Dim row As Integer, col As Integer
    Dim matrix(NUM_ROWS, NUM_COLS) As String
    On Error Resume Next
    
    ' Get the selected date
    If is_hebrew Then
        hebrewDate = HDateNew(0, 0, 0, 8, 30, 0, 0, 0) 'HDateNew(CLng(sel_year), CLng(sel_month), 1, 1, 1, 1, 1, 0)
        Call HDateAddYear(hebrewDate, CLng(sel_year - 1))
        If sel_month < 7 Then Call HDateAddYear(hebrewDate, CLng(1))
        Call HDateAddMonth(hebrewDate, CLng(sel_month))
        Call HDateAddDay(hebrewDate, 1)
        hebrewDate.offset = 3600 * (2 + IsDST(currentDate))
        Call SetEY(hebrewDate, 1)
        currentDate = (HDateGregorian(hebrewDate))
    Else
        currentDate = DateSerial(sel_year, sel_month, 1)
        hebrewDate = (ConvertDate(currentDate))
        hebrewDate.offset = 3600 * (2 + IsDST(currentDate))
        Call SetEY(hebrewDate, 1)
    End If
    
    ' Get the first day of the month (as a gregorian date)
    If is_hebrew Then
        checkdate = hebrewDate
        Call HDateAddDay(checkdate, 1 - hebrewDate.day)
        firstDayOfMonth = (HDateGregorian(checkdate))
    Else
        firstDayOfMonth = DateSerial(year(currentDate), month(currentDate), 1)
    End If

    ' Get the last day of the month (as a gregorian date)
    If is_hebrew Then
        checkdate = hebrewDate
        Call HDateAddDay(checkdate, LastDayOfHebrewMonth(hebrewDate.month, hebrewDate.year) - hebrewDate.day)
        lastDayOfMonth = (HDateGregorian(checkdate))
    Else
        lastDayOfMonth = DateSerial(year(currentDate), month(currentDate) + 1, 0)
    End If
    
    startDate = firstDayOfMonth
    
    ' Populate calendar with dates
    currentDay = startDate
        'add days from previous month's last sunday, when applicable
    col = Weekday(currentDay) - 1
    For I = col To 0 Step -1
        matrix(row, I) = DateSerial(year(startDate), month(startDate), day(startDate)) & ";0;" & firstDayOfMonth
        startDate = DateAdd("d", -1, startDate)
    Next I
        'add selected month's days
    startDate = currentDay
    Do
        matrix(row, col) = DateSerial(year(currentDay), month(currentDay), day(currentDay)) & ";1;" & firstDayOfMonth
        col = (col + 1) Mod NUM_COLS ' Increment column counter, wrapping around at NUM_COLS
        If col = 0 Then
            row = row + 1 ' Move to the next row after completing a week
        End If
        currentDay = DateAdd("d", 1, currentDay) ' Move to the next day
    Loop Until currentDay > lastDayOfMonth
        'add first days from next month to fill the remained days
    For I = col To 6
        matrix(row, I) = DateSerial(year(currentDay), month(currentDay), day(currentDay)) & ";0;" & firstDayOfMonth
        currentDay = DateAdd("d", 1, currentDay)
    Next I
    If row < 5 Then
        For I = 0 To 6
            matrix(row + 1, I) = DateSerial(year(currentDay), month(currentDay), day(currentDay)) & ";0;" & firstDayOfMonth
            currentDay = DateAdd("d", 1, currentDay)
        Next I
    End If

        'pass as Variant to overcome VB limitation
    calendar_utils_get_month_matrix = CVar(matrix)
End Function

'get molad string for hebrew month of selected date (might also return for the month after)
Function calendar_utils_get_month_molad_info(date_in As Date) As String
    Dim hebrewDate As hdate
    Dim checkdate As hdate
    Dim currentDate As Date
    Dim firstDayOfMonth As Date
    Dim lastDayOfMonth As Date
    Dim molad As hdate
    
    Dim is_hebrew As Boolean
    Dim sel_year As Integer
    Dim sel_month As Integer
    is_hebrew = False
    sel_year = year(date_in)
    sel_month = month(date_in)
    
    ' Get the selected date
    If is_hebrew Then
        hebrewDate = HDateNew(0, 0, 0, 8, 30, 0, 0, 0) 'HDateNew(CLng(sel_year), CLng(sel_month), 1, 1, 1, 1, 1, 0)
        Call HDateAddYear(hebrewDate, CLng(sel_year))
        Call HDateAddMonth(hebrewDate, CLng(sel_month))
        Call HDateAddDay(hebrewDate, 1)
        hebrewDate.offset = 3600 * (2 + IsDST(currentDate))
        Call SetEY(hebrewDate, 1)
        currentDate = (HDateGregorian(hebrewDate))
    Else
        currentDate = DateSerial(sel_year, sel_month, 1)
        hebrewDate = (ConvertDate(currentDate))
        hebrewDate.offset = 3600 * (2 + IsDST(currentDate))
        Call SetEY(hebrewDate, 1)
    End If

    ' Get the first day of the month (as a gregorian date)
    If is_hebrew Then
        checkdate = hebrewDate
        Call HDateAddDay(checkdate, 1 - hebrewDate.day)
        firstDayOfMonth = (HDateGregorian(checkdate))
    Else
        firstDayOfMonth = DateSerial(year(currentDate), month(currentDate), 1)
    End If

    ' Get the last day of the month (as a gregorian date)
    If is_hebrew Then
        checkdate = hebrewDate
        Call HDateAddDay(checkdate, LastDayOfHebrewMonth(hebrewDate.month, hebrewDate.year) - hebrewDate.day)
        lastDayOfMonth = (HDateGregorian(checkdate))
    Else
        lastDayOfMonth = DateSerial(year(currentDate), month(currentDate) + 1, 0)
    End If

    'if selected cell is not a date then show molad info
    checkdate = ConvertDate(DateAdd("d", 15, firstDayOfMonth))
    molad = GetMolad(checkdate.year, checkdate.month)
    calendar_utils_get_month_molad_info = "מולד חודש " & NumToHMonth(checkdate.month, checkdate.leap) & ":" _
    & vbCrLf & MoladFormat(molad)
    
    hebrewDate = ConvertDate(lastDayOfMonth)
    'if relevant, show molad for next chodesh as well (when its roch_chodesh or shabbos mevorchim appear on current view)
    If hebrewDate.month <> checkdate.month Or GetRoshChodesh(hebrewDate) = ROSH_CHODESH Or IsShabbosMevorchim(hebrewDate) Then
        If hebrewDate.month = checkdate.month Then Call HDateAddMonth(hebrewDate, 1)
        checkdate = hebrewDate
        molad = GetMolad(checkdate.year, checkdate.month)
        calendar_utils_get_month_molad_info = calendar_utils_get_month_molad_info & vbCrLf & vbCrLf & "מולד חודש " & NumToHMonth(checkdate.month, checkdate.leap) & ":" _
        & vbCrLf & MoladFormat(molad)
    End If

End Function

'get day info string suited for calendar display, including:
'hebrew & loazi dates
'4 candle lighting times, when applicable
'sefirat haomer, when applicable
'day title (shabbos-parsha, moed, taanis etc)
Function calendar_utils_get_day_info(date_in As Date) As String
    Dim hebrewDate As hdate
    Dim checkdate As hdate
    Dim currentDay As Date
    Dim shabbosstr As String
    Dim mozashstr As String
    Dim parsh As parshah
    Dim daystr As String
    Dim limudstr As String
    Dim hereJ As location
    Dim hereT As location
    Dim hereH As location
    Dim hereB As location
    Dim erevshabbos As hdate
    Dim ytov As yomtov
    
    'init locations
    With hereJ ' =Yerushalayim (CL - 40 mins)
    .latitude = 31.788
    .longitude = 35.218
    .elevation = 800
    End With
    
    With hereT ' =Tel Aviv (CL - 22 mins)
    .latitude = 32.06
    .longitude = 34.77
    .elevation = 20
    End With
    
    With hereH ' =Haifa (CL - 30 mins)
    .latitude = 32.8
    .longitude = 34.991
    .elevation = 300
    End With
    
    With hereB ' =Beer Sheva (CL - 20 mins)
    .latitude = 31.24
    .longitude = 34.79
    .elevation = 0
    End With
        
    currentDay = date_in
    
    'get hebrew date of current date
    hebrewDate = (ConvertDate(currentDay))
    hebrewDate.offset = 3600 * (2 + IsDST(currentDay))
    Call SetEY(hebrewDate, 1)
    checkdate = hebrewDate
    
    'when current day is shabbos, then get its details (parsha, shabbos mevorchim)
    shabbosstr = ""
    parsh = GetParshah(hebrewDate)
    If parsh <> NOPARSHAH Then
        shabbosstr = "שבת פרשת " & ParshahFormat(parsh)
    Else
        'if not a standard shabbos, check if is a yomtov / moed and get its title, if relevant
        ytov = GetYomTov(hebrewDate)
        If ytov <> CHOL Then shabbosstr = YomTovFormat(ytov)
    End If
    ytov = GetSpecialShabbos(hebrewDate)
    shabbosstr = shabbosstr & IIf(ytov <> CHOL, vbCrLf & "(" & YomTovFormat(ytov) & ")", "")
    If IsShabbosMevorchim(hebrewDate) Then shabbosstr = shabbosstr & vbCrLf & "שבת מברכים"
    
    'collect date details for currentDay:
    '1. hebrew and loazi dates
    daystr = Format(currentDay, "d mmmm") & vbCrLf & NumToHChar(hebrewDate.day) & " " & NumToHMonth(hebrewDate.month, hebrewDate.leap) & vbCrLf & shabbosstr
    '2. sefirat haomer, if relevant
    If GetOmer(hebrewDate) Then daystr = daystr & vbCrLf & "(" & (GetOmer(hebrewDate)) & " בעומר)"
    '3. rosh chodesh, if relevant
    Call HDateAddDay(checkdate, 1)
    If GetRoshChodesh(hebrewDate) = ROSH_CHODESH Then daystr = daystr & vbCrLf & "ראש חודש " & IIf(hebrewDate.day = 1, NumToHMonth(hebrewDate.month, hebrewDate.leap), NumToHMonth(checkdate.month, checkdate.leap))
    '4. candle lighting times for erev shabbos / yomtov / 2nd yomtov
    If IsCandleLighting(hebrewDate) = 1 Then
        'in case the CL time is before shekiah
        daystr = daystr & vbCrLf & "הדלקת נרות:" & vbCrLf & _
        "ירושלים " & Format(DateAdd("n", -40, tround((HDateGregorian(getelevationsunset(hebrewDate, hereJ))))), "hh:mm") & vbCrLf & _
        "תל אביב " & Format(DateAdd("n", -22, tround((HDateGregorian(getelevationsunset(hebrewDate, hereT))))), "hh:mm") & vbCrLf & _
        "חיפה " & Format(DateAdd("n", -30, tround((HDateGregorian(getelevationsunset(hebrewDate, hereH))))), "hh:mm") & vbCrLf & _
        "באר שבע " & Format(DateAdd("n", -20, tround((HDateGregorian(getelevationsunset(hebrewDate, hereB))))), "hh:mm")
    ElseIf IsCandleLighting(hebrewDate) = 2 Then
        'in case CL time is after zais (e.g. after shabbos or before 2nd yomtov)
        daystr = daystr & vbCrLf & "הדלקת נרות (צה""כ):" & vbCrLf & _
        "ירושלים " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereJ)))), "hh:mm") & vbCrLf & _
        "תל אביב " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereT)))), "hh:mm") & vbCrLf & _
        "חיפה " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereH)))), "hh:mm") & vbCrLf & _
        "באר שבע " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereB)))), "hh:mm")
    ElseIf IsAssurBeMelachah(hebrewDate) And Not IsAssurBeMelachah(checkdate) Then
        'havdalah times, if relevant
        'get type of day (shabbos, yomtov, kippur)
        ytov = GetYomTov(hebrewDate)
        If ytov = YOM_KIPPUR Then
            mozashstr = "מוצאי יוה""כ:"
        ElseIf hebrewDate.wday = 0 Then
            mozashstr = "מוצאי שבת:"
        Else
            mozashstr = "מוצאי יו""ט:"
        End If
        'get havdalah times
        daystr = daystr & vbCrLf & mozashstr & vbCrLf _
        & "ירושלים " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereJ)))), "hh:mm") & vbCrLf _
        & "תל אביב " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereT)))), "hh:mm") & vbCrLf _
        & "חיפה " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereH)))), "hh:mm") & vbCrLf _
        & "באר שבע " & Format(tround((HDateGregorian(gettzais8p5(hebrewDate, hereB)))), "hh:mm")
    Else
    
    End If
            
    calendar_utils_get_day_info = daystr
End Function

'get common zmanim info string for selected date & location
Function calendar_utils_get_zmanim_info(date_in As Date, here As location) As String

    Dim hebrewDate As hdate
    Dim currentDay As Date
    Dim daystr As String


    currentDay = date_in
    hebrewDate = (ConvertDate(currentDay))
    hebrewDate.offset = 3600 * (2 + IsDST(currentDay))
    Call SetEY(hebrewDate, 1)

    'get zmanin info for the selected day and location
    daystr = HDateFormat(hebrewDate) & vbCrLf & vbCrLf & _
        "עלות השחר (72 דק'): " & " " & Format((HDateGregorian(getalos72(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "הנץ החמה: " & " " & Format((HDateGregorian(getsunrise(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "סוף זמן קריאת שמע (מג""א): " & " " & Format((HDateGregorian(getshmamga(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "סוף זמן קריאת שמע (גר""א): " & " " & Format((HDateGregorian(getshmagra(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "סוף זמן תפילה (מג""א): " & " " & Format((HDateGregorian(gettefilamga(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "סוף זמן תפילה (גר""א): " & " " & Format((HDateGregorian(gettefilagra(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "חצות היום: " & " " & Format((HDateGregorian(getchatzosgra(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "מנחה גדולה: " & " " & Format((HDateGregorian(getminchagedolagra(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "מנחה קטנה: " & " " & Format((HDateGregorian(getminchaketanagra(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "שקיעה: " & " " & Format((HDateGregorian(getelevationsunset(hebrewDate, here))), "hh:mm:ss") & vbCrLf & _
        "צאת הכוכבים: " & " " & Format((HDateGregorian(gettzais8p5(hebrewDate, here))), "hh:mm:ss")

    calendar_utils_get_zmanim_info = daystr

End Function

'get daily limud string for selected date
Function calendar_utils_get_limud_info(date_in As Date) As String
    Dim limudstr As String
    Dim hebrewDate As hdate
    Dim currentDay As Date

    currentDay = date_in
    hebrewDate = (ConvertDate(currentDay))
    hebrewDate.offset = 3600 * (2 + IsDST(currentDay))
    Call SetEY(hebrewDate, 1)
    
    limudstr = limudstr & "דף יומי בבלי: " & vbCrLf & GetDafYomiFormat((HDateGregorian(hebrewDate))) & vbCrLf
    limudstr = limudstr & vbCrLf & "דף יומי ירושלמי: " & vbCrLf & GetDafYomiFormat((HDateGregorian(hebrewDate)), True) & vbCrLf
    limudstr = limudstr & vbCrLf & "משנה יומית: " & vbCrLf & GetMishnaYomiFormat((HDateGregorian(hebrewDate))) & vbCrLf
    limudstr = limudstr & vbCrLf & "הלכה יומית: " & vbCrLf & GetHalacha(hebrewDate) & vbCrLf
    limudstr = limudstr & vbCrLf & "פרק רמבם יומי: " & vbCrLf & Replace(GetRambam(hebrewDate, 1), ";", ", ") & vbCrLf
    limudstr = limudstr & vbCrLf & "רמבם יומי (3 פרקים): " & vbCrLf & Replace(GetRambam(hebrewDate, 0), ";", ", ") & vbCrLf
    limudstr = limudstr & vbCrLf & " תניא יומי: " & vbCrLf & GetTanya(hebrewDate)
    calendar_utils_get_limud_info = limudstr
    
End Function

