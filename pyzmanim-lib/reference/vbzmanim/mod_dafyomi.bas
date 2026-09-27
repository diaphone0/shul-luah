Attribute VB_Name = "mod_dafyomi"
'dafyomi list & functions ported from https://github.com/NykUser/MyZman/

Public mishnayos_arr(63) As String
Public masechtosBavli(39) As String
Public masechtosBavliTransliterated(39) As String
Public masechtosYerushalmi(38) As String
Public masechtosYerushalmiTransliterated(38) As String
Public blattPerMasechta(39) As Integer
Public blattPerMasechtaYerushalmi(38) As Integer

Public Type Daf
    masechtaNumber As Integer
    Page As Integer 'used as mishnah for mishnah yomi
    HasSecondaryMesechta As Boolean
    SecondaryMesechtaNumber As Integer
    PerekNumber As Integer ' used as perek for mishnah yomi
End Type
Private dafYomiStartDate As Date
Private shekalimChangeDate As Date
Dim dafyom_ready As Boolean
Dim moshnayos_ready As Boolean
Sub init_mishnayos()
    mishnayos_arr(0) = "ברכות,berachos,5,8,6,7,5,8,5,8,5"
    mishnayos_arr(1) = "פאה,peah,6,8,8,11,8,11,8,9"
    mishnayos_arr(2) = "דמאי,demai,4,5,6,7,11,12,8"
    mishnayos_arr(3) = "כלאיים,kilayim,9,11,7,9,8,9,8,6,10"
    mishnayos_arr(4) = "שביעית,sheviis,8,10,10,10,9,6,7,11,9,9"
    mishnayos_arr(5) = "תרומות,terumos,10,6,9,13,9,6,7,12,7,12,10"
    mishnayos_arr(6) = "מעשרות,maasros,8,8,10,6,8"
    mishnayos_arr(7) = "מעשר שני,maaser_sheni,7,10,13,12,15"
    mishnayos_arr(8) = "חלה,chalah,9,8,10,11"
    mishnayos_arr(9) = "ערלה,orlah,9,17,9"
    mishnayos_arr(10) = "ביכורים,bikurim,11,11,12,5"
    mishnayos_arr(11) = "שבת,shabbos,11,7,6,2,4,10,4,7,7,6,6,6,7,4,3,8,8,3,6,5,3,6,5,5"
    mishnayos_arr(12) = "עירובין,eruvin,10,6,9,11,9,10,11,11,4,15"
    mishnayos_arr(13) = "פסחים,pesachim,7,8,8,9,10,6,13,8,11,9"
    mishnayos_arr(14) = "שקלים,shekalim,7,5,4,9,6,6,7,8"
    mishnayos_arr(15) = "יומא,yoma,8,7,11,6,7,8,5,9"
    mishnayos_arr(16) = "סוכה,sukkah,11,9,15,10,8"
    mishnayos_arr(17) = "ביצה,beitzah,10,10,8,7,7"
    mishnayos_arr(18) = "ראש השנה,rosh_hashanah,9,8,9,9"
    mishnayos_arr(19) = "תענית,taanis,7,10,9,8"
    mishnayos_arr(20) = "מגילה,megillah,11,6,6,10"
    mishnayos_arr(21) = "מועד קטן,moed_katan,10,5,9"
    mishnayos_arr(22) = "חגיגה,chagigah,8,7,8"
    mishnayos_arr(23) = "יבמות,yevamos,4,10,10,13,6,6,6,6,6,9,7,6,13,9,10,7"
    mishnayos_arr(24) = "כתובות,kesubos,10,10,9,12,9,7,10,8,9,6,6,4,11"
    mishnayos_arr(25) = "נדרים,nedarim,4,5,11,8,6,10,9,7,10,8,12"
    mishnayos_arr(26) = "נזיר,nazir,7,10,7,7,7,11,4,2,5"
    mishnayos_arr(27) = "סוטה,sotah,9,6,8,5,5,4,8,7,15"
    mishnayos_arr(28) = "גיטין,gitin,6,7,8,9,9,7,9,10,10"
    mishnayos_arr(29) = "קידושין,kiddushin,10,10,13,14"
    mishnayos_arr(30) = "בבא קמא,bava_kamma,4,6,11,9,7,6,7,7,12,10"
    mishnayos_arr(31) = "בבא מציעא,bava_metzia,8,11,12,12,11,8,11,9,13,6"
    mishnayos_arr(32) = "בבא בתרא,bava_basra,6,14,8,9,11,8,4,8,10,8"
    mishnayos_arr(33) = "סנהדרין,sanhedrin,6,5,8,5,5,6,11,7,6,6,6"
    mishnayos_arr(34) = "מכות,makkos,10,8,16"
    mishnayos_arr(35) = "שבועות,shevuos,7,5,11,13,5,7,8,6"
    mishnayos_arr(36) = "עדויות,eduyos,14,10,12,12,7,3,9,7"
    mishnayos_arr(37) = "עבודה זרה,avodah_zarah,9,7,10,12,12"
    mishnayos_arr(38) = "אבות,avos,18,16,18,22,23,11"
    mishnayos_arr(39) = "הוריות,horiyos,5,7,8"
    mishnayos_arr(40) = "זבחים,zevachim,4,5,6,6,8,7,6,12,7,8,8,6,8,10"
    mishnayos_arr(41) = "מנחות,menachos,4,5,7,5,9,7,6,7,9,9,9,5,11"
    mishnayos_arr(42) = "חולין,chullin,7,10,7,7,5,7,6,6,8,4,2,5"
    mishnayos_arr(43) = "בכורות,bechoros,7,9,4,10,6,12,7,10,8"
    mishnayos_arr(44) = "ערכין,arachin,4,6,5,4,6,5,5,7,8"
    mishnayos_arr(45) = "תמורה,temurah,6,3,5,4,6,5,6"
    mishnayos_arr(46) = "כריתות,kerisos,7,6,10,3,8,9"
    mishnayos_arr(47) = "מעילה,meilah,4,9,8,6,5,6"
    mishnayos_arr(48) = "תמיד,tamid,4,5,9,3,6,4,3"
    mishnayos_arr(49) = "מדות,midos,9,6,8,7,4"
    mishnayos_arr(50) = "קנים,kinnim,4,5,6"
    mishnayos_arr(51) = "כלים,keilim,9,8,8,4,11,4,6,11,8,8,9,8,8,8,6,8,17,9,10,7,3,10,5,17,9,9,12,10,8,4"
    mishnayos_arr(52) = "אהלות,ohalos,8,7,7,3,7,7,6,6,16,7,9,8,6,7,10,5,5,10"
    mishnayos_arr(53) = "נגעים,negaim,6,5,8,11,5,8,5,10,3,10,12,7,12,13"
    mishnayos_arr(54) = "פרה,parah,4,5,11,4,9,5,12,11,9,6,9,11"
    mishnayos_arr(55) = "טהרות,taharos,9,8,8,13,9,10,9,9,9,8"
    mishnayos_arr(56) = "מקוואות,mikvaos,8,10,4,5,6,11,7,5,7,8"
    mishnayos_arr(57) = "נדה,niddah,7,7,7,7,9,14,5,4,11,8"
    mishnayos_arr(58) = "מכשירין,machshirin,6,11,8,10,11,8"
    mishnayos_arr(59) = "זבים,zavim,6,4,3,7,12"
    mishnayos_arr(60) = "טבול יום,tevul_yom,5,8,6,7"
    mishnayos_arr(61) = "ידים,yadayim,5,4,5,8"
    mishnayos_arr(62) = "עוקצין,uktzin,6,10,12"
    mishnayos_ready = True
    
End Sub
Sub init_dafyomi()
    masechtosBavli(0) = "ברכות"
    masechtosBavli(1) = "שבת"
    masechtosBavli(2) = "עירובין"
    masechtosBavli(3) = "פסחים"
    masechtosBavli(4) = "שקלים"
    masechtosBavli(5) = "יומא"
    masechtosBavli(6) = "סוכה"
    masechtosBavli(7) = "ביצה"
    masechtosBavli(8) = "ראש השנה"
    masechtosBavli(9) = "תענית"
    masechtosBavli(10) = "מגילה"
    masechtosBavli(11) = "מועד קטן"
    masechtosBavli(12) = "חגיגה"
    masechtosBavli(13) = "יבמות"
    masechtosBavli(14) = "כתובות"
    masechtosBavli(15) = "נדרים"
    masechtosBavli(16) = "נזיר"
    masechtosBavli(17) = "סוטה"
    masechtosBavli(18) = "גיטין"
    masechtosBavli(19) = "קידושין"
    masechtosBavli(20) = "בבא קמא"
    masechtosBavli(21) = "בבא מציעא"
    masechtosBavli(22) = "בבא בתרא"
    masechtosBavli(23) = "סנהדרין"
    masechtosBavli(24) = "מכות"
    masechtosBavli(25) = "שבועות"
    masechtosBavli(26) = "עבודה זרה"
    masechtosBavli(27) = "הוריות"
    masechtosBavli(28) = "זבחים"
    masechtosBavli(29) = "מנחות"
    masechtosBavli(30) = "חולין"
    masechtosBavli(31) = "בכורות"
    masechtosBavli(32) = "ערכין"
    masechtosBavli(33) = "תמורה"
    masechtosBavli(34) = "כריתות"
    masechtosBavli(35) = "מעילה"
    masechtosBavli(36) = "תמיד"
    masechtosBavli(37) = "קינים"
    masechtosBavli(38) = "מידות"
    masechtosBavli(39) = "נדה"
    masechtosBavliTransliterated(0) = "Berachos"
    masechtosBavliTransliterated(1) = "Shabbos"
    masechtosBavliTransliterated(2) = "Eruvin"
    masechtosBavliTransliterated(3) = "Pesachim"
    masechtosBavliTransliterated(4) = "Shekalim"
    masechtosBavliTransliterated(5) = "Yoma"
    masechtosBavliTransliterated(6) = "Sukkah"
    masechtosBavliTransliterated(7) = "Beitzah"
    masechtosBavliTransliterated(8) = "Rosh Hashana"
    masechtosBavliTransliterated(9) = "Taanis"
    masechtosBavliTransliterated(10) = "Megillah"
    masechtosBavliTransliterated(11) = "Moed Katan"
    masechtosBavliTransliterated(12) = "Chagigah"
    masechtosBavliTransliterated(13) = "Yevamos"
    masechtosBavliTransliterated(14) = "Kesubos"
    masechtosBavliTransliterated(15) = "Nedarim"
    masechtosBavliTransliterated(16) = "Nazir"
    masechtosBavliTransliterated(17) = "Sotah"
    masechtosBavliTransliterated(18) = "Gitin"
    masechtosBavliTransliterated(19) = "Kiddushin"
    masechtosBavliTransliterated(20) = "Bava Kamma"
    masechtosBavliTransliterated(21) = "Bava Metzia"
    masechtosBavliTransliterated(22) = "Bava Basra"
    masechtosBavliTransliterated(23) = "Sanhedrin"
    masechtosBavliTransliterated(24) = "Makkos"
    masechtosBavliTransliterated(25) = "Shevuos"
    masechtosBavliTransliterated(26) = "Avodah Zarah"
    masechtosBavliTransliterated(27) = "Horiyos"
    masechtosBavliTransliterated(28) = "Zevachim"
    masechtosBavliTransliterated(29) = "Menachos"
    masechtosBavliTransliterated(30) = "Chullin"
    masechtosBavliTransliterated(31) = "Bechoros"
    masechtosBavliTransliterated(32) = "Arachin"
    masechtosBavliTransliterated(33) = "Temurah"
    masechtosBavliTransliterated(34) = "Kerisos"
    masechtosBavliTransliterated(35) = "Meilah"
    masechtosBavliTransliterated(36) = "Kinnim"
    masechtosBavliTransliterated(37) = "Tamid"
    masechtosBavliTransliterated(38) = "Midos"
    masechtosBavliTransliterated(39) = "Niddah"
    masechtosYerushalmi(0) = "ברכות"
    masechtosYerushalmi(1) = "פאה"
    masechtosYerushalmi(2) = "דמאי"
    masechtosYerushalmi(3) = "כלאיים"
    masechtosYerushalmi(4) = "שביעית"
    masechtosYerushalmi(5) = "תרומות"
    masechtosYerushalmi(6) = "מעשרות"
    masechtosYerushalmi(7) = "מעשר שני"
    masechtosYerushalmi(8) = "חלה"
    masechtosYerushalmi(9) = "ערלה"
    masechtosYerushalmi(10) = "ביכורים"
    masechtosYerushalmi(11) = "שבת"
    masechtosYerushalmi(12) = "ערובין"
    masechtosYerushalmi(13) = "פסחים"
    masechtosYerushalmi(14) = "ביצה"
    masechtosYerushalmi(15) = "ראש השנה"
    masechtosYerushalmi(16) = "יומא"
    masechtosYerushalmi(17) = "סוכה"
    masechtosYerushalmi(18) = "תענית"
    masechtosYerushalmi(19) = "שקלים"
    masechtosYerushalmi(20) = "מגילה"
    masechtosYerushalmi(21) = "חגיגה"
    masechtosYerushalmi(22) = "מועד קטן"
    masechtosYerushalmi(23) = "יבמות"
    masechtosYerushalmi(24) = "כתובות"
    masechtosYerushalmi(25) = "סוטה"
    masechtosYerushalmi(26) = "נדרים"
    masechtosYerushalmi(27) = "נזיר"
    masechtosYerushalmi(28) = "גיטין"
    masechtosYerushalmi(29) = "קידושין"
    masechtosYerushalmi(30) = "בבא קמא"
    masechtosYerushalmi(31) = "בבא מציעא"
    masechtosYerushalmi(32) = "בבא בתרא"
    masechtosYerushalmi(33) = "שבועות"
    masechtosYerushalmi(34) = "מכות"
    masechtosYerushalmi(35) = "סנהדרין"
    masechtosYerushalmi(36) = "עבודה זרה"
    masechtosYerushalmi(37) = "הוריות"
    masechtosYerushalmi(38) = "נידה"
    masechtosYerushalmiTransliterated(0) = "Berachos"
    masechtosYerushalmiTransliterated(1) = "Peah"
    masechtosYerushalmiTransliterated(2) = "Demai"
    masechtosYerushalmiTransliterated(3) = "Kilayim"
    masechtosYerushalmiTransliterated(4) = "Sheviis"
    masechtosYerushalmiTransliterated(5) = "Terumos"
    masechtosYerushalmiTransliterated(6) = "Maasros"
    masechtosYerushalmiTransliterated(7) = "Maaser Sheni"
    masechtosYerushalmiTransliterated(8) = "Chalah"
    masechtosYerushalmiTransliterated(9) = "Orlah"
    masechtosYerushalmiTransliterated(10) = "Bikurim"
    masechtosYerushalmiTransliterated(11) = "Shabbos"
    masechtosYerushalmiTransliterated(12) = "Eruvin"
    masechtosYerushalmiTransliterated(13) = "Pesachim"
    masechtosYerushalmiTransliterated(14) = "Beitzah"
    masechtosYerushalmiTransliterated(15) = "Rosh Hashanah"
    masechtosYerushalmiTransliterated(16) = "Yoma"
    masechtosYerushalmiTransliterated(17) = "Sukah"
    masechtosYerushalmiTransliterated(18) = "Taanis"
    masechtosYerushalmiTransliterated(19) = "Shekalim"
    masechtosYerushalmiTransliterated(20) = "Megilah"
    masechtosYerushalmiTransliterated(21) = "Chagigah"
    masechtosYerushalmiTransliterated(22) = "Moed Katan"
    masechtosYerushalmiTransliterated(23) = "Yevamos"
    masechtosYerushalmiTransliterated(24) = "Kesuvos"
    masechtosYerushalmiTransliterated(25) = "Sotah"
    masechtosYerushalmiTransliterated(26) = "Nedarim"
    masechtosYerushalmiTransliterated(27) = "Nazir"
    masechtosYerushalmiTransliterated(28) = "Gitin"
    masechtosYerushalmiTransliterated(29) = "Kidushin"
    masechtosYerushalmiTransliterated(30) = "Bava Kama"
    masechtosYerushalmiTransliterated(31) = "Bava Metzia"
    masechtosYerushalmiTransliterated(32) = "Bava Basra"
    masechtosYerushalmiTransliterated(33) = "Shevuos"
    masechtosYerushalmiTransliterated(34) = "Makos"
    masechtosYerushalmiTransliterated(35) = "Sanhedrin"
    masechtosYerushalmiTransliterated(36) = "Avodah Zarah"
    masechtosYerushalmiTransliterated(37) = "Horayos"
    masechtosYerushalmiTransliterated(38) = "Nidah"
    blattPerMasechta(0) = 64
    blattPerMasechta(1) = 157
    blattPerMasechta(2) = 105
    blattPerMasechta(3) = 121
    blattPerMasechta(4) = 22
    blattPerMasechta(5) = 88
    blattPerMasechta(6) = 56
    blattPerMasechta(7) = 40
    blattPerMasechta(8) = 35
    blattPerMasechta(9) = 31
    blattPerMasechta(10) = 32
    blattPerMasechta(11) = 29
    blattPerMasechta(12) = 27
    blattPerMasechta(13) = 122
    blattPerMasechta(14) = 112
    blattPerMasechta(15) = 91
    blattPerMasechta(16) = 66
    blattPerMasechta(17) = 49
    blattPerMasechta(18) = 90
    blattPerMasechta(19) = 82
    blattPerMasechta(20) = 119
    blattPerMasechta(21) = 119
    blattPerMasechta(22) = 176
    blattPerMasechta(23) = 113
    blattPerMasechta(24) = 24
    blattPerMasechta(25) = 49
    blattPerMasechta(26) = 76
    blattPerMasechta(27) = 14
    blattPerMasechta(28) = 120
    blattPerMasechta(29) = 110
    blattPerMasechta(30) = 142
    blattPerMasechta(31) = 61
    blattPerMasechta(32) = 34
    blattPerMasechta(33) = 34
    blattPerMasechta(34) = 28
    blattPerMasechta(35) = 22
    blattPerMasechta(36) = 4
    blattPerMasechta(37) = 9
    blattPerMasechta(38) = 5
    blattPerMasechta(39) = 73
    blattPerMasechtaYerushalmi(0) = 68
    blattPerMasechtaYerushalmi(1) = 37
    blattPerMasechtaYerushalmi(2) = 34
    blattPerMasechtaYerushalmi(3) = 44
    blattPerMasechtaYerushalmi(4) = 31
    blattPerMasechtaYerushalmi(5) = 59
    blattPerMasechtaYerushalmi(6) = 26
    blattPerMasechtaYerushalmi(7) = 33
    blattPerMasechtaYerushalmi(8) = 28
    blattPerMasechtaYerushalmi(9) = 20
    blattPerMasechtaYerushalmi(10) = 13
    blattPerMasechtaYerushalmi(11) = 92
    blattPerMasechtaYerushalmi(12) = 65
    blattPerMasechtaYerushalmi(13) = 71
    blattPerMasechtaYerushalmi(14) = 22
    blattPerMasechtaYerushalmi(15) = 22
    blattPerMasechtaYerushalmi(16) = 42
    blattPerMasechtaYerushalmi(17) = 26
    blattPerMasechtaYerushalmi(18) = 26
    blattPerMasechtaYerushalmi(19) = 33
    blattPerMasechtaYerushalmi(20) = 34
    blattPerMasechtaYerushalmi(21) = 22
    blattPerMasechtaYerushalmi(22) = 19
    blattPerMasechtaYerushalmi(23) = 85
    blattPerMasechtaYerushalmi(24) = 72
    blattPerMasechtaYerushalmi(25) = 47
    blattPerMasechtaYerushalmi(26) = 40
    blattPerMasechtaYerushalmi(27) = 47
    blattPerMasechtaYerushalmi(28) = 54
    blattPerMasechtaYerushalmi(29) = 48
    blattPerMasechtaYerushalmi(30) = 44
    blattPerMasechtaYerushalmi(31) = 37
    blattPerMasechtaYerushalmi(32) = 34
    blattPerMasechtaYerushalmi(33) = 44
    blattPerMasechtaYerushalmi(34) = 9
    blattPerMasechtaYerushalmi(35) = 57
    blattPerMasechtaYerushalmi(36) = 37
    blattPerMasechtaYerushalmi(37) = 19
    blattPerMasechtaYerushalmi(38) = 13

    dafYomiStartDate = DateSerial(1923, 9, 11)
    shekalimChangeDate = DateSerial(1975, 6, 24)
    dafyom_ready = True
    
End Sub
Private Function IsCurrentBlattWithNextMasechta(ByVal masechta As Integer, ByVal blatt As Integer) As Boolean
    IsCurrentBlattWithNextMasechta = (masechta = 35 And blatt = 22) Or (masechta = 36 And blatt = 25)
End Function

Public Function GetDafYomiBavli(ByVal date_in As Date) As Daf
    ''
    ' * The number of daf per masechta. Since the number of blatt in Shekalim changed on the 8th Daf Yomi cycle
    ' * beginning on June 24, 1975 from 13 to 22, the actual calculation for blattPerMasechta[4] will later be
    ' * adjusted based on the cycle.

    
    Dim dafYomi As Daf ' = Nothing

    Dim cycleNo ' = 0
    Dim dafNo As Integer ' = 0
    If dafyom_ready = False Then init_dafyomi
    
    If date_in < dafYomiStartDate Then
        ' TODO: should we return a null or throw an IllegalArgumentException?
        'Throw New ArgumentException([date] & " is prior to organized Daf Yomi Bavli cycles that started on " & dafYomiStartDate)
            dafYomi.Page = -1
            GetDafYomiBavli = dafYomi
        Exit Function
    End If

    'Changed in Vb to use DateDiff and not JulianDay as it was off, probably issue with time in GetJulianDay
    Dim DaysFromDafYomiStart As Long: DaysFromDafYomiStart = DateDiff("d", dafYomiStartDate, date_in)
    Dim DaysFromShekalimChange As Long: DaysFromShekalimChange = DateDiff("d", shekalimChangeDate, date_in)

    If date_in >= shekalimChangeDate Then
        'cycleNo = 8 + ((julianDay - shekalimJulianChangeDay) / 2711)
        cycleNo = 8 + (DaysFromShekalimChange / 2711)
        'dafNo = ((julianDay - shekalimJulianChangeDay) Mod 2711)
        dafNo = DaysFromShekalimChange Mod 2711
    Else
        'cycleNo = 1 + ((julianDay - dafYomiJulianStartDay) / 2702)
        cycleNo = 1 + (DaysFromDafYomiStart / 2702)
        'dafNo = ((julianDay - dafYomiJulianStartDay) Mod 2702)
        dafNo = DaysFromDafYomiStart Mod 2702
    End If

    'Debug.Print(dafNo)

    Dim total ' = 0
    Dim masechta: masechta = -1
    Dim blatt ' = 0

    ' Fix Shekalim for old cycles.
    If cycleNo <= 7 Then
        blattPerMasechta(4) = 13
    Else
        blattPerMasechta(4) = 22 ' correct any change that may have been changed from a prior calculation
    End If

    For J = 0 To 39
        masechta = masechta + 1
        total = total + blattPerMasechta(J) - 1

        If dafNo < total Then
            blatt = 1 + blattPerMasechta(J) - (total - dafNo)
            ' Fiddle with the weird ones near the end.
            If masechta = 36 Then
                blatt = blatt + 21
            ElseIf masechta = 37 Then
                blatt = blatt + 24
            ElseIf masechta = 38 Then
                blatt = blatt + 32
            End If

            Dim isWithNextMasechta: isWithNextMasechta = IsCurrentBlattWithNextMasechta(masechta, blatt)
            'dafYomi= New Daf(masechta, blatt, isWithNextMasechta)
            dafYomi.masechtaNumber = masechta
            dafYomi.Page = blatt
            dafYomi.HasSecondaryMesechta = isWithNextMasechta
            dafYomi.SecondaryMesechtaNumber = masechta + 1
            Exit For
        End If
    Next

    GetDafYomiBavli = dafYomi
End Function
    
'Returns the Daf Yomi Yerusalmi page for a given date (as Daf type).
'The first Daf Yomi cycle started on 15 Shevat (Tu Bishvat), 5740 (February, 2, 1980) and calculations
'prior to this date will result in Daf.page=-1, which means no-daf for the date.
'Similar value would be returned on Tisha B'Av or Yom Kippur.
Public Function GetDafYomiYerushalmi(ByVal date_in As Date) As Daf
    Const WHOLE_SHAS_DAFS As Integer = 1554
    Const DAF_YOMI_START_DAY As Date = #2/2/1980#
    Dim nextCycle As Date
    Dim prevCycle As Date
    Dim result As Daf
    Dim hdate_check As hdate, dafNo As Integer
    If dafyom_ready = False Then init_dafyomi

    hdate_check = ConvertDate(date_in)
    
    'There isn't Daf Yomi on Yom Kippur or Tisha B'Av, or before the start of the first cycle (February 2, 1980)
    If GetYomTov(hdate_check) = YOM_KIPPUR Or GetYomTov(hdate_check) = TISHA_BAV Or date_in < DAF_YOMI_START_DAY Then
        result.Page = -1
        result.masechtaNumber = 0
        result.HasSecondaryMesechta = False
        result.SecondaryMesechtaNumber = 0
        GetDafYomiYerushalmi = result
        Exit Function
    End If
    
    'Start to calculate current cycle. init the start day
    nextCycle = DAF_YOMI_START_DAY
    prevCycle = DAF_YOMI_START_DAY
    
    'Go cycle by cycle, until we get the next cycle (after date_in)
    Do While date_in > nextCycle
        prevCycle = nextCycle
        
        'Adds the number of whole shas dafs. and the number of days that not have daf.
        nextCycle = DateAdd("d", WHOLE_SHAS_DAFS, nextCycle)
        nextCycle = DateAdd("d", getNumOfSpecialDays(prevCycle, nextCycle), nextCycle)
    Loop

    'Get the number of days from cycle start until request.
    dafNo = DateDiff("d", prevCycle, date_in)

    'Get the number of special day to subtract, and subtract them
    dafNo = dafNo - getNumOfSpecialDays(prevCycle, date_in)
    
    'Finally find the daf.
    For I = 0 To 38
        If dafNo < blattPerMasechtaYerushalmi(I) Then
            result.masechtaNumber = I
            result.Page = dafNo
            Exit For
        End If
        dafNo = dafNo - blattPerMasechtaYerushalmi(I)
    Next I

    result.Page = result.Page + 1
    GetDafYomiYerushalmi = result

End Function
        
'Return the number of special days (Yom Kippur and Tisha Beav, where there are no dafim on these days),
'from the start date given until the end date.
Public Function getNumOfSpecialDays(date_start As Date, date_end As Date) As Integer
    Dim start_year As Long
    Dim end_year As Long
    Dim yom_kipur As hdate
    Dim tisha_beav As hdate
    Dim result As Integer
    Dim date_check As Date
        
    'Value to return
    result = 0
    
    'Find the start and end Jewish years
    start_year = ConvertDate(date_start).year
    end_year = ConvertDate(date_end).year
    
    'Instant of special Dates
    yom_kipur = HDateNew(0, 0, 0, 0, 0, 0, 0, 0)
    tisha_beav = HDateNew(0, 0, 0, 0, 0, 0, 0, 0)
    Call HDateAdd(yom_kipur, start_year, 7, 10, 0, 0, 0, 0)
    Call HDateAdd(tisha_beav, start_year, 5, 9, 0, 0, 0, 0)
    
    'Go over the years and find special dates
    For I = start_year To end_year
        date_check = HDateGregorian(yom_kipur)
        If date_check >= date_start And date_check <= date_end Then result = result + 1
        
        date_check = HDateGregorian(tisha_beav)
        If date_check >= date_start And date_check <= date_end Then result = result + 1
        
        Call HDateAddYear(yom_kipur, 1)
        Call HDateAddYear(tisha_beav, 1)
    Next I
    getNumOfSpecialDays = result
End Function

Public Function GetDafYomiFormat(ByVal date_in As Date, Optional yerushalmi As Boolean = False)
    Dim result As Daf
    If hdateformat_ready = False Then init_hdateformat
    If yerushalmi Then
        result = GetDafYomiYerushalmi(date_in)
        If result.Page = -1 Then
            GetDafYomiFormat = "-"
        Else
            GetDafYomiFormat = masechtosYerushalmi(result.masechtaNumber) & " דף " & NumToHChar(result.Page)
        
        End If
    Else
        result = GetDafYomiBavli(date_in)
        If result.Page = -1 Then
            GetDafYomiFormat = "-"
        Else
            GetDafYomiFormat = masechtosBavli(result.masechtaNumber) & " דף " & NumToHChar(result.Page)
        End If
    End If
    
End Function


Public Function GetMishnaYomi(ByVal date_in As Date) As Daf
    Const MISHNAH_YOMI_START_DAY As Date = #5/20/1947# 'note - officially it was 6th of sivan 5707 (Shavuos 1947), but since the first cycle was 5 days shorter then alef sivan is being used instead.
    Const DaysPerCycle = 2096 'total of 4192 Mishanyos / 2 per day
    
    Dim result As Daf
    Dim day_in_cycle As Integer
    Dim str_parts() As String
    Dim sum As Integer
    Dim nxt_sum As Integer
    Dim m_since_cycle As Integer
    Dim m_since_masechet As Integer
    Dim m_since_perek As Integer
    
    If date_in < MISHNAH_YOMI_START_DAY Then
        result.Page = -1
        result.masechtaNumber = 0
        result.HasSecondaryMesechta = False
        result.SecondaryMesechtaNumber = 0
        GetMishnaYomi = result
        Exit Function
    End If

    'elapsed days since the start of the current cycle
    day_in_cycle = (DateDiff("d", MISHNAH_YOMI_START_DAY, date_in)) Mod DaysPerCycle
    
    For I = 0 To 62 'I iterates over masechtos
        'get array parts from string (each number represents the num of mishnaos in a perek)
        str_parts = Split(mishnayos_arr(I), ",")
        nxt_sum = 0
        'J iterates over the chapters and counts the total mishanyos since the beginning of the masechet
        For J = 2 To UBound(str_parts) 'array strats from index 2, since 0 & 1 are titles of the masechtos
            m_since_cycle = day_in_cycle * 2 'mishanyos since the beginning of the cycle
            m_since_masechet = m_since_cycle - sum 'mishanyos since the beginning of the masechet
            m_since_perek = m_since_masechet - nxt_sum 'mishanyos since the beginning of the perek
            'return the result if current mishnayos count since the beginning of the cycle is between current perek and the next one
            If m_since_cycle >= sum And m_since_cycle < sum + nxt_sum + str_parts(J) Then
                result.masechtaNumber = I
                result.PerekNumber = J - 1 'perek
                result.Page = m_since_perek + 1 'mishnah
                GetMishnaYomi = result
                Exit Function
            End If
            nxt_sum = nxt_sum + Val(str_parts(J)) ' nxt_sum counts the mishnayos since the begining of the masechet
        Next J
        sum = sum + nxt_sum 'the currnet amount of mishnayos since first masechet
    Next I
     
    
End Function

Public Function GetMishnaYomiFormat(ByVal date_in As Date) As String
    Dim result As Daf
    If hdateformat_ready = False Then init_hdateformat
    If mishnayos_ready = False Then init_mishnayos

    Dim str_parts() As String
        
    result = GetMishnaYomi(date_in)
    If result.Page = -1 Then
        GetMishnaYomiFormat = "-"
        Exit Function
    End If
    str_parts = Split(mishnayos_arr(result.masechtaNumber), ",")
        
    GetMishnaYomiFormat = str_parts(0) & " פרק " & NumToHChar(result.PerekNumber) & " משנה " & NumToHChar(result.Page)
End Function

