Attribute VB_Name = "mod_zmanim"
Public Const GEOMETRIC_ZENITH As Double = 90#
Public Const ZENITH_AMITIS As Double = GEOMETRIC_ZENITH + 1.583

Public Const ZENITH_26_D As Double = GEOMETRIC_ZENITH + 26#
Public Const ZENITH_19_P_8 As Double = GEOMETRIC_ZENITH + 19.8
Public Const ZENITH_18_D As Double = GEOMETRIC_ZENITH + 18
Public Const ZENITH_16_P_9 As Double = GEOMETRIC_ZENITH + 16.9
Public Const ZENITH_16_P_1 As Double = GEOMETRIC_ZENITH + 16.1
Public Const ZENITH_13_P_24 As Double = GEOMETRIC_ZENITH + 13.24

Public Const ZENITH_11_P_5 As Double = GEOMETRIC_ZENITH + 11.5
Public Const ZENITH_11_D As Double = GEOMETRIC_ZENITH + 11
Public Const ZENITH_10_P_2 As Double = GEOMETRIC_ZENITH + 10.2

Public Const ZENITH_3_P_65 As Double = GEOMETRIC_ZENITH + 3.65
Public Const ZENITH_3_P_676 As Double = GEOMETRIC_ZENITH + 3.676
Public Const ZENITH_3_P_7 As Double = GEOMETRIC_ZENITH + 3.7
Public Const ZENITH_3_P_8 As Double = GEOMETRIC_ZENITH + 3.8
Public Const ZENITH_4_P_37 As Double = GEOMETRIC_ZENITH + 4.37
Public Const ZENITH_4_P_61 As Double = GEOMETRIC_ZENITH + 4.61
Public Const ZENITH_4_P_8 As Double = GEOMETRIC_ZENITH + 4.8
Public Const ZENITH_5_P_88 As Double = GEOMETRIC_ZENITH + 5.88
Public Const ZENITH_5_P_95 As Double = GEOMETRIC_ZENITH + 5.95
Public Const ZENITH_6_D As Double = GEOMETRIC_ZENITH + 6
Public Const ZENITH_7_P_083 As Double = GEOMETRIC_ZENITH + 7.083
Public Const ZENITH_8_P_5 As Double = GEOMETRIC_ZENITH + 8.5

Public Const MINUTES60 As Double = 60 * 60000
Public Const MINUTES72 As Double = 72 * 60000
Public Const MINUTES90 As Double = 90 * 60000
Public Const MINUTES96 As Double = 96 * 60000
Public Const MINUTES120 As Double = 120 * 60000

Public Const MINUTES18 As Double = 18 * 60000
Public EmptyHdate As hdate

Public Function getLocalMeanTimeOffset(now As hdate, here As location) As Long
    getLocalMeanTimeOffset = CLng(here.longitude * 4 * 60 - now.offset)
End Function

Public Function getAntimeridianAdjustment(now As hdate, here As location) As Long
    Dim localHoursOffset As Double
    localHoursOffset = getLocalMeanTimeOffset(now, here) / 3600
    If localHoursOffset >= 20 Then
        getAntimeridianAdjustment = 1
    ElseIf localHoursOffset <= -20 Then
        getAntimeridianAdjustment = -1
    Else
        getAntimeridianAdjustment = 0
    End If
End Function

Public Function getDateFromTime(current As hdate, time As Double, here As location, isSunrise As Long) As hdate
    Dim result As hdate
    result.year = current.year
    result.EY = current.EY
    result.offset = current.offset
    result.month = current.month
    result.day = current.day
    
    Dim adjustment As Long
    adjustment = getAntimeridianAdjustment(current, here)
    If adjustment <> 0 Then
        HDateAddDay result, adjustment
    End If
    
    Dim hours As Long
    hours = Int(time)
    time = (time - hours) * 60
    Dim minutes As Long
    minutes = Int(time)
    time = (time - minutes) * 60
    Dim seconds As Long
    seconds = Int(time)
    time = (time - seconds) * 1000
    Dim milliseconds As Long
    milliseconds = Int(time)
    
    Dim localTimeHours As Long
    localTimeHours = Int(here.longitude / 15)
    If isSunrise <> 0 And localTimeHours + hours > 18 Then
        HDateAddDay result, -1
    ElseIf isSunrise = 0 And localTimeHours + hours < 6 Then
        HDateAddDay result, 1
    End If
    
    result.hour = hours
    result.min = minutes
    result.sec = seconds
    result.msec = milliseconds
    HDateAddSecond result, current.offset
    
    getDateFromTime = result
End Function

Public Function calcsunrise(date_in As hdate, here As location, zenith As Double, adjustForElevation As Boolean) As hdate
    Dim sunrise As Double
    sunrise = getUTCSunrise(HDateJulian(date_in), here, zenith, adjustForElevation)
    calcsunrise = getDateFromTime(date_in, sunrise, here, 1)
End Function

Public Function calcsunset(date_in As hdate, here As location, zenith As Double, adjustForElevation As Boolean) As hdate
    Dim sunset As Double
    sunset = getUTCSunset(HDateJulian(date_in), here, zenith, adjustForElevation)
    calcsunset = getDateFromTime(date_in, sunset, here, 0)
End Function

Public Function calcshaahzmanis(startday As hdate, iEndday As hdate) As Long
    Dim start As Long
    Dim iend As Long
    start = HebrewCalendarElapsedDays(startday.year) + (startday.dayOfYear - 1)
    iend = HebrewCalendarElapsedDays(iEndday.year) + (iEndday.dayOfYear - 1)
    Dim diff As Long
    diff = iend - start
    diff = (diff * 24) + (iEndday.hour - startday.hour)
    diff = (diff * 60) + (iEndday.min - startday.min)
    diff = (diff * 60) + (iEndday.sec - startday.sec)
    diff = (diff * 1000) + (iEndday.msec - startday.msec)
    If startday.year = 0 Or iEndday.year = 0 Then
        calcshaahzmanis = 0
        Exit Function
    End If
    calcshaahzmanis = diff \ 12
End Function

Public Function calctimeoffset(time As hdate, offset As Long) As hdate
    Dim result As hdate
    If time.year = 0 Or offset = 0 Then
        calctimeoffset = result
        Exit Function
    End If
    result = time
    HDateAddMSecond result, offset
    calctimeoffset = result
End Function

Public Function getalos(date_in As hdate, here As location) As hdate
    getalos = calcsunrise(date_in, here, ZENITH_16_P_1, 0)
End Function

Public Function getalosbaalhatanya(date_in As hdate, here As location) As hdate
    getalosbaalhatanya = calcsunrise(date_in, here, ZENITH_16_P_9, 0)
End Function

Public Function getalos26degrees(date_in As hdate, here As location) As hdate
    getalos26degrees = calcsunrise(date_in, here, ZENITH_26_D, 0)
End Function

Public Function getalos19p8degrees(date_in As hdate, here As location) As hdate
    getalos19p8degrees = calcsunrise(date_in, here, ZENITH_19_P_8, 0)
End Function

Public Function getalos18degrees(date_in As hdate, here As location) As hdate
    getalos18degrees = calcsunrise(date_in, here, ZENITH_18_D, 0)
End Function

Public Function getalos120(date_in As hdate, here As location) As hdate
    getalos120 = calctimeoffset(getsunrise(date_in, here), -MINUTES120)
End Function

Public Function getalos120zmanis(date_in As hdate, here As location) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = getshaahzmanisgra(date_in, here)
    If shaahzmanis = 0 Then
        getalos120zmanis = EmptyHdate
        Exit Function
    End If
    getalos120zmanis = calctimeoffset(getsunrise(date_in, here), shaahzmanis * -2)
End Function

Public Function getalos96(date_in As hdate, here As location) As hdate
    getalos96 = calctimeoffset(getsunrise(date_in, here), -MINUTES96)
End Function

Public Function getalos96zmanis(date_in As hdate, here As location) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = getshaahzmanisgra(date_in, here)
    If shaahzmanis = 0 Then
        getalos96zmanis = EmptyHdate
        Exit Function
    End If
    getalos96zmanis = calctimeoffset(getsunrise(date_in, here), shaahzmanis * -1.6)
End Function

Public Function getalos90(date_in As hdate, here As location) As hdate
    getalos90 = calctimeoffset(getsunrise(date_in, here), -MINUTES90)
End Function

Public Function getalos90zmanis(date_in As hdate, here As location) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = getshaahzmanisgra(date_in, here)
    If shaahzmanis = 0 Then
        getalos90zmanis = EmptyHdate
        Exit Function
    End If
    getalos90zmanis = calctimeoffset(getsunrise(date_in, here), shaahzmanis * -1.5)
End Function

Public Function getalos72(date_in As hdate, here As location) As hdate
    getalos72 = calctimeoffset(getsunrise(date_in, here), -MINUTES72)
End Function

Public Function getalos72zmanis(date_in As hdate, here As location) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = getshaahzmanisgra(date_in, here)
    If shaahzmanis = 0 Then
        getalos72zmanis = EmptyHdate
        Exit Function
    End If
    getalos72zmanis = calctimeoffset(getsunrise(date_in, here), shaahzmanis * -1.2)
End Function

Public Function getalos60(date_in As hdate, here As location) As hdate
    getalos60 = calctimeoffset(getsunrise(date_in, here), -MINUTES60)
End Function

Public Function getmisheyakir11p5degrees(date_in As hdate, here As location) As hdate
    getmisheyakir11p5degrees = calcsunrise(date_in, here, ZENITH_11_P_5, 0)
End Function

Public Function getmisheyakir11degrees(date_in As hdate, here As location) As hdate
    getmisheyakir11degrees = calcsunrise(date_in, here, ZENITH_11_D, 0)
End Function

Public Function getmisheyakir10p2degrees(date_in As hdate, here As location) As hdate
    getmisheyakir10p2degrees = calcsunrise(date_in, here, ZENITH_10_P_2, 0)
End Function

Public Function getsunrise(date_in As hdate, here As location) As hdate
    getsunrise = calcsunrise(date_in, here, GEOMETRIC_ZENITH, 0)
End Function

Public Function getsunrisebaalhatanya(date_in As hdate, here As location) As hdate
    getsunrisebaalhatanya = calcsunrise(date_in, here, ZENITH_AMITIS, 0)
End Function

Public Function getelevationsunrise(date_in As hdate, here As location) As hdate
    getelevationsunrise = calcsunrise(date_in, here, GEOMETRIC_ZENITH, 1)
End Function

Public Function calcshma(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calcshma = calctimeoffset(startday, shaahzmanis * 3)
End Function

Public Function getshmabaalhatanya(date_in As hdate, here As location) As hdate
    getshmabaalhatanya = calcshma(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function
Public Function getshmagra(date_in As hdate, here As location) As hdate
    getshmagra = calcshma(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getshmamga(date_in As hdate, here As location) As hdate
    getshmamga = calcshma(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function calctefila(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calctefila = calctimeoffset(startday, shaahzmanis * 4)
End Function

Public Function gettefilabaalhatanya(date_in As hdate, here As location) As hdate
    gettefilabaalhatanya = calctefila(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function gettefilagra(date_in As hdate, here As location) As hdate
    gettefilagra = calctefila(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function gettefilamga(date_in As hdate, here As location) As hdate
    gettefilamga = calctefila(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function getachilaschometzbaalhatanya(date_in As hdate, here As location) As hdate
    getachilaschometzbaalhatanya = calctefila(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getachilaschometzgra(date_in As hdate, here As location) As hdate
    getachilaschometzgra = calctefila(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getachilaschometzmga(date_in As hdate, here As location) As hdate
    getachilaschometzmga = calctefila(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function calcbiurchometz(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calcbiurchometz = calctimeoffset(startday, shaahzmanis * 5)
End Function

Public Function getbiurchometzbaalhatanya(date_in As hdate, here As location) As hdate
    getbiurchometzbaalhatanya = calcbiurchometz(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getbiurchometzgra(date_in As hdate, here As location) As hdate
    getbiurchometzgra = calcbiurchometz(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getbiurchometzmga(date_in As hdate, here As location) As hdate
    getbiurchometzmga = calcbiurchometz(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function calcchatzos(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calcchatzos = calctimeoffset(startday, shaahzmanis * 6)
End Function

Public Function getchatzosbaalhatanya(date_in As hdate, here As location) As hdate
    getchatzosbaalhatanya = calcchatzos(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getchatzosgra(date_in As hdate, here As location) As hdate
    getchatzosgra = calcchatzos(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function calcminchagedola(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calcminchagedola = calctimeoffset(startday, shaahzmanis * 6.5)
End Function

Public Function getminchagedolabaalhatanya(date_in As hdate, here As location) As hdate
    getminchagedolabaalhatanya = calcminchagedola(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getminchagedolagra(date_in As hdate, here As location) As hdate
    getminchagedolagra = calcminchagedola(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getminchagedolamga(date_in As hdate, here As location) As hdate
    getminchagedolamga = calcminchagedola(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function calcminchagedola30min(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calcminchagedola30min = calctimeoffset(startday, (shaahzmanis * 6) + 1800000)
End Function

Public Function calcminchagedolagreater30min(startday As hdate, endday As hdate) As hdate
    If (calcshaahzmanis(startday, endday) * 0.5) >= 1800000 Then
        calcminchagedolagreater30min = calcminchagedola(startday, endday)
    Else
        calcminchagedolagreater30min = calcminchagedola30min(startday, endday)
    End If
End Function

Public Function getminchagedolabaalhatanyag30m(date_in As hdate, here As location) As hdate
    getminchagedolabaalhatanyag30m = calcminchagedolagreater30min(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getminchagedolagrag30m(date_in As hdate, here As location) As hdate
    getminchagedolagrag30m = calcminchagedolagreater30min(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getminchagedolamgag30m(date_in As hdate, here As location) As hdate
    getminchagedolamgag30m = calcminchagedolagreater30min(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function calcminchaketana(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calcminchaketana = calctimeoffset(startday, shaahzmanis * 9.5)
End Function

Public Function getminchaketanabaalhatanya(date_in As hdate, here As location) As hdate
    getminchaketanabaalhatanya = calcminchaketana(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getminchaketanagra(date_in As hdate, here As location) As hdate
    getminchaketanagra = calcminchaketana(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getminchaketanamga(date_in As hdate, here As location) As hdate
    getminchaketanamga = calcminchaketana(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function calcplag(startday As hdate, endday As hdate) As hdate
    Dim shaahzmanis As Long
    shaahzmanis = calcshaahzmanis(startday, endday)
    calcplag = calctimeoffset(startday, shaahzmanis * 10.75)
End Function

Public Function getplagbaalhatanya(date_in As hdate, here As location) As hdate
    getplagbaalhatanya = calcplag(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getplaggra(date_in As hdate, here As location) As hdate
    getplaggra = calcplag(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getplagmga(date_in As hdate, here As location) As hdate
    getplagmga = calcplag(getalos72(date_in, here), gettzais72(date_in, here))
End Function

Public Function getcandlelighting(date_in As hdate, here As location) As hdate
    getcandlelighting = calctimeoffset(calcsunset(date_in, here, GEOMETRIC_ZENITH, 0), -MINUTES18)
End Function

Public Function getsunset(date_in As hdate, here As location) As hdate
    getsunset = calcsunset(date_in, here, GEOMETRIC_ZENITH, 0)
End Function

Public Function getsunsetbaalhatanya(date_in As hdate, here As location) As hdate
    getsunsetbaalhatanya = calcsunset(date_in, here, ZENITH_AMITIS, 0)
End Function

Public Function getelevationsunset(date_in As hdate, here As location) As hdate
    getelevationsunset = calcsunset(date_in, here, GEOMETRIC_ZENITH, 1)
End Function

Public Function gettzaisbaalhatanya(date_in As hdate, here As location) As hdate
    gettzaisbaalhatanya = calcsunset(date_in, here, ZENITH_6_D, 1)
End Function

Public Function gettzais8p5(date_in As hdate, here As location) As hdate
    gettzais8p5 = calcsunset(date_in, here, ZENITH_8_P_5, 1)
End Function

Public Function gettzais72(date_in As hdate, here As location) As hdate
    gettzais72 = calctimeoffset(getsunset(date_in, here), MINUTES72)
End Function

Public Function calcmoladoffset(date_in As hdate, offsetsec As Long) As hdate
    Dim result As hdate
    Dim tz As Long
    Dim adjustment As Long
    result = GetMolad(date_in.year, date_in.month)
    tz = (-result.offset) + date_in.offset
    adjustment = ((result.sec * 10) / 3) + tz + offsetsec
    result.sec = 0
    HDateAddSecond result, adjustment
    result.EY = date_in.EY
    result.offset = date_in.offset
    calcmoladoffset = result
End Function

Public Function getmolad7days(date_in As hdate) As hdate
    getmolad7days = calcmoladoffset(date_in, 604800)
End Function

Public Function getmoladhalfmonth(date_in As hdate) As hdate
    getmoladhalfmonth = calcmoladoffset(date_in, 1275722)
End Function
Public Function getmolad15days(date_in As hdate) As hdate
    getmolad15days = calcmoladoffset(date_in, 1296000)
End Function

Public Function getshaahzmanisbaalhatanya(date_in As hdate, here As location) As Long
    getshaahzmanisbaalhatanya = calcshaahzmanis(getsunrisebaalhatanya(date_in, here), getsunsetbaalhatanya(date_in, here))
End Function

Public Function getshaahzmanisgra(date_in As hdate, here As location) As Long
    getshaahzmanisgra = calcshaahzmanis(getsunrise(date_in, here), getsunset(date_in, here))
End Function

Public Function getshaahzmanismga(date_in As hdate, here As location) As Long
    getshaahzmanismga = calcshaahzmanis(getalos72(date_in, here), gettzais72(date_in, here))
End Function

