Attribute VB_Name = "mod_NOAACalculator"
Public Const refraction As Double = 34 / 60#
Public Const solarradius As Double = 16 / 60#
Public Const earthradius As Double = 6356.9
Public Const PI_CONSTANT As Double = 3.14159265358979
Type location
    latitude As Double
    longitude As Double
    elevation As Double
End Type
Public Function ASin( _
      ByVal x As Double _
      ) As Double

   'Const PIover2 As Double = 1.5707963267949
    ASin = ArcSin(x)
   'If (x = 1) Then
   '   ASin = PIover2
   'ElseIf (x = -1) Then
   '   ASin = -PIover2
   'Else
   '   ASin = Atn(x / Sqr(-x * x + 1))
   'End If

End Function

Public Function Acos(x As Double) As Double
 If x = 1 Then 'x = 0.99999999999
    Acos = 0
 ElseIf x = -1 Then ' x = -0.99999999999
    Acos = Pi
 Else
'    Acos = ArcCos(x) '
    Acos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
 End If
End Function

Public Function Pi() As Double
  Pi = 4 * Atn(1#)
End Function

Public Function ArcSin(ByVal x As Double) As Double
 If x = 1 Then
    ArcSin = 1.5707963267949 ' x = 0.99999999999
 ElseIf x = -1 Then
 ArcSin = -1.5707963267949 'x = -0.99999999999
  Else
  ArcSin = Atn(x / Sqr(-x * x + 1))
 End If
End Function

Public Function ArcCos(ByVal x As Double) As Double
  ArcCos = ArcSin(x) + Pi / 2
End Function

Public Function radToDeg(ByVal angleRad As Double) As Double
    radToDeg = (180# * angleRad / PI_CONSTANT)
End Function

Public Function degToRad(ByVal angleDeg As Double) As Double
    degToRad = (PI_CONSTANT * angleDeg / 180#)
End Function

Public Function calcTimeJulianCent(ByVal JD As Double) As Double
    Dim jcent As Double
    jcent = (JD - 2451545#) / 36525#
    calcTimeJulianCent = jcent
End Function

Public Function calcJDFromJulianCent(ByVal jcent As Double) As Double
    Dim JD As Double
    JD = jcent * 36525# + 2451545#
    calcJDFromJulianCent = JD
End Function

Public Function calcGeomMeanLongSun(ByVal jcent As Double) As Double
    Dim gmls As Double
    gmls = 280.46646 + jcent * (36000.76983 + 0.0003032 * jcent)
    While gmls > 360#
        gmls = gmls - 360#
    Wend
    While gmls < 0#
        gmls = gmls + 360#
    Wend
    calcGeomMeanLongSun = gmls
End Function

Public Function calcGeomMeanAnomalySun(ByVal jcent As Double) As Double
    Dim gmas As Double
    gmas = 357.52911 + jcent * (35999.05029 - 0.0001537 * jcent)
    calcGeomMeanAnomalySun = gmas
End Function

Public Function calcEccentricityEarthOrbit(ByVal jcent As Double) As Double
    Dim eeo As Double
    eeo = 0.016708634 - jcent * (0.000042037 + 0.0000001267 * jcent)
    calcEccentricityEarthOrbit = eeo
End Function

Public Function calcSunEqOfCenter(ByVal jcent As Double) As Double
    Dim m As Double
    Dim mrad As Double
    Dim sinm As Double
    Dim sin2m As Double
    Dim sin3m As Double
    Dim seoc As Double

    m = calcGeomMeanAnomalySun(jcent)
    mrad = degToRad(m)
    sinm = Sin(mrad)
    sin2m = Sin(mrad + mrad)
    sin3m = Sin(mrad + mrad + mrad)
    
    seoc = sinm * (1.914602 - jcent * (0.004817 + 0.000014 * jcent)) + _
            sin2m * (0.019993 - 0.000101 * jcent) + _
            sin3m * 0.000289
    calcSunEqOfCenter = seoc
End Function

Public Function calcSunTrueLong(ByVal jcent As Double) As Double
    Dim gmls As Double
    Dim seoc As Double
    Dim stl As Double
    
    gmls = calcGeomMeanLongSun(jcent)
    seoc = calcSunEqOfCenter(jcent)
    
    stl = gmls + seoc
    calcSunTrueLong = stl
End Function

Public Function calcSunApparentLong(ByVal jcent As Double) As Double
    Dim stl As Double
    Dim omega As Double
    Dim sal As Double
    
    stl = calcSunTrueLong(jcent)
    
    omega = 125.04 - 1934.136 * jcent
    sal = stl - 0.00569 - 0.00478 * Sin(degToRad(omega))
    
    calcSunApparentLong = sal
End Function

Public Function calcMeanObliquityOfEcliptic(ByVal jcent As Double) As Double
    Dim seconds As Double
    Dim mooe As Double
    
    seconds = 21.448 - jcent * (46.815 + jcent * (0.00059 - jcent * 0.001813))
    mooe = 23# + (26# + (seconds / 60#)) / 60#
    
    calcMeanObliquityOfEcliptic = mooe
End Function


Public Function calcObliquityCorrection(ByVal jcent As Double) As Double
    Dim mooe As Double
    mooe = calcMeanObliquityOfEcliptic(jcent)
    Dim omega As Double
    omega = 125.04 - 1934.136 * jcent
    Dim oc As Double
    oc = mooe + 0.00256 * Cos(degToRad(omega))
    calcObliquityCorrection = oc
End Function

Public Function calcSunDeclination(ByVal jcent As Double) As Double
    Dim oc As Double
    oc = calcObliquityCorrection(jcent)
    Dim sal As Double
    sal = calcSunApparentLong(jcent)
    Dim sint As Double
    sint = Sin(degToRad(oc)) * Sin(degToRad(sal))
    Dim sd As Double
    sd = radToDeg(ASin(sint))
    calcSunDeclination = sd
End Function

Public Function calcEquationOfTime(ByVal jcent As Double) As Double
    Dim oc As Double
    oc = calcObliquityCorrection(jcent)
    Dim gmls As Double
    gmls = calcGeomMeanLongSun(jcent)
    Dim eeo As Double
    eeo = calcEccentricityEarthOrbit(jcent)
    Dim gmas As Double
    gmas = calcGeomMeanAnomalySun(jcent)
    Dim y As Double
    y = Tan(degToRad(oc) / 2#)
    y = y * y
    Dim sin2gmls As Double
    sin2gmls = Sin(2# * degToRad(gmls))
    Dim singmas As Double
    singmas = Sin(degToRad(gmas))
    Dim cos2gmls As Double
    cos2gmls = Cos(2# * degToRad(gmls))
    Dim sin4gmls As Double
    sin4gmls = Sin(4# * degToRad(gmls))
    Dim sin2gmas As Double
    sin2gmas = Sin(2# * degToRad(gmas))
    Dim Etime As Double
    Etime = y * sin2gmls - 2# * eeo * singmas + 4# * eeo * y * singmas * cos2gmls - 0.5 * y * y * sin4gmls - 1.25 * eeo * eeo * sin2gmas
    calcEquationOfTime = radToDeg(Etime) * 4#
End Function

Public Function calcHourAngleSunrise(ByVal lat As Double, ByVal solarDec As Double, ByVal zenith As Double) As Double
    Dim latRad As Double
    latRad = degToRad(lat)
    Dim sdRad As Double
    sdRad = degToRad(solarDec)
    Dim HA As Double
    HA = (Acos(Cos(degToRad(zenith)) / (Cos(latRad) * Cos(sdRad)) - Tan(latRad) * Tan(sdRad)))
    calcHourAngleSunrise = HA
End Function

Public Function calcHourAngleSunset(ByVal lat As Double, ByVal solarDec As Double, ByVal zenith As Double) As Double
    Dim latRad As Double
    latRad = degToRad(lat)
    Dim sdRad As Double
    sdRad = degToRad(solarDec)
    Dim HA As Double
    HA = (Acos(Cos(degToRad(zenith)) / (Cos(latRad) * Cos(sdRad)) - Tan(latRad) * Tan(sdRad)))
    calcHourAngleSunset = -HA
End Function

Public Function calcSolNoonUTC(ByVal JD As Double, ByVal longitude As Double) As Double
    Dim jcent As Double
    jcent = calcTimeJulianCent(JD)
    
    Dim tnoon As Double
    tnoon = calcTimeJulianCent(calcJDFromJulianCent(jcent) + longitude / 360#)
    Dim eqTime As Double
    eqTime = calcEquationOfTime(tnoon)
    Dim solNoonUTC As Double
    solNoonUTC = 720 + (longitude * 4) - eqTime
    
    Dim newt As Double
    newt = calcTimeJulianCent(calcJDFromJulianCent(jcent) - 0.5 + solNoonUTC / 1440#)
    
    eqTime = calcEquationOfTime(newt)
    solNoonUTC = 720 + (longitude * 4) - eqTime
    
    calcSolNoonUTC = solNoonUTC
End Function

Public Function calcSunriseUTC(ByVal JD As Double, ByVal latitude As Double, ByVal longitude As Double, ByVal zenith As Double) As Double
    Dim jcent As Double
    jcent = calcTimeJulianCent(JD)
    
    Dim noonmin As Double
    noonmin = calcSolNoonUTC(jcent, longitude)
    Dim tnoon As Double
    tnoon = calcTimeJulianCent(JD + noonmin / 1440#)
    
    Dim eqTime As Double
    eqTime = calcEquationOfTime(tnoon)
    Dim solarDec As Double
    solarDec = calcSunDeclination(tnoon)
    Dim hourAngle As Double
    hourAngle = calcHourAngleSunrise(latitude, solarDec, zenith)
    
    Dim delta As Double
    delta = longitude - radToDeg(hourAngle)
    Dim timeDiff As Double
    timeDiff = 4 * delta
    Dim timeUTC As Double
    timeUTC = 720 + timeDiff - eqTime
    
    Dim newt As Double
    newt = calcTimeJulianCent(calcJDFromJulianCent(jcent) + timeUTC / 1440#)
    eqTime = calcEquationOfTime(newt)
    solarDec = calcSunDeclination(newt)
    hourAngle = calcHourAngleSunrise(latitude, solarDec, zenith)
    delta = longitude - radToDeg(hourAngle)
    timeDiff = 4 * delta
    timeUTC = 720 + timeDiff - eqTime
    
    calcSunriseUTC = timeUTC
End Function

Public Function calcSunsetUTC(ByVal JD As Double, ByVal latitude As Double, ByVal longitude As Double, ByVal zenith As Double) As Double
    Dim jcent As Double
    jcent = calcTimeJulianCent(JD)
    
    Dim noonmin As Double
    noonmin = calcSolNoonUTC(jcent, longitude)
    Dim tnoon As Double
    tnoon = calcTimeJulianCent(JD + noonmin / 1440#)
    
    Dim eqTime As Double
    eqTime = calcEquationOfTime(tnoon)
    Dim solarDec As Double
    solarDec = calcSunDeclination(tnoon)
    Dim hourAngle As Double
    hourAngle = calcHourAngleSunset(latitude, solarDec, zenith)
    
    Dim delta As Double
    delta = longitude - radToDeg(hourAngle)
    Dim timeDiff As Double
    timeDiff = 4 * delta
    Dim timeUTC As Double
    timeUTC = 720 + timeDiff - eqTime
    
    Dim newt As Double
    newt = calcTimeJulianCent(calcJDFromJulianCent(jcent) + timeUTC / 1440#)
    eqTime = calcEquationOfTime(newt)
    solarDec = calcSunDeclination(newt)
    hourAngle = calcHourAngleSunset(latitude, solarDec, zenith)
    
    delta = longitude - radToDeg(hourAngle)
    timeDiff = 4 * delta
    timeUTC = 720 + timeDiff - eqTime
    
    calcSunsetUTC = timeUTC
End Function

Public Function getElevationAdjustment(ByVal elevation As Double) As Double
    Dim elevationAdjustment As Double
    elevationAdjustment = radToDeg(Acos(earthradius / (earthradius + (elevation / 1000))))
    getElevationAdjustment = elevationAdjustment
End Function

Public Function adjustZenith(ByVal zenith As Double, ByVal elevation As Double) As Double
    Dim adjustedZenith As Double
    adjustedZenith = zenith
    If zenith = 90# Then
        adjustedZenith = zenith + (solarradius + refraction + getElevationAdjustment(elevation))
    End If
    adjustZenith = adjustedZenith
End Function

Public Function getUTCSunrise(ByVal JD As Double, here As location, ByVal zenith As Double, ByVal adjustForElevation As Integer) As Double
    Dim elevation As Double
    elevation = IIf(adjustForElevation <> 0, here.elevation, 0)
    Dim adjustedZenith As Double
    adjustedZenith = adjustZenith(zenith, elevation)
    
    Dim sunrise As Double
    sunrise = calcSunriseUTC(JD, here.latitude, -here.longitude, adjustedZenith)
    sunrise = sunrise / 60
    
    Do While sunrise < 0#
        sunrise = sunrise + 24#
    Loop
    Do While sunrise >= 24#
        sunrise = sunrise - 24#
    Loop
    getUTCSunrise = sunrise
End Function

Public Function getUTCSunset(ByVal JD As Double, here As location, ByVal zenith As Double, ByVal adjustForElevation As Integer) As Double
    Dim elevation As Double
    elevation = IIf(adjustForElevation <> 0, here.elevation, 0)
    Dim adjustedZenith As Double
    adjustedZenith = adjustZenith(zenith, elevation)
    
    Dim sunset As Double
    sunset = calcSunsetUTC(JD, here.latitude, -here.longitude, adjustedZenith)
    sunset = sunset / 60
    
    Do While sunset < 0#
        sunset = sunset + 24#
    Loop
    Do While sunset >= 24#
        sunset = sunset - 24#
    Loop
    getUTCSunset = sunset
End Function

