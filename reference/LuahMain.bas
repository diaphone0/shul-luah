Attribute VB_Name = "LuahMain"
Option Explicit
Dim oEH As New clsAppEvents

Dim AnClkCount As Integer
Dim DiClkCount As Integer
Dim FullZmanimCount As Integer
Dim DataShapeCount As Integer

Dim time_first As Date
Dim time_current As Date
Dim dirname As String
Public date_ff As Double

Dim hebrewDate As hdate
Dim shabbos As hdate
Dim erevshabbos As hdate
Dim here As location
Dim today_is_chol As Boolean
Dim multi_mode As Boolean

Dim data_tag_col As New Collection
Dim timecount As Integer
Dim advancecount As Integer
Dim luahready As Boolean
Dim long_cycle As Integer
#If VBA7 Then
    Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, _
        ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Public TimerID As LongPtr
    Public TimerID2 As LongPtr
#Else
    Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, _
        ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Public TimerID As Long
    Public TimerID2 As Long
#End If

Function StartTimer()
    If Not InStr(1, ActivePresentation.FullName, ".pptm", vbTextCompare) > 0 Then
    If TimerID = 0& Then
        TimerID = SetTimer(0&, 0&, 1000&, AddressOf myTimer)  ' 1000 = 1sec
    End If
    If TimerID2 = 0& Then
        TimerID2 = SetTimer(0&, 0&, 10000&, AddressOf watchdog)   ' 1000 = 1sec
    End If
    End If
End Function

Function StopTimer()
    On Error Resume Next
    If TimerID <> 0& Then
        KillTimer 0&, TimerID
        TimerID = 0&
        Debug.Print "timer stopped"
    End If
    If TimerID2 <> 0& Then
        KillTimer 0&, TimerID2
        TimerID2 = 0&
    End If

End Function

'main timer process : this sub-routine CANNOT be interrupted by any error or itself
Sub myTimer()
     On Error Resume Next
    Dim sld As Slide
'    If Pause Then Exit Sub
    luah
    
    If Application.SlideShowWindows.Count > 0 Then
        advancecount = advancecount + 1
        Set sld = Application.SlideShowWindows(1).Presentation.Slides(Application.SlideShowWindows(1).View.CurrentShowPosition)
        If sld.SlideShowTransition.AdvanceTime = 0 Then
            advancecount = 0
        ElseIf advancecount >= sld.SlideShowTransition.AdvanceTime Then
            advancecount = 0
            Application.SlideShowWindows(1).View.Next
        End If
    End If
End Sub

Public Sub reset_shapes(Optional hideshapes As Boolean = False, Optional tmp_prs As Presentation = Nothing)
    'For Each SSW In SlideShowWindows
    Dim prs As Presentation
    Dim shp_cnt As Integer
    Dim shp As Shape
    Dim sld As Slide
    Dim delshape As Boolean
    Dim datestr As String
    Dim clkHour As Shape
    Dim clkMinute As Shape
    Dim clkSecond As Shape

    shp_cnt = 0
    If tmp_prs Is Nothing Then Set prs = Application.Presentations.Item(1) Else Set prs = tmp_prs
    
        Debug.Print "BEFORE RS"
        For Each sld In prs.Slides  'SSW.Presentation.Slides
            For Each shp In sld.Shapes
                If shp.Name Like "clkHour_*" Then
                    Set clkHour = shp
                    clkHour.Rotation = 0 '(hour(time) Mod 12) * 30 + ((minute(time) + second(time) / 60) / 60 * 30)
                    AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                    shp.GroupItems.Item(1).Visible = IIf(hideshapes, msoFalse, msoTrue)
                    shp.GroupItems.Item(2).Visible = msoFalse
                    shp.GroupItems.Item(3).Visible = IIf(hideshapes, msoFalse, msoTrue)
                    If Not hideshapes Then delshape = True
                ElseIf shp.Name Like "clkMinute_*" Then
                    Set clkMinute = shp
                    clkMinute.Rotation = 90 '(minute(time)) * 6 + (second(time) / 60 * 6)
                    AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                    shp.GroupItems.Item(1).Visible = IIf(hideshapes, msoFalse, msoTrue)
                    shp.GroupItems.Item(2).Visible = msoFalse
                    shp.GroupItems.Item(3).Visible = IIf(hideshapes, msoFalse, msoTrue)
                    If Not hideshapes Then delshape = True
                ElseIf shp.Name Like "clkSecond_*" Then
                    Set clkSecond = shp
                    AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                    clkSecond.Rotation = 15 '(second(time)) * 6
                    shp.GroupItems.Item(1).Visible = IIf(hideshapes, msoFalse, msoTrue)
                    shp.GroupItems.Item(2).Visible = msoFalse
                    shp.GroupItems.Item(3).Visible = IIf(hideshapes, msoFalse, msoTrue)
                    If Not hideshapes Then delshape = True
                ElseIf shp.Name Like "clkFrame_*" Then
                    'Set clkFrame = shp
                    AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                    If Not hideshapes Then delshape = True
                ElseIf shp.Name Like "AnClk_*" Then
                    shp.TextFrame.TextRange.Text = IIf(Not hideshapes, "#ANCLOCK", "")
                    AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                ElseIf shp.Name Like "clkDigital_*" Then
                    DiClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                    shp.TextFrame.TextRange.Text = IIf(Not hideshapes, "#DICLOCK", "")  'times
                ElseIf shp.Name Like "DataTagShape_*" Or shp.Tags.Item("ORG") <> "" Then '"dtHebDate_*" Then
                        'HebDateCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                        shp.Visible = IIf(hideshapes, msoFalse, msoTrue)
                        If Not hideshapes Then
                            shp.TextFrame.TextRange.Text = shp.Tags.Item("ORG") 'IIf(Not hideshapes, "#HEBDATE", "טוען") 'HDateOrFormat(hebrewDate, here)
                            shp.Tags.Delete "ORG"
                        End If
                    ElseIf shp.Name Like "dtFullZmanim_*" Then
                        FullZmanimCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                        datestr = IIf(Not hideshapes, "#FULLZMANIM", "טוען")
                        shp.TextFrame.TextRange.Text = datestr
                        shp.Visible = IIf(hideshapes, msoFalse, msoTrue)
                    ElseIf shp.HasTextFrame Then
                        If shp.TextFrame.TextRange.Text = "#MULTIMODE" Then
                            shp.Visible = IIf(hideshapes, msoFalse, msoTrue)
                            multi_mode = True
                        End If
                    End If
'                End If
                If Not hideshapes Then
                    shp.Name = "genshp" & shp_cnt
                    shp_cnt = shp_cnt + 1
                End If
                If delshape Then
                    shp.Delete
                    delshape = False
                End If
            Next shp
        Next sld
  'Next SSW
        Debug.Print "AFTER RS"

End Sub

Sub init_shapes(prs As Presentation, Optional skip_reload As Boolean = False)
    Dim clkShp As Shape
    Dim clkFrame As Shape
    Dim clkHour As Shape
    Dim clkMinute As Shape
    Dim clkSecond As Shape
    
    Dim clkHour1 As Shape
    Dim clkHour2 As Shape
    Dim clkHour3 As Shape
    Dim clkMinute1 As Shape
    Dim clkMinute2 As Shape
    Dim clkMinute3 As Shape
    Dim clkSecond1 As Shape
    Dim clkSecond2 As Shape
    Dim clkSecond3 As Shape
    Dim reload_needed As Boolean
    Dim Wnd As DocumentWindow
    Dim shp As Shape
    Dim sld As Slide
    Dim SSW As SlideShowWindow
    Dim i As Integer
    multi_mode = False
    For i = prs.Slides.Count To 1
        If prs.Slides(i).SlideShowTransition.Hidden = msoTrue Then prs.Slides(i).Delete
    Next i
    For Each sld In prs.Slides
        'Debug.Print "shp_cnt:" & sld.Shapes.Count
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If data_tag_func(shp, True) Then
                    reload_needed = True
                ElseIf shp.TextFrame.TextRange.Text = "#FULLZMANIM" Then
                    shp.TextFrame.TextRange.Text = ""
                    shp.Name = "dtFullZmanim_" & FullZmanimCount
                    reload_needed = True
                'ElseIf shp.TextFrame.TextRange.Text = "#DICLOCK" Then
                '    shp.TextFrame.TextRange.Text = ""
                '    shp.Name = "clkDigital_" & DiClkCount
                '    DiClkCount = DiClkCount + 1
                '    reload_needed = True
                ElseIf shp.TextFrame.TextRange.Text = "#MULTIMODE" Then
                    shp.Visible = msoFalse
                    multi_mode = True
                ElseIf shp.TextFrame.TextRange.Text = "#ANCLOCK" Then
                   Set clkShp = shp
                   clkShp.TextFrame.TextRange.Text = ""
    
                   'Set clkFrame = sld.Shapes.AddShape(msoShapeOval, clkShp.Left, clkShp.Top, clkShp.Width, clkShp.Height)
                   'clkFrame.Fill.Visible = msoFalse
                   shp.Name = "AnClk_" & AnClkCount
                   Set clkHour1 = sld.Shapes.AddLine(clkShp.Left + clkShp.Width / 2, clkShp.Top + (clkShp.Height / 3.8), clkShp.Left + clkShp.Width / 2, clkShp.Top + clkShp.Height / 2)
                   clkHour1.Line.Weight = 5
                   clkHour1.Line.ForeColor.RGB = RGB(0, 0, 0)
                   clkHour1.Visible = msoFalse
                   'clkHour1.Line.EndArrowheadStyle = msoArrowheadOval
                   Set clkHour2 = clkHour1.Duplicate.Item(1)
                   clkHour2.Top = clkHour1.Top + clkHour1.Height
                   clkHour2.Left = clkHour1.Left
                   clkHour2.Visible = msoFalse
                   Set clkHour3 = sld.Shapes.AddLine(clkShp.Left + clkShp.Width / 2, clkHour2.Top, clkShp.Left + clkShp.Width / 2, clkHour2.Top + clkShp.Height / 36)
                   clkHour3.Line.Weight = 5
                   clkHour3.Line.ForeColor.RGB = RGB(0, 0, 0)
                   clkHour3.Visible = msoFalse
                   
                   Set clkHour = sld.Shapes.Range(Array(clkHour1.Name, clkHour2.Name, clkHour3.Name)).Group
                   clkHour.Name = "clkHour_" & AnClkCount
                    
                   Set clkMinute1 = sld.Shapes.AddLine(clkShp.Left + clkShp.Width / 2, clkShp.Top + (clkShp.Height / 7), clkShp.Left + clkShp.Width / 2, clkShp.Top + clkShp.Height / 2)
                   clkMinute1.Line.Weight = 3
                   clkMinute1.Line.ForeColor.RGB = RGB(0, 0, 0)
                   clkMinute1.Visible = msoFalse
                   'clkMinute1.Line.EndArrowheadStyle = msoArrowheadOval
                   Set clkMinute2 = clkMinute1.Duplicate.Item(1)
                   clkMinute2.Top = clkMinute1.Top + clkMinute1.Height
                   clkMinute2.Left = clkHour1.Left
                   clkMinute2.Visible = msoFalse
                   Set clkMinute3 = sld.Shapes.AddLine(clkShp.Left + clkShp.Width / 2, clkMinute2.Top, clkShp.Left + clkShp.Width / 2, clkMinute2.Top + clkShp.Height / 32)
                   clkMinute3.Line.Weight = 3
                   clkMinute3.Line.ForeColor.RGB = RGB(0, 0, 0)
                   clkMinute3.Visible = msoFalse
                   
                   Set clkMinute = sld.Shapes.Range(Array(clkMinute1.Name, clkMinute2.Name, clkMinute3.Name)).Group
                   clkMinute.Name = "clkMinute_" & AnClkCount
                   
                   Set clkSecond1 = sld.Shapes.AddLine(clkShp.Left + clkShp.Width / 2, clkShp.Top + (clkShp.Height / 9), clkShp.Left + clkShp.Width / 2, clkShp.Top + clkShp.Height / 2)
                   clkSecond1.Line.Weight = 2
                   clkSecond1.Line.ForeColor.RGB = RGB(255, 0, 0)
                   clkSecond1.Visible = msoFalse
                   'clkSecond1.Line.EndArrowheadStyle = msoArrowheadOval
                   Set clkSecond2 = clkSecond1.Duplicate.Item(1)
                   clkSecond2.Top = clkSecond1.Top + clkSecond1.Height
                   clkSecond2.Left = clkHour1.Left
                   clkSecond2.Visible = msoFalse
                   Set clkSecond3 = sld.Shapes.AddLine(clkShp.Left + clkShp.Width / 2, clkSecond2.Top, clkShp.Left + clkShp.Width / 2, clkSecond2.Top + clkShp.Height / 19)
                   clkSecond3.Line.Weight = 4
                   clkSecond3.Line.ForeColor.RGB = RGB(255, 0, 0)
                   clkSecond3.Visible = msoFalse
                   
                   Set clkSecond = sld.Shapes.Range(Array(clkSecond1.Name, clkSecond2.Name, clkSecond3.Name)).Group
                   clkSecond.Name = "clkSecond_" & AnClkCount
                                                         
                   Set clkFrame = sld.Shapes.AddLine(clkShp.Left + clkShp.Width / 2 - 1, clkShp.Top + clkShp.Height / 2 - 1, clkShp.Left + clkShp.Width / 2, clkShp.Top + clkShp.Height / 2)
                   clkFrame.Line.Weight = 4
                   clkFrame.Line.ForeColor.RGB = RGB(0, 0, 0)
                   clkFrame.Line.EndArrowheadStyle = msoArrowheadOval
                   clkFrame.Name = "clkFrame_" & AnClkCount
                   clkFrame.Visible = msoFalse
                   AnClkCount = AnClkCount + 1
                    reload_needed = True
                End If
                
            End If
        Next shp
    Next sld
    If reload_needed And Not skip_reload Then
    '    plog "rneede!d"
    '    StopTimer
    '    plog "timer stopped"
    '    Call reset_shapes(True)
    '    plog "shapes rese!t"
    '    prs.Save
    '    plog "psave!d"
    '    plog "RUN:" & dirname & "reload.cmd 1"
    '    Shell dirname & "reload.cmd 1"
    '    plog "shell_done"
    '    Application.Quit
    '    Application.SlideShowWindows(1).View.Exit
    '    Application.SlideShowWindows(1).Presentation.Close
    '    For Each Wnd In Application.Windows
    '        Wnd.Close
    '    Next Wnd
    '    For Each prs1 In Application.Presentations
    '        prs1.Close
    '    Next prs1
    '    reload_needed = False
    '    End
    End If
    
    With here ' =Eli
    .latitude = 32.06
    .longitude = 35.26
    .elevation = 720
    End With
    hebrewDate = ConvertDate(mktm(now))
    hebrewDate.offset = 3600 * 3
'    If Not InStr(1, ActivePresentation.FullName, ".pptm", vbTextCompare) > 0 Then date_ff = 0
    'If date_ff <> 0 Then HDateAddHour hebrewDate, CLng(date_ff)
    Call SetEY(hebrewDate, 1)
    
    shabbos = hebrewDate
    erevshabbos = hebrewDate
    
    If IsAssurBeMelachah(erevshabbos) Then
        Call HDateAddDay(erevshabbos, -1)
    Else
        Do While IsCandleLighting(erevshabbos) = 0
            Call HDateAddDay(erevshabbos, 1)
        Loop 'TODO check chanuka CL on wday=1
        shabbos = erevshabbos
        Call HDateAddDay(shabbos, 1)
    End If

    
    If (hebrewDate.wday = 6 And mkdate(HDateGregorian(hebrewDate)) > mkdate(HDateGregorian(getchatzosgra(hebrewDate, here)))) Or hebrewDate.wday = 0 Then
        today_is_chol = False
    Else
        today_is_chol = True
    End If
    today_is_chol = today_is_chol
    'Dim i As Integer
    If multi_mode And prs.Slides.Count > 1 Then
        prs.Slides(1).SlideShowTransition.Hidden = IIf((today_is_chol) Or (Not multi_mode), msoTrue, msoFalse)
        prs.Slides(prs.Slides.Count).SlideShowTransition.Hidden = IIf((Not today_is_chol) Or (Not multi_mode), msoTrue, msoFalse)
        If multi_mode And prs.Slides.Count >= 3 Then
            For i = 2 To prs.Slides.Count - 1
            prs.Slides(i).SlideShowTransition.Hidden = msoTrue
            Next i
        End If
        
        'If Application.SlideShowWindows.Count > 0 Then
        '    Set SSW = Application.SlideShowWindows(1)
        '    Select Case SSW.View.CurrentShowPosition
        '        Case 1
        '            If today_is_chol Then
        '                SSW.View.GotoSlide SSW.Presentation.Slides.Count, msoTrue
        '            End If
        '        Case SSW.Presentation.Slides.Count
        '            If Not today_is_chol Then
        '                SSW.View.GotoSlide 1, msoTrue
        '            End If
        '    End Select
        'End If
    End If
    luahready = True
    long_cycle = 2
End Sub

Sub update_shapes(sld As Slide, Optional clkonly As Boolean = False)
    Dim clkShp As Shape
    Dim clkFrame As Shape
    Dim clkHour As Shape
    Dim clkMinute As Shape
    Dim clkSecond As Shape
    
    Dim clkHour1 As Shape
    Dim clkHour2 As Shape
    Dim clkHour3 As Shape
    Dim clkMinute1 As Shape
    Dim clkMinute2 As Shape
    Dim clkMinute3 As Shape
    Dim clkSecond1 As Shape
    Dim clkSecond2 As Shape
    Dim clkSecond3 As Shape
    Dim datestr As String
    
    Dim shp As Shape
            
    Dim i As Integer
    'Debug.Print "shp_cnt:" & sld.Shapes.Count
    For i = 1 To sld.Shapes.Count
        Set shp = sld.Shapes(i)
        'Debug.Print shp.Name & " - " & i  ' shp.TextFrame.TextRange.Text
plog "CHK7"
        If shp.Name Like "clkHour_*" Then
            Set clkHour = shp
            clkHour.Rotation = (hebrewDate.hour Mod 12) * 30 + ((hebrewDate.min + hebrewDate.sec / 60) / 60 * 30)
            shp.GroupItems.Item(1).Visible = msoTrue
            shp.GroupItems.Item(2).Visible = msoFalse
            shp.GroupItems.Item(3).Visible = msoTrue
            AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
        ElseIf shp.Name Like "clkMinute_*" Then
            Set clkMinute = shp
            clkMinute.Rotation = (hebrewDate.min) * 6 + (hebrewDate.sec / 60 * 6)
            shp.GroupItems.Item(1).Visible = msoTrue
            shp.GroupItems.Item(2).Visible = msoFalse
            shp.GroupItems.Item(3).Visible = msoTrue
            AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
        ElseIf shp.Name Like "clkSecond_*" Then
            Set clkSecond = shp
            clkSecond.Rotation = (hebrewDate.sec) * 6
            shp.GroupItems.Item(1).Visible = msoTrue
            shp.GroupItems.Item(2).Visible = msoFalse
            shp.GroupItems.Item(3).Visible = msoTrue
            AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
        ElseIf shp.Name Like "clkFrame_*" Then
            Set clkFrame = shp
            clkFrame.Visible = True
            clkFrame.ZOrder msoBringToFront
            AnClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
        ElseIf clkonly = False Then
            If shp.Name Like "clkDigital_*" Then
                DiClkCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                shp.Visible = msoTrue
                shp.TextFrame.TextRange.Text = time
            ElseIf shp.Name Like "DataTagShape_*" Then
                Call data_tag_func(shp, False)
                DataShapeCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                shp.Visible = msoTrue
            ElseIf shp.Name Like "dtFullZmanim_*" Then
                FullZmanimCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                shp.Visible = msoTrue
                datestr = _
                "עלות השחר (72 דק'): " & String(2, Asc(vbTab)) & Format(mkdate(HDateGregorian(getalos72(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "הנץ החמה: " & String(2, Asc(vbTab)) & Format(mkdate(HDateGregorian(getsunrise(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "סוף זמן קריאת שמע (מג""א): " & String(1, Asc(vbTab)) & Format(mkdate(HDateGregorian(getshmamga(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "סוף זמן קריאת שמע (גר""א): " & String(1, Asc(vbTab)) & Format(mkdate(HDateGregorian(getshmagra(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "סוף זמן תפילה (מג""א): " & String(1, Asc(vbTab)) & Format(mkdate(HDateGregorian(gettefilamga(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "סוף זמן תפילה (גר""א): " & String(1, Asc(vbTab)) & Format(mkdate(HDateGregorian(gettefilagra(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "חצות היום: " & String(2, Asc(vbTab)) & Format(mkdate(HDateGregorian(getchatzosgra(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "מנחה גדולה: " & String(2, Asc(vbTab)) & Format(mkdate(HDateGregorian(getminchagedolagra(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "מנחה קטנה: " & String(2, Asc(vbTab)) & Format(mkdate(HDateGregorian(getminchaketanagra(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "שקיעה: " & String(3, Asc(vbTab)) & Format(mkdate(HDateGregorian(getelevationsunset(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "צאת הכוכבים: " & String(2, Asc(vbTab)) & Format(mkdate(HDateGregorian(gettzais8p5(hebrewDate, here))), "hh:mm")
                shp.TextFrame.TextRange.Text = datestr
            End If
        End If
    Next i

End Sub

Sub watchdog()
    'Debug.Print "WD"
    dirname = "C:\luah\" ' Mid(pathname, 1, InStrRev(pathname, "\"))
'    If InStr(1, ActivePresentation.FullName, ".pptm", vbTextCompare) > 0 Then GoTo skip_check
    If year(time_first) <> year(now) Then
        time_first = FileDateTime(dirname & "luah.pptx")
    Else
        'If luahready And timecount > 9 Then
            time_current = FileDateTime(dirname & "luah.pptx")
        
            If time_current > time_first Then
                time_first = time_current
                'reload_ssw
                Shell dirname & "reload.cmd 1"
            End If
        'End If
    End If
'skip_check:


End Sub
Sub reload_ssw()
    Dim tmp_prs As Presentation
    Dim prs1 As Presentation
    dirname = "C:\luah\"
    StopTimer
    If SlideShowWindows.Count > 0 Then
        Application.SlideShowWindows(1).Presentation.Save
        Application.SlideShowWindows(1).View.Exit
    End If
    FileCopy dirname & "luah.pptx", Environ("temp") & "\luah.tmp.pptx"
    Set tmp_prs = Application.Presentations.Open(Environ("temp") & "\luah.tmp.pptx", False, False, False)
    init_shapes tmp_prs, True
    reset_shapes True, tmp_prs
    tmp_prs.Save
    tmp_prs.Close
    luahready = False
    'MsgBox "before3"
    If Dir(Environ("temp") & "\modaot.pps") <> "" Then Kill Environ("temp") & "\modaot.pps"
    Name Environ("temp") & "\luah.tmp.pptx" As Environ("temp") & "\modaot.pps"
    Set prs1 = Application.Presentations.Open(Environ("temp") & "\modaot.pps", False, False, False)
    'StartTimer
'    prs1.Slides (3) 'IIf((Not today_is_chol) Or (Not multi_mode), 1, prs1.Slides.Count - 1)
    'prs1.SlideShowSettings.RangeType = ppShowSlideRange
    'prs1.SlideShowSettings.StartingSlide = 3
    'prs1.SlideShowSettings.EndingSlide = 2
    prs1.SlideShowSettings.Run
End Sub
Sub luah()
    Dim i As Integer, j As Integer
    
    Dim sld As Slide
    Dim SSW As SlideShowWindow
    Dim currentSlideIndex As Long
    Dim reload_needed As Boolean
    Dim Wnd As DocumentWindow
    reload_needed = False
    plog "CHK1"
    plog "CHK3"
    If Application.SlideShowWindows.Count > 0 Then
        Set SSW = Application.SlideShowWindows(1)
    Else
        Exit Sub
    End If
    
    If luahready = False Then
        Debug.Print "first luah"
        Call init_shapes(ActivePresentation, IIf(InStr(1, ActivePresentation.FullName, ".pptm", vbTextCompare) > 0, True, False))
        If SSW.Presentation.Slides.Count >= 3 Then
            For i = 2 To SSW.Presentation.Slides.Count - 1
              SSW.Presentation.Slides(i).SlideShowTransition.Hidden = msoFalse
            Next i
        End If
    End If
    
    
    'If shabbos.wday <> 0 Then Call HDateAddDay(shabbos, (7 - shabbos.wday))
    'If erevshabbos.wday <> 6 Then Call HDateAddDay(erevshabbos, (6 - erevshabbos.wday))
    
    plog "CHK4"
    With here ' =Eli
    .latitude = 32.06
    .longitude = 35.26
    .elevation = 720
    End With
    hebrewDate = ConvertDate(mktm(now))
    hebrewDate.offset = 3600 * 3
    If Not InStr(1, ActivePresentation.FullName, ".pptm", vbTextCompare) > 0 Then date_ff = 0
    If date_ff <> 0 Then HDateAddHour hebrewDate, CLng(date_ff)
    Call SetEY(hebrewDate, 1)
'    shabbos = hebrewDate
'    erevshabbos = hebrewDate
    
'    If IsAssurBeMelachah(erevshabbos) Then
'        Call HDateAddDay(erevshabbos, -1)
'    Else
'        Do While IsCandleLighting(erevshabbos) = 0
'            Call HDateAddDay(erevshabbos, 1)
'        Loop 'TODO check chanuka CL on wday=1
'        shabbos = erevshabbos
'        Call HDateAddDay(shabbos, 1)
'    End If

    currentSlideIndex = SSW.View.CurrentShowPosition
    If currentSlideIndex = 0 Then Exit Sub
    plog "CHK5"
    If luahready And timecount > long_cycle Then

        shabbos = hebrewDate
        erevshabbos = hebrewDate
        
        If IsAssurBeMelachah(erevshabbos) Then
            Call HDateAddDay(erevshabbos, -1)
        Else
            Do While IsCandleLighting(erevshabbos) = 0
                Call HDateAddDay(erevshabbos, 1)
            Loop 'TODO check chanuka CL on wday=1
            shabbos = erevshabbos
            Call HDateAddDay(shabbos, 1)
        End If


        If (hebrewDate.wday = 6 And mkdate(HDateGregorian(hebrewDate)) > mkdate(HDateGregorian(getchatzosgra(hebrewDate, here)))) Or hebrewDate.wday = 0 Then
            today_is_chol = False
        Else
            today_is_chol = True
        End If
        today_is_chol = today_is_chol
        If multi_mode And SSW.Presentation.Slides.Count > 1 Then
            Select Case currentSlideIndex
                Case 1
                    If today_is_chol Then
                        SSW.View.GotoSlide SSW.Presentation.Slides.Count, msoTrue
                        SSW.Presentation.Slides(1).SlideShowTransition.Hidden = msoTrue
                        SSW.Presentation.Slides(SSW.Presentation.Slides.Count).SlideShowTransition.Hidden = msoFalse
                    End If
                Case SSW.Presentation.Slides.Count
                    If Not today_is_chol Then
                        SSW.View.GotoSlide 1, msoTrue
                        SSW.Presentation.Slides(1).SlideShowTransition.Hidden = msoFalse
                        SSW.Presentation.Slides(SSW.Presentation.Slides.Count).SlideShowTransition.Hidden = msoTrue
                    End If
            End Select
        End If
    End If
    Set sld = SSW.Presentation.Slides(currentSlideIndex)
    update_shapes sld, IIf(luahready And timecount > long_cycle, False, True)

    If luahready And timecount > long_cycle Then
        timecount = 0
        long_cycle = 9
    End If
    timecount = timecount + 1

End Sub


Public Sub data_tag_init()
    data_tag_col.Add "HEBDATE"
    data_tag_col.Add "DAYZMANIM"
    data_tag_col.Add "DAFYOMI"
    data_tag_col.Add "PARSHA"
    data_tag_col.Add "SHABBOS"
'    data_tag_col.Add "NEROT"
'    data_tag_col.Add "MINHA"
'    data_tag_col.Add "SHIR"
    data_tag_col.Add "MOZASH"
    data_tag_col.Add "LIMUDYOMI"
'    data_tag_col.Add "FULLZMANIM"
    data_tag_col.Add "DICLOCK"
    data_tag_col.Add "ALOS72"
    data_tag_col.Add "SUNRISE"
    data_tag_col.Add "SHMAMGA"
    data_tag_col.Add "SHMAGRA"
    data_tag_col.Add "TFILAMGA"
    data_tag_col.Add "TFILAGRA"
    data_tag_col.Add "CHAZOT"
    data_tag_col.Add "BGMINHA"
    data_tag_col.Add "LTMINHA"
    data_tag_col.Add "SUNSET"
    data_tag_col.Add "TZAIS"
End Sub


Function data_tag_func(shp As Shape, do_init As Boolean) As Boolean
    If data_tag_col.Count = 0 Then Call data_tag_init
    Dim stag
    Dim resultstr As String
    data_tag_func = False
    
    If do_init Then
        If shp.HasTextFrame Then
            For Each stag In data_tag_col
                If InStr(1, shp.TextFrame.TextRange.Text, "#" & stag) > 0 And (shp.Tags.Count = 0 Or (shp.Tags.Count <> 0 And Not shp.Name Like "DataTagShape_*")) And InStr(1, shp.TextFrame.TextRange.Text, "טוען!") <> 1 Then
                    shp.Name = "DataTagShape" & "_" & DataShapeCount
                    DataShapeCount = CInt(Mid(shp.Name, InStr(1, shp.Name, "_") + 1)) + 1
                    'shp.Tags.Add "ORG", shp.TextFrame.TextRange.Text
                    shp.TextFrame.TextRange.Text = "טוען!" & shp.TextFrame.TextRange.Text
                    shp.Visible = msoFalse
                    data_tag_func = True
                    Exit Function
                End If
            Next stag
            'shp.TextFrame.TextRange.Text
        End If
    Else
        If InStr(1, shp.TextFrame.TextRange.Text, "טוען!") = 1 Then
            shp.TextFrame.TextRange.Text = Mid(shp.TextFrame.TextRange.Text, 6)
            resultstr = shp.TextFrame.TextRange.Text ''Mid(shp.TextFrame.TextRange.Text, 6)
            shp.Tags.Add "ORG", resultstr
            shp.TextFrame.TextRange.Text = "טוען!"
        End If
        resultstr = shp.Tags("ORG")
        For Each stag In data_tag_col
            Call data_tag_parse(stag, resultstr)
        Next stag
        shp.TextFrame.TextRange.Text = resultstr
        shp.Visible = msoTrue
        data_tag_func = True
    End If
End Function

Sub data_tag_parse(ByVal stag As String, ByRef resultstr) ' As Boolean
    Dim tag_start As Integer
    Dim tag_len As Integer
    Dim tag_val As String
    Dim after_tag As String
    Dim before_tag As String
    Dim add_offset As String
    Dim rambamstr As String
    Dim parsh As parshah
    Dim ytov As yomtov
    Dim i As Integer
    Dim j As Integer

    tag_start = InStr(1, resultstr, "#" & stag)
    If tag_start > 0 Then
        tag_len = Len(stag) + 1
        before_tag = Mid(resultstr, 1, tag_start - 1)
        after_tag = Mid(resultstr, tag_start + tag_len)
        If after_tag Like "+[0-9]*" Or after_tag Like "-[0-9]*" Then
            add_offset = Mid(after_tag, 1, Len(str(Val(after_tag))))
            after_tag = Replace(after_tag, add_offset, "", , 1)
        Else
             add_offset = ""
        End If
        Select Case stag
            Case "HEBDATE": tag_val = HDateOrFormat(hebrewDate, here)
            Case "DAYZMANIM"
                tag_val = _
                "זריחה: " & Format(mkdate(HDateGregorian(getsunrise(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "שקיעה: " & Format(mkdate(HDateGregorian(getelevationsunset(hebrewDate, here))), "hh:mm") & vbCrLf & _
                "צאת הכוכבים: " & Format(mkdate(HDateGregorian(gettzais8p5(hebrewDate, here))), "hh:mm")
            Case "DAFYOMI": tag_val = "דף יומי:" & vbCrLf & GetDafYomiFormat(mkdate(HDateGregorian(hebrewDate)))
            Case "PARSHA"
                parsh = GetParshah(shabbos)
                If parsh <> NOPARSHAH Then
                    tag_val = "שבת פרשת " & parshahformat(parsh)
                Else
                    ytov = GetYomTov(shabbos)
                    If ytov <> CHOL Then tag_val = YomTovFormat(ytov)
                End If
                ytov = GetSpecialShabbos(shabbos)
                tag_val = tag_val & IIf(ytov <> CHOL, vbCrLf & "(" & YomTovFormat(ytov) & ")", "")
            'Case "NEROT": tag_val = Format(DateAdd("n", -30, tround(mkdate(HDateGregorian(getelevationsunset(erevshabbos, here))))), "hh:mm")
            'Case "MINHA": tag_val = Format(DateAdd("n", -25, tround(mkdate(HDateGregorian(getelevationsunset(erevshabbos, here))))), "hh:mm")
            'Case "SHIR": tag_val = Format(DateAdd("n", -15, tround(mkdate(HDateGregorian(getelevationsunset(erevshabbos, here))))), "hh:mm")
            Case "SHABBOS": tag_val = Format(DateAdd("n", Val(add_offset), tround(mkdate(HDateGregorian(getelevationsunset(erevshabbos, here))))), "hh:mm")
            Case "MOZASH": tag_val = Format(DateAdd("n", Val(add_offset), tround(mkdate(HDateGregorian(gettzais8p5(shabbos, here))))), "hh:mm")
            Case "LIMUDYOMI"
                rambamstr = GetRambam(hebrewDate, True)
                i = InStr(1, rambamstr, ";")
                j = InStr(i + 1, rambamstr, ";")
                rambamstr = Mid(rambamstr, 1, i - 1) & " - " & Mid(rambamstr, i + 1, j - i - 1) & vbCrLf & Mid(rambamstr, j + 1)
                tag_val = "דף יומי:" & vbCrLf & GetDafYomiFormat(mkdate(HDateGregorian(hebrewDate))) & vbCrLf & vbCrLf & _
                "רמב""ם יומי:" & vbCrLf & rambamstr & vbCrLf & vbCrLf & _
                "תהילים:" & vbCrLf & Tehillim(hebrewDate)
            Case "ALOS72": tag_val = Format(mkdate(HDateGregorian(getalos72(hebrewDate, here))), "hh:mm")
            Case "SUNRISE": tag_val = Format(DateAdd("n", Val(add_offset), mkdate(HDateGregorian(getsunrise(hebrewDate, here)))), "hh:mm")
            Case "SHMAMGA": tag_val = Format(mkdate(HDateGregorian(getshmamga(hebrewDate, here))), "hh:mm")
            Case "SHMAGRA": tag_val = Format(mkdate(HDateGregorian(getshmagra(hebrewDate, here))), "hh:mm")
            Case "TFILAMGA": tag_val = Format(mkdate(HDateGregorian(gettefilamga(hebrewDate, here))), "hh:mm")
            Case "TFILAGRA": tag_val = Format(mkdate(HDateGregorian(gettefilagra(hebrewDate, here))), "hh:mm")
            Case "CHAZOT": tag_val = Format(mkdate(HDateGregorian(getchatzosgra(hebrewDate, here))), "hh:mm")
            Case "BGMINHA": tag_val = Format(mkdate(HDateGregorian(getminchagedolagra(hebrewDate, here))), "hh:mm")
            Case "LTMINHA": tag_val = Format(mkdate(HDateGregorian(getminchaketanagra(hebrewDate, here))), "hh:mm")
            Case "SUNSET": tag_val = Format(DateAdd("n", Val(add_offset), mkdate(HDateGregorian(getelevationsunset(hebrewDate, here)))), "hh:mm")
            Case "TZAIS": tag_val = Format(DateAdd("n", Val(add_offset), mkdate(HDateGregorian(gettzais8p5(hebrewDate, here)))), "hh:mm")
            Case "DICLOCK": tag_val = time
        End Select
        resultstr = before_tag & tag_val & after_tag
        Call data_tag_parse(stag, resultstr)
        Exit Sub
    End If
'    data_tag_parse = False

End Sub

'the timer must be stopped after finishing the show
Public Sub OnSlideShowTerminate(SSW As SlideShowWindow)
Debug.Print "teminate!"
    StopTimer
'    Call reset_shapes(True)
'    reset_shapes
End Sub

'To start the clock automactically
Sub OnSlideShowPageChange(ByVal SSW As SlideShowWindow)
'  Set oEH.App = Application
'If SSW.View.CurrentShowPosition = 1 Then SSW.View.GotoSlide 3, msoFalse
    Dim weekday As Integer
    Dim nextslide As Integer
    Dim prs As Presentation
    Set prs = SSW.Presentation
    Dim i As Integer
'    Dim pathname As String
'    Dim dirname As String
'    Debug.Print year(time_first)

    If prs.Slides.Count >= 3 Then
        For i = 2 To prs.Slides.Count - 1
          prs.Slides(i).SlideShowTransition.Hidden = msoFalse
        Next i
    End If


'plog "CHK0.2"
'Debug.Print "begin!"
'    If SSW.View.CurrentShowPosition = Default Then
'plog "CHK0.3"
        luah
        luah
        luah
        StartTimer
'plog "CHK0.4"
'    Else
        'StopTimer
'    End If
'    Debug.Print "im at:" & currentSlideIndex
    
'    If nextslide > 0 And nextslide <> currentSlideIndex Then SSW.View.GotoSlide nextslide, msoTrue
End Sub

Sub Auto_Open()
plog "CHK0.1"
'StartTimer

End Sub

Sub test_luah()
    luah
    init_shapes ActivePresentation, True
    update_shapes ActivePresentation.Slides(1)
'    update_shapes ActivePresentation.Slides(3)

End Sub

Sub test_zmanim_eli()
  Dim offset As Long
  Dim i As Integer
  'Dim here As location
  With here
  .latitude = 32.06
  .longitude = 35.26
  .elevation = 720
  End With
  offset = 3600 * 3
  Dim tm As TMStruct
  Dim datestr As String
'  Dim hebrewDate As hdate
'  Dim shabbos As hdate
'  Dim erevshabbos As hdate
  Dim orgdate As Date
  Dim ytov As yomtov
  Dim parsh As parshah
  Dim shabbosstr As String
  Dim zmanstr As String
  
  'Dim zman As hdate
  
  'Open "C:\Users\ngabbay\OneDrive - Intel Corporation\Documents\vbout2.log" For Output As #1
  'orgdate = #4/15/2031 1:00:00 PM# '10/7/2024 1:00:00 PM#
  orgdate = DateAdd("d", -2, now)
  'For i = 1 To 6581
  tm = mktm(orgdate)
  hebrewDate = ConvertDate(tm)
  hebrewDate.offset = 3600 * 3
  Call SetEY(hebrewDate, 1)
'  shabbos = hebrewDate
'  erevshabbos = hebrewDate
'  If shabbos.wday <> 0 Then Call HDateAddDay(shabbos, (7 - shabbos.wday))
'  If erevshabbos.wday <> 6 Then Call HDateAddDay(erevshabbos, (6 - erevshabbos.wday))
  
  

    shabbos = hebrewDate
    erevshabbos = hebrewDate
    
    If IsAssurBeMelachah(erevshabbos) Then
        Call HDateAddDay(erevshabbos, -1)
    Else
        Do While IsCandleLighting(erevshabbos) = 0
            Call HDateAddDay(erevshabbos, 1)
        Loop 'TODO check chanuka CL on wday=1
        shabbos = erevshabbos
        Call HDateAddDay(shabbos, 1)
    End If
  
  
  'If GetShabbosMevorchim(shabbos) Then
  'Debug.Print "מולד: " & MoladFormat(GetMolad(hebrewDate.year, hebrewDate.month + 1))
  'End If

'Debug.Print HDateOrFormat(hebrewDate, here)
                    parsh = GetParshah(shabbos)
                    If parsh <> NOPARSHAH Then
                        shabbosstr = parshahformat(parsh)
                    Else
                        ytov = GetYomTov(shabbos)
                        If ytov <> CHOL Then shabbosstr = YomTovFormat(ytov)
                    End If
                    ytov = GetSpecialShabbos(shabbos)
                    shabbosstr = shabbosstr & IIf(ytov <> CHOL, " (" & YomTovFormat(ytov) & ")", "")

                    datestr = _
                    "0" & "," & NumToHChar(hebrewDate.day) & vbCrLf & _
                    "1" & "," & NumToHMonth(hebrewDate.month, hebrewDate.leap) & vbCrLf & _
                    "3" & "," & Format(tround(mkdate(HDateGregorian(getalos72(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "6" & "," & Format(tround(mkdate(HDateGregorian(getsunrise(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "8" & "," & Format(tround(mkdate(HDateGregorian(getshmamga(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "9" & "," & Format(tround(mkdate(HDateGregorian(getshmagra(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "11" & "," & Format(tround(mkdate(HDateGregorian(gettefilamga(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "12" & "," & Format(tround(mkdate(HDateGregorian(gettefilagra(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "13" & "," & Format(tround(mkdate(HDateGregorian(getchatzosgra(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "15" & "," & Format(tround(mkdate(HDateGregorian(getminchagedolagra(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "16" & "," & Format(tround(mkdate(HDateGregorian(getminchaketanagra(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "19" & "," & Format(tround(mkdate(HDateGregorian(getelevationsunset(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "20" & "," & Format(tround(mkdate(HDateGregorian(gettzais8p5(hebrewDate, here)))), "h:mm") & vbCrLf & _
                    "22" & "," & month(mkdate(HDateGregorian(hebrewDate))) & vbCrLf & _
                    "23" & "," & day(mkdate(HDateGregorian(hebrewDate))) & vbCrLf & _
                    "24" & "," & weekday(mkdate(HDateGregorian(hebrewDate))) & vbCrLf & _
                    "26" & "," & shabbosstr 'parshahformat(GetParshah(shabbos)) & IIf(GetSpecialShabbos(shabbos) <> CHOL, " (" & YomTovFormat(GetSpecialShabbos(shabbos)) & ")", "")
                    
                    'Print #1, datestr
                    orgdate = DateAdd("d", 1, orgdate)
'                    Debug.Print datestr
'Debug.Print "זריחה: " & Format(mkdate(HDateGregorian(getsunrise(hebrewDate, here))), "hh:mm:ss")
'Debug.Print "שקיעה: " & Format(mkdate(HDateGregorian(getelevationsunset(hebrewDate, here))), "hh:mm:ss")
'Debug.Print "צאת הכוכבים: " & Format(mkdate(HDateGregorian(gettzais8p5(hebrewDate, here))), "hh:mm:ss")
'Debug.Print "דף יומי: " & GetDafYomiFormat(mkdate(HDateGregorian(hebrewDate)))
'Debug.Print "שבת " & shabbosstr
'Debug.Print "הדלקת נרות: " & Format(DateAdd("n", -30, mkdate(HDateGregorian(getelevationsunset(erevshabbos, here)))), "hh:mm:ss")
'Debug.Print "4צאת שבת: " & Format(mkdate(HDateGregorian(gettzais8p5(shabbos, here))), "hh:mm:ss")
    zmanstr = "שקיעה: #SUNSET !!!"
    Call data_tag_parse("SUNSET", zmanstr)
    Debug.Print zmanstr
    
    'Next i
    'Close #1
End Sub
Sub plog(str As String)
    ''Open "sbgout.txt" For Append As #1
    ''Print #1, str
   '' Close #1
  '' Debug.Print str
End Sub
Function tround(t As Date) As Date
    tround = IIf(second(t) > 29, DateAdd("n", 1, t), t)
End Function


Public Sub fast_f()
    date_ff = date_ff + 24
    luah
    update_shapes ActivePresentation.Slides(1)
    update_shapes ActivePresentation.Slides(3)
End Sub
