Class TimelineInfo
'
'    Used to hold Timeline data, i.e. start & end dates, no. of nodes, time inteval unit, time interval width (pixels)
'
    Private IsUpdated_, Unit_, Pixels_, DateMin_, DateMax_, Nodes_, DateLimit_, ID_

    Private Sub Class_Initialize()
        Nodes_ = 0
        DateMin_ = DateValue("31 " & MonthName(12) & " 9999")
        DateMax_ = DateValue("01 " & MonthName(1) & " 100")
        DateLimit_ = DateMax_
        IsUpdated_ = False
    End Sub

    Public Property Get DateMax
        DateMax = DateMax_
    End Property

    Public Property Get DateMin
        DateMin = DateMin_
    End Property
    
    Public Property Get ID
        ID = ID_
    End Property
    
    Public Property Get Nodes
        Nodes = Nodes_
    End Property

    Public Property Get Unit
        If Not IsUpdated_ Then Update_
        Unit = Unit_
    End Property

    Public Property Get Pixels
        If Not IsUpdated_ Then Update_
        Pixels = Pixels_
    End Property

    Private Sub CheckDate_(strDate)
        Dim nDate
        If IsDate(strDate) Then
            nDate = DateValue(strDate)
            If nDate < DateMin_ Then DateMin_ = nDate 
            If nDate  > DateMax_ Then DateMax_ = nDate
        End If
    End Sub

    Private Sub Update_
    '    Dim nYears, nDensity
    '    nYears = Year(DateMax_) - Year(DateMin_)
    '    If nYears > 0 And Nodes_ > 0 Then
    '        nDensity=Nodes_ * 100 / nYears
    '        Unit_ = Int(Log(nDensity)/Log(10))+7
    '        Pixels_ = Round(((Log(nDensity)/Log(10))-Unit_ - 7) * 100, 0)
    '        If Pixels_ >= 50 Then Pixels_ = 200
    '        If Pixels_ < 50 Then Pixels_ = 100
    '    Else
            Unit_ = 8
            Pixels_ = 50
    '    End If
        IsUpdated_ = True
    End Sub    

    Public Function AddEvent(oType, oStart, oEnd,fYearOnly, strTitle, strDesc)
        Dim strStartDate, strEndDate, strEnd, strSep, strIcon, strImage, strImageX, oBegin, oFinish, strEvent, strLink, strColor, fIsCurrent, nDate
        strStartDate = StrFullDate(oStart)
        strEndDate = StrFullDate(oEnd)
        If strTitle = "" Or (strStartDate = "" And strEndDate = "") Then
            AddEvent = 0
            Exit Function
        End If
        strSep = " - "
        strEnd = ""
        strLink =""
        strIcon=""
        If IsObject(oEnd) Then
            Set oFinish = oEnd
        Else
            oFinish = oEnd
        End If
        fIsCurrent = False
        If IsObject(oType) Then
            strEvent = oType.Class
            Select Case strEvent
            Case "Individual"
                strEvent = strEvent & "_" & oType.Gender.ID
                If strDesc <> "" Then strLink = ", 'link' : '" & oType.Href & "'"
                fIsCurrent = Not oType.IsDead
            Case "Family"
                strEvent = strEvent & "_" & oType.Parents(0).Gender.ID & oType.Parents(1).Gender.ID
                If strDesc <> "" Then strLink = ", 'link' : '" & oType.Href & "'"
                If strEndDate = "" Then
                    fIsCurrent = oType.AreTogether
                    If Not fIsCurrent Then
                        If Not oType.Parents(0).IsDead = True Then
                            strEndDate = StrFullDate(oType.Parents(1).Death.Date)
                            Set oFinish = oType.Parents(1).Death.Date
                        ElseIf Not oType.Parents(1).IsDead = True Then
                            strEndDate = StrFullDate(oType.Parents(0).Death.Date)
                            Set oFinish = oType.Parents(0).Death.Date
                        ElseIf StrFullDate(oType.Parents(0).Death.Date) <> "" And StrFullDate(oType.Parents(1).Death.Date) <> "" Then
                            If DateValue(StrFullDate(oType.Parents(0).Death.Date)) < DateValue(StrFullDate(oType.Parents(1).Death.Date)) Then
                                strEndDate = StrFullDate(oType.Parents(0).Death.Date)
                                Set oFinish = oType.Parents(0).Death.Date
                            Else
                                strEndDate = StrFullDate(oType.Parents(1).Death.Date)
                                Set oFinish = oType.Parents(1).Death.Date
                            End If
                        End If
                    End If
                End If
            Case "Marriage"
                If oType.Session("Event") <> "" Then strEvent = "Event"
            Case "Occupation"
                fIsCurrent = oType.Termination.ID="StillWorking"
                If oType.Session("Event") <> "" Then strEvent = "Event"
            Case "Education"
                fIsCurrent = oType.Termination.ID="StillAttending"
                Case "SocialEntity"
                If strDesc <> "" Then strLink = ", 'link' : '" & oType.Href & "'"
                        strEvent="Occupancy"
            End Select
        Else
            strEvent = oType
        End If
        
        strImage = Dic.Peek("Timeline" & strEvent)
        If strImage <> "" Then
            If strEndDate = strStartDate Then        ' end only 'point' event so get special icon if available.
                strImageX =StrDicAttribute("Timeline" & strEvent, "P")
                If strImageX <> "" Then strImage=strImageX    ' check if P attribute present
                oBegin=""
                strEndDate = ""
            ElseIf strStartDate <> "" Then
                Set oBegin = oStart
            End If
            strIcon = Util.FormatString(", 'icon' : 'images/{}.gif'", strImage)
            strColor = Util.IfElse(Session("TimelineShowDuration"),", 'color' : '", ", 'textColor' : '") & Dic.PlurialCardinal("Timeline" & strEvent, 1) & "'"
        End If
        If strEndDate = "" And Session("TimelineContemporary") and fIsCurrent Then strEndDate = Session("Today")
        If strEndDate = "" Then
            strSep = ""
        Else
            strEnd = Util.FormatString(", '{}' : '{}', 'isDuration' : 'false'", Util.IfElse(Session("TimelineShowDuration"), "end", "finish"), StrEndDate)
        End If
        If strStartDate = "" Then
                nDate = DateAdd("d", -1, strEndDate)
            strStartDate = Day(nDate) & " " & MonthName(Month(nDate), True) & " " & Year(nDate)
            strEndDate = ""
            strSep = "< "
        End If
        If strDesc <> "" Then strDesc = Util.FormatPhrase(",'description' : '{0}'",Util.FormatString("{&j}", Util.StrFirstCharUCase(strDesc)))
        Report.WriteLn
        Report.WriteFormatted "{} 'start' : '{}'{}, 'title' : '{} \u202D{}{}{}\u202C'{}{}{}{}{},", "{", StrStartDate, strEnd, Replace(Replace(Replace(Replace(Replace(strTitle, "\","\\"),"'","\'"), """","\"""), "&#32;"," "), "&#39;","\'"), StrGnoDate(oBegin, fYearOnly), strSep, StrGnoDate(oFinish, fYearOnly), strIcon, strColor, strLink, strDesc, "}"

        Nodes_ = Nodes_ + 1
        IsUpdated_ = False
        CheckDate_ strStartDate
        CheckDate_ strEndDate
        AddEvent = 1
    End Function

    Function AddHeader(ID, Inline)
        ' turn ID into a valid variable name
        ID_ = Replace(Replace(Replace("tl_" & ID,".","$"),"-","_"),":","$")
        If Inline Then
            Report.WriteLn "var " & ID_ & " = new Object();"
            Report.WriteFormattedLn ID_ & ".unit = 8; " & ID_ & ".pixels=70; " & ID_ & ".duration = {}", Util.IfElse(Session("TimelineShowDuration"), "true", "false")
            Report.WriteFormattedLn ID_ & ".nowTag = ' {&j}';", Dic("TimelineNowTag")
            Report.WriteFormattedLn ID_ & ".wrapEvents = {};", Util.IfElse(Session("TimelineWrapEvents"), "true", "false")
            Report.Write ID_ & ".json0 = "
        End If
        Report.WriteLn "{"
        Report.WriteLn "'events' : ["
    End function

    Function AddTrailer(Inline, strTitle)
        Dim nDateNoteStart, nDateNoteEnd, nDate
        If Not Inline Then
            On Error Resume Next
            nDateNoteStart = DateAdd("d",1, DateMax_)
            If Err.Number > 0 Then nDateNoteStart = DateMax_
            nDateNoteEnd   = DateAdd("yyyy",50, DateMax_)
            If Err.Number > 0 Then nDateNoteEnd = DateMax_
            nDate   = DateAdd("yyyy",-50, DateMax_)
            If Err.Number > 0 Then nDate = DateMax_
            On Error Goto 0
            Report.WriteFormattedLn "{}'title' : '{&j}', 'start' : '{&j} {&j} {&j}', 'end' : '{&j} {&j} {&j}', 'textColor' : 'black', 'color' : 'white', 'isDuration' : 'false', 'icon' : 'images/info.gif', 'description' : '{&j}'{}],",_
                 "{", Dic("TimelineHelp"), _
                Day(nDateNoteStart), MonthName(Month(nDateNoteStart),True), Year(nDateNoteStart), _
                Day(nDateNoteEnd), MonthName(Month(nDateNoteEnd),True), Year(nDateNoteEnd), _
                Dic("TimelineHelpDetails"), "}"
            Report.WriteFormattedLn "'nowTag' : ' {&j}',", Dic("TimelineNowTag")
            Report.WriteFormattedLn "'subtitle': '{&j}',", strTitle
            Report.WriteFormattedLn "'date' : '{} {} {}'{}", Day(nDate), MonthName(Month(nDate),True), Year(nDate), "}"
        Else
            On Error Resume Next
            nDate = DateAdd("yyyy",-10, DateMin_)
            If Err.Number > 0 Then nDate = DateMin_
            On Error Goto 0
            Report.BufferLength = Report.BufferLength - 1 ' remove last comma
            Report.WriteLn "]};"
            Report.WriteFormattedLn ID_ & ".date = '{} {} {}';", Day(nDate), MonthName(Month(nDate),True), Year(nDate)
            Report.WriteLn "addEvent(window, 'load', function () {"
            Report.WriteLn "    " & ID_ & ".div = document.getElementById('" & ID_ & "');"
            Report.WriteLn "    timeLineOnLoad(" & ID_ & ");"
            Report.WriteLn "    addEvent(window, 'resize', function() {"
            Report.WriteLn "        timeLineOnResize(" & ID_ & ");"
            Report.WriteLn "    });"
            Report.WriteLn "});"
        End If
    End Function
End Class
