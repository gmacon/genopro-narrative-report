' Generic utility routines that can be used anywhere.
' The routines in this files are language independent.
'
' HISTORY
' Aug-2005        GenoPro            Creation
' Sep 2005 -      Ron            Development & Maintenance

'===========================================================

' following 2 functions are to simplify changes when GenoPro supports boolean Custom Tags
Function IsTrue(YorN, Default_)
    Select Case VarType(YorN)
    Case 11        ' boolean
        IsTrue = YorN
    Case Else    ' string etc.
        IsTrue = Util.IfElse(Util.IsNothing(YorN),Default_,Left(UCase(YorN), 1) = "Y")
    End Select
End Function

Function IsFalse(YorN, Default_)
    Select Case VarType(YorN)
    Case 11        ' boolean
        IsFalse = YorN
    Case Else    ' string etc.
        IsFalse = Util.IfElse(Util.IsNothing(YorN), Not Default_, Left(UCase(YorN), 1) = "N")
    End Select
End Function

Function ConfigMessage(strKey)
    ConfigMessage = ConfigMsg(strKey, "", "")
End Function

Function ConfigMsg(strKey, strDefault, strVersion)
    Dim oCfg, oNode, strPhrase
    Set oCfg = Session("ConfigMessages")
    If Util.IsNothing(oCfg) Then
        ConfigMsg = Dic(strKey)
    Else
        Set oNode = oCfg.selectSingleNode(strKey)
        If Util.IsNothing(oNode) Then
            If strDefault <> "" Then
                Set oNode = oCfg.selectSingleNode("WarningTagChange")
                If Util.IsNothing(oNode) Then
                    strPhrase= "Warning: tag {0} introduced at version {2} is missing from Config.xml so defaulted to {1}"
                Else strPhrase=StrNormalizeSpace(oNode.getAttribute("T"))
                End If
                Report.LogWarning Util.FormatPhrase(strPhrase, strKey, strDefault, strVersion)
            Else
                Set oNode = oCfg.selectSingleNode("ErrorConfigMessageNotFound")
                If Not Util.IsNothing(oNode) Then Report.LogError StrNormalizeSpace(oNode.getAttribute("T")) & ": " & strKey
            End If
            ConfigMsg = strDefault
        Else
            ConfigMsg = StrNormalizeSpace(oNode.getAttribute("T"))
        End If
    End If
End Function

' initialise patterns for GetCoord Function (see below)
    Dim g_RegExp1, g_RegExp2
    Set g_RegExp1 = New RegExp
    g_RegExp1.Pattern = "[WSws]"
    Set g_RegExp2 = New RegExp
    g_RegExp2.Pattern = "[^\d\.]"
    g_RegExp2.Global = True

    Function GetCoord(strCoord)
    ' normalize grid coordinate to a signed decimal e.g. for sorting
    dim coord, i,dec,negate,multiplier, strLocale
    strLocale = GetLocale()
    SetLocale("en-gb")
    If (strCoord <> "") Then
        dec=0: negate = 1: multiplier = 1
        coord = g_RegExp1.Replace(strCoord, "-") ' change W or S to -
        If InStr(coord,"-") > 0 Then negate = -1
        coord = g_RegExp2.Replace(coord," ") ' change any other non decimal to space
        coord=Split(coord," ")
        For i=0 to UBound(coord)
            If coord(i) <> "" Then
               If Isnumeric(coord(i)) Then
                   dec = dec + coord(i) * multiplier
                   multiplier = multiplier / 60
                  Else
                    Report.LogError ConfigMsg("ErrorLatLang", "Error: Unknown Latitude/Longitude format: " , "2011.01.07") & strCoord
                    Exit For
           End If
            End If
        Next
        GetCoord = dec * negate
    Else
        GetCoord = ""
    End If
    SetLocale(strLocale)
    End Function

Function GetFileTimestamp(strPath)
    ' get 'date last modified' of a picture and convet to timestamp
    ' N.B. requires objects oHttp and oFso to be initialised and available
    Dim strPictureDate, strLocale
    strLocale = GetLocale
    SetLocale("en-gb")    'level playing field for date conversion
    Err.Clear
    On Error Resume Next
    If LCase(Left(strPath,5)) = "http:" Then
        oHttp.Open "HEAD", strPath, False
        oHttp.Send
        On Error Goto 0
        If Err.Number = 0 Then strPictureDate = oHttp.getResponseHeader("Last-Modified")
        If Err.Number <> 0 Or strPictureDate = "" Then
            Report.LogError Util.FormatString(ConfigMessage("ErrorFetchHeader"), strPath) & _
                Util.IfElse(Err.Number > 0, "(Error " & (Err.Number And (256*256-1)) & " " & Err.Description & ")", "")
            GetFileTimeStamp=""
        Else    ' convert date & time to timestamp (100s of nanoseconds since 01/01/1601 !!!
            GetFileTimestamp = HexTimestamp(Mid(strPictureDate,6,20))
            If Err.Number <> 0 Then
                Report.LogError Util.FormatString(ConfigMessage("ErrorVB"), (Err.Number And (256*256-1)), Err.Description, strPath)
                GetFileTimeStamp=""
            End If
        End If
    Else
        GetFileTimestamp = HexTimestamp(oFso.GetFile(strPath).DateLastModified)
        If Err.Number <> 0 Then
            Report.LogError Util.FormatString(ConfigMessage("ErrorVB"), (Err.Number And (256*256-1)), Err.Description, strPath)
            GetFileTimeStamp=""
        End If
    End If
    SetLocale(strLocale)
End Function

    Sub GMapAddPlace(oPlaces, oParent, oEvent, oStart, oEnd, strType, strText)
        Dim cnt, idx, oEvents, oDate, strName, strEnd
        If oEvent.Place.Latitude <> "" And oEvent.Place.Longitude <> "" Then
             cnt=oPlaces.Count
             idx = oPlaces.AddUnique(oEvent.Place)
             If oPlaces.Count > cnt Then
             oEvent.Place.Session("events") = Util.NewStringDictionary()
          End If
          Set oEvents=oEvent.Place.Session("events")
             If Util.IsNothing(oEnd) Then
                 Set oDate=oStart
                 strEnd=""
          Else
              Set oDate=Util.IfElse(Util.IsNothing(oStart),oEnd, oStart)
              strEnd=oEnd.ToString
          End If
          strName=Util.IfElse(Util.IsNothing(oParent),"", oParent)
          oEvents.Add Util.FormatPhrase("{0} {1} {2}: {3}[ {4}][->{5}]", oDate.Year, strName, strType, strText, oStart.ToString, strEnd), GetDate(oDate)
         End If
    End Sub

Sub GMapCollateFamilyEvents(oPlaces, f, NameReqd)
    Dim e, oFam
      If NameReqd Then Set oFam = f Else Set oFam = Nothing End If
      For Each e In f.Unions
          strEvent = e.Session("Event")
        If strEvent <> "" Then
            If Dic.Peek("Ph" & strEvent) <> "" Then strEvent = StrDicAttribute("Ph" & strEvent, "N")
              GMapAddPlace oPlaces, oFam, e, e.Date, Nothing, strEvent, ""
        Else
              GMapAddPlace oPlaces, oFam, e, e.Date, Nothing, Util.IfElse(e.Type.ID = "", Dic("Marriage"), e.Type), ""
              GMapAddPlace oPlaces, oFam, e.Divorce, e.Divorce.Date, Nothing, StrDicExt("Divorce","","Divorce/Separation","","2010.12.10"), ""
        End If
    Next
End Sub

Sub GMapCollateIndividualEvents(oPlaces, i, NameReqd)
    Dim e, oInd
      If NameReqd Then Set oInd = i Else Set oInd = Nothing End If
      GMapAddPlace oPlaces, oInd, i.Birth, i.Birth.Date, Nothing, Dic("Birth"), ""
      GMapAddPlace oPlaces, oInd, i.Birth.Baptism, i.Birth.Baptism.Date, Nothing, "Baptism", ""
      GMapAddPlace oPlaces, oInd, i.Death, i.Death.Date, Nothing, Dic("Death"), ""
      GMapAddPlace oPlaces, oInd, i.Death.Funerals, i.Death.Funerals.Date, Nothing, Dic("Funeral"), ""
      GMapAddPlace oPlaces, oInd, i.Death.Disposition, i.Death.Disposition.Date, Nothing, StrDicExt("Disposition", "", "Body disposition", "", "2010.12.10"), ""
      For Each e In i.Contacts
          GMapAddPlace oPlaces, oInd, e, e.DateStart, e.DateEnd, Util.IfElse(e.Type.ID = "", Dic("Occupancy"), e.Type), ""
    Next
      For Each e In i.Educations
          GMapAddPlace oPlaces, oInd, e, e.DateStart, e.DateEnd, Dic("Education"), Util.FormatPhrase("{0}[ {1}][ {2}]", e.Institution, e.Level, e.Program)
    Next
      For Each e In i.Occupations
          strEvent = e.Session("Event")
        If strEvent <> "" Then
            If Dic.Peek("Ph" & strEvent) <> "" Then strEvent = StrDicAttribute("Ph" & strEvent, "N")
              GMapAddPlace oPlaces, oInd, e, e.DateStart, e.DateEnd, strEvent, Util.FormatPhrase("{0}[ {1}]", e.Session("Title"), e.Session("Agency"))
        Else
              GMapAddPlace oPlaces, oInd, e, e.DateStart, e.DateEnd, Dic("Occupation"), Util.FormatPhrase("{0}[ ({1})][ {2}]", e.Title, e.WorkType, e.Company)
        End If
    Next
End Sub

Sub GMapWriteEvents(oPlaces)
' Write Google Map Marker data to output buffer
    Dim oCoord, oCoords, oCnt, oEvents, e, p, strPlaces, strSep, strSep2
    ' amalgamate places at same coords for one pin on map
    Set oCoords = Util.NewObjectRepertory()
    For Each p in oPlaces
        oCoords.Add GetCoord(p.Latitude) &"/" & GetCoord(p.Longitude), p
    Next
    Report.WriteLn "var gMapData = {markers: ["
      strSep = ""
    For Each oCoord in oCoords
        cCnt = 0
          For Each p In oCoord
              cCnt = cCnt + p.Session("events").Count
          Next
        strSep2="": strPlaces=""
          Report.Write strSep & "{"
               Report.WriteFormatted """lat"": ""{&j}"", ""lng"": ""{&j}"", ""html"":""<div class='infoWindow{}'>", oCoord.Object(0).Latitude, oCoord.Object(0).Longitude, Util.IfElse(cCnt+oCoord.Count>10," infoScroll","")
          For Each p In oCoord
               Set oEvents=p.Session("events")
              oEvents.SortByValue() ' date order
                  Report.WriteFormattedBr "<b>{&j}</b>", JoinPlaceNames(p, p.Name, 0), Util.IfElse(oEvents.Count>10," infoScroll","")
            For e=0 To oEvents.Count-1
                Report.WriteBr oEvents.Key(e)
            Next
            strPlaces = strPlaces & strSep2 & JoinPlaceNames(p, p.Name, 0)
            strSep2 = "; "
            Set oEvents = Nothing
          Next
              Report.WriteFormattedLn "</div>"",""label"":""{&j} ({})""", strPlaces, Dic.PlurialCount("Reference",cCnt)
              Report.WriteLn "}"
        strSep=","
       Next
    Report.WriteLn "]}"
End Sub
Function HexTimestamp(strDate)
' Convert date & time to timestamp (hexadecimal string representing 'hundreds of nanoseconds' since 1st Jan 1601
    Dim n1,n2,n3, n3a, n3b
    n1 = CDbl(0)
    n1 = n1 + DateDiff("s", #1/1/1601#, DateValue(strDate)) + (Hour(strDate) * 60 + Minute(strDate)) * 60 + Second(strDate)' timestamp in seconds since 01/01/1601
    n2  = n1 / 2^32 * 10000000    ' MS 32 bits of "hundreds of nanoseconds"
    n3 = (n2 - Int(n2)) * 2^32    ' LS 32 bits
    n3a = Int(n3 / 2^16)
    n4a = (n3 - n3a) * 2^16
    HexTimestamp = Right("00000000" & Hex(n2),8) & Right("0000" & hex(n3a), 4) & Right("0000" & hex(n3b), 4)
End Function


Sub WriteHtmlButtonToggle(style)
    Select Case style
    Case "Notes"
        Report.WriteFormattedLn "<img src='images/{}.gif' class='toggle24' name='toggle' onclick='javascript:ToggleTree(this.name,"""");' alt='{1}' title='{1}'/>", Util.IfElse(Session("fCollapseNotes"),"expand","collapse"), StrDicOpt("TocExpandCollapseAll", Dic(style), "{0} {1}")
    Case "References", "OtherDetails"
        Report.WriteFormattedLn "<img src='images/{}2.gif' class='toggle24' name='toggle2' onclick='javascript:ToggleTree(this.name,""2"");' alt='{1}' title='{1}'/>", Util.IfElse(Session("fCollapseReferences"),"expand","collapse"), StrDicOpt("TocExpandCollapseAll", Dic(style), "{0} {1}")
    Case "Entries"
    Report.WriteFormattedLn "<div class='floatright buttontoc'><img name='tocStateButton' onclick='tocToggle();' class='control24' src='images/blank.gif' alt='' title='' style='margin-right:3px;'/><img class='control24' src='images/close.gif' alt='{}' title='{0}' onclick='javascript:tocHide();' /></div>", StrDicExt("AltCloseIndex", "", "Close this index", "", "2.0.1.7")
        Report.WriteFormattedLn "<div class='buttontoc'><img src='images/{}2.gif' class='toggle24' name='toggle2' onclick='javascript:ToggleTree(this.name,""2"");' alt='{1}' title='{1}'/></div>", Util.IfElse(fTreeOpen,"collapse","expand"), StrDicOpt("TocExpandCollapseAll", Dic(style), "{0} {1}")
    Case "HidePopUp"  ' close 'popup' frame
        Report.WriteFormattedLn "<img src='images/close.gif' class='togglepopup control24' onclick='javascript:hidePopUpFrame(event);' alt='{0}' title='{0}'/>", Dic("AltHidePopUpFrame")
    End Select
End Sub

' Write the HTML code to display all the pictures of an individual, family or place
Sub WriteHtmlPictures(i, cxWidthMax, cyHeightLimit, cxyPadding, strAlign, strHeading, fNoName, fBr)
    Dim p, pic, fExclude, pic1, pCnt
    pCnt = 0
    Set pic1 = i.Pictures.Primary    ' Get the primary picture of the individual or family
    If Not Util.IsNothing(pic1) Then If pic1.PictureDimension = "" Or pic1.Session("IsExcluded") Then Set pic1 = Nothing
'    find first non-excluded, count of non-excluded & tallest
    For Each pic in i.Pictures
        If Not pic.Session("IsExcluded") And pic.PictureDimension <> "" Then
            strDimension = pic.PictureDimension(cxWidthMax & "x" & cyHeightLimit)
            cyHeight = Util.GetHeight(strDimension)
            If cyHeight > cyHeightMax Then cyHeightMax = cyHeight
            pCnt = pCnt + 1
            If pic1 Is Nothing Then Set pic1 = pic
        End If
    Next
    If pCnt > 0 Then
        Dim strId, strDimension,  cyHeightMax, cyHeight, pic1Start, pic1End, picStart, picEnd, strpic
        i.Session("PictureMain") = pic1
        strId = i.ID
        Report.WriteFormattedLn "{}<div id='idPVp_{}' class='queue'>", Util.IfElse(Session("Book"), "", strHeading), strId
		pic1Start = Report.BufferLength
		Report.WriteLn "<div class='queue-entry'>"
        WriteHtmlPicture i, pic1, cxWidthMax, cyHeightMax, cxyPadding, strAlign, cxWidthMax, cyHeightLimit
        WriteHtmlNamePicture i, pic1
		If (pCnt > 1) And Not Session("Book") Then WriteHtmlButtonsPV i.ID
		If Session("fShowPictureDetails") Then
			WriteHtmlDetailsPicture i, pic1
			WriteHtmlAnnotationPicture i, pic1
        End If
		Report.WriteLn "</div>"
		pic1End = Report.BufferLength

        If (pCnt > 1) And Not Session("Book") Then
            For Each pic In i.Pictures
                If Not pic.Session("IsExcluded") And pic.ID <> pic1.ID Then
					Report.WriteLn "<div class='queue-entry hide'>"
					WriteHtmlPicture i, pic, cxWidthMax, cyHeightMax, cxyPadding, strAlign, cxWidthMax, cyHeightLimit
					WriteHtmlNamePicture i, pic
					WriteHtmlButtonsPV i.ID
					If Session("fShowPictureDetails") Then
						WriteHtmlDetailsPicture i, pic
						WriteHtmlAnnotationPicture i, pic
					End If
					Report.WriteLn "</div>"
                End If
            Next
        End If
		Report.WriteLn "</div>"
    End If
End Sub

Sub WriteHtmlPicturesLarge(i, strAlign, strHeading, fNoName, fBr)
    WriteHtmlPictures i, Session("cxPictureSizeLarge"), Session("cyPictureSizeLarge"), Session("cxyPicturePadding"), strAlign, strHeading, fNoName, fBr
End Sub

Sub WriteHtmlPicturesSmall(i, strAlign, fNoName)
    Report.WriteLn "<table class='photo floatleft aligncenter widthpaddedsmall'><tr><td>"
    WriteHtmlPictures i, Session("cxPictureSizeSmall"), Session("cyPictureSizeSmall"), Session("cxyPicturePadding"), strAlign, "", fNoName, ""
    Report.WriteLn "</td></tr></table>"
End Sub

Sub WriteHtmlPictureSmall(i, p, strAlign)
    Report.WriteLn "<table class='photo floatleft aligncenter widthpaddedsmall'><tr><td>"
	Report.WriteLn "<div class='queue'>"
    WriteHtmlPicture i, p, Session("cxPictureSizeSmall"), Session("cyPictureSizeSmall"), Session("cxyPicturePadding"), strAlign, Session("cxPictureSizeSmall"), Session("cyPictureSizeSmall")
    WriteHtmlNamePicture i, p
	Report.WriteLn "</div>"
    Report.WriteLn "</td></tr></table>"
End Sub

Sub WriteHtmlPicture(obj, p, cxWidthMax, cyHeightMax, cxyPadding, strAlign, cxWidthLimit, cyHeightLimit)
    Dim strDimension, cxWidth, cyHeight, cxMargin, cyMargin, cxMarginL, cxMarginR, strAttributesHTML, strFloat, strThumbnail, i, strAreaMap, strExt, strPath

    If Not p.Session("IsExcluded") Then
        strPath = p.Path.Report
        If (Session("BasePath") <> "" And Instr(strPath, Session("BasePath")) = 8) Then strPath = "..\.." & Mid(strPath, 8+Len(Session("BasePath")))
        strThumbnail = strPath
        If Session("fUsePictureThumbnails") Then
            strThumbnail = p.Session("Thumbnail")
        End If
        strDimension = p.PictureDimension(cxWidthMax & "x" & cyHeightMax)
        cxWidth = Util.GetWidth(strDimension)
        cyHeight = Util.GetHeight(strDimension)

        cxMargin = (cxWidthMax - cxWidth) \ 2
        If (cxMargin < 0) Then cxMargin = 0
        cyMargin = (cyHeightMax - cyHeight) \ 2
        If (cyMargin < 0) Then cyMargin = 0
        strFloat=""
        cxMarginL = cxMargin
        If (strAlign = "right") Then cxMarginL = cxMarginL + cxyPadding
        cxMarginR = cxMargin
        If (strAlign = "left") Then
            cxMarginR = cxMarginR + cxyPadding
        End If
        strAttributesHTML = Util.FormatString(" width='{}px' height='{}px' class='pic{}' style='margin:{}px {}px {}px {}px;'", cxWidth, cyHeight, strAlign, cyMargin, cxMarginR, cyMargin, cxMarginL)
        strAreaMap = Util.FirstNonEmpty(CustomTag(p, "_AreaMap"), CustomTag(p, "AreaMap"))
        If strAreaMap = "" Then
            Report.WriteFormattedLn "<a class='gallery' rel='{5}' href='{2}' title='{4&t}'><img src='{0}'{1} alt='{3&t}' title='{3&t}'/></a>", strThumbnail, strAttributesHTML, strPath, Util.FormatPhrase(Dic("PhPictureTooltip"), "", p.Date.ToStringNarrative, p.Place.Session("Locative"), StrParseText(p.Source, True)), p, obj.id
        Else
            Report.WriteFormattedLn "<a href='picture-{}.htm' onclick='showPopUpFrame();' target='popup' title='{&t}'><img src='{}'{} title='{1&t}' alt='{&t}'/></a>", p.ID, Dic("PictureDetailsTip"), strThumbnail, strAttributesHTML, Util.FormatPhrase(Dic("PhPictureTooltip"), "", p.Date.ToStringNarrative, p.Place.Session("Locative"), StrParseText(p.Source, True))
        End If
    End If
End Sub

Sub WriteHtmlGenericPicture(i, PorC, size, cxWidthMax, cyHeightMax, cxyPadding, strAlign)
    Dim cxWidth, cyHeight, cxMargin, cyMargin, cxMarginL, cxMarginR, strAttributesHTML, strGender
    If size = "l" Then
        cxWidth = 138
        cyHeight = 175
    Else
        cxWidth = 59
        cyHeight = 75
    End If

    cxMargin = (cxWidthMax - cxWidth) \ 2
    If (cxMargin < 0) Then cxMargin = 0
    cyMargin = (cyHeightMax - cyHeight) \ 2
    If (cyMargin < 0) Then cyMargin = 0
    cxMarginL = cxMargin
    If (strAlign = "right") Then cxMarginL = cxMarginL + cxyPadding
    cxMarginR = cxMargin
    If (strAlign = "left") Then
        cxMarginR = cxMarginR + cxyPadding
    End If
    strAttributesHTML = Util.FormatString(" width='{}px' height='{}px' class='pic{}' style='margin:{}px {}px {}px {}px;'", cxWidth, cyHeight, strAlign, cyMargin, cxMarginR, cyMargin, cxMarginL)
    strGender = i.Gender.ID
    If strGender="M" Or strGender="F" Then
        Report.WriteFormattedLn "<img src='images/profile{}{}.jpg'{} alt='{&t}' title='{3&t}'/>", PorC, strGender, strAttributesHTML, Dic("AltGenericPicture")
    End If
End Sub

Sub WriteHtmlNamePicture(o, p)
    Dim strId
    strId = o.ID
    If Not (Session("fHidePictureName") Or Session("fShowPictureDetails") Or p.Session("IsExcluded")) Then
        Report.WriteFormattedLn "<div class='caption' id='idPVn_{}'>{}</div>", strId, StrHtmlHyperlink(p)
    End If
End Sub

Function StrFullDate(xDate)

' Make a complete date (dd mmm yyyy) from a either a GenoDate object or a (partial) date string by adding month (Jan) and/or day (01)

Dim strMonth, oDate, nMonth, nDay, strDate, nDate, strEra

If IsObject(xDate) Then
'    On Error Resume Next
'    If IsObject(xDate.Date) Then
'     Set oDate = CustomDate(xDate.Date)    ' check if parent of Date object
'  ElseIf IsObject(xDate) Then
      Set oDate = CustomDate(xDate)
''  Else
''      Set oDate = xDate
''  End If
        Err.Clear
    strDate = Replace(oDate.ToString(""),"th,","")
    On Error Goto 0
    strDate = Replace(strDate,"st,","")
    strDate = Replace(strDate,"nd,","")
    strDate = Replace(strDate,"rd,","")

    If IsDate(strDate) Then
        nDate = CDate(strDate)
        StrFullDate = Day(nDate) & " " & MonthName(Month(nDate), True) & " " & Year(nDate)
        Exit Function
    End If
    On Error Resume Next
    strEra = " " & oDate.Era
    If Err.Number > 0 Then strEra = ""
    On Error Goto 0
    If oDate.Year <> "" And strEra = "" Then
        nMonth = oDate.NMonth
        If IsNumeric(nMonth) Then
            If nMonth > 0 And nMonth < 13 Then
                strMonth = MonthName(nMonth, True)
            Else
                strMonth = MonthName(1, True)
            End If
        Else
            strMonth = MonthName(1, True)
        End If
        nDay = oDate.NDay
        If IsNumeric(nDay) Then
            strDate = Util.IfElse(nDay > 0 And nDay < 32, nDay, "01") & " " & strMonth & " " & oDate.Year
            StrFullDate = Util.IfElse(IsDate(strDate), strDate, "")
        Else
            strDate = "01 " & strMonth & " " & oDate.Year
            StrFullDate = Util.IfElse(IsDate(strDate), strDate, "")
        End If
    ElseIf oDate.Year <> "" Then
           StrFullDate = oDate.Year
    Else
        StrFullDate = ""
    End If
Else
    StrFullDate = ""
End If
End Function

Function StrGnoDate(xDate, fYearOnly)
Dim oDate

If IsObject(xDate) Then
    Set oDate = xDate
    On Error Resume Next
    If IsObject(xDate.Date) Then Set oDate = xDate.Date    ' check if parent of Date object
        Err.Clear
    On Error Goto 0
    If fYearOnly Then
        StrGnoDate = oDate.ToString("[|~|<|>]yyyy")
    Else
        StrGnoDate = oDate.Approximation & Trim(oDate.ToString("dd MMM yyyy"))
    End If
Else
    StrGnoDate = ""
End If

End Function

' Output Html Content-Language e.g. "en-us" "fr" etc.
Sub WriteHtmlLang()
    Report.Write Session("HtmlLang")
End Sub

Sub WriteHtmlExtraNarrative(obj)
    Dim strNarrative
    strNarrative = Util.FirstNonEmpty(CustomTag(obj, "_Narrative"),CustomTag(obj, "Narrative"))
    If strNarrative <> "" Then
        Report.WriteFormattedLn "<div>{}</div>", StrFormatText(obj, StrParseText(strNarrative, True))
    End If
End Sub

' Same as WriteHtmlPicture() however returns a string containing the HTML code from what should be written to the output buffer
Function StrHtmlPicture(p, cxWidthMax, cyHeightMax, cxyPadding, strAlign, cxWidthLimit, cyHeightLimit)
    Dim cchBuffer
    cchBuffer = Report.BufferLength
    WriteHtmlPicture p, cxWidthMax, cyHeightMax, cxyPadding, strAlign, cxWidthLimit, cyHeightLimit
  Report.BufferLength = Report.BufferLength -2
    StrHtmlPicture = Report.Buffer(cchBuffer)    ' Get the text from the output buffer
    Report.BufferLength = cchBuffer    ' Truncate the buffer to its original length
End Function


' Write the HTML code to generate a button (Play, Pause, Next, Prev) of the Picture Viewer (PV).
Sub WriteHtmlButtonPV(strId, strVerb)
    Report.WriteFormattedLn "<a href=""#"" onclick=""PV_{1}(this, '{0}');""><img src='images/PV_{1}.gif' border='0' alt='{2&t}' title='{2&t}'/></a>", strId, strVerb, Dic("PV_" & strVerb)
End Sub
' Generate the buttons to create a slide show of pictures
Sub WriteHtmlButtonsPV(strId)
			Report.Write "<div>"
            WriteHtmlButtonPV strId, "Play"
            WriteHtmlButtonPV strId, "Pause"
            WriteHtmlButtonPV strId, "Prev"
            WriteHtmlButtonPV strId, "Next"
            Report.WriteFormattedLn "<img id='idPVslider_{2}' src='images/slider.gif' style='position:relative;' width='30' height='16' alt='{0&t}' title='{0&t}'/>" & _
                    "<img onmousemove='sliderMouseMove(event);' onmousedown='sliderMouseDown(event);' onmouseout='sliderMouseUpOut(event);' onmouseup='sliderMouseUpOut(event);' " & _
                        "style='cursor:pointer;position:relative;left:{1}px;' class='knob' src='images/knob.gif' width='9' height='15' alt='{0&t}' title='{0&t}'/>", Dic("PV_Speed"), Session("PictureInterval"), strId
			Report.Write "</div>"
End Sub 

' Return a string containing the HTML code to display the gender of the individual.
Function StrHtmlImgGender(i)
    StrHtmlImgGender = Util.FormatString("<img src='images/gender_{}.gif' class='icon' alt='{&t}' title='{1&t}' />&nbsp;", i.Gender.Id, i.Gender)
End Function

Function StrHtmlImgFamily(f)
    Dim strGender0, strGenders
    strGenders = "MF"
    strGender0 = f.Parents(0).Gender.ID
    If (strGender0 = f.Parents(1).Gender.ID) Then
        ' Same-sex couples
        Select Case strGender0
            Case "M"
                strGenders = "MM"
            case "F"
                strGenders = "FF"
        End Select
    End If
    ' strGenders = ""    ' Enable this line if you want a generic icon for a family
    StrHtmlImgFamily = Util.FormatString("<img src='images/family_{}.gif' class='icon' alt='{}' title='{1}'/> ", strGenders, Dic("Family"))
End Function

Function StrHtmlImgPhoto(obj)
    Select Case obj.Class
    Case "Picture"
        If Not obj.Session("IsExcluded") Then
            StrHtmlImgPhoto = Util.FormatString("<img src='images/space.gif' width='16' alt=''/>&nbsp;<a href='picture-{}.htm' onclick='tocExit();' title='{&t}'><img src='images/photo.gif' class='icon' alt='' title='' />&nbsp;{}</a>", _
                obj.ID, Dic("PictureDetailsTip"), Util.IfElse(Trim(obj.Name) <> "", StrFormatText(obj, StrParseText(Trim(obj.Name), True)), "(" & obj.ID &")"))
        End If
    Case Else
        cPictures = obj.Session("PicturesIncluded")
        If (cPictures > 0) Then
            StrHtmlImgPhoto = Util.FormatString("&nbsp;<a href='{}'><img src='images/photo.gif' class='icon' alt='{}' title='{1}' /></a>", _
                obj.Href, Dic.PlurialCount("Picture", cPictures))
        End If
    End Select
End Function

Function StrHtmlImgDescendantTreeChart(i)
    If IsTrue(CustomTag(i, "DescendantTreeChart"), False) Then
        StrHtmlImgDescendantTreeChart = Util.FormatString("&nbsp;<a href='descendants/DescendantTree.htm?tree={}' target='popup'><img src='images/descendants.gif' class='icon' alt='{}' title='{1}' /></a>", _
            i.ID & ".xml", Dic("AltDescendantTreeChart"))
    End If
End Function

Function StrHtmlImgPhotoLink(i, target)
    cPictures = i.Session("PicturesIncluded")
    If (cPictures > 0) Then
        StrHtmlImgPhotoLink = Util.FormatString("&nbsp;<a href='{}#{}'><img src='images/photo.gif' class='icon' alt='{}' title='{2}' /></a>", _
            target, i.ID, Dic.PlurialCount("Picture", cPictures))
    End If
End Function

' Link to the .gno file if present in the report
Function StrHtmlImgFileGno(obj)
    Dim strName
    strFileFamilyTreeGno = ReportGenerator.ExtraFiles("FamilyTree.gno")
    If (strFileFamilyTreeGno <> "") Then
        If obj.Class="Individual" Then
            strName = obj.Session("NameFull")
        Else
            strName = obj.Name
        End If
        StrHtmlImgFileGno = Util.FormatString(" <a href='{}?id={}'><img src='images/gno.gif' class='icon' alt='{&t}' title='{2&t}'/> </a> ", strFileFamilyTreeGno, obj.Id, Dic.FormatString("FmtAltViewInGnoFile", strName))
    End If
End Function

Function StrHtmlImgMap(o)
  If o.Session("gMap") = True Then ' i.e. there are some geo-tagged PLaces for this individual/family
     Dim strAlt
     strAlt=StrDicExt("AltgMap" & o.Class, "", "Display a Google Map showing geo-tagged places associated with this " & Dic(o.Class),"", "2010.12.10")
     StrHtmlImgMap = Util.FormatString(" <a href='{}_map-{}.htm' target='popup'><img src='images/pin.gif' class='icon' alt='{&t}' title='{2&t}'/> </a> ", LCase(o.Class), o.ID, strAlt) 'ShowPopup removed from google map initiation code, to allow map of all people to appear in _detail
  End If
End Function

Function StrHtmlImgFileSvg(obj)
    Dim fFlag, strFileGenoMap, xyTopRight, xyTopLeft, xyStartRight, xyStartLeft, xTopCenter, pos
    fFlag = ReportGenerator.NegateAxisY
    ReportGenerator.NegateAxisY = True
    strFileGenoMap = obj.Position.GenoMap.Session("PathGenoMap")
    If ((Session("Svg")) And strFileGenoMap <> "") Then
        Select Case obj.Class
        Case "Individual"
            StrHtmlImgFileSvg = Util.FormatString(" <a href='{}?x={},y={},highlight=true,toggle={},name={}' target='popup'><img src='images/svg.gif' class='icon' alt='{&t}' title='{4&t}'/> </a> ", strFileGenoMap, obj.Position.x, obj.Position.y, Util.IfElse(Session("SvgDefault"), "SVG", "PDF"), obj.ID, Dic.FormatString("FmtAltViewInSvgFile", obj.Name))
        Case "Family"
            Set pos = obj.Position
            xyTopRight=split(pos.Top.Right,",")
            xyTopLeft=split(pos.Top.Left,",")
            xyStartRight=split(pos.Bottom.Right,",")
            xTopCenter=xyTopLeft(0)+ ( xyTopRight(0) - xyTopLeft(0) ) / 2
            If uBound(xyStartRight) > 0 Then    ' bottom line present, so centre highlight on top of vertical line
                StrHtmlImgFileSvg = Util.FormatString(" <a href='{}?x={},y={},highlight=true, toggle={}, name={}' target='popup'><img src='images/svg.gif' class='icon' alt='{&t}' title='{4&t}' /> </a> ", strFileGenoMap, obj.Position.x, obj.Position.y, Util.IfElse(Session("SvgDefault"), "SVG", "PDF"), obj.ID, Dic.FormatString("FmtAltViewInSvgFile", obj.Name))
            Else                    ' centre highlight on mid-point of family relationship
                StrHtmlImgFileSvg = Util.FormatString(" <a href='{}?x={},y={},highlight=true,toggle={},name={}' target='popup'><img src='images/svg.gif' class='icon' alt='{&t}' title='{4&t}' /> </a> ", strFileGenoMap, xTopCenter, xyTopRight(1), Util.IfElse(Session("SvgDefault"), "SVG", "PDF"), obj.ID, Dic.FormatString("FmtAltViewInSvgFile", obj.Name))
            End If
        End Select
    End If
    ReportGenerator.NegateAxisY = fFlag
End Function

Function StrHtmlImgTimeline(obj)
    Dim strDate, strData, strLocale, oDate, strAlt, collUnions, u
    strData = obj.Position.GenoMap.Session("TLData")
    If Session("Timelines") And strData <> "" Then
        strLocale=GetLocale
        SetLocale("en-gb")
        Select Case obj.Class
        Case "Individual"
            strDate = Util.FirstNonEmpty(StrFullDate(obj.Birth.Date), StrFullDate(obj.Birth.Baptism.Date), StrFullDate(obj.death.Date), StrFullDate(obj.Death.Disposition.Date))
            strAlt=Dic("AltTimelineImageInd")
        Case "Family"
            Set collUnions = obj.Unions.ToGenoCollection
            For Each u in collUnions
                strDate = Util.FirstNonEmpty(StrFullDate(u.Date), StrFullDate(u.Divorce.Date))
                If IsDate(strDate) Then Exit For
            Next
            strAlt=Dic("AltTimelineImageFam")
        End Select
        On Error Resume Next
        oDate=DateAdd("d",-1,strDate)
        If Err.Number = 0 Then StrHtmlImgTimeline = Util.FormatString(" <a href='timeline{}.htm?{},date={} {} {}'><img src='images/timeline.gif' class='icon' alt='{&t}' title='{4&t}'/> </a> ", _
                          obj.Position.GenoMap.Index, strData, Day(oDate), MonthName(Month(oDate),True), Year(oDate), strAlt)
        On Error Goto 0
        SetLocale(strLocale)
    End If
End Function

' Return individuals name if not blank otherwise if not null generate substitute text for unknown individual
Function StrNameNN(i)
Dim strName
strName = i.Session("NameFull")
If strName <> "" Then
    StrNameNN = strName
ElseIf not Util.IsNothing(i) Then
    StrNameNN = StrDicMFU("_NoName", i.Gender.ID)
Else
    StrNameNN = ""
End If
End Function

' Like ToHtmlHyperlinkNN method except returns blank if Individual object is empty or nothing
Function StrHtmlHyperlinkNN(obj)
    Select Case obj.class
    Case "Individual"
        If obj.Name <> "" Then
            StrHtmlHyperlinkNN = Util.FormatString("<a href='{0}' onclick='javascript:hidePopUpFrame("""");' target='detail'>{1}</a>", obj.Href, StrHtmlHighlightName(obj.Session("NameFull")))
        Else
            StrHtmlHyperlinkNN = Util.HtmlEncode(StrNameNN(obj))
        End If
    End Select
End Function

Function StrHtmlHyperlink(obj)
    Dim strName, arrLocative
    Select Case obj.Class
    Case "Individual"
        StrHtmlHyperlink=Util.FormatString("<a href='{0}' onclick='javascript:hidePopUpFrame("""");' target='detail'>{1}</a>", obj.Href, StrHtmlHighlightName(obj.Session("NameFull")))
    Case "Family"
        StrHtmlHyperlink=Util.FormatString("<a href='{0}' onclick='javascript:hidePopUpFrame("""");' target='detail'>{1}</a>", obj.Href, StrHtmlHighlightName(obj.Session("Name")))
    Case "Place"
        arrLocative = Split(Replace(obj.Session("LocativeRaw"),"]","["),"[") ' break into preposition (if any), name & postpostion(if any)
        StrHtmlHyperlink=Util.FormatString("{0}<a href='place-{1}.htm?popup' onclick='showPopUpFrame();' target='popup'>{2&t}</a>{3}", arrLocative(0), obj.ID, arrLocative(1), arrLocative(2))
    Case "Picture"
        StrHtmlHyperlink=Util.FormatString("<a href='picture-{1}.htm?popup' onclick='openPopUpFrame(""picture-{1}.htm"");' target='popup' title='{2&t}'>{0}</a>", Util.IfElse(Session("fUsePictureId"), obj.ID, Util.FirstNonEmpty(StrFormatText(obj, Trim(StrParseText(obj.Name, True))), Dic("PictureDetails"))), obj.ID, Dic("PictureDetailsTip"))
    Case "SourceCitation"
        StrHtmlHyperlink=Util.FormatString(" <a href='source-{1}.htm?popup' onclick='showPopUpFrame();' target='popup'>{0&t}</a>", JoinSourceCitationNames(obj, obj.title, true), obj.ID)
    Case "SocialEntity"
        StrHtmlHyperlink=Util.FormatString(" <a href='entity-{1}.htm?popup' onclick='showPopUpFrame();' target='popup'>{0&t}</a>", obj.Session("Name"), obj.ID)
    Case Else
            StrHtmlHyperlink = ""
    End Select
End Function

Function StrHtmlHyperlinkTag(obj, strTag)
    Select Case obj.Class
    Case "Individual"
        StrHtmlHyperlinkTag = Replace(obj.ToHtmlHyperlink,">" & obj.Name & "<"," target='detail' onclick='javascript:hidePopUpFrame();' >" & Util.HtmlEncode(obj.TagValue(strTag)) & "<",1,1)
    End Select
End Function


Function StrHtmlHyperlinkPlace(obj)
' Like StrHtmlHyperlink except no preposition included
    Dim arrLocative
    Select Case obj.Class
    Case "Place"
        ' arrLocative = Split(Replace(obj.Session("LocativeRaw"),"]","["),"[") ' break into preposition (if any), name & postpostion(if any)
        'StrHtmlHyperlinkPlace=Util.FormatString(" <a href='place-{1}.htm' onclick='showPopUpFrame();' target='popup'>{0&t}</a>", arrLocative(1), obj.ID)
		StrHtmlHyperlinkPlace=Util.FormatString(" <a href='place-{1}.htm' onclick='showPopUpFrame();' target='popup' alt='{2&t}' title='{2&t}'>{0&t}</a>", obj.Session("NameShort"), obj.ID, obj.Session("Address"))
    Case Else
            StrHtmlHyperlinkPlace = ""
    End Select
End Function

Function StrHtmlTag(strValue, strTag)
    If strValue <> "" Then
        StrHtmlTag = "<" & strTag & ">" & strValue & "</" & strTag & ">"
    End If
End Function

Function StrViaBuffer(strData)
    Dim cch
    cch = Report.BufferLength
    Report.WriteText strData
    StrViaBuffer = Report.Buffer(cch)
    Report.BufferLength = cch
End Function

' Generate the code to display text at the bottom of the status bar when the mouse is over a link or an image
Function StrMouseOver(s)
    StrMouseOver = Util.FormatString("onmouseover=""return ss('{&j}')"" onmouseout=""cs()"" title='{0&t}'", s)
End Function

Function WriteOccupancyEvents (oTLInfo, obj, strTitle, fYearOnly)
    Dim collContacts, o, cchBegin, cchStart, strYear, oStart, oEnd, nEvent, strGender, strRelative, strPnP, fExtant, nPlural
    Set collContacts = obj.Contacts.ToGenoCollection
    For Each o in collContacts
        oStart = ""
        oEnd = ""
        cchStart = Report.BufferLength
        If (o.DateStart <> "") Then Set oStart = o.DateStart
        If (o.DateEnd <> "") Then Set oEnd = o.DateEnd
        strGender = Util.FirstNonEmpty(CustomTag(obj,"Gender.ID"), "N")
        nPlural = Util.IfElse(IsFalse(CustomTag(obj,"Plural"), False),1,2)
        strPnP = Dic.Plurial(Util.IfElse(Dic.Peek("PnP_" & strGender)<>"", "PnP_" & strGender,"PnP_"), nPlural)
        strRelative = Dic.Plurial(Util.IfElse(Dic.Peek("PnR_" & strGender)<>"", "PnR_" & strGender,"PnR_"), nPlural)
        fExtant = (CustomTag(obj, "Extant") <> "N")
        oTLInfo.AddEvent obj, oStart, oEnd, fYearOnly, strTitle & " - " & Util.FormatString(StrDicOrTag("Timeline_" & o.Type.ID, CustomTag(o, "NarrativeStyle")), o.Place.Session("NameSort")), _
            StrViaBuffer(Util.FormatPhrase(Util.FirstNonEmpty(StrDicOrTag("PhOT_" & o.Type.ID, CustomTag(o, "NarrativeStyle")), _
                               Dic.Lookup2("PhOT_" & o.Type.ID & "_" & obj.Class, "PhOT_" & o.Type.ID)), _
                               strTitle, strRelative, StrDateSpan(o.DateStart, o.DateEnd), _
                               Util.IfElse(o.DateStart <> o.DateEnd, StrTimeSpan(o.Duration), ""), _
                               (Not fExtant Or (o.DateEnd.ToStringNarrative<>""))=False, _
                               o.Place.Session("HlinkLocative"), o.Summary, o.Place.Session("Hlink")))
    Next
End Function
Function WriteFamilyEvents (oTLInfo, f, strSuffix, fYearOnly, fFull)
    Dim collUnions, c, u, d, cchBegin, cchStart, strEvent, strYear, collChildren, nEvent, arrOfficiatorTitle, strTitle
    Dim oLink, oRepertoryNonBio, oStart, oEnd, oCnt
    Set oRepertoryNonBio = Session("oRepertoryNonBio")
    Set collUnions = f.Unions.ToGenoCollection
    For Each u in collUnions
        oStart = ""
        oEnd = ""
        Set d = u.Divorce
        cchStart = Report.BufferLength
        If (u.Date <> "") Then Set oStart = u.Date
        If (d.Date <> "") Then Set oEnd = d.Date
        strEvent = u.Session("Event")
        If strEvent <> "" Then 
            oCnt = oCnt + 1
            If Dic.Peek("Ph" & strEvent) <> "" Then strEvent = StrDicAttribute("Ph" & strEvent, "N")
            If fFull Then oTLInfo.AddEvent u, oStart, oEnd, False, StrPlainText(u, Util.FirstNonEmpty(strEvent, u.Session("Title"), u.Session("Agency"), "", "")), u.Session("Title") & " - " & u.Session("Agency")
        Else
             strTitle = Util.FirstNonEmpty(StrDicOrTag("Union", CustomTag(u, "NarrativeStyle")), Dic("Marriage")) & " " & strSuffix
            If IsObject(oStart) Then
                arrOfficiatorTitle = Split(u.Officiator.Title & "|", "|")
                If arrOfficiatorTitle(1) = "" Then arrOfficiatorTitle(1) = arrOfficiatorTitle(0)
                If IsObject(oEnd) Then
                    oTLInfo.AddEvent f, oStart, oEnd, fYearOnly, strTitle, _
                            Util.FormatPhrase(StrDicOrTag("PhTL_Union", CustomTag(u,"NarrativeStyle")), Dic.Plurial("PnP_", 2), u.Type, u.Date.ToStringNarrative, u.Place.Session("Locative"), arrOfficiatorTitle(0), u.Officiator, u.Witnesses, StrVerb("ToBe", f.AreTogether, False, "2", ""), arrOfficiatorTitle(1)) & " " & _
                            Util.FormatPhrase(StrDicOrTag("PhTL_Divorce", CustomTag(u,"NarrativeStyle")), u.IsAnnulled, d.Date.ToStringNarrative, d.Place.Session("Locative"), d.RequestedBy.ID, _
                              d.RequestedBy, d.Attorney.Husband, d.Attorney.Wife, d.Officiator)
                Else
                    oTLInfo.AddEvent f, oStart, "", fYearOnly, strTitle, _
                            Util.FormatPhrase(StrDicOrTag("PhTL_Union", CustomTag(u,"NarrativeStyle")), Dic.Plurial("PnP_", 2), u.Type, u.Date.ToStringNarrative, u.Place.Session("Locative"), arrOfficiatorTitle(0), u.Officiator, u.Witnesses, StrVerb("ToBe", f.AreTogether, False, "2", ""), arrOfficiatorTitle(1))
                End If
            End If
            If IsObject(oEnd) And Not IsObject(oStart) Then
                    oTLInfo.AddEvent f, oEnd, oEnd, fYearOnly, strTitle & " " & Dic("DivorceAbbr"), _
                            Util.FormatPhrase(StrDicOrTag("PhTL_Divorce", CustomTag(u,"NarrativeStyle")), u.IsAnnulled, d.Date.ToStringNarrative, d.Place.Session("Locative"), d.RequestedBy.ID, _
                              d.RequestedBy, d.Attorney.Husband, d.Attorney.Wife, d.Officiator)
            End If
        End If
    Next
    If fFull Then
        Set collChildren = f.Children.ToGenoCollection
        For Each c in collChildren
            If Not c.IsAdopted = "Y" Then
                WriteIndividualEvents oTLInfo, c, Dic.LookupEx("Child_", c.Gender.ID) & " " & Replace(c.Session("NameFull"), Session("MarkerFirstName"), ""), False
            Else
                If oRepertoryNonBio.KeyCounter("I" & c.ID & "F" & f.ID) > 0 Then
                    Set oLink = oRepertoryNonBio.Entry("I" & c.ID & "F" & f.ID).Object(0)
                    WriteIndividualEvents oTLInfo, c, Dic.FormatString("TimelineAdopted", StrDicMFU("Pedigree" & oLink.PedigreeLink.ID,c.Gender.ID), Dic.LookupEx("Child_", c.Gender.ID), Replace(c.Session("NameFull"), Session("MarkerFirstName"), "")), False
                    If oLink.Adoption.Date <> "" Then
                        oTLInfo.AddEvent "", oLink.Adoption.Date, , fYearOnly, Dic("Pedigree" & oLink.PedigreeLink.ID & "Event"),""
                    End If
                End If
            End If
        Next
    End If
End Function

Sub WriteHtmlAdditionalInformation(obj)
' Write details of this object's Custom Tags to the report

    Dim oCustomTagRepertory, oCustomTagDictionary, Layout, Layouts, i, j, cchBufferStart, cchBufferNow
    Dim strCustomTagData, strCustomTagDesc, strPrivate, strLink, strTag, strFmtTemplate, Args(), strSubHead, strGender
    strPrivate = StrDicOrTag("", "Private")
    Session("BufferBegin") = Report.BufferLength
    Set oCustomTagRepertory = Session("oCustomTagRepertory")

    Report.WriteLn    "<br /><div class='clearleft no-break'><ul class='xT'>"
    Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
    Report.WriteFormattedLn "<a name='Additional_Information'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Util.StrFirstCharUCase(Util.FormatPhrase(Dic("FmtAdditionalInformation"), LCase(Dic(obj.Class))))

    If obj.Class = "Individual" Then ' List any External Hyperlink or Blood Type here
        cchBufferNow = Report.BufferLength
        If obj.Hyperlink <> "" Then
			strLink = Mid(Lcase(obj.Hyperlink.Target),1,5)
			If  strLink = "http:" or  strLink = "file:" or  mid(Lcase(obj.Hyperlink.Target),1,7) = "mailto:" Then
				strLink = obj.Hyperlink.Target
			Else
				i = InStrRev(obj.Hyperlink.Target, "\")
				strLink = "media/" & Mid(obj.Hyperlink.Target, i+1)
				ReportGenerator.FileUpload obj.Hyperlink.Target, strLink
			End If
			Report.Write Util.FormatPhrase(Dic("PhExternalLink"), strLink)
		End If
        Report.WritePhraseDic("PhBloodType"), obj.Name.short, obj.Birth.BloodType, obj.IsDead=false
        If Report.BufferLength > cchBufferNow Then ' some phrase written
            Session("BufferBegin")= -1      ' indicate data present in Additional Information section
        End If
    End If
    If Session("Flag_T") Then
        If oCustomTagRepertory.KeyCounter(obj.Class) > 0 Then ' Custom tags exist for this object
            Layouts = oCustomTagRepertory.Entry(obj.Class).Count
            Set oCustomTagDictionary = oCustomTagRepertory(obj.Class)
        End If

        For i = 1 to Layouts - 1

            Layout = oCustomTagRepertory.Entry(obj.Class).Object(i)
            cchBufferStart = Report.BufferLength


            ' check if this Custom Tag Dialog Layout has an associated phrase in the Language Dictionary
            strFmtTemplate = Util.FirstNonEmpty(Layout(2), LanguageDictionary.Peek("PhCT_" & Replace(Layout(0)," ","")))

            If strFmtTemplate <> "" Then            ' create a custom phrase
                Report.WriteLn "<li class='xT2-o xT-h' onclick='xTclk(event,""2"")'>"
                Report.WriteFormattedLn "<span class='xT-i subhead bold'>{&t}</span><ul class='xT-h'>", Layout(1)
                ReDim Args(Ubound(Layout)+8)
                Args(0) = obj & ""  ' default property e.g. name
                Select Case obj.Class
                Case "Individual"
                    Args(0) = obj.Name.short
                    Args(1) = PnP(obj)
                    Args(2) = PnR(obj)
                    Args(3) = PnO(obj)
                    Args(4) = Not obj.IsDead
                Case "Family"
                    Args(1) = Dic.Plurial("PnP_", 2)
                    Args(2) = Dic.Plurial("PnR_", 2)
                    Args(3) = Dic.Plurial("PnO_", 2)
                    Args(4) = obj.AreTogether
                Case "Marriage"
                    Args(1) = Dic.Plurial("PnP_", 2)
                    Args(2) = Dic.Plurial("PnR_", 2)
                    Args(3) = Dic.Plurial("PnO_", 2)
                    Args(4) = Not obj.AreTogether
                Case Else
                    strGender = Util.FirstNonEmpty(CustomTag(obj, "Gender.ID"), CustomTag(obj, "Name.Gender.ID"), "N")
                    Args(1) = Dic("PnP_" & strGender)
                    Args(2) = Dic("PnR_" & strGender)
                    Args(3) = Dic("PnO_" & strGender)
                    Args(4) = (CustomTag(obj, "Extant") <> "N")
                End Select
                ' create fixed part of Report.WritePhrase statement. N.B. params 8 & 9 are reserved but not used at present
                strExecute = "Report.WritePhrase strFmtTemplate, Args(0), Args(1), Args(2), Args(3), " & Util.IfElse(Args(4),"True","False") & _
                                         ",""<"", "">"", ""&"", Chr(39), Chr(34)"
                ' now add Custom Tag contents as param 10 onwards.
                For j = 3 to Ubound(Layout)
                    strTag = Layout(j)
                    If strTag <> "" Then
                        Args(j+7) = StrFormatText(obj, CustomTag(obj,strTag))
                    End If
                    strExecute = strExecute & ", Args(" & j+7 & ")"
                Next
                cchBufferNow = Report.BufferLength
                Execute strExecute
                If Report.BufferLength > cchBufferNow Then
                    cchBufferStart = -1      ' indicate at least one value present in this Custom Tag Dialog Layout
                End If

                Report.WriteLn "</ul></li>"
            Else                        ' no custom template
                                    ' so create a collapse/expand section with the
                                    ' custom tags in a table.

                Report.WriteFormattedLn "<li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
                Report.WriteFormattedLn "<span class='xT-i subhead bold'>{&t}</span><ul class='xT-h'><table class='customtagtable'>", Layout(1)
                For j = 3 to Ubound(Layout)
                    strTag = Trim(Layout(j))
                    strCustomTagData = Util.IfElse(strTag <> "", CustomTag(obj,strTag), "")

                    If strCustomTagData <> "" Then
                        strCustomTagDesc = oCustomTagDictionary.KeyValue(strTag)
                        Report.WriteFormattedLn "<tr><td>{&t}</td><td>{}</td></tr>",strCustomTagDesc, StrFormatText(obj, strCustomTagData)
                        cchBufferStart = -1      ' indicate at least one value present in this Custom Tag Dialog Layout
                    End If
                Next

                Report.WriteLn "</table></ul></li>"
            End If

            If cchBufferStart > 0 Then ' no Custom Tags set in this Dialog Layout Section so remove phrase or table
                Report.BufferLength = cchBufferStart
            Else
                Session("BufferBegin") = -1 ' indicate at least one custom tag set
            End If
        Next
    End If
    Report.WriteLn "</ul></li></ul></div>"
    If Session("BufferBegin") > 0 Then    ' no Custom tags present so remove the Additional Information section
        Report.BufferLength = Session("BufferBegin")
    Else
        Session("ReferencesStart") = -1    ' indicate at least one expand/collapse non-note item present
    End If
End Sub

Sub WriteHtmlRelationships(obj)
    Dim strPnO, strPnp, strTitle, fAlive, fExtant, r, oRepertoryEntity, oRepertoryEntity1, oRepertoryEntity2, strConnection, strConnections, strName, strNameO, strNameP
    strPnP = PnP(obj)
    strPnO = Util.StrFirstCharUCase(PnO(obj))

    Set oRepertoryEntity1 = Session("oRepertoryEntity1")
    Set oRepertoryEntity2 = Session("oRepertoryEntity2")

    If Session("Flag_R") And _
        (oRepertoryEntity1.KeyCounter(obj.ID) > 0 Or oRepertoryEntity2.KeyCounter(obj.ID) > 0) Then
        Report.WriteLn    "<div class='clearleft'><br /><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='Relationships'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Dic("Relationships")
        If oRepertoryEntity1.KeyCounter(obj.ID) > 0 Then
            Set oRepertoryEntity = oRepertoryEntity1.Entry(obj.ID)
            strName = obj.Session("NameShort")
            strNameO = strName
            strNameP = obj.Session("NamePossessive")
            For Each r In oRepertoryEntity
                Report.WriteLn "<div class='clearleft'>"
                If (r.Session("PicturesIncluded") > 0) Then WriteHtmlPicturesSmall r, "left", Session("fHidePictureName")
                WriteHtmlRelationship r, strName, strNameO, "", "", False, True, strNameP, ""
                WriteHtmlAdditionalInformation(r)
                WriteHtmlAnnotation r, Dic("AnnotationRelationship"), r.Comment
                strName = strPnP
                strNameO = strPnO
                Report.WriteLn "</div>"
            Next
        End If
        If oRepertoryEntity2.KeyCounter(obj.ID) > 0 Then
            Set oRepertoryEntity = oRepertoryEntity2.Entry(obj.ID)
            strName = obj.Session("NameShort")
            strNameO = strName
            strNameP = obj.Session("NamePossessive")
            For Each r In oRepertoryEntity
                Report.WriteLn "<div class='clearleft'>"
                If (r.Session("PicturesIncluded") > 0) Then WriteHtmlPicturesSmall r, "left", Session("fHidePictureName")
                strConnection = ""
                If r.Class = "SocialRelationship" Then
                    strConnection = r.Connection.ID
                Else
                    strConnection = r.EmotionalLink.ID
                End If
                Select Case strConnection
                Case "Roommate", "LivesWith", "Neighbor", "Acquaintance", "Associate", "Know", "Relative", "WorkWith", "" ' two way relationship so keep individual as subject (i.e. switch)
                    WriteHtmlRelationship r,  r.entity1.Session("HlinkNN"),"", strName, strNameO, True, True, "", StrNameP
                Case Else
                    If strConnection = "" Then strConnection = " "
                    Select Case Left(strConnection,3)
                    Case "Abu", "Man", "Con", "Foc", "Fan", "Lim", "Jea"
                        WriteHtmlRelationship r, "", "", strName, strNameO, False, True, "", strNameP
                    Case Else
                        If r.Class = "EmotionalRelationship" Then
                            WriteHtmlRelationship r,  r.Entity1.Session("HlinkNN"),"", strName, strNameO, True, True, "", strNameP
                        Else
                            WriteHtmlRelationship r, "", "", strName, strNameO, False, True, "", strNameP
                        End If
                    End Select
                End Select
                WriteHtmlAdditionalInformation(r)
                WriteHtmlAnnotation r, Dic("AnnotationRelationship"), r.Comment
                Report.WriteLn "</div>"
                strName = strPnP
                strNameO = strPnO
            Next
        End If

        Session("BufferBegin") = -1      ' indicate data present in other details section
        Session("ReferencesStart") = -1

        Report.WriteLn "</ul></li></ul><div class='clearleft'></div></div>"
    End If
End Sub

' Write the HTML code to display notes and comments for birth, death, education, occupation, etc.
Sub WriteHtmlAnnotation(obj, strAnnotationType, strAnnotationCommentFull)
    Dim strAnnotationComments, strAnnotationComment
    strAnnotationComment = StrFormatText(obj, strAnnotationCommentFull)
    If Session("Notes") And (strAnnotationComment <> "") Then
        Session("NotesStart") = -1    ' Use a negative value to indicate the presence of an annotation
        Report.WriteFormattedLn    "<ul class='xT note'>" & vbNewline & _
                         "  <li class='xT-{} xT-h' onclick='xTclk(event,"""")'><h5 class='inline'>{&t}</h5>" & vbNewline & _
                         "    <ul class='xT-n'>" & vbNewline & _
                         "     <div>{}" & _
                         "     </div>" & vbNewline & _
                         "    </ul>" & vbNewline & _
                    "  </li>" & vbNewline & _
                    "</ul>" & vbNewline, _
                    Util.IfElse(Session("fCollapseNotes"), "c", "o"), strAnnotationType, strAnnotationComment
    End If
End Sub
' as above but get result as string
Function StrHtmlAnnotation(obj, strAnnotationType, strAnnotationCommentFull)
    Dim cchBuffer
    cchBuffer = Report.BufferLength
    WriteHtmlAnnotation obj, strAnnotationType, strAnnotationCommentFull
    StrHtmlAnnotation = Report.Buffer(cchBuffer)    ' Get the text from the output buffer
    Report.BufferLength = cchBuffer    ' Truncate the buffer to its original length
End Function

' Write the HTML code to display the comment of the primary picture (if present)
' If the picture does not have a description, the code must be written, otherwise
' the slideshow will abort by a script error
Sub WriteHtmlAnnotationPicture(obj, p)
    Dim strId, strClass, strDescription
    If IsNull(p) Then Set p = obj.Session("PictureMain")    ' Get the main picture of the object
    If (Not Util.IsNothing(p)) Then
        strId = obj.ID
        strClass = "show"
        strDescription = StrFormatText(p, p.Comment)
        If (strDescription = "") Then
            strDescription = "&#32;"
            strClass = "hide"
        Else
            Session("NotesStart") = -1    ' Use a negative value to indicate the presence of an annotation
        End If
        Report.WriteFormattedLn "<ul id='idPVv_{}' class='xT {} note'>", strId, strClass
        Report.WriteFormattedLn "    <li class='xT-{} xT-h' onclick='xTclk(event,"""")'>{&t}", Util.IfElse(Session("fCollapseNotes"), "c", "o"), Dic("AnnotationPicture")
        Report.WriteFormattedLn "        <ul class='xT-n'><li id='idPVd_{}' class='xT-b xT-n'>{}</li></ul>", strId, strDescription
        Report.WriteLn          "    </li>"
        Report.WriteLn          "</ul>"
    End If
End Sub

' Write details associated with a main picture with a specific HTML id so that it can be overwritten for each picture in a slideshow if required
Sub WriteHtmlDetailsPicture(obj,p)
	if IsNull(p) Then Set p = obj.Session("PictureMain")
    If Not (p Is Nothing) And p.PictureDimension <> "" Then
        Report.WriteFormattedLn "<div class='caption' id='idPVc_{}'>{}</div>", obj.ID, LanguageDictionary.FormatPhrase("PhPictureDetailsHtml", Util.IfElse(Session("fHidePictureName"), "",StrHtmlHyperLink(p)), p.Date.ToStringNarrative, p.Place.Session("HlinkLocative"), p.Source.Session("Hlink"))
    End If
End Sub

' Write a reference to a footnote item
' This routine requires the presence of a GenoCollection object referenced by Session("Footnotes") to store
' the unique footnotes.
Sub WriteHtmlFootnoteRef(oFootnote)
    If Session("Flag_S") And (Not Util.IsNothing(oFootnote)) Then
        Dim iFootnote        ' Zero-based index representing the position where the footnote is found in the collection
        iFootnote = Session("Footnotes").AddUnique(oFootnote)
        Report.WriteFormatted "<sup><a href='#{0}' title='{}'>{0}</a>&nbsp;</sup>", iFootnote + 1, Dic("SourceFootnote")
    End If
End Sub

' Write a reference for two footnotes.  If both footnotes are identical,
' then write only a single footnote
Sub WriteHtmlFootnoteRef2(oFootnote1, oFootnote2)
'    If (Not Util.AreObjectsEqual(oFootnote1, oFootnote2)) Then
    If oFootnote1 <> oFootnote2 Then
        WriteHtmlFootnoteRef oFootnote1
    End If
    WriteHtmlFootnoteRef oFootnote2
End Sub

Sub WriteHtmlFootnoteRefs(oFootnotes)
    Dim oFootNote
    For Each oFootnote in oFootnotes.ToGenoCollection
        WriteHtmlFootnoteRef(oFootnote)
    Next
End Sub

' Write all the footnotes to the report
Sub WriteHtmlAllFootnotes(oSources, asNote)
  Dim collSources, litag
  Set collSources = Util.NewGenoCollection()
  If Not Util.IsNothing(oSources) Then collSources.Add(oSources.ToGenoCollection)
  If Session("Flag_S") And (Session("Footnotes").Count > 0) Or (collSources.Count > 0) Then
      If Session("NestSourceRefs") Then
		  Report.WriteLn "<br />"
          If asNote then
              Report.WriteLn           "<div class='clearleft'>"
              Report.WriteLn           "  <ul class='xT note'>"
              Report.WriteFormattedLn  "    <li class='xT-o xT-h' onclick='xTclk(event,"""")'>Source", Dic.Plurial("SourceCitation", 1)
              Report.WriteLn           "      <ul class='xT-n'>"
              litag =           "        <li class='xT-b xT-n'>"
          Else
              Report.WriteLn           "<div class='clearleft'>"
              Report.WriteLn           "  <ul class='xT'>"
              Report.WriteFormattedLn  "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
              Report.WriteFormattedLn  "      <a name='SourcesCitations'></a><h4 class='xT-i inline'>{&t}</h4>", Dic.Plurial("SourceCitation", 2)
              Report.WriteLn           "      <ul class='xT-h'>"
              litag =           "        <li>"
          End If
      Else
          Report.WriteLn "<div class='footnote clearleft'>"
      End If
          If (Session("Footnotes").Count > 0) Then
              Dim iFootnote, oFootnote, strTitle, strSep, s
              iFootnote = 0
              For Each oFootnote In Session("Footnotes")
                  iFootnote = iFootnote + 1
                  strTitle = StrPlainText(oFootnote, Util.FirstNonEmpty(oFootnote.Subtitle, oFootnote.Description, oFootnote.QuotedText, Dic("SourceInformation")))
                  If (strTitle <> "") Then
                      strTitle = Util.FormatString(" title='{}'", strTitle)
                  End If
                  Report.WriteFormattedBr "<sup><a name='{0}'></a>{0} </sup> <a href='source-{&t}.htm' {}><i>{}</i></a>", iFootnote, oFootnote.ID, strTitle, JoinSourceCitationNames(oFootnote, StrFormatText(oFootnote, StrParseText(oFootnote.title, True)), true)
                  ' remove referenced source from list
                  collSources.Remove(oFootnote)
              Next
          End If
          If collSources.Count > 0 Then ' write details of remaining source/citations
              If Not Session("NestSourceRefs") Then Report.WriteBr Dic.Plurial("SourceCitation" & Util.IfElse(Session("Footnotes").Count > 0, "_Other",""), collSources.Count)
              strSep = ""
              For Each oFootnote in collSources
				strTitle = StrPlainText(oFootnote, Util.FirstNonEmpty(oFootnote.Subtitle, oFootnote.Description, oFootnote.QuotedText, Dic("SourceInformation")))
				If (strTitle <> "") Then strTitle = Util.FormatString(" title='{}'", strTitle)
				If Session("NestSourceRefs") Then Report.WriteLn litag
				Report.WriteFormatted strSep & " <a href='source-{&t}.htm' {}><i>{}</i></a>", oFootnote.ID, strTitle, JoinSourceCitationNames(oFootnote, StrFormatText(oFootnote, StrParseText(oFootnote.title, True)), true)
				strSep = Util.IfElse(Session("NestSourceRefs"), "", ", ")
				If Session("NestSourceRefs") Then Report.WriteLn "        </li>"
              Next
          End If
          If Session("NestSourceRefs") Then
          Report.WriteLn "      </ul>"
          Report.WriteLn "    </li>"
          Report.WriteLn "  </ul>"
      End If
      Report.WriteBr "</div>"
             Session("Footnotes").Clear
      End If
End Sub


' This method writes the HTML code to make sure the entire report is accessible from any page.
' Do not remove this safeguard; it may affect the navigation of the report.
' The text is written in English because it is very rare someone will actually read it.  This text is not
' the most eloquent English, however it contains a wealth of keywords for surfers to quickly find your
' genealogy report if published on the web.
' The text should be translated to include keywords in your language.
Sub WriteHtmlFramesetSafeguardK(s, k, o)
    Dim strHref, strName, strClass, strLink
    strHref = "toc_individuals.htm"
    If Session("Trees") And Not Util.IsNothing(o) Then ' page last updated only if publishing to familytrees.genopro.com
    Report.WriteLn StrDicExt("FmtHtmlLastModified", "", "<div class='small floatright'><span>Page last modified </span><span  id='lastModified'></span></div>", "", "2.0.1.6 (>01.09.2010)")
        strClass = o.Class
        If (strClass = "Individual") Then
            strName = o.Name.Last
            k = ""
        ElseIf (strClass = "Family") Then
            strName =  o
            strHref = "toc_families.htm"
            k = ""
        End If
    End If
    If Not Session("ForceFrames") Then Exit Sub
    Report.WriteLn "<div id='divFrameset' class='hide'> <hr /><p><br />"
    Report.WriteBr Util.FormatPhrase(Dic("FmtFramesetSafeguard"), strName, strHref, Session("Title"), strClass="Individual",  strClass="Family")
    Report.WriteT3Br "(", k, ")"
    strLink =  Replace(Dic.SearchKeywordHyperlink, "Learn how to build your family tree", Session("SKAltDefault"))
    Report.WriteFormatted "{}.</p><hr /><br /><p class='aligncenter'><small>{}</small></p></div>", strLink, Util.FormatString(Dic("FmtCopyright"), "© 2011 GenoPro Inc.")
End Sub

Sub WriteHtmlFramesetSafeguard(s)
    WriteHtmlFramesetSafeguardK s, "", nothing
End Sub
Sub WriteHtmlPedigreeChart(i)
    If Not Session("Flag_P") Then Exit Sub
    If  (i.Mother.Name<>"") Or (i.Father.Name<>"") Then
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present

        Dim ChartMap
        Set ChartMap = Session("ChartMap")
        Dim Ancestors, fMore, depth
        fMore =False
        depth=0
        Set Ancestors = Util.NewObjectRepertory

        SetAncestor i, Ancestors, "", depth

        Dim e, r, s, c, obj, k
        s = Array(0)
        If depth > 0 Then
            Report.WriteLn "<br /><div class='clearleft no-break'><ul class='xT'>"
            Report.WriteFormattedLn "    <li class='xT2-{} xT-h XT-clr clear{}' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o"), Util.IfElse(Session("fCollapseReferences"), "left", "")
            Report.WriteFormattedLn "<a name='PedigreeChart'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'><li>", Dic("PedigreeChart")
            Report.WriteBr
            Report.WriteLn "<div class='chartblock'>"
            For e = 0 to 62
                r = ChartMap.Entry(e).Object(0)
                k = r(r(0))
                If (Ancestors.KeyCounter(k)>0) Then
                    Set obj = Ancestors.Object(k)
                        If Not Util.IsNothing(obj) Then
                            fMore = fMore Or PrintChartLine(r, obj, s)
                            s = r
                        End If
                End If
            Next
            Report.WriteLn "</div>"
            If fMore Then Report.Write "<a name='chartnote'></a><sup>*</sup>" & Dic("PedigreeChartFootnote")
            Report.WriteLn "</li></ul></li></ul></div>"
        End If
    End If
End Sub
Function PrintChartLine(r, i, s)
    Dim c, l, obj, strFill, strLink, fMore, gender, fPadding, strName
    fMore = False
    fPadding = False
    gender = Right(r(r(0)),1)
    strName = Util.HtmlEncode(Util.IfElse(Session("OriginalNamesCharts"), i.TagValue(Session("TagNameFull")), Replace(i.Session("NameFull"), Session("MarkerFirstName"), "")))
    If gender = "i" Then
        gender = gender & i.Gender.ID
        For c = 1 to s(0)-1
            If s(c)="I" Or s(c)="T" Then
                Report.WriteFormatted "<img class='chart' src='images/line_I.gif' alt='' title=''/>"
                fPadding = True
            End If
        Next
        If fPadding Then Report.WriteBr
    End If
    For c = 1 to r(0)-1
        Report.WriteFormatted "<img class='chart' src='images/line_{}.gif' alt='' title=''/>", r(c)
    Next
    if i.Href <> "" Then
        strLink=Util.FormatString("<a href='{}' target='detail'>{}</a>", i.Href, Replace(strName," ","&nbsp;"))
    Else
      strName = StrDicMFU("_NoName", i.Gender.ID)
        strLink=Util.HtmlEncode(strName)
    End If
    l=10 - Len(strName)

    If c=6 Then
        If i.Mother.Name <> "" Or i.Father.Name <> "" Then
            strLink = strLink & "<a href='#chartnote'><sup>* </sup></a>"
            fMore = True
        End If
    End If
    If l > 0 Then 
        For c= 1 To l
            strLink = strLink & "&nbsp;"
        Next
    End If
    Report.WriteFormattedLn "<div class='chartbox'><div class='charttext {0}'>{1}</div></div><div class='chartdates'>{2}{3}</div>"+Report.TagBr, _
                    gender, strLink, "&nbsp;",_
                    Replace(Util.HtmlEncode(Util.FormatPhrase(Dic("PedigreeChartDetails"), _
                        Trim(CustomDate(i.Birth.Date).ToString(Dic("PedigreeChartDate"))), i.Birth.Place.Session("Locative"), _
                        Trim(CustomDate(i.Death.Date).ToString(Dic("PedigreeChartDate"))), i.Death.Place.Session("Locative")))," ","&nbsp;")
    If fMore Then
        PrintChartLine = True
    Else
        PrintChartLine = False
    End If
End Function
Function SetAncestor(i, Ancestors, p, ByRef depth)
    Ancestors.Object(Util.FirstNonEmpty(p,"i")) = i
    If Len(p) < 5 Then
        If Not Util.IsNothing(i.Father) Then SetAncestor i.Father, Ancestors,p & "f", depth
        If Not Util.IsNothing(i.Mother) Then SetAncestor i.Mother, Ancestors,p & "m", depth 
    End If
    If Len(p) > depth Then depth = Len(p)
End Function
Sub WriteHtmlEducations(i)
        Dim strPrefix, e, p, collEducations, strTermination, oRegExp
    Set collEducations = i.Educations.ToGenoCollection
    If Session("Flag_E") And (collEducations.Count > 0) Then
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present
        Report.WriteLn    "<div class='clearleft'><br /><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='Education'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Dic("HeaderEducation")
        strPrefix = i.Session("NameShort")
        Set oRegExp = New RegExp
        oRegExp.Global=True
        oRegExp.Pattern="[\.\ \(\)\\\/,']"
        For Each e In collEducations 
            Report.WriteFormattedLn "<a name='{}'></a><div class='clearleft'>", e.ID
            If (e.Pictures.Count > 0) Then WriteHtmlPicturesSmall e, "left", True
            Set p = e.Place

            If e.Termination.ID <> "" Then
                strTermination = StrDicMFU2("PhET_" & e.Termination.ID, i.Gender.ID, "=" & e.Termination, "2.0.1.?")
                If Instr(strTermination,"{") > 0 Then strTermination = Util.FormatPhrase(strTermination, i.Gender.ID)
            Else
                strTermination = ""
            End If
            Report.WritePhrase StrEventPhrase(e), _
                            strPrefix, _
                            Util.IfElse(e.Institution <> p.Name, p.Session("HlinkLocative"), p.Session("Hlink")), _
                            e.StudyType, _
                            StrDateSpan(e.DateStart, Util.IfElse(e.Termination = "", e.DateEnd, Nothing)), _
                             StrTimeSpan(e.Duration), _
                            Util.FirstNonEmpty(Dic.Peek("PhEL_" & oRegExp.Replace(e.Level,"")),Util.FormatPhrase(Dic("PhEL"),e.Level)), _
                            StrFormatText(e, e.Program), _
                            Util.FirstNonEmpty(Dic.Peek("PhEA_" & oRegExp.Replace(e.Achievement,"")), Util.FormatPhrase(Dic("PhEA"), StrFormatText(e, e.Achievement))), _
                            e.Termination.ID = "StillAttending", _
                            Util.IfElse(e.Institution <> p.Name, StrParseText(StrPlaceTranslate(e.Institution), True), ""), _
                            PnP(i), _
                            CustomDate(e.DateEnd).ToStringNarrative, _
                            strTermination, _
                            Dic.Peek("PrefixInstitution_" & CustomTag(e, "Institution.Prefix")), _
                            Dic.Peek("PrefixProgram_" & CustomTag(e, "Program.Prefix"))
            WriteHtmlFootnoteRef(e.Source)
            WriteHtmlFootnoteRefs(e.Sources)
            WriteHtmlAdditionalInformation(e)
            WriteHtmlAnnotation e, Dic("AnnotationEducation"), e.Comment
            If Session("fShowPictureDetails") Then
                WriteHtmlDetailsPicture e, Null
                WriteHtmlAnnotationPicture e, Null
            End If
            'Report.WriteBr "</div>"
            strPrefix = PnP(i)
        Next
        Report.WriteLn "</div>"
        Report.WriteLn "</ul></li></ul></div>"
    End If
End Sub

Sub WriteHtmlEventsAndAttributes(i, collOccupations)
    If Session("Flag_A") And (collOccupations.Count > 0) Then
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present
        Report.WriteLn    "<div class='clearleft'><br /><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='AttributesEvents'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Dic("HeaderAttributesEvents")
        Dim strRelative, strPnP, strGender, strName, strPhrase
        strGender = i.Gender.ID
        strRelative = i.Session("NamePossessive")
        strPnp = PnP(i)
        strName = i.Session("NameShort")
        For Each o In collOccupations
          If o.Session("Event") <> "" Then
              Report.WriteFormattedLn "<a name='{}'></a><div class='clearleft'>", o.ID
              If (o.Pictures.Count > 0) Then WriteHtmlPicturesSmall o, "left", True
              strPhrase = StrEventPhrase(o)
              Report.WritePhrase strPhrase, _
                              StrDateSpan(o.DateStart, Util.IfElse(o.Termination.ID = "", o.DateEnd, Nothing)), _
                              i.Session("NameShort"), _
                              strRelative, _
                              StrFormatText(o, StrParseText(o.Session("Title"), True)), _
                              "<h6>","</h6>", _
                              StrFormatText(o, StrParseText(o.Session("Company"), True)), _
                              StrHtmlHyperlink(o.Place), _
                              Util.FirstNonEmpty(Dic.Peek("PrefixCompany_" & o.Session("Company.Prefix")), o.Session("Company.Prefix")), _
                              strName, _
                              Util.FirstNonEmpty(Dic.Peek("PrefixTitle_" & o.Session("Title.Prefix")), o.Session("Title.Prefix"))
              Report.WritePhrase StrDicMFU("PhJT_" & o.Session("Event") & o.Termination.ID, i.Gender.ID), strPnP, CustomDate(o.DateEnd).ToStringNarrative, strGender
              WriteHtmlFootnoteRef(o.Source)
              WriteHtmlFootnoteRefs(o.Sources)
              WriteHtmlAdditionalInformation(o)
              WriteHtmlAnnotation o, Dic("AnnotationEventAttribute"), o.Comment
              If Session("fShowPictureDetails") Then
                  WriteHtmlDetailsPicture o, Null
                  WriteHtmlAnnotationPicture o, Null
              End If
              If InStr(strPhrase, "{9}") > 0 Or InStr(strPhrase, "{!9}") > 0 Then strName = strPnP
              If InStr(strPhrase, "{2}") > 0 Or InStr(strPhrase, "{!2}") > 0 Then strRelative = PnR(i)
              Report.WriteLn "</div>"
      End If
        Next
        Report.WriteLn "</ul></li></ul></div>"
    End If
End Sub

Sub WriteHtmlOccupations(i, collOccupations)
    If Session("Flag_O") And (collOccupations.Count > 0) Then
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present
        Report.WriteLn    "<div class='clearleft'><br /><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='Occupation'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Dic("HeaderOccupation")
        Dim strRelative, strPnP, strGender, strName
        strGender = i.Gender.ID
        strRelative = i.Session("NamePossessive")
        strPnp = PnP(i)
        strName = i.Session("NameShort")
        For Each o In collOccupations
          If o.Session("Event") = "" Then
              Report.WriteFormattedLn "<a name='{}'></a><div class='clearleft'>", o.ID
              If (o.Pictures.Count > 0) Then WriteHtmlPicturesSmall o, "left", True
              Report.WritePhrase StrDicOrTag("PhOccupation", CustomTag(o, "NarrativeStyle")), _
                              StrDateSpan(o.DateStart, Util.IfElse(o.Termination.ID = "" , o.DateEnd, Nothing)), _
                              StrTimeSpan(o.Duration), _
                              strRelative, _
                              StrFormatText(o, StrParseText(o.Session("Title"), True)), _
                              o.WorkType, o.Industry, _
                              StrFormatText(o, StrParseText(o.Session("Company"), True)), _
                              StrHtmlHyperlink(o.Place), _
                              o.Termination.ID = "StillWorking", _
                              strName, _
                              Util.FirstNonEmpty(Dic.Peek("PrefixTitle_" & o.Session("Title.Prefix")), o.Session("Title.Prefix")), _
                              Util.FirstNonEmpty(Dic.Peek("PrefixCompany_" & o.Session("Company.Prefix")), o.Session("Company.Prefix"))
              Report.WritePhrase StrDicMFU("PhJT_" & o.Termination.ID, i.Gender.ID), strPnP, CustomDate(o.DateEnd).ToStringNarrative, strGender
              WriteHtmlFootnoteRef(o.Source)
              WriteHtmlFootnoteRefs(o.Sources)
              WriteHtmlAdditionalInformation(o)
              WriteHtmlAnnotation o, Dic("AnnotationOccupation"), o.Comment
              If Session("fShowPictureDetails") Then
                  WriteHtmlDetailsPicture o, Null
                  WriteHtmlAnnotationPicture o, Null
              End If
              strName = strPnP
              strRelative = PnR(i)
              Report.WriteLn "</div>"
          End If
        Next
        Report.WriteLn "</ul></li></ul></div>"
    End If
End Sub

Function WriteIndividualEvents (oTLInfo, i, strTitle, fYearOnly)
    Dim cchBegin, cchStart, oStart, oEnd, fLink, nEvent, strGender, strPnR
    Dim b, ba, c, d, f
    cchStart = Report.BufferLength
    strGender = i.Gender.ID
    strPnR = PnR(i)
    Set b = i.Birth
    Set ba = b.Baptism
    Set d = i.Death
    Set f = d.Funerals
    Set c = d.Cause
    If b.Date <> "" Then
        Set oStart = b.Date
    ElseIf ba.Date <> "" Then
        Set oStart = ba.Date
    End If
    If d.Date <> "" Then
        Set oEnd = d.Date
    ElseIf f.Date <> "" Then
        Set oEnd = f.Date
    End If
    If IsObject(oStart) Then
        If IsObject(oEnd) Then
            oTLInfo.AddEvent i, oStart, oEnd, fYearOnly, strTitle, _
                    Util.FormatPhrase(StrDicMFU("PhTL_Birth", strGender), PnP(i), b.Date.ToStringNarrative, b.Place.Session("Locative"), ba.Date.ToStringNarrative, _
                                ba.Place.Session("Locative"), ba.Officiator.Title, ba.Officiator, (b.Place.ID = ba.Place.ID), _
                                b.Doctor, b.PregnancyLength.Months,  b.PregnancyLength.Weeks, strPnR, b.CeremonyType, i.Session("NamePossessive"), strGender) & " " & _
                               Util.StrFirstCharUCase(Util.FormatPhrase(StrDicMFU("PhTL_Died", strGender), PnP(i), i.Death.Age.Years, i.Death.Age.Months, i.Death.Age.Days, _
                                d.Date.ToStringNarrative, d.Place.Session("Locative"), c, c.Description, i.Birth.Date.Approximation & i.Death.Date.Approximation, strGender))
        Else
            oTLInfo.AddEvent i, oStart, , fYearOnly, strTitle, _
                    Util.FormatPhrase(StrDicMFU("PhTL_Birth", strGender), PnP(i), b.Date.ToStringNarrative, b.Place.Session("Locative"), ba.Date.ToStringNarrative, _
                                ba.Place.Session("Locative"), ba.Officiator.Title, ba.Officiator, (b.Place.ID = ba.Place.ID), _
                                b.Doctor, b.PregnancyLength.Months,  b.PregnancyLength.Weeks, strPnR, b.CeremonyType, i.Session("NamePossessive"), strGender)
        End If
    End If
    If IsObject(oEnd) And Not IsObject(oStart) Then
        oTLInfo.AddEvent i, oEnd, oEnd, fYearOnly, strTitle & " " & Dic("DiedAbbr"), _
            Util.FormatPhrase(StrDicMFU("PhTL_Died", strGender), PnP(i), i.Death.Age.Years, i.Death.Age.Months, i.Death.Age.Days, d.Date.ToStringNarrative, d.Place.Session("Locative"), c, c.Description, i.Birth.Date.Approximation & i.Death.Date.Approximation, strGender)
    End If
End Function

' Write event details as a JSON array element for the MIT Simile Timeline 'widget'.
Sub WriteIndividualTimelineData(i)
    ' create timeline data in JSON format.

    Dim nEvents, strLocale, strBuffer, collEvents, oEvent, strEvent, fTimeline, cchStart, oCnt

    strLocale = GetLocale
    ' force Locale to be English so that dates are in english. 
    SetLocale("en-gb")
    'oLinks.Clear()
    cchStart = Report.BufferLength

    oTLInfo.AddHeader i.ID, True
    If CustomDate(i.Birth.Date).Year <> "" and CustomDate(i.Death.Date).Year <> "" Then
        oTLInfo.AddEvent i, i.Birth.Date, i.Death.Date, False, Dic("Lived"), ""
    Else
        oTLInfo.AddEvent i, i.Birth.Date, , False, Dic("Birth"), ""
        oTLInfo.AddEvent i, i.Death.Date ,i.Death.Date, False, Dic("Death"), ""
    End If
    oTLInfo.AddEvent "", i.Birth.Baptism.Date, , False, Util.FirstNonEmpty(i.Birth.CeremonyType, Dic("BirthCeremonyTypeDefault")), ""
    oTLInfo.AddEvent "", i.Death.Funerals.Date, , False, Dic("Funeral"), ""
    oTLInfo.AddEvent "", i.Death.Disposition.Date, , False, Dic.LookupEx("PhBD_", i.Death.Disposition.Type),""

    Set collEvents = i.Educations.ToGenoCollection
    For Each oEvent In collEvents
        oTLInfo.AddEvent oEvent, oEvent.DateStart, oEvent.DateEnd, False, StrPlainText(oEvent, Util.FirstNonEmpty(oEvent.Institution, oEvent.Place.Session("NameFull"), oEvent.Level)), ""
    Next

    Set collEvents = i.Occupations.ToGenoCollection
    For Each oEvent In collEvents
          oTLInfo.AddEvent oEvent, oEvent.DateStart, oEvent.DateEnd, False, StrPlainText(oEvent, Util.FirstNonEmpty(oEvent.Session("EventName"), oEvent.Session("Title"), oEvent.Company, oEvent.Industry, oEvent.WorkType)), oEvent.Session("Title") & " - " & oEvent.Session("Company")
    Next

    Set collEvents = i.Contacts.ToGenoCollection
    For Each oEvent In collEvents
        oTLInfo.AddEvent oEvent, oEvent.DateStart, oEvent.DateEnd, False, Util.IfElse(oEvent.Place.Name <> "", Util.FormatString(StrDicOrTag("Timeline_" & oEvent.Type.ID, CustomTag(oEvent, "NarrativeStyle")), oEvent.Place.Session("NameFull")),""), ""
    Next

    Set collEvents = i.Families.ToGenoCollection
    For Each oEvent In collEvents
        WriteFamilyEvents oTLInfo, oEvent, Replace(i.FindMate(oEvent).Session("NameFull"), Session("MarkerFirstName"), ""), False, True
    Next

    oTLInfo.AddTrailer True, ""

    If oTLInfo.Nodes >= Session("TimelineMinEventsIndividual") Then 
        i.Session("Timeline") = True
    Else
        Report.BufferLength = cchStart
    End If

    setLocale(strLocale)
End Sub
Sub WriteIndividualBody(i)
    Dim level
    Report.WriteLn "<div>"
    Report.WriteFormattedLn  "<a id='{}'></a>", i.ID
    Report.WriteLn "<h3>"
    Report.WriteFormatted "{}{}", StrHtmlImgGender(i), StrHtmlHighlightName(i.Session("NameAlternative"))
    If Not Session("Book") Then
        Report.WriteFormatted "{}{}{}{}{}</h3>", StrHtmlImgFileGno(i), StrHtmlImgFileSvg(i), StrHtmlImgTimeline(i), StrHtmlImgMap(i), StrHtmlImgDescendantTreeChart(i)
    Else
        Report.WriteLn "</h3>"
    End If
        '    put pictures in a 'div' floating right. If no pictures added then remove the 'div' 

    Dim cchBufferStart    ' Start of the write operations

    cchBufferStart = Report.BufferLength

    Report.WriteLn "<table class='photo floatright aligncenter widthpaddedlarge'><tr><td>"

    Session("BufferBegin") = Report.BufferLength
    If i.Pictures.Count > 0 Then
        WriteHtmlPicturesLarge i, "right", "", Session("fHidePictureName") Or Session("fShowPictureDetails"), False
    Else
        If Session("fAddGenericImage") Then WriteHtmlGenericPicture i, "P", "l", Session("cxPictureSizeLarge"), Session("cyPictureSizeLarge"), Session("cxyPicturePadding"), "right"
    End If
    If Report.BufferLength = Session("BufferBegin") Then        ' no picture information written
        Report.BufferLength = cchBufferStart        ' so remove the 'div' by stepping back the buffer
    Else
        Report.WriteLn "</td></tr></table>"                ' close the 'div' with end tag
    End If

    WriteNarrativeIndividual i, false, Null
    Report.WriteLn "</div>"
    WriteHtmlAnnotation i, Dic("AnnotationGeneral"), i.Comment
    WriteHtmlExtraNarrative i
    WriteHtmlEducations(i)
    Dim oCnt, collOccupations
    oCnt = i.Session("Events")
    Set collOccupations = i.Occupations.ToGenoCollection
    If oCnt < collOccupations.Count Then WriteHtmlOccupations i, collOccupations
    WriteHtmlOccupancies(i)
    If oCnt > 0  Then WriteHtmlEventsAndAttributes i, collOccupations
    WriteHtmlRelationships i
    If Session("Flag_P") Then WriteHtmlPedigreeChart(i)
    if i.Session("Timeline") Then
        Report.WriteLn          "<div class='clearleft'><br /><ul class='xT'>"
        Report.WriteLn          "    <li class='xT3-o xT-h XT-clr clear' onclick='xTclk(event,""3"")'>"
        Report.WriteFormattedLn "       <a name='TimeLine'></a><h4 class='xT-i inline'>{}</h4><ul class='xT-h'><li>", Dic("TimelineHeadingIndividual")
        Report.WriteFormattedLn "       <div id='tl_{}' class='clear timeline' style='height: 50px; border: 1px solid #aaa;direction:ltr;' >", i.ID
        Report.WriteLn          "       </div>"
        Report.WriteLn          "   </li></ul></li></ul>"
        Report.WriteLn          "</div>"
    End If

    WriteHtmlAdditionalInformation(i)
    
    Dim collFamilies, iFamily, nFamily, iFamilyLast, f, spouse, strSpouseName, strHyperlinkText, p, s
    
    If i.Families.Count > 0 And i.Families.Order.Count > 0 Then
        Set collFamilies = i.Families.Order.ToGenoCollection
    Else
        Set collFamilies = i.Families.ToGenoCollection
    End If
    If Not Session("fHideFamilyDetails") Then
        If collFamilies.Count > 0 Then
            Report.WriteLn "<a name='Family'></a>"
            nFamily = 0
            iFamilyLast = collFamilies.Count - 1
            For iFamily = 0 To iFamilyLast
                If (iFamilyLast > 0) Then
                    nFamily = iFamily + 1
                End If
                Set f = collFamilies(iFamily)
                Set p = f.Parents.ToGenoCollection
                If Session("OnlyPrincipalSpouse") Then
                    s = CustomTag(f, "PrincipalSpouse")
                    If Util.IsNothing(s) Or Not IsNumeric(s) Then s = 0
                    If p(CInt(s)).ID = i.ID Then ' this is the principal spouse
                        WriteHtmlFamily f, nFamily, i
                    Else ' refer to principal
                        Set spouse = i.FindMate(f)    ' Find the other spouse
                        strSpouseName = Util.IfElse(f.Parents.Count > 1, spouse.Session("NameFull"), "")
                        If (strSpouseName = "") And Not Util.IsNothing(spouse) Then
                            strSpouseName = StrDicMFU("_NoName", Util.FirstNonEmpty(spouse.Gender.ID,Util.IfElse(i.Gender.ID = "M", "F", "M")))
                        End If
                        strHyperlinkText = Util.FormatPhrase(StrDicExt("PhFamilyReference","","See {2} for details of {0} [{1} ]family with {4}.","","2013.10.21"), _
                                                        i.Session("NamePossessive"), _
                                                        StrDicMFU("_Ordinal_" & nFamily, StrDicAttribute("PhFamilyWith", "G1")), _
                                                        strSpouseName, _
                                                        Dic("DetailLink"), _
                                                        PnO(spouse))
                        Report.WriteFormatted "<h4>{}{}</h4>", StrHtmlImgFamily(f), Util.FormatHtmlHyperlink(spouse.Href, strHyperlinkText)
                    End If
                Else 
                    WriteHtmlFamily f, nFamily, i
                End If
            Next
        End If
    Else
      Report.Writeln "<div class='clear'></div>"
        For iFamily = 0 To collFamilies.Count - 1
            Set f = collFamilies(iFamily)
            Set spouse = i.FindMate(f)    ' Find the other spouse
            strSpouseName = Util.IfElse(f.Parents.Count > 1, spouse.Session("NameFull"), "")
            If (strSpouseName = "") And Not Util.IsNothing(spouse) Then
                strSpouseName = StrDicMFU("_NoName", Util.FirstNonEmpty(spouse.Gender.ID,Util.IfElse(i.Gender.ID = "M", "F", "M")))
            End If
            strHyperlinkText = Dic.FormatPhrase("PhFamilyWith", i.Session("NamePossessive"), Dic.Ordinal(nFamily), strSpouseName, Dic("DetailLink"))
            Report.WriteFormatted "<h4>{}{}</h4>", StrHtmlImgFamily(f), Util.FormatHtmlHyperlink(f.Href, strHyperlinkText)
        Next
    End If
    Report.WriteLn "<span class='clear'></span>"
    WriteHtmlAllFootnotes i.Sources, False
End Sub


' Write the report description for the META tag.
' This routine is used by default.htm, header.htm and home.htm.
Sub WriteMetaDescriptionReport

    Report.WritePhraseDic "FmtMetaDescReport1", Session("Title"), Individuals.Count, Families.Count
    Report.WriteText " " & Dic("FmtMetaDescReport2") & " "

    Dim oStringDictionaryNames, iName, iNameLast, strSep
    Set oStringDictionaryNames = Session("oStringDictionaryNames")
    If (Not Util.IsNothing(oStringDictionaryNames)) Then
        iNameLast = oStringDictionaryNames.Count - 1
        If (iNameLast > 9) Then
            iNameLast = 9        ' Keep only the first 10 families
        End If
        strSep = ""
        For iName = 0 To iNameLast
            Report.WriteText strSep & oStringDictionaryNames.Key(iName)
            strSep = ", "
        Next
    End If
    Report.WriteText "."
End Sub
' Write the report keywords for the META tag.
Sub WriteMetaKeywordsReport
    Report.Write Dic("FmtMetaKeyWordsReport")
End Sub
' ===============================================================================
Function StrFormatText(obj, strRawText)
'
' Handles text 'processing instructions' embedded in Comment, Descriptions, Custom Tags etc.
'
    Dim strArgs, strPart, strParts, strSubParts, strMap, strParams, strParam, strResult, strText
    Dim i, j, strUID, strMode, strTemp, strClass, arrArgs(15)
    strText = StrParseText(strRawText, True)
    If strText = "" or strText = "<?text?>" Then
        StrFormatText = ""
        Exit Function
    End If
    strClass = "Document"
    If Not IsNull(obj) Then strClass = obj.Class
    If Left(strText,7)="<?off?>" Then
        StrFormatText=Util.HtmlEncode(Mid(strText,8))
        Exit Function
    End If
    strMode = "<?text?>"
    strParts = split(strText, "<?")
    strSubParts = split(strParts(0),StrDicOrTag("", "Private"))
	strResult=""
    If strParts(0) <> "" Then strResult = Util.HtmlEncode(strSubParts(0))        ' get any leading text less any old format private section
    For i=1 to Ubound(strParts)
        strSubParts = split(strParts(i), "?>")
        If Ubound(strSubParts) <> 1 Then Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), obj, strRawText)
        strParams = split(strSubParts(0), " ")
        Select Case LCase(strParams(0))
            Case "off" ' turn off custom markup processing i.e. treat everything after <?off?> directive as standard text
                strResult = strResult + Util.HtmlEncode(mid(strText, Instr(strText, "<?off?>")+7))
                Exit For
            Case "html"
                strResult = strResult & strSubParts(1)
                strMode = "<?" & strSubParts(0) & "?>"
            Case "text"
                strResult = strResult & Util.HtmlEncode(strSubParts(1))
                strMode = "<?" & strSubParts(0) & "?>"
            Case "hide", "plain"
            Case "image"
                strParams = Split(Trim(Mid(strSubParts(0),6)),"""")
                If Ubound(strParams)  <> 2 And UBound(strParams) <> 0 Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                Else
                    If UBound(strParams) = 2 Then
                        If Trim(strParams(0)) <> "" Then Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), obj, strSubParts(0))
                        strTemp = strParams(1)
                        strParam = Session("eSpace").Replace(strParams(2), " ") & "  "
                        strParams = Split(strParam ," ")
                        strParams(0) = strTemp
                        strTemp = ""
                    Else
                        strTemp = Trim(Session("eSpace").Replace(strParams(0), " "))
                        strParams = Split(strTemp & "  "," ")
                        strTemp=Session("oPicMaps").KeyValue(strParams(0))
                        strParams(0) = Session("oPicIndex")(strParams(0)).Path.Report
                    End If
                    strMap = ""
                    If strTemp <> "" Then strMap = StrNewUID
                    strResult = strResult & Util.FormatString("<div width='95%' style='text-align:center; overflow:auto;'>" & _
                                "<img src='{}' class='pic' {}/>", _
                                strParams(0), Util.FormatPhrase("[ width='{0}'][ height='{1}'][ usemap='#map{2}']", strParams(1), strParams(2), strMap))
                    If strTemp<>"" Then strResult = strResult & Util.FormatString("<map name='map{}'>{}</map><div class='note'>{}</div>", strMap, strTemp, Dic("PictureMapHint"))
                    strResult = strResult & "</div>"
                    If UBound(strSubParts) > 0 Then strResult = strResult & StrFormatText(obj, strMode & strSubParts(1))
                End If
            Case "merge"
                strParams = Split(Trim(Mid(strSubParts(0),6)),"""")
                strResult = strResult & StrMergeText(obj, strMode, strParams, 0)
                If UBound(strSubParts) > 0 Then strResult = strResult & StrFormatText(obj, strMode & strSubParts(1))
            Case "note"
                strParams = Split(Trim(Mid(strSubParts(0),5)),"""")
                If Ubound(strParams)  < 2 Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"),strClass, strSubParts(0))
                ElseIf Trim(strParams(0)) <> "" Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"),strClass, strSubParts(0))
                Else
                    strResult = strResult & StrHtmlAnnotation(obj, strParams(1), "<?html?>" & StrMergeText(obj, strMode, strParams, 2))
                    If UBound(strSubParts) > 0 Then strResult = strResult & StrFormatText(obj, strMode & strSubParts(1))
                End If
            Case "popup"
                strParams = Split(Trim(Mid(strSubParts(0),6)),"""")
                If Ubound(strParams)  < 4 Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                ElseIf (Trim(strParams(0)) & Trim(strParams(2))) <> "" Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                Else
                    strUID = StrNewUid
                    strResult = strResult & Util.FormatString("<div class='std'><a href=""javascript:savePopupContent('popup{}', '{&j}');displayPopup();showPopUpFrame('');"">{}</a></div>", strUID, strParams(1), strParams(3))
                    strResult = strResult & Util.FormatString("<div id=""popup{}"" style=""display:none"">{}</div>", strUID, StrMergeText(obj, strMode, strParams, 4))
                    If UBound(strSubParts) > 0 Then strResult = strResult & StrFormatText(obj, strMode & strSubParts(1))
                End If
            Case "subsection"
                strParams = Split(Trim(Mid(strSubParts(0),11)),"""")
                If Ubound(strParams)  < 2 Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                ElseIf Trim(strParams(0)) <> "" Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                Else
                    strResult = strResult & StrHtmlSubSection(obj, strParams(1), StrMergeText(obj, strMode, strParams, 2))
                    If UBound(strSubParts) > 0 Then strResult = strResult & StrFormatText(obj, strMode & strSubParts(1))
                End If
            Case "url"    ' with thanks & acknowledgement to 'BobWebster' on the GenoPro forum
                 if Ubound(strParams) >= 1 Then
                        strParams = Split(Trim(Mid(strSubParts(0),4)) & "\", "\")
                    strTemp = strParams(0)
                    If instr(strParams(0), "://") = 0 Then strParams(0) = "http://" & strParams(0)
                    strResult = strResult & "<a target=""_blank"" href=""" & strParams(0) & """>" & Util.HtmlEncode(Util.FirstNonEmpty(strParams(1), strTemp)) & "</a>" & Util.HtmlEncode(strSubParts(1))
                Else
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                End If
            Case "email", "mail"
                If Ubound(strParams) >= 1 Then
                    strParams = Split(Trim(Mid(strSubParts(0),6)) & "\", "\")
                    strTemp = strParams(0)
                    strResult = strResult & "<a href=""mailto:" & strParams(0) & """>" & Util.HtmlEncode(Util.FirstNonEmpty(strParams(1), strTemp)) & "</a>" & Util.HtmlEncode(strSubParts(1))
                Else
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                End If
            Case "movie"
                strParams=Split(Trim(Mid(strSubParts(0),7))," ")
                If Ubound(strParams)  < 2 Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                ElseIf Trim(strParams(0)) = "" Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                Else
                    strParams = Split(Trim(Mid(strSubParts(0),7)),"""")
                    strParam = ""
                    For j = 0 to Ubound(strParams) Step 2
                        strTemp = Trim(strParams(j))
                        If strTemp <> "" Then strParam = strParam & "|" & Join(split(strTemp," "), "|")
                        If j<Ubound(strParams) Then strParam = strParam & "|" & strParams(j+1)
                    Next
                    strArgs = Split(strParam & "|||||||||", "|")
                    strTemp=Split(Trim(Mid(strSubParts(0),7)) & "       "," ")
                    If Left(strTemp(1),1) <> """" Then
						strArgs(2) = Session("oPicIndex")(strArgs(2)).Path.Report
					Else
						If Not Instr(strArgs(2), ":") > 0 Then strArgs(2) = ReportGenerator.Document.BasePath & strArgs(2) ' relative link
						On Error Resume Next
						Dim oFso
						Set oFso = CreateObject("Scripting.FileSystemObject")
						On Error GoTo 0
						If  oFso.FileExists(strArgs(2)) Then
							strFile = strArgs(2)
							strArgs(2) = "media/" & Session("UUID") & "_" & Util.HrefEncode(oFso.GetFile(strArgs(2)).Name) ' make filename valid and unique
							Session("UUID") = Session("UUID") + 1
							ReportGenerator.FileUpload strFile, Util.UrlDecode(strArgs(2))
						End If
					End If
                    strResult = strResult & Util.FormatPhrase(Dic("PhMovie_" & strArgs(1)), strArgs(2), strArgs(3), strArgs(4), strArgs(5), strArgs(6), strArgs(7), strArgs(8), strArgs(9)) & Util.HtmlEncode(strSubParts(1))
                End If
            Case "phrase"
                strParams = Split(Trim(Mid(strSubParts(0),7)),"""")
                If Ubound(strParams)  < 2 Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                ElseIf Trim(strParams(0)) <> "" Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, strSubParts(0))
                Else
                    strTemp = ""
                    For j = 1 to Ubound(strParams) Step 2
                        strTemp = strTemp & ",""" & strParams(j) & """"
                        strParam = Trim(strParams(j+1))
                        If strParam <> "" Then strTemp = strTemp & ",CustomTag(obj, """ & Join(split(strParam, " "), """),CustomTag(obj, """) & """)"
                    Next
                    On Error Resume Next
                    strResult = strResult & eval("Util.FormatPhrase(" & Mid(strTemp,2) & ")")
                    If Err Then Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), strClass, "Util.FormatPhrase(" & Mid(strTemp,2) & ")" & ": " & Err.Description)
                    Err.Clear
                    On Error Goto 0
                End If
            Case Else Report.LogError Util.FormatString(ConfigMessage("ErrorTextToken"), strClass, strParams(0))
        End Select
    Next
    StrFormatText = strResult
End Function

Function StrHtmlHighlightName(strName)
    Dim strMarker, nStart, nEnd
    strMarker = Session("MarkerFirstName")
    nStart = Instr(strName, strMarker)
    If nStart > 0 Then
        nEnd = Instr(nStart, strName, " ")
        If nEnd > 0 Then
            StrHtmlHighlightName = Util.HtmlEncode(Left(strName, nStart-1)) & "<span class='namehighlight'>" & Util.HtmlEncode(Mid(strName, nStart+1, nEnd-nStart-1)) & "</span>" & StrHtmlHighlightName(Mid(strName,nEnd))
        Else
            StrHtmlHighlightName = Util.HtmlEncode(Left(strName, nStart)) & "<span class='namehighlight'>" & Util.HtmlEncode(Mid(strName, nStart+1)) & "</span>"
        End If
    Else
        StrHtmlHighlightName = Util.HtmlEncode(strName)
    End If
End Function

Function StrHtmlSubSection(obj, strSectionType, strContent)
    If (strContent <> "") Then
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present
        StrHtmlSubSection = Util.FormatString("<ul class='xT std'>" & vbNewline & _
                         "  <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'><h4 class='xT-i inline'>{&t}</h4>" & vbNewline & _
                         "    <ul class='xT-h'>" & vbNewline & _
                         "     <li>{}" & vbNewline & _
                         "     </li>" & vbNewline & _
                         "    </ul>" & vbNewline & _
                    "  </li>" & vbNewline & _
                    "</ul>" & vbNewline, _
                    Util.IfElse(Session("fCollapseReferences"), "c", "o"), strSectionType, strContent)
    End If
End Function

Function StrJavaScriptEncode(strValue)
' do proper JavaScript encoding, unlike Util.JavaScriptEncode
    StrJavaScriptEncode = Replace(Replace(Replace(Replace(Replace(strValue, "\", "\\"),"""","\"""),"'","\'"), vbTab, "\t"), vbLineFeed, "\n")
End Function


Function StrMergeText(obj, strMode, strParams, nFirst)
    Dim strResult, i
    For i=nFirst To Ubound(strParams) Step 2
        If strParams(i) <> "" Then strResult = strResult & StrMergeTags(obj, strMode, Split(strParams(i)))
        If i < Ubound(strParams) Then strResult = strResult & strParams(i+1)
    Next
    StrMergeText = strResult
End Function

Function StrMergeTags(obj, strMode, strParams)
'=============================================
    Dim strResult, strCode
    For Each strCode In strParams
        If strCode <> "" Then strResult = strResult & StrFormatText(obj, strMode & CustomTag(obj, strCode))
    Next
    StrMergeTags = strResult
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
Function StrNormalizeSpace(strText)
    Dim oMatches, oMatch, strResult, oRegExP
    strResult = strText
    Set oRegExp = New RegExp
    oRegExp.Global = True
    oRegExp.Pattern = "\s\s+"
    Set oMatches = oRegExp.Execute(strText)
    For Each oMatch in oMatches
        strResult = Replace(strResult, oMatch.Value, Left(oMatch.Value,1))
    Next
    StrNormalizeSpace = strResult
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
Function StrParseText(strRawText, fUseLangShowOthers)
'=============================
' check for language specific sections, making skin language sections visible by default
Dim arrParts, arrSubParts, strResult, i, strLang, strTemp, strText, strSep, fLangShowOthers
fLangShowOthers = False
If fUseLangShowOthers Then fLangShowOthers = Session("LangShowOthers")
If VarType(fLangShowOthers) <> vbBoolean Then
   Report.LogError "Error: Configuration parameter 'LangShowOthers' is invalid " & VarType(fLangShowOthers) & " / " & Session("LangShowOthers")
   fLangShowOthers = False
End If
If Instr(strRawText,"<?off?>") = 1 Then
    StrParseText = strRawText
    Exit Function
End If
strText=Replace(strRawText,"¶",vbLf)
If Instr(strText,"{?") > 0 Or Instr(strText,"{¿") > 0 Then
    arrParts = GetLanguageParts(strText)
    strResult = ""
    If fLangShowOthers And (UBound(arrParts) > 3) Then    ' swap if needed to put skin language first
        For i=2 To UBound(arrParts) Step 2
            If arrParts(i) <> "" Then Exit For    ' stop if not consecutive phrases
            If i < UBound(arrParts) Then
                If Left(arrParts(i+1), Instr(arrParts(i+1),":")-1) = Session("ReportLanguage") Then
                    strTemp = arrParts(1)
                    arrParts(1) = arrParts(i+1)
                    arrParts(i+1) = strTemp
                End If
            End If
        Next
    End If
    For i=0 To UBound(arrParts) Step 2
        strResult = strResult & arrParts(i)
        If i < UBound(arrParts) Then
            If Instr(arrParts(i+1),":") < 3 Then
                Report.LogError Util.FormatString(ConfigMessage("ErrorLangMarkup"), strText)
                Exit Function
            Else
                arrSubParts = Split(arrParts(i+1),":")
                If arrSubParts(0) = Session("ReportLanguage") Then
                    strResult = strResult & Mid(arrParts(i+1), 2+Len(arrSubParts(0)))
                ElseIf Session("LangShowOthers") Then
                    strLang = CustomTag(Null, "_" & arrSubParts(0))
                    If strLang = "" Then strLang = arrSubParts(0)
                    strResult=strResult & " (<?html?><span class='langtoggle' onclick=""javascript:this.nextSibling.style.display = (this.nextSibling.style.display == 'none' ? 'inline' : 'none');"">" & strLang & "</span><span style='display:none' onclick=""javascript:this.style.display = (this.style.display == 'none' ? 'inline' : 'none');"">: " & Util.HtmlEncode( Mid(arrParts(i+1), 2+Len(arrSubParts(0)))) & "</span><?text?>) "
                End If
            End If
        End If
    Next
    If strResult = "" And Session("LangShowDefault") <> "" Then
        For i=0 To UBound(arrParts) Step 2
            If i < UBound(arrParts) Then
                strResult = strResult & arrParts(i)
                arrSubParts = Split(arrParts(i+1),":")
                If arrSubParts(0) = Session("LangShowDefault") Then strResult = strResult & Mid(arrParts(i+1), 2+Len(arrSubParts(0)))
            End If
        Next
    End If
    StrParseText = strResult
Else
    StrParseText = strText
End If
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
Function StrPlainText(obj, strRawText)
'==================================
    ' check for language specific sections, making skin language sections visible by default
    Dim arrParts, arrSubParts, strResult, i, strLang, strText
    strText=Replace(strRawText,"¶",vbLf)

    If Instr(strText,"{?") > 0 Or Instr(strText,"{¿") > 0 Then
        arrParts = GetLanguageParts(strText)
        strResult = ""
        If Session("LangShowOthers") And (UBound(arrParts) > 3) Then    ' swap if needed to put skin language first
            For i=2 To UBound(arrParts) Step 2
                If arrParts(i) <> "" Then Exit For    ' stop if not consecutive phrases
                If i < UBound(arrParts) Then
                    If Left(arrParts(i+1), Instr(arrParts(i+1),":")-1) = Session("ReportLanguage") Then
                        strTemp = arrParts(1)
                        arrParts(1) = arrParts(i+1)
                        arrParts(i+1) = strTemp
                    End If
                End If
            Next
        End If
        For i=0 To UBound(arrParts) Step 2
            strResult = strResult & arrParts(i)
            If i < UBound(arrParts) Then
                If Instr(arrParts(i+1),":") < 3 Then
                    Report.LogError Util.FormatString(ConfigMessage("ErrorLangMarkup"), strText)
                    Exit Function
                Else
                    arrSubParts = Split(arrParts(i+1),":")
                    If arrSubParts(0) = Session("ReportLanguage") Then
                        strResult = strResult & Mid(arrParts(i+1), 2+Len(arrSubParts(0)))
                    ElseIf Session("LangShowOthers") Then
                        strLang = CustomTag(Null, "_" & arrSubParts(0))
                        If strLang = "" Then strLang = arrSubParts(0)
                        strResult=strResult & " (" & strLang & ": " & Mid(arrParts(i+1), 2+Len(arrSubParts(0))) & ") "
                    End If
                End If
            End If
        Next
        If strResult = "" And Session("LangShowDefault") <> "" Then
            For i=0 To UBound(arrParts) Step 2
                If i < UBound(arrParts) Then
                    strResult = strResult & arrParts(i)
                    arrSubParts = Split(arrParts(i+1),":")
                    If arrSubParts(0) = Session("LangShowDefault") Then strResult = strResult & Mid(arrParts(i+1), 2+Len(arrSubParts(0)))
                End If
            Next
        End If
    Else
        StrResult = strText
    End If
    '
    ' Remove text 'processing instructions' embedded in Comment, Descriptions, Custom Tags etc.
    '
    Dim strPart, strParts, strSubParts, strParams, strUID, strMode, strTemp
    If strResult = "" or strResult = "<?text?>" Then
        StrPlainText = ""
        Exit Function
    End If
    strMode = "<?text?>"
    strParts = split(StrParseText(strResult, True), "<?")
    strSubParts = split(strParts(0),StrDicOrTag("", "Private"))
    strResult = Util.FormatString("{&x}",strSubParts(0))        ' get any leading text less any old format private section
    For i=1 to Ubound(strParts)
        strSubParts = split(strParts(i), "?>")
        If Ubound(strSubParts) <> 1 Then Report.LogError Util.FormatString(ConfigMessage("ErrorTextFormat"), obj, strText)
        strParams = split(strSubParts(0), " ")
        Select Case LCase(strParams(0))
            Case "text"
                strResult = strResult & Util.FormatString("{&x}",strSubParts(1))
                strMode = "<?" & strSubParts(0) & "?>"
        End Select
    Next
    StrPlainText = strResult

End Function

Function StrPlainName(strName)
    StrPlainName = Replace(strName, Session("MarkerFirstName"), "")
End Function

Function StrNewUID()
'===================
    StrNewUID = Session("UUID") & ""
    Session("UUID") = Session("UUID") + 1
End Function

Function CustomTag(obj,tag)
'==========================
On Error Resume Next
If Not IsNull(obj) Then        ' ordinary custom tag
    CustomTag = obj.TagValue(tag)
Else                ' Document custom Tag
    CustomTag = Session("oGlobal").selectSingleNode(tag).text
End If
End Function

Function GetDate(oDate)
'======================
' format date for sorting
  Dim Year, NYear
    If Not Util.IsNothing(oDate) Then
       Year = CustomDate(oDate).Year
       NYear = Right("0000" & CustomDate(oDate).NYear, 4)
     GetDate = Util.FormatPhrase(Dic("FmtDateSort"), Util.IfElse(Right(Year,2)="BC","-","+") & NYear, Util.IfElse(oDate.Month<>"",Right("0" & oDate.Month,2),"00"), Util.IfElse(oDate.Day<>"",Right("0" & oDate.Day,2),"00"))
  Else
     GetDate = "+0000"
  End If
End Function

Function GetDateString(oDate)
    If Not Util.IsNothing(oDate) Then
         GetDateString = CustomDate(oDate).ToString("")
  Else
         GetDateString = ""
  End If
End Function

Function GetLanguageParts(strText)
'=================================
' Break a Multiple Language formatted field into an array its component parts
    Dim oMatches, oRegExp, i
    Set oRegExp = New RegExp
    If  Instr(strText,"{¿") > 0 Then            ' type 1 Multiple Language Format
        oRegExp.Global = True
        oRegExp.Pattern = "\{¿[A-Z]{2}:([^¿]|¿[^\}])*¿\}"
        Set oMatches = oRegExp.Execute(strText)
        i = oMatches.Count
        oRegExp.Pattern="\{¿"
        Set oMatches = oRegExp.Execute(strText)
        If i <> oMatches.Count Then Report.LogError Util.FormatString(ConfigMessage("ErrorLangMarkup"), strText)
        oRegExp.Pattern="¿\}"
        Set oMatches = oRegExp.Execute(strText)
        If i <> oMatches.Count Then Report.LogError Util.FormatString(ConfigMessage("ErrorLangMarkup"), strText)
        GetLanguageParts = split(Replace(strText,"¿}","{¿"),"{¿")
    Else                            ' type 2 Multiple Language Format
        oRegExp.Global = True
        oRegExp.Pattern = "\{\?[A-Z]{2}:[^\}]*\}"
        Set oMatches = oRegExp.Execute(strText)
        i = oMatches.Count
        oRegExp.Pattern="\{\?"
        Set oMatches = oRegExp.Execute(strText)
        If i <> oMatches.Count Then Report.LogError Util.FormatString(ConfigMessage("ErrorLangMarkup"), strText)
        oRegExp.Pattern="\}"
        Set oMatches = oRegExp.Execute(strText)
        If i <> oMatches.Count Then Report.LogError Util.FormatString(ConfigMessage("ErrorLangMarkup"), strText)
        GetLanguageParts = split(Replace(strText,"}","{?"),"{?")
    End If
End Function

Function MatchDate(o, p, obj)
'============================
    If o.Id = p.ID And obj.ToStringNarrative <> "" Then
        Set MatchDate = obj
    Else
        Set MatchDate = Nothing
    End If
End Function

Function GetFile(path, localpath)
        If LCase(Left(path,5)) <> "http:" Then
            GetFile = path
        Else
            Dim strTempFldr
            Err.Clear
            On Error Resume Next
            oHttp.Open "Get", path & "?Now=" & Now, False ' add param Now to avoid cache (hopefully!)
            oHttp.Send
            If Err.Number <> 0 Then
                Report.LogError Util.FormatString(ConfigMessage("ErrorHttpGet"),(Err.Number And (256*256-1)), Err.Description, path)
                Exit Function
            End If
            oBinaryStream.Write oHttp.ResponseBody
            If localpath = "" Then
                strTempFldr = oFso.GetSpecialFolder(2).Path & "\"
                localpath = oFso.GetTempName
                localpath = strTempFldr & Mid(localpath, 1, InstrRev(localpath, ".")-1) & Mid(path, InstrRev(path, "."))
            End If
            oBinaryStream.SaveToFile localpath, 2
            GetFile = localpath
        End If
End Function

Function PicResize(p, strNewPath, nMaxWidth, nMaxHeight, fFixRatio, fForce)
    Dim strPath, nWidth, hHeight, nDPI, nRatioPic, nRatioLab, nErrorCode, strDim
    PicResize = ""
    If Not Util.IsNothing(p.Cache.DPI) Then
        strDim = PicRedim(p, nMaxWidth, nMaxHeight, fFixRatio, fForce)
        nHeight = Util.GetHeight(strDim) + 0
        nWidth = Util.GetWidth(strDim) + 0
        nDPI = p.Cache.DPI
        If Not IsNumeric(nDPI) Then nDPI = 0
        If nDPI = 0 Or nDPI > Session("ThumbnailDpi") Then nDPI = Session("ThumbnailDpi")
        strPath = GetFile(p.Path, "")
        nErrorCode = oShell.Run("""" & Session("IrfanViewPath") & """ """ & strPath & """ /resize=(" & nWidth & "," & nHeight & ") /resample /dpi=(" & nDPI & "," & nDPI & ")" & Util.IfElse(fFixRatio, " /aspectratio", "") & " /jpgq=" & Session("ThumbnailQuality") & " /convert=""" & strNewPath & """", 0, True)
        if nErrorCode <> 0 Then
            Report.LogWarning Util.FormatString(ConfigMessage("ErrorPictureConvert"), nErrorCode, p.Path, strPath)
        Else
            PicResize = nWidth & "x" & nHeight
        End If
    Else
        Report.LogError Util.FormatString(ConfigMessage("ErrorNoCachedSize"), strOldPath)
    End If
End Function

Function PicRedim(p,nMaxWidth, nMaxHeight, fFixRatio,fForce)
    Dim nHeight, nWidth, nRationPic, nRatioLab
    nHeight = Util.GetHeight(strDim) + 0
    nWidth = Util.GetWidth(p.Cache.Dimension) + 0
    nHeight = Util.GetHeight(p.Cache.Dimension) + 0
    If fForce Or nWidth > nMaxWidth Or nHeight > nMaxHeight Then
        nRatioPic = nWidth / nHeight
        nRatioLab = nMaxWidth / nMaxHeight
        nWidth=nMaxWidth
        nHeight=nMaxHeight
        If fFixRatio = True Then
            If nRatioLab > nRatioPic Then    ' wider
                nHeight = nMaxHeight
                nWidth = Round(nHeight * nRatioPic, 0)
            Else
                nWidth = nMaxWidth
                nHeight = Round(nWidth / nRatioPic, 0)
            End If
        End If
    End If
    PicRedim = nWidth & "x" & nHeight
End Function

Function CustomDate(oDate)
' Check date for text only date e.g. 400 BC or 'nothing; 
' If true return a CustomDate object else return gnoDate object
         Dim DateObject
         Set DateObject = New CustomDateClass
         If Not DateObject.Test(oDate) Then
            Set CustomDate = oDate                   ' standard GenoPro Date
         Else
             DateObject.Parse oDate
             Set CustomDate = DateObject
         End If
End Function

Class CustomDateClass
      Private oDate_, Approximation_, Year_, NYear_, Era_, Value_
      Private BCDate, oMatches, oMatch

      Private Sub Class_Initialize()
         Set BCDate = New RegExp
         BCDate.IgnoreCase = True
         BCDate.Pattern = "(|~|<|>)(\d+) (BC)"
      End Sub

      Public Function Test(oDate)
        If oDate Is Nothing Then
            Test = True
        Else
            Test = BCDate.Test(oDate) Or oDate Is Nothing
        End If
      End Function

      Public Sub Parse(oDate)
        Set oDate_ = oDate
        If oDate Is Nothing Then
            Value_ = ""
            Approximation_ = ""
            Era_ = ""
            Year_ = ""
        Else
            Value_ = oDate
            Set oMatches = BCDate.Execute(oDate_)
            Set oMatch = oMatches(0)
            Approximation_ = oMatch.SubMatches(0)
            Era_ = oMatch.SubMatches(2)
            Year_ = oMatch.SubMatches(1)
        End If
      End Sub

      Public Default Property Get Value
        Value=Value_
      End Property
      Public Property Get Approximation
        Approximation=Approximation_
      End Property
      Public Property Get Year
        Year=Year_ & " " & Era_
      End Property
      Public Property Get Era
        Era=Era_
      End Property
      Public Property Get NYear
        NYear=Int("0" & Year_)
      End Property
      Public Property Get NMonth
        NMonth=0
      End Property
      Public Property Get NDay
        NDay=0
      End Property
      Public Property Get ToStringNarrative
        ToStringNarrative=DateFormat(GetDateFormat("Narrative", "YG"))
      End Property
      Public Function ToString(fmt)
        ToString=DateFormat(Util.IfElse(fmt<>"",Replace(fmt, "yyyy", "yyyy GG") ,GetDateFormat("Default", "YG")))
      End Function
      Private Function DateFormat(fmt)
        Dim start, fini, prepositions, index
        If Value_ = "" Then
            DateFormat = ""
        Else
            start=Instr(fmt,"[")
            fini=Instr(fmt,"]")
            prepositions = split(Mid(fmt, start+1,fini-start),"|")
            index=0
            If Approximation_ <> "" Then index=Instr("~<>", Approximation_)
            DateFormat=" " & Trim(Replace(Replace(Left(fmt,start-1) & prepositions(index) & Mid(fmt,fini+1),"yyyy",Year_), "GG", Era_))
        End If
      End Function
End Class

Sub GoogleAnalytics()
If Session("GoogleAnalyticsAccount") <> "" Then
   Report.WriteLn "<script type=""text/javascript"">"
   Report.WriteLn "var _gaq = _gaq || [];"
   Report.WriteFormattedLn "_gaq.push(['_setAccount', '{}']);", Session("GoogleAnalyticsAccount")
   If Session("GoogleAnalyticsDomain") <> "" Then
      Report.WriteFormattedLn "_gaq.push(['_setDomainName', '{}']);", Session("GoogleAnalyticsDomain")
   End If
   Report.WriteLn "_gaq.push(['_trackPageview']);"
   Report.WriteLn "(function () {"
   Report.WriteLn "    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;"
   Report.WriteLn "    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';"
   Report.WriteLn "    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);"
   Report.WriteLn "})();"
   Report.WriteLn "</script>"
End If
End Sub

Sub CurvyBoxOpen()
    Report.WriteLn "<div class='curvyboxbackground'>"
    Report.WriteLn " <div class='curvycorners_box'>"
    Report.WriteLn "  <div class='curvycorners_top'><div></div></div>"
    Report.WriteLn "   <div class='curvycorners_content'>"
End Sub

Sub CurvyBoxClose()
    Report.WriteLn "   </div>"
    Report.WriteLn "  <div class='curvycorners_bottom'><div></div></div>"
    Report.WriteLn " </div>"
    Report.WriteLn "</div>"
End Sub

