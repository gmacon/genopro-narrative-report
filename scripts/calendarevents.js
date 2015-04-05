<%['  Module: calendarevents.js
'     Version: 2013.07.30
'     Creates a JSON-like javascript file for the Narrative Report Calendar page
'     with birthday and wedding anniversary events
'     based on Javascript Event Calendar http://calendar.pikesys.com
 ]%>
<%[If Session("Book") Then Report.AbortTemplate]%>
<%[@ IncludeFile "Code/Util.vbs" ]%>
<%[@ IncludeFile "Code/Lang.vbs" ]%>
<%[



If Not Session("Calendar") Then Report.AbortPage

' Create an calendar data file

Dim i, f, nAge, nYear, strOrdinal, oDicCache, strType

Set oDicCache = Session("DicCache")

Report.WriteFormattedln "months=['',{}];", oDicCache("Months")
Report.WriteFormattedln "weekdays=[{}];", oDicCache("Weekdays")
Report.WriteFormattedln "SpecialDay={};", oDicCache("FirstDay")
Report.Writeln "oDic = {"
Report.WriteFormattedLn "'view':'{&j}',", StrDicExt("CalendarView","","View","","")
Report.WriteFormattedLn "'changes':'{&j}',", StrDicExt("CalendarChanges","","Apply changes","","")
Report.WriteFormattedLn "'print':'{&j}',",StrDicExt("CalendarPrint","","Printer friendly","","")
Report.WriteFormattedLn "'export':'{&j}',", StrDicExt("CalendarExport","","Export events displayed","","")
Report.WriteFormattedLn "'filter':'{&j}',", StrDicExt("CalendarFilter","","Filter events","","")
Report.WriteFormattedLn "'jump':'{&j}'}", StrDicExt("CalendarJump","","Jump to month","","")

For Each i in Individuals
	If Not i.IsDead And i.Birth.Date.NDay > 0 And i.Birth.Date.NMonth > 0 Then
		nAge = i.Age.NYears
		nYear = Year(Date)
		If (i.Birth.Date.NMonth < Month(Date)) Or (i.Birth.Date.NMonth = Month(Date) And i.Birth.Date.NDay <= Day(Date)) Then
			Report.WriteFormattedLn "AddEvent({}{}{},'{&j}','{&j}','','{&j}','','','','');", nYear, Mid(i.Birth.Date.NMonth + 100,2), Mid(i.Birth.Date.NDay+100,2), Util.FormatPhrase(StrDicExt("PhCalendarEvent","","{0} [{?6}[{?1}{3}{1}{4} {2}][{?!1}{5}]][{?!6}{2}]","",""), StrPlainName(i.Session("NameFull")), StrRank(nAge, i.Gender.ID), StrDicExt("Birthday","","birthday","",""), StrDicExt("CalendarFix1","","<bdo dir=\'ltr\'>","",""), StrDicExt("CalendarFix2","","</bdo>","",""), StrDicExt("Born","","born","",""), i.Birth.Date.NYear<>0), StrDicExt("Birthday","","birthday","",""), i.Href
			nYear = nYear + 1
		End If
		nAge = i.Age.NYears + 1
		Report.WriteFormattedLn "AddEvent({}{}{},'{&j}','{&j}','','{&j}','','','','');", nYear, Mid(i.Birth.Date.NMonth + 100,2), Mid(i.Birth.Date.NDay+100,2), Util.FormatPhrase(StrDicExt("PhCalendarEvent","","{0} [{?6}[{?1}{3}{1}{4} {2}][{?!1}{5}]][{?!6}{2}]","",""), StrPlainName(i.Session("NameFull")), StrRank(nAge, i.Gender.ID), StrDicExt("Birthday","","birthday","",""), StrDicExt("CalendarFix1","","<bdo dir=\'ltr\'>","",""), StrDicExt("CalendarFix2","","</bdo>","",""), StrDicExt("Born","","born","",""), i.Birth.Date.NYear<>0), StrDicExt("Birthday","","birthday","",""), i.Href
	End If
Next

For Each f in Families
  ' 2013.07.30 exclude parent-less familes see http://support.genopro.com/Topic32062.aspx
	If f.AreTogether And f.Marriage.Date.NDay > 0 And f.Marriage.Date.NMonth > 0  And f.Parents.Count > 0Then
		nYear = Year(Date)
		nAge = nYear - f.Marriage.Date.NYear
		strType = ""
		' If f.Unions.Count > 0 Then strType = f.Unions(f.Unions.Count-1).Type
		strType = Util.FirstNonEmpty(strType, Dic("Married"))
		If (f.Marriage.Date.NMonth < Month(Date)) Or (f.Marriage.Date.NMonth = Month(Date) And f.Marriage.Date.NDay <= Day(Date)) Then
			Report.WriteFormattedLn "AddEvent({}{}{},'{&j}','{&j}','','{&j}','','','','');", nYear, Mid(f.Marriage.Date.NMonth + 100,2), Mid(f.Marriage.Date.NDay+100,2),Util.FormatPhrase(Dic("PhCalendarEvent"), f.Session("Name"), StrRank(nAge, ""), Dic("Anniversary"), Dic("CalendarFix1"), Dic("CalendarFix2"), strType, f.Marriage.Date.NYear<>0), Dic("Anniversary") , f.Href
			nYear = nYear + 1
			nAge = nAge + 1
		End If
		Report.WriteFormattedLn "AddEvent({}{}{},'{&j}','{&j}','','{&j}','','','','');", nYear, Mid(f.Marriage.Date.NMonth + 100,2), Mid(f.Marriage.Date.NDay+100,2),Util.FormatPhrase(Dic("PhCalendarEvent"), f.Session("Name"), StrRank(nAge, ""), Dic("Anniversary"), Dic("CalendarFix1"), Dic("CalendarFix2"), strType, f.Marriage.Date.NYear<>0), Dic("Anniversary") , f.Href
	End If
Next

Function StrRank(nAge, strGender)
	Dim strFmt, nUnit
	If nAge = 0 Then Exit Function
	strFmt = Dic.Peek("_OrdinalFormat_" & nAge & "_" & strGender)
	If strFmt = "" Then strFmt = Dic.Peek("_OrdinalFormat_" & nAge)
	nUnit = Right(nAge, 1)
	If strFmt = "" Then strFmt = Dic.Peek("_OrdinalFormat_x" & nUnit & "_" & strGender)
	If strFmt = "" Then strFmt = Dic.Peek("_OrdinalFormat_x" & nUnit)
	If strFmt = "" Then strFmt = Dic.Lookup2("_OrdinalFormat_" & strGender, "_OrdinalFormat_")
	StrRank = Util.FormatString(strFmt, nAge)
End Function

]%>
