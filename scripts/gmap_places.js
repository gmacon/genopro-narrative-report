<%[@ IncludeFile "Code/Util.vbs" ]%>
<%[@ IncludeFile "Code/Lang.vbs" ]%>
<%[If (Session("Book")) Then Report.AbortTemplate]%>
<%[

' generate Google Maps marker data for each Place with coordinates set as a javascript object.

Dim p, pCnt, strSep
If Session("fGoogleMapsOverview") Then
	Report.WriteLn "gMapData = {""markers"": ["
	strSep=""
	For Each p In Places
		pCnt = p.References.Count
		If pCnt > 0 And p.Latitude <> "" And p.Longitude <> "" Then
			If strSep <> "" Then
				Report.WriteLn strSep
			Else
				strSep=","
			End If
			Report.Write "{"
			Report.WriteFormatted """lat"": ""{&j}"", ""lng"": ""{&j}"", ""n"":{}, ""html"":""<b>{&j}</b><div class='infoWindow{}'>", p.Latitude, p.Longitude, pCnt, JoinPlaceNames(p, p.Name, true), Util.IfElse(pCnt>10," infoScroll","")
			WriteHtmlReferenceList p
			Report.WriteFormatted "</div>"",""label"":""{&j} ({})""", JoinPlaceNames(p, p.Name, true), Dic.PlurialCount("Reference",pCnt)
			Report.Write "}"
		End If
	Next
	Report.WriteLn "]"
	Report.WriteLn "}"
End If
Sub WriteHtmlReferenceList(o)
	Dim collReferences, oParent, d, oDate, oDateYears, strSep
	Set oParent = Nothing
	Set collReferences = o.References
	Set oDateYears=Util.NewStringDictionary
	Select Case o.Class
	Case  "SourceCitation"
		Set oParent = o.Parent
	Case  "Place"
		Set oParent = o.Parent
		Set collReferences = Util.NewDataSorter
		For Each r in o.References
			Set oDate = Nothing
			Select Case r.Class
			Case "Individual"
				Set oDate = MatchDate(o, r.Birth.Place, r.Birth.Date)
				If Util.IsNothing(oDate) Then Set oDate = MatchDate(o, r.Birth.Baptism.Place, r.Birth.Baptism.Date)
				If Util.IsNothing(oDate) Then Set oDate = MatchDate( o, r.Death.Place, r.Death.Date)
				If Util.IsNothing(oDate) Then Set oDate = MatchDate(o, r.Death.Funerals.Place, r.Death.Funerals.Date)
				If Util.IsNothing(oDate) Then Set oDate = MatchDate(o, r.Death.Disposition.Place, r.Death.Disposition.Date)
			Case "Education", "Occupation", "Contact"
				If Not r.DateStart.ToStringNarrative = "" Then
					Set oDate = r.DateStart
				Else
					Set oDate = r.DateEnd
				End If
			Case "Marriage"
				Set oDate = MatchDate(o, r.Place, r.Date)
				If Util.IsNothing(oDate) Then Set oDate = MatchDate(o, r.Divorce.Place, r.Divorce.Date)
			End Select
			collReferences.Add r, GetDate(oDate), r.Class
			If Not Util.IsNothing(oDate) Then oDateYears.Add r.Class & r.ID, oDate.Year
		Next
		collReferences.SortByKey()
		Set collReferences = collReferences.ToGenoCollection
	End Select
	If collReferences.Count > 0 Or Not Util.IsNothing(oParent) Then
		If Not Util.IsNothing(oParent) Then
			Session("ReferencesStart") = -1
			Report.WriteFormattedBr "{&t}: {}", Dic(o.Class & "_Parent"), oParent
		End If
		For Each d In collReferences
			Session("ReferencesStart") = -1
			strHref = d.Href
			If (strHref <> "") Then
				strHref = d
				If strHref = "" Then StrHref = "'" & d.Href & "'"
				If d.Class = "Individual" And o.Class = "Place" Then
					strHref = strHref & Util.FormatPhrase(Dic("FmtDatesIndividual"), _
						Util.IfElse(o.ID = d.Birth.Place.ID, d.Birth.Date.ToString,""), _
						Dic("PhBC_" & d.Birth.CeremonyType.ID), _
						Util.IfElse(o.ID = d.Birth.Baptism.Place.ID, d.Birth.Baptism.Date.ToString,""), _
						Util.IfElse(o.ID = d.Death.Place.ID, d.Death.Date.ToString,""), _
						Util.IfElse(o.ID = d.Death.Funerals.Place.ID, d.Death.Funerals.Date.ToString,""), _
						Util.IfElse(o.ID = d.Death.Disposition.Place.ID, d.Death.Disposition.Date.ToString,""))
				End If
			Else
				strSep = ""	
				Select Case d.Class
					Case "Occupation", "Education"
						strHref = d & " "
						If d.Class = "Occupation" Then strHref = d.Session("Title") & " "
						For Each dobj in d.References
							strHref = strHref & strSep & dobj
							strSep = ", "
						Next
						If o.Class = "Place" And o.ID = d.Place.ID Then
							strHref = strHref & Util.FormatPhrase(Dic("FmtDatesFromTo"), _
								d.DateStart.ToString, _
								d.DateEnd.ToString)
						End If
					Case "PedigreeLink"
						strHref = d.PedigreeLink & " " & strSep & d.individual
					Case "Marriage", "Contact"
						If d.Class = "Contact" And d.Type.ID <> "" Then strHref = d.Type & " "
						For Each dobj in d.References
							strHref = strHref & strSep & dobj
							strSep = ", "
						Next
						If o.Class = "Place" Then
							If d.Class = "Contact" And o.Id = d.Place.ID Then
								strHref = strHref & Util.FormatPhrase(Dic("FmtDatesFromTo"), _
									d.DateStart.ToString, _
									d.DateEnd.ToString)
							Else
								strHref = strHref & Util.FormatPhrase(Dic("FmtDatesUnion"), _
									Util.IfElse(o.ID = d.Place.ID, d.Date.ToString,""), _
									Util.IfElse(o.ID = d.Divorce.Place.ID, d.Divorce.Date.ToString,""))
							End If
						End If
					Case Individual
						strHref = d.Session("HlinkNN")

					Case Else
						strHref = d & ""
				End Select
			End If
			Select Case d.Class
			Case "SocialRelationship", "EmotionalRelationship"
				Report.WriteFormatted "{&t}: ", Dic(d.Class)
				WriteHtmlRelationship d, "","","", "", False, False
				Report.WriteBr
			Case "Occupation"
			     If d.Session("Event") <> "" Then
				    Report.WriteFormattedBr "{} {}: {&j}", oDateYears.KeyValue(d.Class & d.ID), d.Session("EventName"), strHref
			     Else
				    Report.WriteFormattedBr "{} {}: {&j}", oDateYears.KeyValue(d.Class & d.ID), Dic(d.Class), strHref
			     End If
			Case Else
				Report.WriteFormattedBr "{} {}: {&j}", oDateYears.KeyValue(d.Class & d.ID), Dic(d.Class), strHref
			End Select
		Next
	End If
End Sub
]%>
