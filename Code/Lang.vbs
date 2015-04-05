' Routines to display the narrative information about an individual.
' Those routines are used for generating individual and family reports.
'
' HISTORY
' Aug-2005        GenoPro            Creation.
' Sep-2005 - Mar 2008    Ron            Development & Maintenance

'===========================================================
' Return Relative Pronoun (his/her/its) according to gender of individual in lowercase
Function PnR(obj)
    Dim strGender
    On Error Resume Next
    StrGender = obj.Gender.ID
    If strGender ="" and obj.Class = "SocialEntity" Then strGender = "N"
    PnR = StrDicExt("PnR_" & strGender,"", Dic("PnR_"),"","")
End Function

'===========================================================
' Return the Personal Pronoun (he/she/it) according to gender of individual, with or without Upper case first letter
Function PnP(obj)
    Dim strGender
    On Error Resume Next
    StrGender = obj.Gender.ID
    If strGender ="" and obj.Class = "SocialEntity" Then strGender = "N"
    PnP = StrDicExt("PnP_" & strGender,"", Dic("PnP_"),"","")
End Function

'===========================================================
' Return Objective Pronoun (him/her/its) according to gender of individual in lowercase
Function PnO(obj)
    Dim strGender
    On Error Resume Next
    StrGender = obj.Gender.ID
    If strGender ="" and obj.Class = "SocialEntity" Then strGender = "N"
    PnO = StrDicExt("PnO_" & strGender,"", Dic("PnO_"),"","")
End Function


' Determine if two objects, typically two individuals or an individual and a collection, are alive or deceased
Function ToBe2(obj1, obj2)
    Dim nStatistics1, nStatistics2
    nStatistics1 = Util.GetStatisticsForIndividuals(obj1)
    nStatistics2 = Util.GetStatisticsForIndividuals(obj2)
    If (nStatistics1 < 4 AND nStatistics2 < 4) Then
        ' single or multiple individuals alive (is/are)
        ToBe2=StrVerb("ToBe", True, nStatistics1 <= 1 AND nStatistics2 <= 1, "", "")
    Else
        ' single or multiple individuals deceased (was/were)
        ToBe2=StrVerb("ToBe", False, nStatistics1 < 8 AND nStatistics2 < 8, "", "")
    End If
End Function

' Determine if an individual is alive or deceased
Function ToBe(i)
    ToBe = StrVerb("ToBe", i.IsDead = False, True, "", i.Gender.ID)
End Function

' Determine correct form of 'to be' verb for (grand)parents of an individual
' N.B. returns blank if one (grand)parent is alive and one is dead, or only one (grand)parent supplied

Function ToBe2a(father, mother)
    Dim nStatisticsI, nStatisticsF, nStatisticsM
    nStatisticsF = Util.GetStatisticsForIndividuals(father) And 5
    nStatisticsM = Util.GetStatisticsForIndividuals(mother) And 5
    Select Case nStatisticsF + nStatisticsM
    Case 10                 ' both (grand)parents dead (were)
        ToBe2a = StrVerb("ToBe", False, False, "", "")    
    Case 2                    ' both (grand)parents alive (are)
        ToBe2a = StrVerb("ToBe", True, False, "", "")
    Case 1, 5, 6                    'only one (grand)parent or one (grand)parent alive, the other dead
        ToBe2a = ""                ' return blank
    End Select
End Function

' Return the keyword from an individual and a collection and/or another individual.
Function ToHave(i, c)
    Dim nStatistics, strGender
    strGender = i.Gender.ID
    nStatistics = Util.GetStatisticsForIndividuals(i)
    If (nStatistics >= 4) Then
        ToHave = StrVerb("ToHave", False, True, "", strGender)     ' The individual deceased
        Exit Function
    End If
    ' The individual is alive, so check if someone is dead in the collection
    nStatistics = Util.GetStatisticsForIndividuals(c)
    If (nStatistics >= 4) Then
        ToHave = StrVerb("ToHave", False, True, "", strGender)
    Else
        ToHave = StrVerb("ToHave", True, True, "", strGender)    ' Everyone is alive
    End If
End Function

Function GetDicMFU(strKey, strGender)
    GetDicMFU=Util.IfElse(Dic.Peek(strKey & "_" & strGender) <> "",strKey & "_" & strGender , strKey)
End Function

Function IsAlive(i)
  If Util.IsNothing(i) Then
    IsAlive = False
  Else
    IsAlive = Not i.IsDead
  End If
End Function 

Function StrDateSpan(oFrom, oTo)
    Dim strTo, strFrom
    If Not CustomDate(oFrom).ToStringNarrative = "" Then
        If Not CustomDate(oTo).ToStringNarrative = "" Then        ' from & to
            If CustomDate(oFrom).ToStringNarrative = CustomDate(oTo).ToStringNarrative Then
                StrDateSpan = Trim(CustomDate(oFrom).ToStringNarrative)
            Else
                StrDateSpan = Trim(Util.FormatString(GetDateFormat("FromAndTo", ""), _
                StrSubstitute(StrDateToString(oFrom, "From"), Session("RegEx_CDS")), _
                StrSubstitute(StrDateToString(oTo, "To"), Session("RegEx_CDS"))))
            End If
        Else                    ' from only
            StrDateSpan = Trim(StrSubstitute(StrDateToString(oFrom, "Since"), Session("RegEx_CDS")))
        End If
    ElseIf Not CustomDate(oTo).ToStringNarrative = ""  Then        ' to only
        StrDateSpan = Trim(StrSubstitute(StrDateToString(oTo, "Until"), Session("RegEx_CDS")))
    Else
        StrDateSpan = ""
    End If
End Function

Function StrTimeSpan(oDuration)
    If Not oDuration = "" Then
        StrTimeSpan = Trim(StrSubstitute(oDuration, Session("RegEx_CTS")))
    Else
        StrTimeSpan = ""
    End If
End Function

Function StrDateToString(oDate, strFormat)
    Dim strYMD, strFmt, nStart, nEnd, arrApprox
    If oDate.Month = "" Then
        strYMD = "Y"
        StrDateToString = CustomDate(oDate).ToString(GetDateFormat(strFormat, strYMD))
    ElseIf oDate.Day = "" Then
        strYMD = "YM"
        ' deal with 'd' bug and approximation string length bug in GenoPro date formatting
        ' pull out the bit in square brackets and replace with {} then replace {} afterwards with required bit.
        
        strFmt = GetDateFormat(strFormat, strYMD)
        nStart = Instr(strFmt, "[")
        nEnd = Instr(strFmt, "]")
        If nStart > 0 Then
            ArrApprox=Split(Mid(strFmt,nStart+1, nEnd - nStart -1),"|")
            strFmt = Left(strFmt, nStart-1) & "@" &Mid(strFmt,nEnd+1)
            StrDateToString = Replace(oDate.ToString(strFmt),"@",ArrApprox(Instr(" ~<>", CustomDate(oDate).Approximation)-1))
        Else
            StrDateToString = oDate.ToString(strFmt)
        End If
    ElseIf oDate.Year = "" Then
        strYMD = "MD"
        StrDateToString = oDate.ToString(GetDateFormat(strFormat, strYMD))
    Else
        strYMD = "YMD"
        StrDateToString = oDate.ToString(GetDateFormat(strFormat, strYMD))
    End If
End Function

Function GetDateFormat(strFormat, strYMD)
    GetDateFormat = Session("DicCache").KeyValue(strFormat & strYMD)
End Function

Sub LogTagChange(strKey, strOldKey, strValue, strAction, strVersion)
        Dim strPhrase, strAlt
        strAlt="Warning: [Dictionary Tag {1} has been replaced from version {!3} by ]Tag {0}[{?!1} introduced at version {3} is missing from Dictionary and has been defaulted to '{2}]'"
        strPhrase = Util.FormatPhrase(Util.FirstNonEmpty(ConfigMsg("WarningDicChange", Dic.Peek("FmtTagChange"),""), strAlt), strKey, strOldKey, strValue, strVersion)
        If Session("Outdated").Added(strPhrase) Then Report.LogWarning(strPhrase)
End Sub

Function StrDicOpt(strRoot, strSuffix, strFmt)
         StrDicOpt = Dic.Peek(strRoot&strSuffix)
         If StrDicOpt="" Then StrDicOpt = Util.FormatString(strFmt, Dic(strRoot), strSuffix)
End Function

Function StrDicExt(strNewTag, strOldTag, strValue, strAction, strVersion)
    Dim strTag
    If (Dic.Peek(strNewTag) <> "") Then
        strTag = strNewTag
    ElseIf strOldTag <> "" Then
        strTag = strOldTag
        LogTagChange strNewTag, strOldTag, strValue, strAction, strVersion
    Else
        strTag = ""
        LogTagChange strNewTag, strOldTag, strValue, strAction, strVersion
        StrDicExt = strValue
        Exit Function
    End If
    Select Case strAction
        Case "PC" : StrDicExt = Dic.PlurialCardinal(strTag, strValue)
        Case "P"  : StrDicExt = Dic.Plurial(strTag, strValue)
        Case Else : StrDicExt = Dic(strTag)
    End Select
End Function

Function StrDicMFU(strKey, strGender)
    If Dic.Peek(strKey & "_" & strGender) <> "" Then
        StrDicMFU=StrDicVariant(strKey & "_" & strGender)
    Else
        StrDicMFU = StrDicVariant(strKey)
    End If
End Function

Function StrDicMFU2(strKey, strGender, strOldKeyOrValue, strVersion)
' as StrDicMFU but with backwards compatibility
Dim strOldKey, strValue

        If Left(strOldKeyOrValue,1) = "=" Then
                strValue = Mid(strOldKeyOrValue, 2)
                strOldKey = ""
        Else
                strValue = ""
                strOldKey = strOldKeyOrValue
        End If
    If StrPeek(strKey & "_" & strGender) <> "" Then
        StrDicMFU2=StrDicVariant(strKey & "_" & strGender)
    ElseIf strOldKey <> "" And StrPeek(strOldKey & "_" & strGender) <> "" Then
                LogTagChange strKey & "_" & strGender, strOldKey & "_" & strGender, "", "", strVersion
        StrDicMFU2=StrDicVariant(strOldKey & "_" & strGender)
    ElseIf StrPeek(strKey) <> "" Then
        StrDicMFU2 = StrDicVariant(strKey)
    ElseIf strOldKey <> "" And StrPeek(strOldKey) <> "" Then
                LogTagChange strKey, strOldKey, "", "", strVersion
        StrDicMFU2 = StrDicVariant(strOldKey)
        ElseIf strValue <> "" Then
                StrDicMFU2 = strValue
                LogTagChange strKey, "", """" & strValue & """", "", strVersion
    Else   ' let it fail'
        StrDicMFU2 = StrDicVariant(strKey)
    End If
End Function

Function StrDicMFUAttribute(strKey, strGender, strID)
    If Dic.Peek(strKey & "_" & strGender) <> "" Then
        StrDicMFUAttribute=StrDicAttribute(strKey & "_" & strGender, strID)
    Else
        StrDicMFUAttribute = StrDicAttribute(strKey, strID)
    End If
End Function

Function StrDicOrTag(strKey, strOption)
' Find entry with optional suffix. If suffix present use a Document Custom Tag if available rather than Dictionary entry
    Dim strTag, strVariant
    If strOption <> "" Then strTag=CustomTag(Null, strKey & strOption)
    If strTag <> "" Then
        If Instr(strTag, "{¿") > 0 Then strTag=StrParseText(strTag, True)
        If strTag <> "" Then
            StrDicOrTag = strTag
            Exit Function
        End If
    End If
    strTag = strKey & strOption
    If Dic.Peek(strTag) = "" Then strTag = strKey
    If Dic.Peek(strTag) <> "" Then StrDicOrTag=StrDicVariant(strTag)
End Function

Function StrDicOrTag2(strKey, strOption, strOldKey, strVersion)
' Find entry with optional suffix. If suffix present use a Document Custom Tag if available rather than Dictionary entry
    Dim strTag, strVariant
    If strOption <> "" Then strTag=CustomTag(Null, strKey & strOption)
    If strTag <> "" Then
        If Instr(strTag, "{¿") > 0 Then strTag=StrParseText(strTag, True)
        If strTag <> "" Then
            StrDicOrTag2 = strTag
            Exit Function
        End If
    End If
    strTag = strKey & strOption
    If StrPeek(strTag) = "" Then strTag = strKey
    If StrPeek(strTag) <> "" Then
                StrDicOrTag2=StrDicVariant(strTag)
                Exit Function
        End If
    strTag = strOldKey & strOption
    If StrPeek(strTag) <> "" Then
                LogTagChange strKey & strOption, strTag, "", "", strVersion
                StrDicOrTag2=StrDicVariant(strTag)
        ElseIf StrPeek(strOldKey) <> "" Then
                LogTagChange strKey, strOldKey, "", "", strVersion
                StrDicOrTag2=StrDicVariant(strOldTag)
        End If
End Function

' get attribute of Dictionary entry.
Function StrDicAttribute(strKey, strID)
    Dim oNode, strValue
    Set oNode = Session("oDicRepGen")
    strValue = ""
    If Not oNode Is Nothing Then Set oNode = oNode.selectSingleNode(strKey)
    If not oNode Is Nothing Then
        strValue = oNode.GetAttribute(strID)
        If IsNull(strValue) Then strValue = ""
    Else
        Report.LogWarning Util.FormatString(ConfigMessage("WarningAttributeKeyMissing"), strKey) & strAttribute
    End If
    StrDicAttribute = strValue
End Function

' get attribute1 if it is presnt otherwise attribute2
Function StrDicAttribute2(strKey, strID1, strID2)
    StrDicAttribute2 = StrDicAttribute(strKey, strID1)
    If StrDicAttribute2 = "" Then StrDicAttribute2 = StrDicAttribute(strKey, strID2)
End Function


' get attribute of Key1 if it exists, otherwise attribute of Key2
Function StrDicLookup2Attribute(strKey1, strKey2, strID)
    If Dic.Peek(strKey1) <> "" Then
        StrDicLookup2Attribute = StrDicAttribute(strKey1, strID)
    Else
        StrDicLookup2Attribute = StrDicAttribute(strKey2, strID)
    End If
End Function

Function StrDicPeekAttribute(strKey, strID)
    If Dic.Peek(strKey) <> "" Then
        StrDicPeekAttribute = StrDicAttribute(strKey, strID)
    Else
        StrDicPeekAttribute = ""
    End If
End Function
Function StrEventPhrase(obj)
' Find a Dictionary phrase for a GenoPro Object, using derived Event custom tag _NarrativeStyle if present for the object
    Dim strTag, strVariant,strKey, strOption
  strKey = obj.Session("Event")
  If strKey = "" Then strKey = obj.Class
  strKey = "Ph" & strKey
  StrOption = CustomTag(obj, "NarrativeStyle")
    If strOption <> "" Then strTag=CustomTag(Null, strKey & strOption)
    If strTag <> "" Then
        If Instr(strTag, "{¿") > 0 Then strTag=StrParseText(strTag, True)
        If strTag <> "" Then
            StrEventPhrase = strTag
            Exit Function
        End If
    End If
    strTag = strKey & strOption
    If Dic.Peek(strTag) = "" Then strTag = strKey
    If Dic.Peek(strTag) <> "" Then 
        StrEventPhrase = StrDicVariant(strTag)
    Else
        Report.LogWarning obj.Class & " " & obj.ID & " Event:" & obj.Session("Event") & " - No Dictionary.xml entry found"
        StrEventPhrase = StrDicVariant("PhEVEN")
    End If
End Function

Function StrDicVariant(strKey)
'    Check for any Phrase variants by checking for X (extra) attribute on main key, if present chose one of the variants of the phrase at random
    Dim nRnd, strRnd
    StrDicVariant = Dic.Peek(strKey)
  If StrDicVariant = "" Then Exit Function
    strRnd = StrDicAttribute(strKey, "X")
    nRnd = (0 & strRnd) + 0
    If nRnd > 0 Then 
        Randomize
        nRnd = Int ((nRnd  + 1 )* Rnd) 
    End If
    If nRnd > 0 Then
        StrDicVariant = Dic(strKey & "_X" & nRnd)
    End If
End Function

Function StrNameTranslate(strName, oNameDic, fBlankIfNone)
    Dim strTemp, strTrans, strPart
    strTemp = strName
    If Not oNameDic Is Nothing And strTemp <> "" Then
        strTrans = oNameDic(strTemp)
        If strTrans <> strTemp Then
            StrNameTranslate = strTrans
            Exit Function
        End If
        Dim oMatches, oMatch
        Session("RegEx").Pattern = "([\w-`' ])+"
        Session("RegEx").Global = True
        Set oMatches = Session("RegEx").Execute(strTemp)
        For Each oMatch In oMatches
            strPart = oMatch.Value
            strTrans = oNameDic(strPart)
            If strTrans <> strPart Then strTemp = Replace(strTemp, strPart, strTrans)
        Next
    End If
    If Not fBlankIfNone Or strTemp <> strName Then
         StrNameTranslate = strTemp
    Else
        StrNameTranslate = ""
    End If
End Function

Function StrPeek(strKey)
    Dim oNode, strValue
    Set oNode = Session("oDicRepGen")
    strValue = ""
    If strKey <> "" And Not oNode Is Nothing Then Set oNode = oNode.selectSingleNode(strKey)
    If not oNode Is Nothing Then
        strValue = oNode.GetAttribute("T")
        If IsNull(strValue) Then strValue = ""
    End If
    StrPeek = strValue
End Function

Function StrPlaceTranslate(strName)
    Dim arrParts, strPart, strTrans, strTranslated, oNameDicPlace
    strTranslated = strName
    Set oNameDicPlace = Session("oNameDicPlace")
    If Not oNameDicPlace Is Nothing Then
        arrParts = split(strName, ",")
        For Each strPart In arrParts
            strTrans = oNameDicPlace(Trim(strPart))
            If strTrans <> strPart Then strTranslated = Replace(strTranslated, Trim(strPart), strTrans)
        Next
    End If
    StrPlaceTranslate = strTranslated            
End Function

Function StrPreferredName(strName)
    nStart = Instr(strName, Session("MarkerFirstName"))
    If nStart > 0 Then
        StrPreferredName = Mid(strName, nStart + 1, Instr(nStart, strName & " ", " ") - nStart -1)
    Else
        StrPreferredName = strName
    End If
End Function

Function StrSubstitute(strValue, arrPatterns)
    Dim i,strTemp, oRegEx
    Set oRegEx = Session("RegEx")
    For i = 0 to Ubound(arrPatterns)-1 Step 2
        oRegEx.Pattern=arrPatterns(i)
        strTemp=oRegEx.Replace(strValue, arrPatterns(i+1))
        If strValue <> strTemp Then Exit For
    Next
    If strTemp <> "" Then
        StrSubstitute = strTemp
    Else
        StrSubstitute = strValue
    End If
End Function


Function StrVerb(strRoot, fPresent, fSingular, strVariant, strGender)
    ' return required form of verb i.e. present/past tense, singular or plural, language variant
    Dim strKey
    strKey = "_" & strRoot & strVariant & Util.IfElse(fPresent, "_Present", "_Past")
    If strGender <> "" Then If Dic.Peek(strKey & "_" & strGender)<> "" Then strKey = strKey & "_" & strGender
    StrVerb = Dic.Plurial(strKey, fSingular + 2)        'n.b. 'fSingluar + 2' equals 1 if fSingular is true and 2 (i.e. plural) if false.
End Function

Sub WriteNarrativeGrandParentsAdopted(i, iParent)
' find any adoptive parents, returning their GenoObject references when present
    Dim iCnt, oLink, oFather, oMother, strPnR, strAlso, strGender, oRepertoryNonBio, strPhrase
    strPhrase="={  }{3h}[{?3^5} and ]{5h} [{?1}{1} [{!7} ]{!0} grandparents][{!}[{?3}{2} [{!7} ]{!0} grandfather][{?5}{4} [{!7} ]{!0} grandmother]] through {0} [{?6=M}father's][{!}mother's] adoption."
    Set oRepertoryNonBio = Session("oRepertoryNonBio")
    If oRepertoryNonBio.KeyCounter("I" & iParent.ID) > 0 Then
        strPnR = PnR(i)
        strGender = iParent.Gender.ID
        strAlso = ""
        For iCnt = 0 to oRepertoryNonBio.Entry("I" & iParent.ID).Count-1
            Set oLink = oRepertoryNonBio.Entry("I" & iParent.ID).Object(iCnt)
            If oLink.PedigreeLink.ID = "Adopted" Then
                Set oFather = oLink.Family.Husband(0)
                Set oMother = oLink.Family.Wife(0)
                If (oMother = iParent.Mother) Or Util.IsNothing(oMother) Then     ' adopted by (step)father
                    Report.WritePhrase StrDicMFU2("PhGrandParentsAdopted", strGender, strPhrase, "2.0.1.?"), strPnR, "", _
                                    ToBe(oFather), oFather.Session("HlinkNN"),"","", strGender, strAlso, IsAlive(i), IsAlive(oFather), ""

                ElseIf (oFather = iParent.Father) Or Util.IsNothing(oFather) Then    ' adopted by (step)mother
                    Report.WritePhrase StrDicMFU2("PhGrandParentsAdopted", strGender, strPhrase, "2.0.1.?"), strPnR, "", _
                                    "", "", _
                                    ToBe(oMother), oMother.Session("HlinkNN"), strGender, strAlso, IsAlive(i), "", IsAlive(oMother)
                ElseIf Not (Util.IsNothing(oFather) Or Util.IsNothing(oMother)) Then                    ' adopted by both
                    Report.WritePhrase StrDicMFU2("PhGrandParentsAdopted", strGender, strPhrase, "2.0.1.?"), strPnR, ToBe2A(oFather, oMother), _
                                    ToBe(oFather), oFather.Session("HlinkNN"), _
                                    ToBe(oMother), oMother.Session("HlinkNN"), strGender, strAlso, IsAlive(i), IsAlive(oFather), IsAlive(oMother)
                End If
                strAlso = Dic("Also")
            End If
        Next
    End If
End Sub

Sub WriteNarrativeIndividual(i, fShortFormat, fam)

    Dim strGender, strName, strNamePossessive, strPnR, strPnP, strToHave, ich, j, arrMotherAdopters, arrFatherAdopters, arrOfficiatorTitle
    strGender = i.Gender.ID
    strName = Util.FormatPhrase(StrDicMFU("PhNameFriendly", strGender) ,i.Session("NameFormal"), i.Session("NameFull"), i.Session("NameKnownAs"), i.Gender.ID)
    strNamePossessive = i.Session("NamePossessive")
    strPnR = PnR(i)
    strPnP = PnP(i)
    
    Dim b, ba, c, d, f
    Set b = i.Birth
    Set ba = b.Baptism
    Set d = i.Death
    Set f = d.Funerals
    ' @ 2.0.1.7RC2 :
    ' allow to forms of Officiator title, separated by a vertical bar. 1st is used when Officiator name is also presnt, the second when it is not.
    ' e.g. 'Reverend' John Smith or just a 'reverend'
    arrOfficiatorTitle=Split(ba.Officiator.Title & "|", "|")
    If arrOfficiatorTitle(1) = "" Then arrOfficiatorTitle(1) = arrOfficiatorTitle(0)
    Report.WritePhrase StrDicMFU("PhBirth",strGender), strName, CustomDate(b.Date).ToStringNarrative, b.Place.Session("HlinkLocative"), CustomDate(ba.Date).ToStringNarrative, _
                ba.Place.Session("HlinkLocative"), arrOfficiatorTitle(0), StrFormatText(i, ba.Officiator), (b.Place.ID = ba.Place.ID), _
                StrFormatText(i, b.Doctor), b.PregnancyLength.Months,  b.PregnancyLength.Weeks, strPnR, b.CeremonyType, strNamePossessive, strGender, arrOfficiatorTitle(1)
    '@
    WriteHtmlFootnoteRef2 b.Source, ba.Source
    Session("RegEx").Pattern = "[,&]"
    Report.WritePhrase StrDicMFU("PhGodparents", strGender), strPnR, StrFormatText(i, ba.Godfather), StrFormatText(i, ba.Godmother), i.IsDead = False, Session("RegEx").Test(ba.Godfather), Session("RegEx").Test(ba.Godmother), strGender
    WriteHtmlAnnotation i, Dic("AnnotationBirth"), b.Comment

    If (Not fShortFormat) Then
        ' Write the parents and grandparents
        Report.WritePhrase StrDicMFU("PhParents", strGender), strNamePossessive, ToBe(i.Father), i.Father.Session("HlinkNN"), _
                                  strPnR, ToBe(i.Mother), i.Mother.Session("HlinkNN"), IsAlive(i), IsAlive(i.Father), IsAlive(i.Mother)
        Report.WritePhrase StrDicMFU("PhGrandParents", strGender), strPnR, ToBe2a(i.Father.Father, i.Father.Mother), _
                                    ToBe(i.Father.Father), i.Father.Father.Session("HlinkNN"), _
                                    ToBe(i.Father.Mother), i.Father.Mother.Session("HlinkNN"), _
                                    ToBe2a(i.Mother.Father, i.Mother.Mother), _
                                    ToBe(i.Mother.Father), i.Mother.Father.Session("HlinkNN"), _
                                    ToBe(i.Mother.Mother), i.Mother.Mother.Session("HlinkNN"), _
                  IsAlive(i), _
                  IsAlive(i.Father.Father), IsAlive(i.Father.Mother), _
                  IsAlive(i.Mother.Father), IsAlive(i.Mother.Mother)

        WriteNarrativeGrandParentsAdopted i, i.Father
        WriteNarrativeGrandParentsAdopted i, i.Mother
            
        WriteNarrativeSiblings i
    End If

    If (i.IsDead) Then
        Set c = d.Cause
        ich = Report.BufferLength
        Report.WritePhrase StrDicMFU("PhDied",strGender), strPnP, i.Death.Age.Years, i.Death.Age.Months, i.Death.Age.Days, CustomDate(d.Date).ToStringNarrative, d.Place.Session("HlinkLocative"), c, StrFormatText(i, c.Description), CustomDate(i.Birth.Date).Approximation & CustomDate(i.Death.Date).Approximation, "", strGender
            Report.WritePhrase StrDicMFU("PhFuneral",strGender), CustomDate(f.Date).ToStringNarrative, f.Place.Session("HlinkLocative"), StrFormatText(i, f.Agency)
        Report.WritePhrase StrDicMFU("PhBurial",strGender) , strPnR, d.Disposition.Type, CustomDate(d.Disposition.Date).ToStringNarrative, d.Disposition.Place.Session("HlinkLocative"), strPnP, d.Disposition.Type.ID, strGender
        If (ich = Report.BufferLength) And (Not i.Father.IsDead = True Or Not i.Mother.IsDead = True) Then Report.WritePhraseDic "PhDead", strName, strGender ' Indicate dead if no other death phrase and one or more parents alive.
    End If
    WriteHtmlFootnoteRef2 f.Source, d.Disposition.Source
    WriteHtmlFootnoteRef d.Source
    WriteHtmlAnnotation i, Dic("AnnotationDeath"), d.Comment
    
    Dim collSpouses, collChildren, strHtmlChildren
    Set collSpouses = i.Mates.ToGenoCollection
    Set collChildren = i.Children.ToGenoCollection.SortByGender

    ' if family report page remove current spouse and this family's children and report on other spouse & children
    If Not Util.IsNothing(fam) Then
        collSpouses.Remove i.FindMate(fam)
        collChildren.Remove fam.Children
        strHtmlChildren = StrHtmlCollectionOtherChildren(collChildren)
    ElseIf fShortFormat Or Session("fHideFamilyDetails") Then
        strHtmlChildren = StrHtmlCollectionChildren(collChildren)
    Else
        strHtmlChildren = StrHtmlCollectionChildrenLocal(collChildren, "")
    End If
    ' {0=Name} {1=has|had} {n spouses} named, ..., and {n sons} and {n daughters}, named ...
    strToHave = StrVerb("ToHave", False, True, "", strGender)
    If (collSpouses.Count = 1) Then
        ' One spouse, therefore check if they are still together
        If (i.Families(0).AreTogether) Then
            strToHave = StrVerb("ToHave", True, True, "", strGender)
        End If
    End If
    ' The following lines have been commented because of linguistic issues such as "spouse", "husband", "wife" and "partner".
    ' Until a solution is developed, those lines will remain commented by default.
    ' if family report page only report other spouse(s) and children
    If Util.IsNothing(fam) Then
        'Report.WritePhraseDic "PhSpousesAndChildren", strName, strToHave, StrHtmlCollectionSpouses(collSpouses), strHtmlChildren
    Else
        'Report.WritePhraseDic "PhOtherSpousesAndChildren", strName, strToHave, StrHtmlCollectionOtherSpouses(collSpouses), strHtmlChildren
    End If
End Sub

Sub WriteNarrativeUnionsAndDivorces(f, nFamily, iParent)
    Dim strPrefix, strToBe, collParents, collUnions, m, d, t, strAlso, fAreAlive, fLast, strLastId
    Dim objLast, strGender, strSpouseGender, arrOfficiatorTitle, strStyleLast, strName, strPhrase
    strAlso = ""
    Set collParents = f.Parents.ToGenoCollection
    strPhrase = Util.IfElse(Session("Book"), StrDicExt("PhTheyBook","","[{0}][{?1}[{?0} and] {1h}][{?!0|1}{0=They}]","","2013.09.22"), Dic("PhThey"))
    If nFamily >= 0 Then
        If iParent.ID = collParents(0).ID Then
            strPrefix = Util.FormatPhrase(strPhrase, Util.HtmlEncode(collParents(0).Session("NameShort")), StrHtmlNarrativeName(collParents(1),"NameShort", "", ""))
            strGender = collParents(0).Gender.ID
            strSpouseGender=collParents(1).Gender.ID
        Else
            strPrefix = Util.FormatPhrase(strPhrase, Util.HtmlEncode(collParents(1).Session("NameShort")), StrHtmlNarrativeName(collParents(0),"NameShort", "", ""))
            strGender = collParents(1).Gender.ID
            strSpouseGender=collParents(0).Gender.ID
        End If
    Else
        ' 1st fix potential issue if outdated translated Dictionary being used due to c
        If InStr(strPhrase, "{0h}")>0 Then ' new style phrase with hyperlink on first name
           strName = StrHtmlNarrativeName(collParents(0),"NameShort", "", "")
        Else  ' old style phrase with no hyperlink for first name
            strName = collParents(0).Session("NameShort")
        End If
        strPrefix = Util.FormatPhrase(strPhrase, strName, StrHtmlNarrativeName(collParents(1),"NameShort", "", ""))
        strGender = collParents(0).Gender.ID
        strSpouseGender=collParents(1).Gender.ID
    End If
    f.Session("Prefix") = strPrefix
    f.Session("Gender") = strGender
    f.Session("GenderSpouse") = strSpouseGender

    fAreAlive = (Util.GetStatisticsForIndividuals(collParents) < 4)

    Set collUnions = f.Unions.ToGenoCollection
    If collUnions.Count > f.Session("Events") Then  ' There are some 'real' Unions
        ' The family has at least one union/marriage
        Session("RegEx").Pattern = "[,&]"
        strStyleLast = ""
        For Each m In collUnions
            If m.Session("Event") = "" Then 'it is a 'real' union, not another event masquerading as one
                arrOfficiatorTitle=Split(m.Officiator.Title & "|", "|")
                If arrOfficiatorTitle(1) = "" Then arrOfficiatorTitle(1) = arrOfficiatorTitle(0)
                Report.WritePhrase StrDicOrTag("PhUnion", CustomTag(m, "NarrativeStyle")), strPrefix, m.Type, (m.Date).ToStringNarrative, m.Place.Session("HlinkLocative"), arrOfficiatorTitle(0), m.Officiator, m.Witnesses, fAreAlive=True, f.AreTogether=True, strAlso, Session("RegEx").Test(m.Witnesses), collParents.Count = 1, arrOfficiatorTitle(1)
                WriteHtmlFootnoteRef m.Source
                Set d = m.Divorce
                Report.WritePhrase Util.FirstNonEmpty(StrDicOrTag("PhUnionEnd", CustomTag(m, "Divorce.NarrativeStyle")), StrDicVariant("PhDivorce")), m.IsAnnulled, CustomDate(d.Date).ToStringNarrative, d.Place.Session("HlinkLocative"), d.RequestedBy.ID, _
                                 d.RequestedBy, StrFormatText(m, d.Attorney.Husband), StrFormatText(m, d.Attorney.Wife), StrFormatText(m, d.Officiator)
                WriteHtmlFootnoteRef m.Divorce.Source
                WriteHtmlAdditionalInformation(m)
                WriteHtmlAnnotation m, Dic("AnnotationUnion"), m.Comment
                strPrefix = Dic.Plurial("PnP_",2)
                If strStyleLast = CustomTag(m, "NarrativeStyle") Then strAlso = Dic("Also")
                strStyleLast = CustomTag(m, "NarrativeStyle")
            End If
        Next
    Else
        ' There is no union/marriage to the family, so display the relationship if any
        If f.Relation.ID <> "" Then Report.WritePhrase StrDicVariant("PhFR_" & f.Relation.ID), strPrefix, fAreAlive=True, f.AreTogether=True, (collParents(0).Name <> "" And collParents(1).Name <> ""), "", strGender, strSpouseGender, (collParents.Count = 1)
    End If
End Sub


' Write details of a family
'
' This routine has two extra parameters (nFamily and iParent) to
' optionally add the spouse name to the report.
' If nFamily is zero, then don't add the spouse name because there is only one spouse (it's redundant)
' If nFamily is negative, it is because we are generating the HTML code for a family report rather than an individual's family
Sub WriteHtmlFamily(f, nFamily, iParent)
    Dim collChildren, collAdopted, collAdoptedByFather, collFamilies, collAdoptedByMother, collFoster, collNatural, collIllegitimate
    Dim oRepertoryNonBio, c, spouse, strSpouseName, strSpousePossessive, strSpouseType, oFamilies, strHaveOrHad
    Dim strHaveOrHad1, strHyperlinkText, strName, strPrefix, ich, strAlso, fSingle, f1
    Dim strOrdinalGender, strOrdinal
    fSingle = (f.Parents.Count = 1)
    If (nFamily >= 0) Then 
        Set spouse = iParent.FindMate(f)    ' Find the other spouse
        strSpouseName = Util.IfElse(Not fSingle, spouse.Session("NameFull"), "")
        strPrefix = Util.IfElse(f.GotMarried, "_Spouse", "_Partner")
        strSpousePossessive = StrDicLookup2Attribute(strPrefix & "_" & iParent.Gender.ID & "_" & spouse.Gender.ID, strPrefix,"G") ' get genetive case of spouse noun
        strOrdinalGender = StrDicAttribute("PhFamilyWith", "G1")    ' optional G1 attribute indicates gender  required for parameter {1} (ie. first, second, third etc) in other language translations
        If nFamily > 0 Then strOrdinal = StrDicMFU("_Ordinal_" & nFamily, strOrdinalGender)
        Report.WriteFormattedLn "<a id='{}'></a>", f.ID
        strHyperlinkText = Dic.FormatPhrase("PhFamilyWith", i.Session("NamePossessive"), strOrdinal, strSpouseName, "")
        Report.WriteFormattedLn "<div class='clear'><br /><{0} class='familyheading'>{1}</{0}>{2}</div>", Util.IfElse(Session("fHideFamilyDetails"), "span", "h3"), Replace(Util.FormatHtmlHyperlink(f.Href, strHyperlinkText),">"," onclick='tocExit();' >",1,1), StrHtmlImgMap(f)
        '   Check for all 'Families' with same spouse e.g. for committed relationship then marriage
        Set collFamilies = Util.NewGenoCollection
        For Each f1 In iParent.Families
            If iParent.FindMate(f1) = spouse Then collFamilies.Add f1
        Next
    End If
'    put pictures in a 'div' floating right. If no pictures added then remove the 'div' 

    Dim cchBufferStart    ' Start of the write operations

    cchBufferStart = Report.BufferLength

    Report.WriteLn "<div class='floatright aligncenter widthpaddedlarge'>"
    Session("BufferBegin") = Report.BufferLength

    If (Not Util.IsNothing(spouse)) Then
        WriteHtmlPicturesLarge spouse, "", Util.StrFirstCharUCase(Dic.FormatString("FmtPictures", strSpousePossessive)), Session("fHidePictureName") Or Session("fShowPictureDetails"), False
        If f.Session("PicturesIncluded") > 0 Then Report.WriteBr
        WriteHtmlPicturesLarge f, "", Util.FirstNonEmpty(Dic.Peek("HeadingPicturesFamily"), Dic.FormatString("FmtPictures", Dic("Family"))), Session("fHidePictureName") Or Session("fShowPictureDetails"), False
     Else
        WriteHtmlPicturesLarge f, "", "", Session("fHidePictureName") Or Session("fShowPictureDetails"), False
    End If

    If Report.BufferLength = Session("BufferBegin") Then        ' no picture information written
        Report.BufferLength = cchBufferStart        ' so remove the 'div' by stepping back the buffer
    Else
Report.WriteLn "</td></tr></table>"
        Report.WriteLn "</div>"                ' close the 'div' with end tag
    End If

    Report.WriteLn "<div>"
    WriteNarrativeUnionsAndDivorces f, nFamily, iParent

    ' They {have|had} {n sons} and {n daughters}, named ...

    Set collChildren = f.Children.ToGenoCollection.SortByGender
    Set collAdopted = Util.NewGenoCollection
    Set collAdoptedByFather = Util.NewGenoCollection
    Set collAdoptedByMother = Util.NewGenoCollection
    Set collFoster = Util.NewGenoCollection
    Set collNatural = Util.NewGenoCollection
    Set collIllegitimate = Util.NewGenoCollection

    Set oRepertoryNonBio = Session("oRepertoryNonBio")

    Dim strType, cnt, oEntry, oLink
    strType = ""
    If fSingle Then strType = "Biological"

    If oRepertoryNonBio.KeyCounter("A:" & f.ID) > 0 Then
        Set oEntry = oRepertoryNonBio.Entry("A:" & f.ID)
        For cnt = 0 To oEntry.Count-1 Step 2
            Set oLink = oEntry.Object(cnt)
            If oLink.Child.Mother.ID = f.Parents(1).ID Then
                collAdoptedByFather.Add oEntry.Object(cnt+1)
            ElseIf oLink.Child.Father.ID = f.Parents(0).ID Then
                collAdoptedByMother.Add oEntry.Object(cnt+1)
            Else
                collAdopted.Add oEntry.Object(cnt+1)
            End If
        Next
        collChildren.Remove collAdopted
        collChildren.Remove collAdoptedByFather
        collChildren.Remove collAdoptedByMother
        collAdopted.SortByGender
        collAdoptedByFather.SortByGender
        collAdoptedByMother.SortByGender
        strType="Biological"
    End If
    If oRepertoryNonBio.KeyCounter("F:" & f.ID) > 0 Then    
        Set oEntry = oRepertoryNonBio.Entry("F:" & f.ID)
        For cnt = 0 To oEntry.Count-1 Step 2
            collFoster.Add oEntry.Object(cnt+1)
        Next
        collChildren.Remove collFoster
        collFoster.SortByGender
        strType="Biological"
    End If

    If oRepertoryNonBio.KeyCounter("I:" & f.ID) > 0 Then    
        Set oEntry = oRepertoryNonBio.Entry("I:" & f.ID)
        For cnt = 0 To oEntry.Count-1 Step 2
            collIllegitimate.Add oEntry.Object(cnt+1)
        Next
        collChildren.Remove collIllegitimate
        collIllegitimate.SortByGender
        strType=""
    End If
    If oRepertoryNonBio.KeyCounter("N:" & f.ID) > 0 Then    
        Set oEntry = oRepertoryNonBio.Entry("N:" & f.ID)
        For cnt = 0 To oEntry.Count-1 Step 2
            collNatural.Add oEntry.Object(cnt+1)
        Next
        collChildren.Remove collNatural
        collNatural.SortByGender
        strType=""
    End If

    strHaveOrHad1 = StrVerb("ToHave", Util.GetStatisticsForIndividuals(iParent, spouse) < 4, fSingle, "", "")

    If (Util.GetStatisticsForIndividuals(collChildren) >= 4) Then
        strHaveOrHad = StrVerb("ToHave", False, fSingle, "", "")    ' One or more children died
    Else
        strHaveOrHad = strHaveOrHad1
    End If
    strPrefix = Util.IfElse(fSingle, StrDicMFU("PnP", f.Parents(0).Gender.ID), Dic.Plurial("PnP_", 2))    'he, she or they
    strAlso = ""
    ich = Report.BufferLength
    Report.WritePhraseDic "PhSpousesAndChildren", strPrefix, strHaveOrHad, StrHtmlCollectionChildrenLocal(collChildren, strType), nothing, strAlso
    If ich <> Report.BufferLength Then
        strAlso = Dic("Also")
    End If
    Report.WritePhraseDic "PhStepChildrenAdoption", f.Parents(0).Session("NameShort"), PnR(f.Parents(0)), StrHtmlCollectionStepChildren(collAdoptedByFather)
    Report.WritePhraseDic "PhStepChildrenAdoption", f.Parents(0).Session("NameShort"), PnR(f.Parents(1)), StrHtmlCollectionStepChildren(collAdoptedByMother)

    If (Util.GetStatisticsForIndividuals(collAdopted) >= 4) Then
        strHaveOrHad = StrVerb("ToHave", False, fSingle, "", "")    ' One or more adopted children died
    Else
        strHaveOrHad = strHaveOrHad1
    End If
    ich = Report.BufferLength
    Report.WritePhraseDic "PhSpousesAndChildren", strPrefix, strHaveOrHad, StrHtmlCollectionChildrenLocal(collAdopted, "Adopted"), nothing, strAlso
    If ich <> Report.BufferLength Then
        strAlso = Dic("Also")
    End If
    If (Util.GetStatisticsForIndividuals(collFoster) >= 4) Then
        strHaveOrHad = StrVerb("ToHave", False, fSingle, "", "")    ' One or more adopted children died
    Else
        strHaveOrHad = strHaveOrHad1
    End If
    Report.WritePhraseDic "PhSpousesAndChildren", strPrefix, strHaveOrHad, StrHtmlCollectionChildrenLocal(collFoster, "Foster"), nothing, strAlso

    If ich <> Report.BufferLength Then
        strAlso = Dic("Also")
    End If
    If (Util.GetStatisticsForIndividuals(collIllegitimate) >= 4) Then
        strHaveOrHad = StrVerb("ToHave", False, fSingle, "", "")    ' One or more adopted children died
    Else
        strHaveOrHad = strHaveOrHad1
    End If
    Report.WritePhraseDic "PhSpousesAndChildren", strPrefix, strHaveOrHad, StrHtmlCollectionChildrenLocal(collIllegitimate, "Illegitimate"), nothing, strAlso

    If ich <> Report.BufferLength Then
        strAlso = Dic("Also")
    End If
    If (Util.GetStatisticsForIndividuals(collNatural) >= 4) Then
        strHaveOrHad = StrVerb("ToHave", False, fSingle, "", "")    ' One or more adopted children died
    Else
        strHaveOrHad = strHaveOrHad1
    End If
    Report.WritePhraseDic "PhSpousesAndChildren", strPrefix, strHaveOrHad, StrHtmlCollectionChildrenLocal(collNatural, "Natural"), nothing, strAlso

    Report.WritePhraseDic "PhFL_" & f.FamilyLine.ID, strHaveOrHad, f.AreTogether And Not f.EndOfFamily, f.Children.ToGenoCollection.Count > 0, ""

    WriteNarrativeTwins f    ' put here because may include notes on each multiple birth set.

    WriteHtmlAnnotation f, Dic("AnnotationFamily"), f.Comment
    Report.WriteBr "</div>"

    WriteHtmlExtraNarrative f

    If (nFamily < 0) Then
                ' add short details of each partner and any other partner and/or children
        Set collParents = f.Parents.ToGenoCollection
        Report.WriteLn    "<div class='clearleft'><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='Personal'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Dic("HeaderPersonal")
        For Each p in collParents
            WriteHtmlIndividual p, f
        Next
        Report.WriteLn "</ul></li></ul></div>"
    End If
    
    '   Write a sub-section with details of Attributes and Events if any
    
    If Session("Flag_A") And f.Session("Events") > 0 Then  ' There are some events masquerading as Unions
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present
        Report.WriteLn    "<div class='clearleft'><br /><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='AttributesEvents'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Dic("HeaderAttributesEvents")

        Dim collUnions, u, strGender, strSpouseGender, strPhrase, strRelative
        Set collUnions = f.Unions.ToGenoCollection
        For Each u In collUnions
            If u.Session("Event") <> "" Then 'it is an event, not a 'real' union
                  Report.WriteFormattedLn "<a name='{}'></a><div class='clearleft'>", u.ID
                    strPrefix=f.Session("Prefix")
                strGender=f.Session("Gender")
                strSpouseGender=f.Session("GenderSpouse")
                strName=f.Session("Name")
                strRelative=Dic.Plurial("PnR_" & f.Session("Gender"), f.Parents.Count)
                 strPhrase = StrEventPhrase(u)
                  Report.WritePhrase strPhrase, _
                                  StrDateSpan(u.Date, u.Divorce.Date), _
                                  f.Session("Prefix"), _
                                  strRelative, _
                                  StrFormatText(u, StrParseText(u.Session("Title"), True)), _
                                  "<h6>","</h6>", _
                                  StrFormatText(u, StrParseText(u.Session("Agency"), True)), _
                                  StrHtmlHyperlink(u.Place), _
                                  u.Session("Agency.Prefix"), _
                                  strName, _
                                  u.Session("Title.Prefix"), _
                                  (f.Parents.Count = 1)
                WriteHtmlFootnoteRef u.Source
                WriteHtmlAdditionalInformation(u)
                WriteHtmlAnnotation u, Dic("AnnotationEventAttribute"), u.Comment
                strPrefix = Dic.Plurial("PnP_" & f.Session("Gender"),f.Parents.Count)
                Report.WriteLn "</div>"
            End If
        Next
        Report.WriteLn "</ul></li></ul></div>"
    End If
    
    Set collChildren = f.Children.ToGenoCollection
    If (collChildren.Count > 0) Then
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present
        Report.WriteLn    "<div class='clearleft'><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' id='Children' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='Children'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'>", Dic("HeaderChildren")
        For Each c In collChildren
            Report.WriteFormattedLn "<a id='{&t}'></a>", c.Href
            WriteHtmlIndividual c, Null
        Next  
        Report.WriteLn "</ul></li></ul></div>"
    End If


    If (nFamily >= 0) Then
        WriteHtmlAdditionalInformation(f)
    End If

End Sub

Sub WriteHtmlIndividual(i, f)
    Report.WriteLn "<div class='clearleft'>"

    Report.Write StrHtmlImgGender(i)
    Report.Write i.Session("HlinkNN")
    Report.WriteBr
    Dim p, pCnt
    If Not IsNull(f) Or Session("Flag_K") Then ' Parent's pic or Show Children's (a.k.a. kid's) pictures
        Set p = Nothing
        pCnt = CustomTag(i, "Pictures.Secondary")
        If IsNull(f) And IsNumeric(pCnt) Then
            pCnt = CInt(pCnt)
            If pCnt > 0 And pCnt < (i.Pictures.Count + 1) Then Set p = i.Pictures(pCnt - 1)
        End If
        If Util.IsNothing(p) Then If Not IsNull(f) Or Not Session("ChildPictureSecondary") Then Set p=i.Pictures.Primary
        If (Not Util.IsNothing(p)) Then
            WriteHtmlPictureSmall i, p, "left"
        Else
            If Session("fAddGenericImage") Then
                Report.WriteLn "<div class='floatleft aligncenter widthpaddedsmall'>"
                WriteHtmlGenericPicture i, Util.IfElse(IsNull(f),"C","P"), "s", Session("cxPictureSizeSmall"), Session("cyPictureSizeSmall"), Session("cxyPicturePadding"), "left"
                Report.WriteLn "</div>"
            End If
        End If
    End If
    WriteNarrativeIndividual i, true, f
    Report.WriteBr "</div>"
End Sub

Sub WriteNarrativeSiblings(i)
    Dim collSiblings, cSiblings, strPrefix, collAdoptedSiblings, collFosterSiblings, collHalfSiblings, collOtherSiblings
    Dim iSibling, oEntry, oLink, c, f, fEndOfLine, fOrdered, oRepertoryTwins, strGender, cMales, cFemales
    strGender = i.Gender.ID
    cMales = 0
    cFemales = 0
    If strGender = "M" Then
        cMales = 1
    ElseIf strGender = "F" Then
        cFemales = 1
    End If
    Set oRepertoryTwins = Session("oRepertoryTwins")
    Set oRepertoryNonBio = Session("oRepertoryNonBio")
    Set collSiblings = Util.NewGenoCollection
    Set collOtherSiblings = Util.NewGenoCollection
    Set collHalfSiblings = Util.NewGenoCollection
    Set f = i.Family
    For Each iSibling In i.Siblings.ToGenoCollection
        If f.ID = iSibling.Family.ID Or oRepertoryNonBio.KeyCounter("I"&iSibling.ID) = 0 Then
            collSiblings.Add iSibling
        Else
            collOtherSiblings.Add iSibling
        End If
        If iSibling.Gender.ID = "M" Then
            cMales = cMales + 1
        ElseIf iSibling.Gender.ID = "F" Then
            cFeMales = cFemales + 1
        End If
    Next
    collHalfSiblings.Add i.Siblings.Half.ToGenoCollection
    Set collOtherSiblings = i.Siblings.other.ToGenoCollection
    cSiblings = i.Siblings.ToGenoCollection.Count
    If (cSiblings = 0) Then
        If (Not Util.IsNothing(i.Family)) Then
            If collHalfSiblings.Count > 0 Then
                WriteNarrativeMFU i, collHalfSiblings, "SiblingHalf"
            Else
                If f.FamilyLine.ID ="NoMoreChildren" Or f.FamilyLine.ID = "" Then Report.WritePhrase StrDicMFU("PhOnlyChild",strGender), PnP(i), ToBe(i), strGender
            End If
        End If
        WriteNarrativeAdoption i
        Exit Sub
    End If
    Dim strChildRank, strChildRank1, nChildRank, cChildren, fOrdinal
    cChildren = cSiblings + 1
    nChildRank = i.FamilyRank
    strChildRank1 = ""
    If (nChildRank <= 1) Then
        If cChildren = 2 Then
            strChildRank = Dic.PlurialCardinal(GetDicMFU("Oldest", strGender), cChildren)'
        Else
            strChildRank = Dic.Plurial(GetDicMFU("Oldest", strGender), cChildren)
        End If
    ElseIf (nChildRank = cChildren) Then
        If cChildren = 2 Then
            strChildRank = Dic.PlurialCardinal(GetDicMFU("Youngest", strGender), cChildren)
        Else
            strChildRank = Dic.Plurial(GetDicMFU("Youngest", strGender), cChildren)
        End If
    Else
            If nChildRank - 1 < cChildren / 2 Then
              strChildRank = StrDicAttribute2(GetDicMFU("_Ordinal_" & nChildRank, strGender), "T1", "T")
              strChildRank1 = Dic("Oldest")
        Else
              strChildRank = StrDicAttribute2(GetDicMFU("_Ordinal_" & (cChildren - nChildRank + 1), strGender), "T1", "T")
              strChildRank1 = Dic("Youngest")
        End If
        fOrdinal = True
    End If
    strPrefix = Pnp(i)
    If Util.IsNothing(i.Family) Then ' individual is adopted or fostered but birth parent(s) not present
        Dim oRepertoryNonBio
        If oRepertoryNonBio.KeyCounter("I" & i.ID) > 0 Then
            Set oLink = oRepertoryNonBio.Entry("I" & i.ID).Object(0)
            strPrefix = Util.FormatPhrase(StrDicExt("PhPL_" & oLink.PedigreeLink.ID & "3", "PhPL_" & LCase(oLink.PedigreeLink.ID) & "3", "", "", "2.0.1.6"),PnR(i), PnP(i), strGender)
            Set f = oLink.Family
        End If
    End If

    WriteNarrativeMFU i, collSiblings, "Sibling"
    Dim collAdoptedByFather, collAdoptedByMother, cnt
    Set collAdoptedByFather = Util.NewGenoCollection
    Set collAdoptedByMother = Util.NewGenoCollection
    If oRepertoryNonBio.KeyCounter("A:" & i.Family.ID) > 0 Then
        collSiblings.Clear
        Set oEntry = oRepertoryNonBio.Entry("A:" & i.Family.ID)
        For cnt = 0 To oEntry.Count-1 Step 2
            Set oLink = oEntry.Object(cnt)
            If oLink.Child.Mother.ID = i.Mother.ID Then
                collHalfSiblings.Add oEntry.Object(cnt+1)
                collAdoptedByFather.Add oEntry.Object(cnt+1)
            ElseIf oLink.Child.Father.ID = i.Father.ID Then
                collHalfSiblings.Add oEntry.Object(cnt+1)
                collAdoptedByMother.Add oEntry.Object(cnt+1)
            Else
                collSiblings.Add oEntry.Object(cnt+1)
            End If
        Next
        WriteNarrativeMFU i, collSiblings, "SiblingAdopted"
    End If

    If oRepertoryNonBio.KeyCounter("F:" & i.Family.ID) > 0 Then    
        collSiblings.Clear
        Set oEntry = oRepertoryNonBio.Entry("F:" & i.Family.ID)
        For cnt = 0 To oEntry.Count-1 Step 2
            collSiblings.Add oEntry.Object(cnt+1)
        Next
        WriteNarrativeMFU i, collSiblings, "SiblingFoster"
    End If

    fOrdered = True
    If (Not Util.IsNothing(f)) Then
    If f.Children.OrderUnknown Then
        fOrdered = False
    ElseIf f.Children.Order.ToGenoCollection.Count = 0 Then
        If  i.Birth.Date.ToStringNarrative = "" Then
            fOrdered = False
        Else
            For Each c in i.Siblings.All.ToGenoCollection
                If c.Birth.Date.ToStringNarrative = "" Then
                    fOrdered = False
                    Exit For
                End If
            Next
        End If
    End If
    fEndOfLine = (f.FamilyLine.ID = "" Or f.FamilyLine.ID = "NoMoreChildren")
    End If

    If fOrdered And oRepertoryNonBio.KeyCounter("A:" & f.ID) = 0 And oRepertoryNonBio.KeyCounter("F:" & f.ID) = 0 And oRepertoryTwins.KeyCounter("F" & f.ID) = 0 Then
            Report.WritePhrase StrDicMFU("PhChildRank", strGender), strPrefix, ToBe(i), strChildRank, strChildRank1, StrDicMFU("_Cardinal_" & cChildren, Util.IfElse(cFemales = cChildren, "F", "")), Util.IfElse(fEndOfLine,"", StrDicLookup2Attribute("Known_" & strGender, "Known", "P")), Dic.Plurial("Child", cChildren), strGender, StrDicMFUAttribute("_Cardinal_" & cChildren, Util.IfElse(cFemales = cChildren, "F", ""), "TE")
    End If
    WriteNarrativeTwins i

    WriteNarrativeAdoption i

    WriteNarrativeMFU i, collHalfSiblings, "SiblingHalf"
    Report.WritePhraseDic "PhAdoptedBy", i.Father.Session("HlinkNN"), StrHtmlNarrativeNamesShort(collAdoptedByFather)
    Report.WritePhraseDic "PhAdoptedBy", i.Mother.Session("HlinkNN"), StrHtmlNarrativeNamesShort(collAdoptedByMother)

End Sub

' Write a narrative phrases, breaking the collection by males, females and unknown gender.
' This routine is used to write the siblings and half siblings.
' Example: "He had a half-brother and two half-sisters, named Benoit, Anne and Estelle."
Sub WriteNarrativeMFU(i, coll, strDicPrefix)
    coll.SortByGender
    Report.WritePhrase StrDicMFU("PhMalesFemalesUnknowns", i.Gender.ID), PnP(i), ToHave(i, coll), StrHtmlCollectionMFU(coll, strDicPrefix, StrHtmlNarrativeNamesShort(coll))
End Sub


Sub WriteNarrativeAdoption(obj)
    Dim oRepertoryNonBio, oLink, oLink1, i,j,f,idFamily,idFamilyBirth, oAdopt, coll, strAlso, strAdopter, strWho, strGender
    Set oRepertoryNonBio = Session("oRepertoryNonBio")
    Set coll = Util.NewGenoCollection
    If oRepertoryNonBio.KeyCounter("I" & obj.ID) > 0 Then
        strGender = obj.Gender.ID
        strAlso = ""
        For i = 0 to oRepertoryNonBio.Entry("I" & obj.ID).Count-1
            Set oLink = oRepertoryNonBio.Entry("I" & obj.ID).Object(i)
            idFamily =oLink.Family.ID
            idFamilyBirth = "B" & obj.Family.ID
            If oRepertoryNonBio.KeyCounter(idFamilyBirth) > 1 Then
                For j = 0 to oRepertoryNonBio.Entry(idFamilyBirth).Count-1
                    Set oLink1 =  oRepertoryNonBio.Entry(idFamilyBirth).Object(j)
                    If oLink1.individual.ID <> obj.ID And oLink1.Family.ID = idFamily Then 'not this individual and the same adopters
                        coll.Add oLink1.individual
                    End If
                Next
            End If
            strAdopter = oLink.Family.Session("Hlink")
            strWho=Dic.Plurial("PnP_",2)
            If oLink.Family.Parents(1) = obj.Mother  Then
                strAdopter=oLink.Family.Parents(0).Session("Hlink")
                strWho=PnP(oLink.Family.Parents(0))
            ElseIf oLink.Family.Parents(0) = obj.Father Then
                strAdopter=oLink.Family.Parents(1).Session("Hlink")
                strWho=PnP(oLink.Family.Parents(1))
            End If
            Set oAdopt = oLink.Adoption
            Report.WritePhrase StrDicMFU2("PhPL_" & oLink.PedigreeLink.ID, strGender, "PhPL_" & LCase(oLink.PedigreeLink.ID),"2.0.1.6"), PnP(obj), strAdopter, oAdopt.Age, CustomDate(oAdopt.Date).ToStringNarrative, oAdopt.Place.Session("HlinkLocative"), oAdopt.Agency, StrHtmlNarrativeNamesShort(coll), strAlso, strWho, strGender
            WriteHtmlFootnoteRef oAdopt.Source
            WriteHtmlAnnotation oLink, Dic("Annotation"+oLink.PedigreeLink.ID), oLink.Comment
            strAlso = Dic("Also")
        Next
    End If
End Sub

Sub WriteNarrativeTwins(obj)
    Dim oRepertoryTwins, strTwinId, oTwin, oTwins, i
    Set oRepertoryTwins = Session("oRepertoryTwins")

    Select Case obj.class
    Case "Family"
        If oRepertoryTwins.KeyCounter("F" & obj.ID) > 0 Then
            For i = 0 to oRepertoryTwins.Entry("F" & obj.ID).Count-1
                strTwinId = oRepertoryTwins.Entry("F" & obj.ID).Object(i)
                WriteNarrativeTwin oRepertoryTwins.Entry(strTwinId).Object(0)
            Next
        End If
    Case "Individual"                
        If oRepertoryTwins.KeyCounter("I" & obj.ID) > 0 Then
            strTwinId = oRepertoryTwins.Entry("I" & obj.ID).Object(0)
            WriteNarrativeTwin oRepertoryTwins.Entry(strTwinId).Object(0)
        End If
    End Select
End Sub

Sub WriteNarrativeTwin(oTwinLink)
    Dim oTwin, oTwins, fSomeAlive, cMales, cFemales, strGender
    set oTwins = oTwinLink.Siblings.ToGenoCollection
    fSomeAlive = True
    cMales = 0
    cFemales = 0
    For Each oTwin in oTwins
        fSomeAlive = fSomeAlive And Not oTwin.IsDead
        If oTwin.Gender.ID = "M" Then cMales = cMales + 1
        If oTwin.Gender.ID = "F" Then cFemales = cFemales + 1
    Next
    If cMales = oTwins.Count then strGender = "M"
    If cFemales = oTwins.Count then strGender = "F"
    Report.WritePhraseDic "PhTwins", StrHtmlNarrativeNamesShort(oTwins), _
                            fSomeAlive = True, Util.FirstNonEmpty(Dic.Peek(oTwinlink.TwinLink.ID & "_" & strGender), oTwinlink.TwinLink), Dic.PlurialCardinal(GetDicMFU("Twins", strGender), oTwins.Count)
    WriteHtmlAnnotation oTwinLink, Util.FormatPhrase(StrDicExt("AnnotationTwins","","{\U}{0} Notes","",""), Dic.PlurialCardinal("Twins", oTwins.Count)), oTwinLink.Comment
    WriteHtmlFootnoteRefs oTwinLink.Sources
End Sub

' Return a string containing the collection summary by male, females and unknown gender, followed by the names in HTML format.
' Example: "a brother and two sisters, named Benoit, Anne and Estelle"
Function StrHtmlCollectionMFU(coll, strDicPrefix, strHtmlNames)
    Dim i, cMales, cFemales, cPets
    For Each i In coll
        Select Case i.Gender.ID
        Case "M"
            cMales = cMales + 1
        Case "F"
            cFemales = cFemales + 1
        Case "P"
            cPets = cPets + 1
        End Select
    Next
    StrHtmlCollectionMFU = Dic.FormatPhrase("PhCollectionMFU", _
        Dic.PlurialCardinal(strDicPrefix &"_M", cMales), _
        Dic.PlurialCardinal(strDicPrefix & "_F", cFemales), _
        Dic.PlurialCardinal(strDicPrefix & "_", coll.Count - cMales - cFemales - cPets), _
    strHtmlNames, _
    coll.Count - cFemales = 0, _
    coll.Count - cPets > 1)
End Function

Function StrHtmlCollectionSpouses(collSpouses)
    StrHtmlCollectionSpouses = StrHtmlCollectionMFU(collSpouses, "Spouse", StrHtmlNarrativeNamesFull(collSpouses))
End Function
Function StrHtmlCollectionOtherSpouses(collSpouses)
    StrHtmlCollectionOtherSpouses = StrHtmlCollectionMFU(collSpouses, "OtherSpouse", StrHtmlNarrativeNamesFull(collSpouses))
End Function
Function StrHtmlCollectionChildren(collChildren)
    StrHtmlCollectionChildren = StrHtmlCollectionMFU(collChildren , "Child", StrHtmlNarrativeNamesShort(collChildren))    
End Function
Function StrHtmlCollectionOtherChildren(collChildren)
    StrHtmlCollectionOtherChildren = StrHtmlCollectionMFU(collChildren , "ChildOther", StrHtmlNarrativeNamesShort(collChildren))    
End Function
Function StrHtmlCollectionStepChildren(collChildren)
    StrHtmlCollectionStepChildren = StrHtmlCollectionMFU(collChildren , "ChildStep", StrHtmlNarrativeNamesShort(collChildren))    
End Function
Function StrHtmlCollectionChildrenLocal(collChildren, strType)
    StrHtmlCollectionChildrenLocal = StrHtmlCollectionMFU(collChildren , "Child" & strType, StrHtmlNarrativeNamesShortLocal(collChildren))    
End Function

Function StrHtmlNarrativeNamesFull(coll)
    StrHtmlNarrativeNamesFull = StrHtmlNarrativeNames(coll, "NameFull", "", "")
End Function
Function StrHtmlNarrativeNamesShort(coll)
    StrHtmlNarrativeNamesShort = StrHtmlNarrativeNames(coll, "NameShort", "", "")
End Function
Function StrHtmlNarrativeNamesShortLocal(coll)
    StrHtmlNarrativeNamesShortLocal = StrHtmlNarrativeNames(coll, "NameShort", "#", "onclick='javascript:explorerTreeOpen(""Children"",""2"");'")
End Function

Function StrHtmlNarrativeName(i, strNameForm, strHrefPrefix, strAttribs)
    Dim strName, strTitle, strNameFull
    strName = i.Session(strNameForm)
    strTitle = ""
    If (strName <> "") Then
        strNameFull = i.Session("NameFull")
        If (strName <> strNameFull) Then
            ' The name is different from the full name
                strTitle = Util.FormatString(" title='{&t}'", strNameFull)
        End If
        StrHtmlNarrativeName = Util.FormatString("<a href='{0}{1}' {2}{3}>{4&t}</a>", strHrefPrefix, i.Href, strAttribs, strTitle, strName)
    ElseIf Util.IsNothing(i) Then
        StrHtmlNarrativeName = ""
    Else
        StrHtmlNarrativeName = Util.HtmlEncode(StrDicMFU("_NoName",i.Gender.ID))
    End If
End Function

' Return a string containing the names of the collection of individuals.
' This string contains HTML tags, so it must be used with Report.Write() or with the WritePhrase's argument "{0h}".
' The implementation uses the Report buffer as a temporary storage, and extract the string from that buffer.
Function StrHtmlNarrativeNames(coll, strNameForm, strHrefPrefix, strAttribs)
    Dim i, iElement, iElementLast
    iElementLast = coll.Count- 1    ' Get the index of the last element in the collection    
    
    Dim cchBuffer    ' Number of characters in the buffer stream before writing the collection
    cchBuffer = Report.BufferLength
    
    Dim fValidName    ' At least one name was valid in the collection
    Dim strSep        ' Separator between the elements
    Dim strName, strNameFull, strTitle
    For iElement = 0 To iElementLast
        Set i = coll(iElement)
        Report.Write strSep
        If (i.Href <> "") Then
            fValidName = True
            Report.Write StrHtmlNarrativeName(i, strNameForm, strHrefPrefix, strAttribs)
        Else
            Report.WriteText StrDicMFU("_NoName",i.Gender.ID)
        End If
        
        If (iElement < iElementLast - 1) Then
            strSep = ", "
        Else
            strSep = Dic("ConjunctionAnd")
        End If
    Next

    If (fValidName) Then
        ' Get the text from the Report buffer from where we started appending
        StrHtmlNarrativeNames = Report.Buffer(cchBuffer)
    End If
    Report.BufferLength = cchBuffer        ' Truncate the buffer to its original size
End Function

Sub WriteHtmlIndex(strClass, strHtmlIndex, strType)
        Dim strFmtTemplate
    If (strHtmlIndex <> "") Then
            strFmtTemplate=Util.FormatString("<li class='xT2-{}'", Util.IfElse(fTreeOpen,"o","c")) & " onclick='xTclk(event,""2"");'><a name='{}'></a><h3 class='xT-i inline'>{}</h3><ul>"
        Report.WriteFormattedLn strFmtTemplate, strClass & strType, StrDicExt("TocIndex" & strType & Util.IfElse(Right(strClass,1) = "y", Left(strClass,Len(strClass) - 1) & "ies",strClass &"s"), "", strClass & " " & strType, "", "2.0.1.6/7")
        Report.Write strHtmlIndex
        Report.WriteLn "</ul></li>"
    End If
End Sub


' Write the individual for the Table of Contents
' The text contains HTML hyperlinks to the individual's report page
Sub WriteHtmlTocIndividual(i, hyperlink)
    Dim strFmtTemplate
    If hyperlink Then
        Report.WriteFormatted "<a href='{}' onclick='tocExit();'>{}</a>", i.Href, StrHtmlHighlightName(i.Session("NameAlternative"))
    Else
        Report.WriteFormatted "{}", StrHtmlHighlightName(i.Session("NameAlternative"))
    End If
    Report.WritePhraseDic "FmtBirth", CustomDate(i.Birth.Date).Approximation, CustomDate(i.Birth.Date).Year, Dic("PhBC_" & i.Birth.CeremonyType.ID), CustomDate(i.Birth.Baptism.Date).Approximation, CustomDate(i.Birth.Baptism.Date).Year
End Sub

Sub WriteHtmlTocIndividualImgGender(i)
    Report.Write StrHtmlImgGender(i)
    WriteHtmlTocIndividual i, True
End Sub

Sub WriteHtmlTocIndividualContact(fTreeOpen, i, fPicturesOnly, fContactsOnly, fChartsOnly)

    Dim p, strSummary, cchBufferIndividual, cchBufferContacts

    If fPicturesOnly And Not fContactsOnly Then
        If i.Session("PicturesIncluded") > 0 Then
            Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"")'>", Util.IfElse(fTreeOpen,"o","c")
            WriteHtmlTocIndividualImgGender i
            Report.Writeln "<ul>"
            For Each p in i.Pictures
                Report.Write3Ln "<li class='xT-b'>", StrHtmlImgPhoto(p), "</li>"
            Next
            Report.Writeln "</ul></li>"
        End If
    ElseIf fContactsOnly Then
        If i.Contacts.Count > 0 Then
            cchBufferIndividual = Report.BufferLength
            Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"")'>", Util.IfElse(fTreeOpen,"o","c")
            Report.Write StrHtmlImgGender(i)
            WriteHtmlTocIndividual i, False
            strFmtTemplate = "<li class='xT-b'><a href='contacts.htm#{&t}' onclick='tocExit();'><img src='images/{}.gif' alt='{}'/> {&t}</a> " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "{}"
            Report.WriteLn "<ul>"
            cchBufferContacts = Report.BufferLength
            For Each c In i.Contacts
                 If Not fPicturesOnly Or c.Session("PicturesIncluded") > 0 Then
                    Report.WriteFormatted strFmtTemplate, c.ID, "contact2", Dic("Occupancy"), c.Session("Name"), c.References.Count, Util.IfElse(fPicturesOnly,"",StrHtmlImgPhoto(c))
                    If fPicturesOnly Then
                        For Each p in c.Pictures
                            Report.WriteBr
                            Report.Write3 "", StrHtmlImgPhoto(p), ""
                        Next
                    End If
                    Report.WriteLn "</li>"
                End If
            Next
            If Report.BufferLength = cchBufferContacts Then
                Report.BufferLength = cchBufferIndividual
            Else
                Report.WriteLn "</ul></li>"
            End If
        End If
    ElseIf fChartsOnly Then
        If IsTrue(CustomTag(i,"DescendantTreeChart"), False) Then
            Report.Write "<li class='xT-b'>"
            Report.Write StrHtmlImgGender(i)
            Report.WriteFormatted "<a href='descendants/DescendantTree.htm?tree={}' target='popup' onclick='tocExit();' alt='{2}' title='{2}'>{1}</a>", i.ID & ".xml", StrHtmlHighlightName(i.Session("NameAlternative")), Dic("AltDescendantTreeChart")
            Report.WritePhraseDic "FmtBirth", CustomDate(i.Birth.Date).Approximation, CustomDate(i.Birth.Date).Year, Dic("PhBC_" & i.Birth.CeremonyType.ID), CustomDate(i.Birth.Baptism.Date).Approximation, CustomDate(i.Birth.Baptism.Date).Year
            Report.WriteLn "</li>"
        End If
    Else
        Report.Write "<li class='xT-b'>"
        WriteHtmlTocIndividualImgGender i
        Report.Write StrHtmlImgPhoto(i)
        Report.WriteLn "</li>"
    End If
End Sub


' Write the HTML code to generate the Table of Contents for the individuals
' The parameter fPictureOnly is to exclude any individual without a picture
Sub WriteHtmlTocIndividuals(fTreeOpen, fPicturesOnly, fContactsOnly, fChartsOnly)

    Dim strFmtTemplate, strFmtTemplate1, strNameLast, cchBufferLetter, cchBufferNameLast, cchBufferNames
    strFmtTemplate = "<a name='{0&t}'></a><span class='xT-i boldu'>{0&t}</span>"
    strFmtTemplate1 = strFmtTemplate
    If Not fPicturesOnly And Not fContactsOnly  And Not fChartsOnly Then
        strFmtTemplate = strFmtTemplate & "&nbsp; " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16")    ' Include the count of individuals
    End If
    
    Dim oRepertoryIndividuals, o, oRepertoryFamilies, oStringDictionaryNames, oFamily, i, iCount, strTemp
    iCount = Session("IndividualsCount")
    Set oRepertoryIndividuals = Session("oRepertoryIndividuals")
    Set oStringDictionaryNames = Session("oStringDictionaryNames")
    If (Not Util.IsNothing(oRepertoryIndividuals)) Then
        strFmtTemplate = Util.FormatString("<li class='xT2-{}' onclick='xTclk(event,""2"")'>&nbsp;", Util.IfElse(fTreeOpen,"o","c")) & strFmtTemplate & "<ul>"
    
        For Each o In oRepertoryIndividuals
            strFirstChar = o.Key
            cchBufferLetter =  Report.BufferLength
            Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'>&nbsp;<span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar        
            cchBufferFamilies =  Report.BufferLength
            Set oRepertoryFamilies = o.Object(0)
            For Each oFamily In oRepertoryFamilies 
                cchBufferNameLast = Report.BufferLength        ' Remember the position where the last name was written
                strNameLast = oFamily.Key
                Report.WriteFormatted strFmtTemplate, Replace(strNameLast & oStringDictionaryNames.KeyValue(strNameLast), " ", "_"), Util.FormatPhrase(StrDicExt("_PhPluralCount","","{0} {1}", "", "2.0.1.6"), oFamily.Count, Dic.Plurial("Individual", oFamily.Count))
                cchBufferNames = Report.BufferLength
                For Each i In oFamily
                    WriteHtmlTocIndividualContact fTreeOpen, i, fPicturesOnly, fContactsOnly, fChartsOnly
                Next
                If (Report.BufferLength = cchBufferNames) Then
                    ' No name have been written, so flush the last name from the output stream.  This situation may happens when we want only individuals with pictures
                    Report.BufferLength = cchBufferNameLast
                Else
                    Report.WriteLn "</ul></li>"
                End If
            Next
            If Report.BufferLength = cchBufferFamilies Then
                ' No families written with this initial letter so flush then letter header from the output stream
                Report.BufferLength = cchBufferLetter
            Else
                Report.WriteLn "</ul></li>"
            End If
        Next
    
        Set oRepertoryNoLastName = Session("oRepertoryNoLastName")
        If (oRepertoryNoLastName.Count > 0) Then
            cchBufferNameLast = Report.BufferLength
            Report.WriteFormattedLn Replace(strFmtTemplate,strFmtTemplate1,"<b>{0&t}</b>"),  Dic("_NoName"), Util.FormatPhrase(StrDicExt("_PhPluralCount","","{0} {1}", "", "2.0.1.6"), oRepertoryNoLastName.Count, Dic.Plurial("Individual", oRepertoryNoLastName.Count))
            cchBufferNames = Report.BufferLength
            For Each o In oRepertoryNoLastName
                For iCount = 0 To o.Count-1
                    Set i = o.Object(iCount)
                    WriteHtmlTocIndividualContact fTreeOpen, i, fPicturesOnly, fContactsOnly, fChartsOnly
                Next
            Next
            If (Report.BufferLength = cchBufferNames) Then
                Report.BufferLength = cchBufferNameLast
            Else
                If Session("fUseTreeIndexes") Then Report.WriteLn "</ul></li>"
            End If            
        End If
    End If
End Sub

Sub WriteHtmlTocFamilies(fTreeOpen, fPicturesOnly)

    Dim collFamiliesSorted, f, strNameParent, strNameParentPrev, strFirstChar, strFirstCharPrev, strName, strPrefix
    Set collFamiliesSorted = Session("collFamiliesSorted")
    
    If (Not Util.IsNothing(collFamiliesSorted)) Then
        Dim cchBufferStart    ' Start of the write operations
        Dim cchBufferValid    ' Previous valid buffer.  If a sequence is not valid, flush the useless family name
        Dim cchBufferLetter    ' Start of letter group
        Dim cchBufferFamliies    ' Start of Family group
        cchBufferStart = Report.BufferLength
        cchBufferValid = cchBufferStart    ' What was already written to the buffer is obviously valid
        Session("BufferBegin") = cchBufferStart

        strFmtTemplateName = "<li>{}<a href='{}' onclick='tocExit();'>{&t}</a>{}"
        strFmtTemplateFamily = "<span class='xT-i boldu'>{&t}</span>"
        strFmtTemplateFamily1 = strFmtTemplateFamily


        strFmtTemplateFamily = Util.FormatString("<li class='xT2-{}' onclick='xTclk(event,""2"");'>", Util.IfElse(fTreeOpen,"o","c")) & strFmtTemplateFamily & "<ul>"

        For Each f In collFamiliesSorted
            Set pp = f.Pictures.Primary
            If Not fPicturesOnly Or Not Util.IsNothing(pp) Then

                strNameParent = f.Parents(0).Session("NameLast")
                If strNameParent = "" Then strNameParent = f.Parents(1).Session("NameLast")
                strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(strNameParent)))
                If strNameParent = "" Then
                    strFirstChar = Dic("_NoName")
                End If
                If Session("fUseTreeIndexes") And (strFirstChar <> strFirstCharPrev) Then
                    ' We have a different initial letter
                    If strFirstCharPrev <> "" Then
                            Report.WriteLn "</ul></li></ul></li>"
                    End If
                    Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar
                    strNameParentPrev = ""
                End If
                strFirstCharPrev = strFirstChar
                    
                If (strNameParent <> strNameParentPrev) Then
                    ' We have a different family name than previous family name
                    If strNameParentPrev <> "" Then
                        If Session("fUseTreeIndexes") Then
                            Report.WriteLn "</ul></li>"
                        Else
                            Report.WriteBr
                            End If
                    End If
                    strNameParentPrev = strNameParent
                    If StrNameParent <> "" Then Report.WriteFormatted strFmtTemplateFamily, strNameParent
                End If
                If Not fPicturesOnly Then
                    Report.WriteFormatted strFmtTemplateName, StrHtmlImgFamily(f), f.Href, f.Session("Name"), StrHtmlImgPhoto(f)
                ElseIf Not Util.IsNothing(pp) Then
                    Report.WriteFormattedLn strFmtTemplateName, StrHtmlImgFamily(f), f.Href, f.Session("Name"), ""
                  Report.WriteLn "<ul>"
                    For Each p in f.Pictures
                        Report.Write3Ln "<li>", StrHtmlImgPhoto(p), "</li>"
                    Next
                  Report.WriteLn "</ul>"
                End If
                Report.WriteLn "</li>"
            End If
        Next
        If    strNameParentPrev <> "" And _
                strNameParent <> Dic("_NoName") Then
            If Session("fUseTreeIndexes") Then
                Report.Write "</ul></li>"
            Else
                Report.WriteBr
            End If
        End If
        If strFirstCharPrev <> "" Then
            If Session("fUseTreeIndexes") Then Report.WriteLn "</ul></li>"
        End If

    End If
    
End Sub
Sub WriteHtmlTocEducations(fTreeOpen, fPicturesOnly)
    Dim strFirstChar, strFirstCharPrev, strName, strNamePrev, e, pic

    Dim oDataSorter, collEducations

    Set oDataSorter = Util.NewDataSorter()
    For Each e in Educations
        oDataSorter.Add e, Util.FirstNonEmpty(e.Institution, e.Place.Session("NameFull"))
    Next

    oDataSorter.SortByKey
    Set collEducations = oDataSorter.ToGenoCollection

    If Session("fUseTreeIndexes") Then
        For Each e in collEducations
            strName = Util.FirstNonEmpty(e.Institution, e.Place.Session("NameFull"))
            If Not fPicturesOnly or e.Session("PicturesIncluded") > 0 Then
                strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(strName)))
                if strFirstChar = "" Then strFirstChar = " "
                if strFirstChar <> strFirstCharPrev Then
                    If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
                    Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar
                    strFirstCharPrev = strFirstChar
                End If

                Report.WriteformattedLn "<li class='xT-b'><img src='images/education.gif' class='icon' />&nbsp;{}" & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "</li><ul>", strName, e.References.Count
                For Each d in e.References
                    Report.Write StrHtmlImgGender(d)
                    Report.WriteFormattedBr"<a href='{}#{}' onclick='tocExit();'>{&t}</a>", d.Href, e.ID, d.Name
                Next
                If fPicturesOnly Then
                    For Each pic in e.Pictures
                        Report.Write3Br "", StrHtmlImgPhoto(pic), ""
                    Next
                End If
                Report.WriteLn "</ul>"
            End If
        Next
        If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
    End If
End Sub

Sub WriteHtmlTocEntities(SocialEntities, fPicturesOnly, fContactsOnly)
    Dim strFirstChar, strFirstCharPrev, strName, strLang, strNamePrev, s, p, cchBufferEntity, cchBufferContacts, cchBufferLetter, cchBufferEnties
    Dim oDataSorter, collSocialEntities

    Set oDataSorter = Util.NewDataSorter()
    For Each s in SocialEntities
		strLang = CustomTag(s,"Language")
        If s.Session("Name") <> "" And IsFalse(CustomTag(s, "IsLabel"), False) And (strLang = "" Or strLang = Session("ReportLanguage")) Then
			oDataSorter.Add s, s.Session("Name")
		End If
    Next

    oDataSorter.SortByKey
    Set collSocialEntities = oDataSorter.ToGenoCollection

    For Each s in collSocialEntities
	   strName = s.Session("Name")
	   If (Not fPicturesOnly And Not fContactsOnly) Or _
					(fPicturesOnly And s.Session("PicturesIncluded") > 0) Or _
			(fContactsOnly And s.Contacts.Count > 0) Then
			strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(strName)))
		End If
		If strFirstChar <> strFirstCharPrev Then
			If strFirstCharPrev <> "" Then
				If Report.BufferLength > cchBufferEntities Then
					Report.WriteLn "</ul></li>"
				Else
					Report.BufferLength = cchBufferLetter
				End If
			End If
			cchBufferLetter = Report.BufferLength
			Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar 
			strFirstCharPrev = strFirstChar
							cchBufferEntities = Report.BufferLength
		End If
	If Not fContactsOnly Then
	   Report.WriteLn "<li class='xT-b'>"
			 Report.WriteFormatted "<img src='images/entity.gif' border='0' width='16' height='16' alt='{}'/> <a href='entities.htm#{}'>{&t}</a>", StrDicExt("AltEntityImage","","Social Entity","",""), s.ID, s.Session("Name")
			 If Not fPicturesOnly Then
					 Report.WriteBr StrHtmlImgPhotoLink(s, "Entities.htm")
			 Else
					 Report.WriteBr
					 For Each p in s.Pictures
				   Report.Write3Br "", StrHtmlImgPhoto(p), ""
			   Next
			 End If
			 Report.WriteLn "</li>"
		Else
			cchBufferEntity = Report.BufferLength
			Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"")'>", Util.IfElse(fTreeOpen,"o","c")
			Report.WriteFormatted "<img src='images/entity.gif' border='0' width='16' height='16' alt='{}'/> <a href='entities.htm#{}'>{&t}</a>", StrDicExt("AltEntityImage","","Social Entity","",""), s.ID, s.Session("Name")
			strFmtTemplate = "<li class='xT-b'><a href='contacts.htm#{&t}' onclick='tocExit();'><img src='images/{}.gif' alt='{}'/> {&t}</a> " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "{}"
			Report.WriteLn "<ul>"
			cchBufferContacts = Report.BufferLength
			For Each c In s.Contacts
				 If Not fPicturesOnly Or c.Session("PicturesIncluded") > 0 Then
							  Report.WriteFormatted strFmtTemplate, c.ID, "contact2", Dic("Occupancy"), c.Session("Name"), c.References.Count, Util.IfElse(fPicturesOnly,"",StrHtmlImgPhoto(c))
								 If fPicturesOnly Then
									  For Each p in c.Pictures
												 Report.WriteBr
											   Report.Write3 "", StrHtmlImgPhoto(p), ""
									  Next
							End If
							Report.WriteLn "</li>"
				 End If
			Next
			If Report.BufferLength = cchBufferContacts Then
							 Report.BufferLength = cchBufferIndividual
					 Else
				 Report.WriteLn "</ul></li>"
			End If
		End If
    Next
    If strFirstCharPrev <> "" Then
            If Report.BufferLength > cchBufferEntities Then
                          Report.WriteLn "</ul></li>"
                Else
                        Report.BufferLength = cchBufferLetter
                End If
        End If
End Sub

Sub WriteHtmlTocLabels(Labels, fPicturesOnly)
    Dim strFirstChar, strFirstCharPrev, strName, strNamePrev, s

    Report.WriteLn "<ul id='names' class='xT'>"
    Labels.Sortby("Text")
    For Each s in Labels
        strLang = CustomTag(s,"Language")
        If (Not fPicturesOnly Or s.Session("PicturesIncluded") > 0) And s.Text <> "" And (strLang = "" Or strLang = Session("ReportLanguage")) Then
            strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(s.Text)))
            If strFirstChar <> strFirstCharPrev Then
                If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
                Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar 
                strFirstCharPrev = strFirstChar
            End If
            Report.WriteFormatted "<img src='images/entity.gif' border='0' width='16' height='16' alt='{}'/> <a href='labels.htm#{}'>{&t}</a>", StrDicExt("AltLabelImage","","Text Label","",""), s.ID, s.Session("Name")
            Report.WriteBr StrHtmlImgPhotoLink(s, "Labels.htm")
        End If
    Next
    If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
     Report.WriteLn "</ul>"
End Sub

Sub WriteHtmlTocRelationships(Relationships, fTreeOpen)
    Dim strFirstChar, strFirstCharPrev, strName, strNamePrev, o, pic, strComments

    If Session("fUseTreeIndexes") Then
        For Each r in Relationships
            If r.Session("PicturesIncluded") > 0 Then
                If r.Class = "EmotionalRelationship" Then
                    strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(r.EmotionalLink.ID)))
                Else
                    strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(r.Connection.ID)))
                End If
                if strFirstChar = "" Then strFirstChar = " "
                if strFirstChar <> strFirstCharPrev Then
                    If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
                        Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar
                    strFirstCharPrev = strFirstChar
                End If
                Select Case r.Class
                Case "EmotionalRelationship"
                    Report.WriteformattedLn "<li class='xT-b'><img src='images/yinyan.gif' class='icon' />&nbsp;{}</li><ul>", r.EmotionalLink
                    Report.Write StrHtmlImgGender(r.Entity1)
                    Report.WriteFormattedLn"<a href='{}' onclick='tocExit();'>{&t}</a>,&nbsp;", r.Entity1.Href, Util.FirstNonEmpty(r.Entity1.Name, StrDicMFU("_NoName", r.Entity1.Gender.ID))
                    Report.Write StrHtmlImgGender(r.Entity2)
                    Report.WriteFormattedBr"<a href='{}' onclick='tocExit();'>{&t}</a>", r.Entity2.Href, Util.FirstNonEmpty(r.Entity2.Name, StrDicMFU("_NoName", r.Entity2.Gender.ID))
                Case Else
                    Report.WriteformattedLn "<li class='xT-b'><img src='images/yinyan.gif' class='icon' />&nbsp;{&t}</li><ul>", Util.FirstNonEmpty(r.Connection, Dic("_NoName"))
                    WriteHtmlRelationshipEntity r.entity1, r
                    Report.Write ",&nbsp;"
                    WriteHtmlRelationshipEntity r.entity2, r
                    Report.WriteBr
                End Select
                For Each pic in r.Pictures
                    Report.Write3Br "", StrHtmlImgPhoto(pic), ""
                Next
                Report.WriteLn "</ul>"
            End If
        Next
        If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
    End If
End Sub

Sub WriteHtmlRelationshipEntity(e, r)
    Select Case e.Class
    Case "Individual"
        Report.Write StrHtmlImgGender(e)
        If e.Name <> "" Then
            Report.WriteFormatted "<a href='{}' onclick='javascript:hidePopUpFrame();'>{&t}</a>", e.Href, e.Name
        Else
            Report.WriteText StrDicMFU("_NoName", e.Gender.ID)
        End If
    Case "SocialEntity"
        Report.Write Replace(e.Text, vbLf, " ")
    End Select
End Sub

Sub WriteHtmlTocOccupations(fTreeOpen, fPicturesOnly)
    Dim strFirstChar, strFirstCharPrev, strName, strNamePrev, o, pic

    Occupations.SortBy("Place Company")

    If Session("fUseTreeIndexes") Then
        For Each o in Occupations
            If Not fPicturesOnly or o.Session("PicturesIncluded") > 0 Then
                strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(o.Place.Name)))
                if strFirstChar = "" Then strFirstChar = " "
                if strFirstChar <> strFirstCharPrev Then
                    If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
                    Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar
                    strFirstCharPrev = strFirstChar
                End If

                Report.WriteformattedLn "<li class='xT-b'><img src='images/occupation.gif' class='icon' />&nbsp;{}" & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "</li><ul>", o.Place & Util.IfElse(o.Company <> "", ", " & o.Company,""), o.References.Count
                For Each d in o.References
                    Report.Write StrHtmlImgGender(d)
                    Report.WriteFormattedBr"<a href='{}#{}' onclick='javascript:hidePopUpFrame();'>{&t}</a>", d.Href, o.ID, d.Name
                Next
                If fPicturesOnly Then
                    For Each pic in o.Pictures
                        Report.Write3Br "", StrHtmlImgPhoto(pic), ""
                    Next
                End If
                Report.WriteLn "</ul>"
            End If
        Next
        If strFirstCharPrev <> "" Then     Report.WriteLn "</ul></li>"
    End If
End Sub

Sub WriteHtmlTocPlaces(fTreeOpen, fPicturesOnly)
    Dim strFirstChar, strFirstCharPrev, strName, strNamePrev, p, oDataSorter, fChild, collPlaces, ptrStart, ptrBegin

    Set oDataSorter = Util.NewDataSorter()
    For Each p in Places
        If  p.Session("References") > 0 Then oDataSorter.Add p, p.Session("NameFull")
    Next

    oDataSorter.SortByKey
    Set collPlaces = oDataSorter.ToGenoCollection

    If Session("fUseTreeIndexes") Then
        For Each p in collPlaces
            If Not fPicturesOnly or p.Session("PicturesIncluded") > 0 Then
                strFirstChar = Util.StrStripAccentsUCase(Util.StrStripPunctuation(Util.StrGetFirstChar(p.Session("NameFull"))))
                If strFirstChar = "" Then strFirstChar = " "
                If strFirstChar <> strFirstCharPrev Then
                    If strFirstCharPrev <> "" Then
            If Report.BufferLength > ptrBegin Then
               Report.WriteLn "</ul></li>"
            Else
                Report.BufferLength = ptrStart
            End If
          End If
                ptrStart = Report.BufferLength
                    Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar
            ptrBegin = Report.BufferLength
                  strFirstCharPrev = strFirstChar
                End If
        fChild =  Not Util.IsNothing(p.Parent)
                If Not fChild Or _
            Session("fIndexChildPlacesAndSources") Or _
            (fPicturesOnly And p.Session("PicturesIncluded") > 0) _
        Then 
             AddPlaceNode p, (fChild And p.Children.Count = 0), fTreeOpen, fPicturesOnly
        End If
            End If
        Next
        If strFirstCharPrev <> "" Then
      If Report.BufferLength > ptrBegin Then
         Report.WriteLn "</ul></li>"
      Else
          Report.BufferLength = ptrStart
      End If
    End If
    Else
        For Each p In collPlaces
            If Not fPicturesOnly or p.Session("PicturesIncluded") > 0 Then
                strName = p.Session("NameFull")
                Report.WriteFormatted "<li class='xT-b'><img src='images/place.gif' class='icon' alt='{}'/> <a href='places.htm#{&t}' title='{&t}' onclick='tocExit();'>{&t}</a> " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "", Dic("Place"), p.ID, p.Category, strName, p.References.Count
                Report.WriteLn StrHtmlImgPhotoLink(p, "places.htm") & "</li>"
            End If
        Next
    End If
End Sub

Sub WriteHtmlTocSources(fTreeOpen, fPicturesOnly)
    Dim strFirstChar, strFirstCharPrev, strName, strNamePrev, s, collSourcesAndCitiations, oDataSorter, ptrStart, ptrBegin
    strFirstCharPrev = ""
    strNamePrev = ""

    Set oDataSorter = Util.NewDataSorter()

    ' Add each source to the DataSorter
    For Each s In SourcesAndCitations
        ' do not report source if no references
        If s.Session("References") > 0 Then
            strName=UCase(Util.StrStripAccentsUCase(Util.StrStripPunctuation(StrPlainText(s, s.Title))))
            oDataSorter.Add s, strName, s.Subtitle, s.Description, s.WhereInSource
        End If
    Next

    ' Sort the sources according the sort keys
    oDataSorter.SortByKey
    Set collSourcesAndCitations = oDataSorter.ToGenoCollection

    If Session("fUseTreeIndexes") Then
        For Each s in collSourcesAndCitations
            If Not fPicturesOnly or s.Session("PicturesIncluded") > 0 Then
                strFirstChar = Util.StrGetFirstChar(Util.StrStripAccentsUCase(Util.StrStripPunctuation(StrFormatText(s, StrParseText(s, True)))))
                if strFirstChar = "" Then strFirstChar = " "
                if strFirstChar <> strFirstCharPrev Then
                    If strFirstCharPrev <> "" Then
            If Report.BufferLength > ptrBegin Then
               Report.WriteLn "</ul></li>"
            Else
                Report.BufferLength = ptrStart
            End If
          End If
                    ptrStart = Report.BufferLength
                    Report.WriteFormattedLn "<li class='xT2-{}' onclick='xTclk(event,""2"");'> <span class='xT-i bold'>{&t}</span><ul>", Util.IfElse(fTreeOpen,"o","c"), strFirstChar
                    ptrBegin = Report.BufferLength
                    strFirstCharPrev = strFirstChar
                End If
        fChild =  Not Util.IsNothing(s.Parent)
                If Not fChild Or _
               Session("fIndexChildPlacesAndSources") Or _
               fPicturesOnly And s.Session("PicturesIncluded") > 0 _
        Then
               AddSourceCitationNode s, (s.Children.Count = 0) And fChild, fTreeOpen, fPicturesOnly
        End If
            End If
        Next
        If strFirstCharPrev <> "" Then
      If Report.BufferLength > ptrBegin Then
         Report.WriteLn "</ul></li>"
      Else
          Report.BufferLength = ptrStart
      End If
    End If
    Else
        For Each s In collSourcesAndCitations
            If Not fPicturesOnly or s.Session("PicturesIncluded") > 0 Then
                strName = JoinSourceCitationNames(s, StrFormatText(s, StrParseText(s.title, True)), true)
                Report.WriteFormatted "<img src='images/source.gif' border='0' width='16' height='16' alt='{}'/> <a href='sources.htm#{&t}' title='{&t}' onclick='hidePopUpFrame();'>{&t}</a> " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "", Dic("Source"), s.ID, s.Subtitle, strName, s.References.Count
                Report.WriteBr StrHtmlImgPhotoLink(s, "sources.htm")
            End If
        Next
    End If
End Sub

Function AddPlaceNode(p, level, fTreeOpen, fPicturesOnly)
    Dim strHtmlPlace, oDataSorter, pchild, pic, strName, strImgLink, collPlaces
    If level = 0 Then
        strName = p.Session("NameFull")
    Else
        strName = p.Session("NameShort")
    End If
    If Not fPicturesOnly Then strImgLink = StrHtmlImgPhotoLink(p, "places.htm")
    strHtmlPlace = Util.FormatString("<img src='images/place.gif' class='icon' alt=''/>&nbsp;<a href='places.htm#{&t}' title='{&t}' onclick='tocExit();'>{&t}</a>" & strImgLink, p.ID, p.Category, strName)
    If Not Util.IsNothing(p.Children) And level < 20 And level > -1 Then
        Report.WriteFormattedLn "<li class='xT-i5 xT2-{}' onclick='xTclk(event,""2"")'>" & strHtmlPlace, Util.IfElse(fTreeOpen,"o","c")
        Report.WriteFormattedLn " " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "", Dic.FormatPhrase("PhUsageCounters",p.Session("References"), p.Children.Count)
        Report.WriteLn "<ul>"
        If fPicturesOnly Then
            For Each pic in p.Pictures
                    Report.Write3Br "<img src='images/space.gif' width='16' alt=''/>", StrHtmlImgPhoto(pic), ""
            Next
        End If
        Set oDataSorter = Util.NewDataSorter()
        For Each pchild in p.Children
            If  pchild.Session("References") > 0 Then oDataSorter.Add pchild, pchild.Session("NameFull")
        Next
        oDataSorter.SortByKey
        Set collPlaces = oDataSorter.ToGenoCollection
        
        For Each pchild in collPlaces
            If pchild.Session("References") > 0 Then AddPlaceNode pchild, level+1, fTreeOpen, fPicturesOnly
        Next
        Report.WriteLn "</ul></li>"
    Else
        Report.WriteformattedLn "<li class='xT-bi'>{} " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "</li>", strHtmlPlace, p.References.Count
        If fPicturesOnly Then
            Report.WriteLn "<ul>"
            For Each pic in p.Pictures
                    Report.Write3Ln "<li>", StrHtmlImgPhoto(pic), "</li>"
            Next
            Report.WriteLn "</ul>"
        End If
    End If
End Function

Function JoinPlaceNames(p, strName, level)
    If Session("fJoinPlaceNames") And Not Util.IsNothing(p.Parent) And level < 11 Then
        JoinPlaceNames = JoinPlaceNames(p.Parent, strName & ", " & StrPlaceTranslate(p.Parent.Name), level+1)
    ElseIf level > 10 Then
        Report.LogError ConfigMsg("ErrorPlaceLoop","Error: Place hierarchy too deep or looping ", "2014.02.25") & strName
        JoinPlaceNames = strName
    Else
        JoinPlaceNames = strName
    End If
End Function

Function JoinTranslatedPlaceNames(p, strName, level)
    If Not Util.IsNothing(p.Parent) And level < 11 Then
        JoinTranslatedPlaceNames = JoinTranslatedPlaceNames(p.Parent, strName & ", " & StrPlaceTranslate(p.Parent.Name), level+1) 
    ElseIf level > 10 Then
        Report.LogError ConfigMsg("ErrorPlaceLoop","Error: Place hierarchy too deep or looping ", "2014.02.25") & strName
        JoinTranslatedPlaceNames = strName
    Else
        JoinTranslatedPlaceNames = strName
    End If
End Function

Function JoinOriginalPlaceNames(p, strName, level)
    If  Not Util.IsNothing(p.Parent) And level < 11 Then
        JoinOriginalPlaceNames = JoinOriginalPlaceNames(p.Parent, strName & "; " & p.Parent.Name, level+1) 
    ElseIf level > 10 Then
        Report.LogError ConfigMsg("ErrorPlaceLoop","Error: Place hierarchy too deep or looping ", "2014.02.25") & strName
        JoinOriginalPlaceNames = strName
    Else
        JoinOriginalPlaceNames = strName
    End If
End Function

Function JoinSourceCitationNames(s, strName, level)
    Dim strNewName, fCitation
    strNewName = strName
    fCitation = false
    If  Not Util.IsNothing(s.Parent) Then
        If s.WhereInSource <> "" and strName = s.Parent.title Then
            strNewName = "'" & s.WhereInSource & "'"
            fCitation = True
        End If
        If Session("fJoinSourceCitationNames") And Not Util.IsNothing(s.Parent) And level < 11 Then
            If fCitation Or (s.title <> s.Parent.title) Then strNewName = strNewName &  " @ " & StrFormatText(s, StrParseText(s.Parent.title, True))
                JoinSourceCitationNames = JoinSourceCitationNames(s.Parent, strNewName, level+1) 
            ElseIf level > 10 Then
                Report.LogError ConfigMsg("ErrorSourceLoop","Error: Source/Citation hierarchy too deep or looping ", "2014.02.25") & strName
                JoinSourceCitationNames = strNewName
            Else
                JoinSourceCitationNames = strNewName
        End If
    Else
        JoinSourceCitationNames = strNewName
    End If
End Function

Function AddSourceCitationNode(s, level, fTreeOpen, fPicturesOnly)
    If (level >= 10) Then
        Exit Function    ' Maximum of 10 levels, to prevent an infinite loop
    End If
    Dim strHtmlSourceCitation, p, schild, strName, strImgLink, collSourcesAndCitations, oDataSorter
    strName = JoinSourceCitationNames(s, StrFormatText(s, StrParseText(s.title, True)), level = 0)
    If Not fPicturesOnly Then strImgLink = StrHtmlImgPhotoLink(s, "sources.htm")
    strHtmlSourceCitation = Util.FormatString("<img src='images/source.gif' class='icon' title='' alt=''/>&nbsp;<a href='sources.htm#{&t}' title='{&t}' onclick='tocExit();'>{}</a>" & strImgLink, s.ID, s.MediaType, strName)
    strHtmlSourceCitation=Replace(Replace(strHtmlSourceCitation,"{","&#123;"),"}","&#125;")
    If Not Util.IsNothing(s.Children) Then
        Report.WriteFormattedLn "<li class='xT-i5 xT2-{}' onclick='xTclk(event,""2"")'>" & strHtmlSourceCitation, Util.IfElse(fTreeOpen,"o","c")
        Report.WriteFormattedLn " " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "", Dic.FormatPhrase("PhUsageCounters",s.Session("References"), s.Children.Count)
        If fPicturesOnly And s.Pictures.Count > 0 Then
            Report.WriteLn "<ul>"
            For Each p in s.Pictures
                Report.Write3Ln "<li>", StrHtmlImgPhoto(p), "</li>"
            Next
            Report.WriteLn "</ul>"
        End If
     ' Add each child source to the DataSorter
      Set oDataSorter = Util.NewDataSorter()
      For Each schild In s.Children
          ' do not report source if no references
          If schild.Session("References") > 0 Then
              strName=UCase(Util.StrStripAccentsUCase(Util.StrStripPunctuation(StrPlainText(schild, schild.Title))))
              oDataSorter.Add schild, strName, schild.Subtitle, schild.Description, schild.WhereInSource
          End If
      Next
  
      ' Sort the sources according the sort keys
      oDataSorter.SortByKey
      Set collSourcesAndCitations = oDataSorter.ToGenoCollection
      If collSourcesAndCitations.Count > 0 Then
        Report.WriteLn "<ul>"
            For Each schild in collSourcesAndCitations
                If (Not fPicturesOnly Or schild.Session("PicturesIncluded")>0) And schild.Session("References") > 0 Then
                    AddSourceCitationNode schild, level+1, fTreeOpen, fPicturesOnly
                End If
            Next
        Report.WriteLn "</ul>"
         End If
         Report.WriteLn "</li>"
    Else
        Report.WriteformattedLn "<li class='xT-bi'>{} " & StrDicExt("FmtCounter", "", "<small><bdo dir'ltr'> ({})</bdo></small>", "", "2011.02.16") & "</li>", strHtmlSourceCitation, s.References.Count
        If fPicturesOnly And s.Pictures.Count > 0 Then
            Report.WriteLn "<ul>"
            For Each p in s.Pictures
                Report.Write3Ln "<li>", StrHtmlImgPhoto(p), "</li>"
            Next
            Report.WriteLn "</ul>"
        End If
    End If
End Function

Sub WriteHtmlEntity(s)
    Dim collReferences, arrText, strLine, fFirst, strLang, i, iMin
    strLang = CustomTag(s, "Language")
    If strLang <> "" And strLang <> Session("ReportLanguage") Then Exit Sub
    Report.WriteFormattedLn "<a name='{&t}'></a><h4>{&t}</h4>", s.ID, s.Session("Name")
    If (s.Session("PicturesIncluded") > 0) Then
        Report.WriteLn "<div class='floatright aligncenter widthpaddedlarge'>"
        WriteHtmlPicturesLarge s, "left", "", Session("fHidePictureName") Or Session("fShowPictureDetails"), False
        Report.WriteLn "</div>"
    End If
    arrText = Split(s.Text, vbLf)
    iMin=0
    If Ubound(arrText)  > -1 Then If s.Session("Name")=arrText(0) Then iMin=1
    For i=iMin To Ubound(arrText)
        Report.WriteBr arrText(i)
    Next
    WriteHtmlExtraNarrative s
    WriteHtmlAnnotation s, StrDicExt("AnnotationSocialEntity","","Notes","",""), s.Comment
    WriteHtmlRelationships s
    WriteHtmlOccupancies(s)
    WriteHtmlAdditionalInformation(s)
    WriteHtmlReferences s, false
End Sub

Sub WriteHtmlOccupancies(obj)
    Dim collOccupancies, o, ich, strName, strRelative, strPnP, strType, fExtant, strGender, nPlural
    Set collOccupancies = obj.contacts.ToGenoCollection
    If Session("Flag_W") And (collOccupancies.Count > 0) Then
        Session("ReferencesStart") = -1 ' indicate at least one collapse/expand non-notes section present
        Report.WriteLn    "<div class='clearleft'><br /><ul class='xT'>"
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>", Util.IfElse(Session("fCollapseReferences"), "c", "o")
        Report.WriteFormattedLn "<a name='Occupancies'></a><h4 class='xT-i inline'>{&t}</h4><ul class='xT-h'><li>", Dic("HeaderOccupancy")
        Select Case obj.Class
            Case "Individual"
                strName = obj.Session("NameShort")
                strRelative = PnR(i)
                strPnP = PnP(i)
                fExtant = Not i.IsDead
                strType = ""
            Case "SocialEntity"
                strName = obj.Session("Name")
        strGender = Util.FirstNonEmpty(CustomTag(obj,"Gender.ID"), "N")
        nPlural = Util.IfElse(IsFalse(CustomTag(obj,"Plural"), False),1,2)
        strPnP = Dic.Plurial(Util.IfElse(Dic.Peek("PnP_" & strGender)<>"", "PnP_" & strGender,"PnP_"), nPlural)
        strRelative = Dic.Plurial(Util.IfElse(Dic.Peek("PnR_" & strGender)<>"", "PnR_" & strGender,"PnR_"), nPlural)
        fExtant = (CustomTag(obj, "Extant") <> "N")
                strType = CustomTag(obj, "Type")
        End Select
        For Each o In collOccupancies
            Report.WriteFormattedLn "<a name='{}'></a><div class='clearleft'>", o.ID
            ich = Report.BufferLength
            If (o.Pictures.Count > 0) Then WriteHtmlPicturesSmall o, "left", True
      ' added  & "[{?!8}{8}]" and end of phrase below for pre 2011.02.06 Dictionaries that will not have param 8 in phrase template
             Report.WritePhrase Util.FirstNonEmpty(StrDicOrTag("PhOT_" & o.Type.ID & "_" & obj.Class, CustomTag(o, "NarrativeStyle")), StrDicOrTag("PhOT_" & o.Type.ID, CustomTag(o, "NarrativeStyle")), _
                                        Dic.Lookup2("PhOT_" & o.Type.ID & "_" & obj.Class, "PhOT_" & o.Type.ID)) & "[{?!8}{8}]", _
                                        strName, strRelative, StrDateSpan(o.DateStart, o.DateEnd), _
                                        Util.IfElse(o.DateStart.ToStringNarrative <> o.DateEnd.ToStringNarrative, StrTimeSpan(o.Duration), ""), _
                                        (Not fExtant Or (o.DateEnd.ToStringNarrative<>""))=False, _
                                        o.Place.Session("HlinkLocative"), o.Summary, o.Place.Session("Hlink"), strType
            If Report.BufferLength > ich And o.Summary <> "" Then Report.WriteBr
            Report.WritePhraseDic "PhContact", o.telephone, o.Fax, Util.FormatHtmlHyperlink(Util.IfElse(o.Email <> "","mailto:" & o.Email,""), o.Email), Util.FormatHtmlHyperlink(o.Homepage, ,"target='_blank'"), o.Type, o.Place, o.Mobile
            WriteHtmlFootnoteRef o.Source
            WriteHtmlFootnoteRefs(o.Sources)
            WriteHtmlExtraNarrative o
            WriteHtmlAdditionalInformation(o)
            WriteHtmlAnnotation o, Dic("AnnotationOccupancy"), o.Comment
            strName = strPnP
            strType = ""
            Report.WriteLn "</div>"
        Next
        Report.WriteLn "</li></ul></li></ul></div>"
    End If
End Sub

Sub WriteHtmlSource(s)
    Dim collReferences, pub, ser, strClassCitation, fTemp, strUrl, strText, strFile
    Set pub = s.Publication
    Set ser = s.Series
    strClassCitation="citation"
    Report.WriteFormattedLn "<a name='{&t}'></a><h4>{}</h4>", s.ID, JoinSourceCitationNames(s, StrFormatText(s, StrParseText(s.title, True)), true)

    If (s.Session("PicturesIncluded") > 0) Then
        Report.WriteLn "<div class='floatright aligncenter widthpaddedlarge'>"
        WriteHtmlPicturesLarge s, "left", "", Session("fHidePictureName") Or Session("fShowPictureDetails"), False
        Report.WriteLn "</div>"
        strClassCitation="citationpic"
    End If

	strUrl = StrParseText(s.Url, False)        ' arg UseLangShowOthers=False so this always sets just a single language version
	strText = strUrl
	If strUrl <> "" And CustomTag(s, "Url.Exclude") = "" Then
		If Not Instr(strUrl, ":") > 0 Then strUrl = ReportGenerator.Document.BasePath & strUrl ' relative link
		If  oFso.FileExists(strUrl) Then
			strFile = strUrl
			strUrl = "media/" & Session("UUID") & "_" & Util.HrefEncode(oFso.GetFile(strUrl).Name) ' make url valid and unique
			Session("UUID") = Session("UUID") + 1
			ReportGenerator.FileUpload strFile, Util.UrlDecode(strUrl)
		End If
	End If
    If s.subtitle <> "" Then Report.WriteFormattedLn "<div class='subhead'>{}</div>", StrFormatText(s, s.subtitle)
    If s.QuotedText <> "" Then Report.WriteFormattedLn "<div class='{0}'>{1}</div>", strClassCitation, StrFormatText(s, s.QuotedText)
    Report.WritePhraseDic "PhSourceEntry", s.Edition, ser.Name, ser.Issue, pub.publisher, pub.Place.Session("HlinkLocative"), _
                CustomDate(pub.Date).ToStringNarrative, s.Originator , StrFormatText(s, s.Description), StrFormatText(s, StrParseText(s.WhereInSource, True)), s.Editor, _
                s.ISBN, Util.FormatHtmlHyperlink(strUrl, strText, "target='_blank'"), s.Repository.Session("HlinkLocative"), s.MediaType, StrFormatText(s, StrParseText(s.ReferenceNumber, True)), s.ConfidenceLevel
    
    WriteHtmlExtraNarrative s

    WriteHtmlAdditionalInformation(s)
    WriteHtmlAnnotation s, Dic("AnnotationSource"), s.Comment
    WriteHtmlReferences s, false
End Sub

Sub WriteHtmlPlace(p, GoogleMaps, fLink)
    Dim strWidth, strHeight, strType, strZoom, strMapLink
    Set collReferences = p.References
    If GoogleMaps And fLink Then ' link to map by loading gmap_place.htm in popup frame
            If Not Session("OriginalNamesGoogleMaps") Then
                strPlace=Util.FormatPhrase("{0}[[{?0};;]{1}][[{?0|1};;]{2}][[{?0|1|2};;]{3}][[{?0|1|2|3};;]{4}][[{?0|1|2|3|4};;]{5}]",p.Street, p.Session("City"), p.Session("County"), p.Session("State"), p.Zip, p.Session("Country"))
                if strPlace = "" Then strPlace=Replace(JoinTranslatedPlaceNames(p,p.Name,0),",",";;")
            Else
                strPlace=Util.FormatPhrase("{0}[[{?0};;]{1}][[{?0|1};;]{2}][[{?0|1|2};;]{3}][[{?0|1|2|3};;]{4}][[{?0|1|2|3|4};;]{5}]",p.Street, p.City, p.County, p.State, p.Zip, p.Country)
                if strPlace = "" Then strPlace=Replace(JoinOriginalPlaceNames(p,p.Name,0),",",";;")
            End If
            strZoom=CustomTag(p, "Map.Google.Zoom")
            if strZoom = "" Then strZoom = Session("GoogleMapsZoom")
            strType=CustomTag(p, "Map.Google.Type")
            if strType = "" Then strType = Session("GoogleMapsType")
            strMapLink = Util.FormatString(" <a href=""gmap_place.htm?lat={&u},lng={&u},place={&u},type={},zoom={}"" target=""popup""><img src=""images/pin.gif"" class=""icon"" alt=""{&t}"" title=""{5&t}""/></a>",p.Latitude, p.Longitude, Util.HtmlEncode(strPlace), strType, strZoom, Dic("gMapLink"))
  Else
      strMapLink = ""
  End If  
    Report.WriteFormattedLn "<a name='{&t}'></a><h4>{&t}{}</h4>", p.ID, p.Session("NameFull"), strMapLink
    If (p.Session("PicturesIncluded") > 0) Then
        Report.WriteLn "<div class='floatright aligncenter widthpaddedlarge'>"
        WriteHtmlPicturesLarge p, "left", "", Session("fHidePictureName") Or Session("fShowPictureDetails"), False
        Report.WriteLn "</div>"
        If GoogleMaps Then
            strWidth = CustomTag(p, "Map.Google.Width")
            If strWidth="" Then strWidth = Util.GetWidth(Session("GoogleMapsSmall"))
            strHeight= CustomTag(p, "Map.Google.Height")
            If strHeight="" Then strHeight = Util.GetHeight(Session("GoogleMapsSmall"))
        End If
    ElseIf GoogleMaps Then
        strWidth = CustomTag(p, "Map.Google.Width")
        If strWidth="" Then strWidth = Util.GetWidth(Session("GoogleMapsLarge"))
        strHeight= CustomTag(p, "Map.Google.Height")
        If strHeight="" Then strHeight = Util.GetHeight(Session("GoogleMapsLarge"))
    End If
    Dim strZip
    If p.Zip <> "" And p.Country = "UK" Then
        strZip = Util.FormatString("<a href='http://www.Streetmap.co.uk/streetmap.dll?postcode2map?{0}&{1&t}' target=_blank>{2}</a>",Replace(p.Zip, " ","+"), Replace(p.Name, " ","+"), p.Zip)
    Else
        strZip = p.Zip
    End If
    Report.WritePhraseDic "PhPlaceDescription", p.Session("NameShort"), Util.IfElse(p.Category <> p.Parent.Category, p.Category, ""), p.Street, p.Session("City"), p.Session("County"), p.Session("State"), strZip, p.Session("Country"), _
                    "", "", "", p.Latitude, p.Longitude, "<b>", "</b>"
    WriteHtmlExtraNarrative p
    WriteHtmlAnnotation p, Dic("Description"),Util.IfElse((p.Description & "") <> (p.Parent.Description & ""), StrFormatText(p, p.Description), "")
  If GoogleMaps And Not fLink Then ' show map on this page
            Report.WriteLn        "<div id='GoogleMapWrapper'><ul class='xT'>"
            Report.WriteFormattedLn "    <li class='xT3-o xT-h' onclick='xTclk(event,""3"")'>{&t}", Dic("GoogleMap")
            Report.WriteLn         "        <ul class='xT-b'>"
            Report.WriteFormattedLn    "        <li><div id='GoogleMap' style='width: {}px; height: {}px;'></div></li>", strWidth, strHeight
            Report.WriteLn         "        </ul>"
            Report.WriteLn        "    </li>"
            Report.WriteLn        "</ul></div>"
    End If
    WriteHtmlAdditionalInformation(p)
    WriteHtmlAnnotation p, Dic("AnnotationLocation"), p.Comment
    WriteHtmlReferences p, false
    WriteHtmlAllFootnotes p.Sources, True

End Sub

Sub WriteHtmlReferences(o, nested)
    Dim collReferences, oParent, d, oDate, oDateYears
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
            Case "Place", "Source"
                If r.Session("References") = 0 Then Set r = Nothing
            End Select
            If Not Util.IsNothing(r) Then collReferences.Add r, GetDate(oDate), r.Class
            If Not Util.IsNothing(oDate) Then oDateYears.Add r.Class & r.ID, oDate.Year
        Next
        collReferences.SortByKey()
        Set collReferences = collReferences.ToGenoCollection
    End Select
    If collReferences.Count > 0 Or Not Util.IsNothing(oParent) Then
        Report.WriteFormattedLn    "<ul class='xT {}'>",Util.IfElse(nested,"","note")    ' bug in IE has wrong interpretation of 'smaller' font
        Report.WriteFormattedLn "    <li class='xT2-{} xT-h' onclick='xTclk(event,""2"")'>{&t} ({})", Util.IfElse(Session("fCollapseReferences"), "c", "o"), Util.StrFirstCharUCase(Dic("References")), collReferences.Count + Util.IfElse(Util.IsNothing(oParent),0,1)
        Report.WriteLn         "        <ul class='xT-n'>"
        If Not Util.IsNothing(oParent) Then
            Session("ReferencesStart") = -1
            Report.WriteFormattedLn "<li>{&t}: {}</li>", Dic(o.Class & "_Parent"), oParent.Session("Hlink")
        End If
        For Each d In collReferences
            Session("ReferencesStart") = -1
            strHref = d.Href
            If (strHref <> "") Then
                Select Case d.Class
                    Case "Place"
                        strHref = d.Session("Hlink")
                     Case "SourceCitation"
                        strHref = d.Session("Hlink")
                    Case "Family"
                        strHref = Util.FormatString("<a href='{&t}' target='detail' onclick='javascript:hidePopUpFrame();'>{&t}</a>", strHref, d.Session("Name"))
                    Case "Individual"
                        strHref = d.Session("Hlink")
                        If o.Class = "Place" Then
                              strHref = strHref & Util.FormatPhrase(Dic("FmtDatesIndividual"), _
                                Util.IfElse(o.ID = d.Birth.Place.ID, GetDateString(d.Birth.Date),""), _
                                Dic("PhBC_" & d.Birth.CeremonyType.ID), _
                                Util.IfElse(o.ID = d.Birth.Baptism.Place.ID, GetDateString(d.Birth.Baptism.Date),""), _
                                Util.IfElse(o.ID = d.Death.Place.ID, GetDateString(d.Death.Date),""), _
                                Util.IfElse(o.ID = d.Death.Funerals.Place.ID, GetDateString(d.Death.Funerals.Date),""), _
                                Util.IfElse(o.ID = d.Death.Disposition.Place.ID, GetDateString(d.Death.Disposition.Date),""))
                        End If
                    Case "Picture"
                        strHref = StrHtmlHyperlink(d)
                End Select    
            Else
                strSep = ""    
                Select Case d.Class
                    Case "Occupation", "Education"
                        strHref = Util.FirstNonEmpty(d.Session("Title"),StrParseText(d, True)) & " "
                        For Each dobj in d.References
                            strHref = strHref & strSep & Util.FormatString("<a href='{&t}' target='detail' onclick='javascript:hidePopUpFrame();'>{&t}</a>", dobj.Href, dobj)
                            strSep = ", "
                        Next
                        If o.Class = "Place" And o.ID = d.Place.ID Then
                            strHref = strHref & Util.FormatPhrase(Dic("FmtDatesFromTo"), _
                                GetDateString(d.DateStart), _
                                GetDateString(d.DateEnd))
                        End If
                    Case "PedigreeLink"
                        strHref = d.PedigreeLink & " " & strSep & Util.FormatString("<a href='{&t}' target='detail' onclick='javascript:hidePopUpFrame();'>{&t}</a>", d.individual.Href, d.individual)
                    Case "Marriage", "Contact"
                        If d.Class = "Contact" And d.Type.ID <> "" Then strHref = d.Type & " "
                        For Each dobj in d.References
                            strHref = strHref & strSep & Util.FormatString("<a href='{&t}' target='detail' onclick='javascript:hidePopUpFrame();'>{&t}</a>", dobj.Href, dobj)
                            strSep = ", "
                        Next
                        If o.Class = "Place" Then
                            If d.Class = "Contact" And o.Id = d.Place.ID Then
                                strHref = strHref & Util.FormatPhrase(Dic("FmtDatesFromTo"), _
                                    GetDateString(d.DateStart), _
                                    GetDateString(d.DateEnd))
                            Else
                                strHref = strHref & Util.FormatPhrase(Dic("FmtDatesUnion"), _
                                    Util.IfElse(o.ID = d.Place.ID, GetDateString(d.Date),""), _
                                    Util.IfElse(o.ID = d.Divorce.Place.ID, GetDateString(d.Divorce.Date),""))
                            End If
                        End If
                    Case Individual
                        strHref = d.Session("Hlink")
                    Case Else
                        strHref = d & ""
                End Select
            End If
            Select Case d.Class
            Case "SocialRelationship", "EmotionalRelationship"
                Report.WriteFormatted "<li>{&t}: ", Dic(d.Class)
                WriteHtmlRelationship d, "", "", "", "", False, False, "", ""
                Report.WriteLn "</li>"
            Case "Occupation"
                If d.Session("Event") <> "" Then
                    Report.WriteFormattedLn "<li>{} {&t}: {}</li>", oDateYears.KeyValue(d.Class & d.ID), d.Session("EventName"), strHref
                Else
                    Report.WriteFormattedLn "<li>{} {&t}: {}</li>", oDateYears.KeyValue(d.Class & d.ID), StrDicExt(d.Class,"",d.Class,"",""), strHref
                End If
            Case Else
                Report.WriteFormattedLn "<li>{} {&t}: {}</li>", oDateYears.KeyValue(d.Class & d.ID), StrDicExt(d.Class,"",d.Class,"",""), strHref
            End Select
        Next
        Report.WriteLn "        </ul>"
        Report.WriteLn "    </li>"
        Report.WriteLn "</ul>"
    End If
End Sub

Function WriteHtmlRelationship(r, strName1,strNameObj1, strName2, strNameObj2, fReversible, fSources, strNameP1a, strNameP2a)
    Dim strHref, strHref2, strTitle, strRelationship, relation,  strTemp
    Dim fExtant, fExtant1, fExtant2, strGender1, strGender2, strRef1, strRef2, strObj1, strObj2
    Dim strNameP1, strNameP2, arrText, fSwitch, e1, e2
    fSwitch = fReversible
    strNameP1 = strNameP1a
    strNameP2 = strNameP2a
    strHref2 = ""
    fExtant = (CustomTag(r, "Extant") <> "N")
    Select Case r.Class
    Case "EmotionalRelationship"
        Set e1 = HyperlinkDataSource(r.Entity1)
        Set e2 = HyperlinkDataSource(r.Entity2)
        Select Case r.EmotionalLink.ID
              Case "Other" :  strRelationship = StrDicOrTag2("PhER_Other", CustomTag(r,"NarrativeStyle"), "PhER_other", "2.0.1.6")
              Case Else : strRelationship = StrDicOrTag("PhER_" & r.EmotionalLink.ID, CustomTag(r,"NarrativeStyle"))
                End Select
    Case "SocialRelationship"
        Set e1 = HyperlinkDataSource(r.entity1)
        Set e2 = HyperlinkDataSource(r.entity2)
        strRelationship = StrDicOrTag("PhSR_" & r.Connection.ID, CustomTag(r,"NarrativeStyle"))
    End Select

    Select Case e1.Class
    Case "Individual"
        fExtant1 = Not e1.IsDead
        strRef1 = Util.FirstNonEmpty(strName1, e1.Session("HlinkNN"))
        strObj1 = strNameObj1
        If strNameP1 = "" Then strNameP1 = e1.Session("NameFullPossessive")
        strGender1 = e1.Gender.ID
    Case "SocialEntity"
        fExtant1 = (CustomTag(e1, "Extant") <> "N")
        strRef1 = e1.Session("Hlink")
        fSwitch = False
    Case Else
        strRef1 = e1 & "" ' use default property
    End Select
    If strObj1 = "" Then strObj1 = strRef1

    Select Case e2.Class
    Case "Individual"
        fExtant2 = Not e2.IsDead
        strRef2 = Util.FirstNonEmpty(strName2, e2.Session("HlinkNN"))
        strObj2 = strNameObj2
        If strNameP2 = "" Then strNameP2 = e2.Session("NameFullPossessive")
        strGender2 = e2.Gender.ID
    Case "SocialEntity"
        fExtant2 = (CustomTag(e2, "Extant")<> "N")
        strRef2 = e2.Session("Hlink")
        fSwitch = False
    Case Else
        strRef2 = e2 & "" ' use default property
    End Select
    If strObj2 = "" Then strObj2 = strRef2

    If Not fSwitch Then
        Report.WritePhrase strRelationship, Util.StrFirstCharUCase(strRef1), (fExtant And fExtant1 And fExtant2) = True, strObj2, strGender1="F", strGender2="F", strGender1 = strGender2, strObj1, strRef2, strNameP2
    Else
        Report.WritePhrase strRelationship, Util.StrFirstCharUCase(strRef2), (fExtant And fExtant1 And fExtant2) = True, strObj1, strGender2="F", strGender1="F", strGender1 = strGender2, strObj2, strRef1, strNameP1
    End If
    If fSources Then WriteHtmlFootnoteRefs(r.Sources)

End Function

Function HyperlinkDataSource(obj)
    If obj.Class = "Individual" Then
        If Not Util.IsNothing(obj.IndividualInternalHyperlink) Then
            Set HyperlinkDataSource = obj.IndividualInternalHyperlink
        Else
            Set HyperlinkDataSource = obj
        End If
    Else
        Set HyperlinkDataSource = obj
    End If
End Function




