<%[@ IncludeFile "Code/Util.vbs" ]%>
<%[@ IncludeFile "Code/Lang.vbs" ]%>
<%[If Session("Book") Then Report.AbortTemplate]%>
	var gMap = new Object();
	var gMapOptions = new Object();
	gMap.reasons=[];
	gMap.reasons['ErrorMessage']           = '@[Report.Write StrJavaScriptEncode(Dic("gMapError"))]@';
	if (typeof google != 'undefined') {		// i.e. GoogleMaps code has loaded OK
		gMap.types=[-1,google.maps.MapTypeId.ROADMAP, google.maps.MapTypeId.SATELLITE, google.maps.MapTypeId.HYBRID, google.maps.MapTypeId.TERRAIN];

//		gMap.reasons[G_GEO_MISSING_ADDRESS]    = '@[Report.Write StrJavaScriptEncode(Dic("gMapMissingAddress"))]@';
		gMap.reasons[google.maps.GeocoderStatus.ZERO_RESULTS]    = '@[Report.Write StrJavaScriptEncode(Dic("gMapUnknownAddress"))]@';
//		gMap.reasons[G_GEO_UNAVAILABLE_ADDRESS]= '@[Report.Write StrJavaScriptEncode(Dic("gMapUnavailableAddress"))]@';
		gMap.reasons[google.maps.GeocoderStatus.INVALID_REQUEST]            = '@[Report.Write StrJavaScriptEncode(Dic("gMapBadKey"))]@';
		gMap.reasons[google.maps.GeocoderStatus.OVER_QUERY_LIMIT]   = '@[Report.Write StrJavaScriptEncode(Dic("gMapTooManyQueries"))]@';
		gMap.reasons[google.maps.GeocoderStatus.REQUEST_DENIED]       = '@[Report.Write StrJavaScriptEncode(Dic("gMapServerError"))]@';
	} else {
		alert('Failed to load Google Maps API - check Internet connection !!');
	}
	gMap.typeDefault=@[Report.WriteText Session("GoogleMapsType")]@;
