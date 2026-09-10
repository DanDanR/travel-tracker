Attribute VB_Name = "SheetMacros"
Private g_airportsJson As String
Private g_airportsLoaded As Boolean

Public Sub ReCalculateCurrentCell()
    On Error Resume Next

    For Each cell In Selection
        cell.Calculate
    Next cell
End Sub

Public Sub ConvertCalculation2StaticVal()
    On Error Resume Next

    For Each cell In Selection
        cell.Value = cell.Value
    Next cell
End Sub

Private Function GeoNamesGet(endpoint As String, queryParams As String) As Variant
    Dim http As Object
    Dim url As String

    Set http = CreateObject("MSXML2.XMLHTTP")
    url = "http://api.geonames.org/" & endpoint & "?" & queryParams & "&username=" & ThisWorkbook.credentials("GEONAMES_USERNAME")
    
    http.Open "GET", url, False
    http.send

    If http.Status <> 200 Then
        GeoNamesGet = CVErr(xlErrNA)
        Exit Function
    End If

    GeoNamesGet = http.responseText
End Function

Public Function GetCityPopulation(cityName As String, countryCode As String) As Variant
    Dim jsonText As Variant
    jsonText = GeoNamesGet("searchJSON", "q=" & WorksheetFunction.EncodeURL(cityName) & "&country=" & countryCode & "&maxRows=1")
    If IsError(jsonText) Then
        GetCityPopulation = jsonText
    Else
        GetCityPopulation = ExtractJsonNumber(jsonText, "population")
    End If
End Function

Public Function GetCountryPopulation(countryCode As String) As Variant
    Dim jsonText As Variant
    jsonText = GeoNamesGet("countryInfoJSON", "country=" & countryCode)
    If IsError(jsonText) Then
        GetCountryPopulation = jsonText
    Else
        GetCountryPopulation = ExtractJsonNumber(jsonText, "population")
    End If
End Function

Public Function GetCountryArea(countryCode As String) As Variant
    Dim jsonText As Variant
    jsonText = GeoNamesGet("countryInfoJSON", "country=" & countryCode)
    If IsError(jsonText) Then
        GetCountryArea = jsonText
    Else
        GetCountryArea = ExtractJsonNumber(jsonText, "areaInSqKm")
    End If
End Function

Public Function GetCountryCapital(countryCode As String) As Variant
    Dim jsonText As Variant
    jsonText = GeoNamesGet("countryInfoJSON", "country=" & countryCode)
    If IsError(jsonText) Then
        GetCountryCapital = jsonText
    Else
        GetCountryCapital = ExtractJsonString(jsonText, "capital")
    End If
End Function

Public Function GetCityLatitude(cityName As String, countryCode As String) As Variant
    Dim jsonText As Variant
    jsonText = GeoNamesGet("searchJSON", "q=" & WorksheetFunction.EncodeURL(cityName) & "&country=" & countryCode & "&maxRows=1")
    If IsError(jsonText) Then
        GetCityLatitude = jsonText
    Else
        GetCityLatitude = ExtractJsonNumber(jsonText, "lat", True)
    End If
End Function

Public Function GetCityLongitude(cityName As String, countryCode As String) As Variant
    Dim jsonText As Variant
    jsonText = GeoNamesGet("searchJSON", "q=" & WorksheetFunction.EncodeURL(cityName) & "&country=" & countryCode & "&maxRows=1")
    If IsError(jsonText) Then
        GetCityLongitude = jsonText
    Else
        GetCityLongitude = ExtractJsonNumber(jsonText, "lng", True)
    End If
End Function

Public Function GetCountryCommonName(countryCode As String) As Variant
    commonName = GetCountryName(countryCode, "common")
    GetCountryCommonName = commonName
End Function

Public Function GetCountryFormalName(countryCode As String) As Variant
    formalName = GetCountryName(countryCode, "official")
    GetCountryFormalName = formalName
End Function

Public Function GetAirportField(iataCode As String, fieldName As String) As Variant
    Dim objText As String
    objText = FindAirportObject(iataCode)

    If objText = "" Then
        GetAirportField = CVErr(xlErrNA)
        Exit Function
    End If

    Select Case fieldName
        Case "lat", "lon"
            GetAirportField = ExtractJsonNumber(objText, fieldName, True)
        Case Else
            GetAirportField = ExtractJsonString(objText, fieldName)
    End Select
End Function

Private Sub EnsureAirportsLoaded()
    If g_airportsLoaded Then Exit Sub

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://raw.githubusercontent.com/mwgg/Airports/master/airports.json", False
    http.send

    If http.Status <> 200 Then
        Err.Raise vbObjectError + 1, , "Could not download airports database (HTTP " & http.Status & ")"
    End If

    g_airportsJson = http.responseText
    g_airportsLoaded = True
End Sub

Private Function FindAirportObject(iataCode As String) As String
    EnsureAirportsLoaded

    Dim posMarker As Long
    posMarker = InStr(1, g_airportsJson, """iata"": """ & iataCode & """", vbTextCompare)
    If posMarker = 0 Then
        posMarker = InStr(1, g_airportsJson, """iata"":""" & iataCode & """", vbTextCompare)
        If posMarker = 0 Then
            FindAirportObject = ""
            Exit Function
        End If
    End If

    Dim posStart As Long, posEnd As Long
    posStart = InStrRev(g_airportsJson, "{", posMarker)
    posEnd = InStr(posMarker, g_airportsJson, "}")

    If posStart = 0 Or posEnd = 0 Then
        FindAirportObject = ""
        Exit Function
    End If

    FindAirportObject = Mid(g_airportsJson, posStart, posEnd - posStart + 1)
End Function

Private Function GetCountryName(countryCode As String, fieldName As String) As Variant
    Dim http As Object
    Dim url As String
    Dim jsonText As String

    url = "https://api.restcountries.com/countries/v5/codes.alpha_2/" & countryCode & "?response_fields=names." & fieldName

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & ThisWorkbook.credentials("RESTCOUNTRIES_API_KEY")
    http.send

    If http.Status <> 200 Then
        GetCountryOfficialName = CVErr(xlErrNA)
        Exit Function
    End If

    jsonText = http.responseText
    GetCountryName = ExtractJsonString(jsonText, fieldName)
End Function

Private Function ExtractJsonNumber(jsonText As Variant, fieldName As String, Optional invertDecimalSeparator As Boolean = False) As Variant
    Dim posStart As Long, posEnd As Long
    Dim marker As String
    marker = """" & fieldName & """:"

    posStart = InStr(jsonText, marker)
    If posStart = 0 Then
        ExtractJsonNumber = CVErr(xlErrNA)
        Exit Function
    End If

    posStart = posStart + Len(marker)
    posEnd = posStart
    Do While Mid(jsonText, posEnd, 1) <> "," And Mid(jsonText, posEnd, 1) <> "}"
        posEnd = posEnd + 1
    Loop

    result = Replace(Mid(jsonText, posStart, posEnd - posStart), Chr(34), "")
    If invertDecimalSeparator = True Then result = Replace(result, ".", ",")
    
    ExtractJsonNumber = CDbl(result)
End Function

Private Function ExtractJsonString(jsonText As Variant, fieldName As String) As Variant
    Dim posStart As Long, posEnd As Long
    Dim marker As String
    
    marker = """" & fieldName & """:"""
    posStart = InStr(jsonText, marker)
    If posStart = 0 Then
        marker = """" & fieldName & """: """
        posStart = InStr(jsonText, marker)
        If posStart = 0 Then
            ExtractJsonString = CVErr(xlErrNA)
            Exit Function
        End If
    End If

    posStart = posStart + Len(marker)
    posEnd = InStr(posStart, jsonText, """")
    If posEnd = 0 Then
        ExtractJsonString = CVErr(xlErrNA)
        Exit Function
    End If

    ExtractJsonString = Mid(jsonText, posStart, posEnd - posStart)
End Function
