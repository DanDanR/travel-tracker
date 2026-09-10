Attribute VB_Name = "MapMacros"
Option Explicit
Const LANG_CODE = "zh"

Public Sub RunMapGenerator(functionName As String, openBrowser As Boolean, Optional langCode As String = "")
    Dim pyBool As String
    pyBool = IIf(openBrowser, "True", "False")

    If langCode = "" Then
        RunPython "import map_generator; map_generator." & functionName & "(" & pyBool & ")"
    Else
        RunPython "import map_generator; map_generator." & functionName & "(" & pyBool & ", '" & langCode & "')"
    End If
End Sub

Public Sub ShowCitiesCountriesMapInInExcel()
    ShowCitiesCountriesMap (False)
End Sub

Public Sub ShowCitiesCountriesMapInBrowser()
    ShowCitiesCountriesMap (True)
End Sub

Public Sub ShowCitiesCountriesMap(Optional showInBrowser As Boolean = True)

    Application.Cursor = xlWait
    Application.StatusBar = "Rebuilding map..."

    On Error GoTo ErrHandler
    
    Call RunMapGenerator("build_cities_countries_map", showInBrowser, LANG_CODE)
    If showInBrowser = False Then
        ShowMapForm
    End If
    
    Application.StatusBar = "Map refreshed."
    Application.Cursor = xlDefault
    Exit Sub

ErrHandler:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Could not refresh the map:" & vbCrLf & Err.Description, vbExclamation
End Sub

Private Sub ShowMapForm()
    ' Load the form once, reuse it on subsequent refreshes so the
    ' window doesn't jump around.
    If Not IsFormLoaded("MapForm") Then
        MapForm.Show vbModeless
    End If
    MapForm.ReloadMap
End Sub

Private Function IsFormLoaded(formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = formName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function

Public Sub ShowAviationMapInBrowser()
    ShowAviationMap (True)
End Sub

Public Sub ShowAviationMapInExcel()
    ShowAviationMap (False)
End Sub

Private Sub ShowAviationMap(Optional showInBrowser As Boolean = True)
    Application.Cursor = xlWait
    Application.StatusBar = "Rebuilding map..."

    On Error GoTo ErrHandler

    Call RunMapGenerator("build_aviation_map", showInBrowser, LANG_CODE)
    If showInBrowser = False Then
        ShowMapForm
    End If

    Application.StatusBar = "Map refreshed."
    Application.Cursor = xlDefault
    Exit Sub

ErrHandler:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Could not refresh the map:" & vbCrLf & Err.Description, vbExclamation
End Sub

Sub ShowYearlyOverviewInBrowser()
    ShowYearlyOverview (True)
End Sub

Sub ShowYearlyOverviewInExcel()
    ShowYearlyOverview (False)
End Sub

Sub ShowYearlyOverview(Optional showInBrowser As Boolean = True)
    Application.Cursor = xlWait
    Application.StatusBar = "Rebuilding page..."

    On Error GoTo ErrHandler

    Call RunMapGenerator("build_years_page", showInBrowser)
    If showInBrowser = False Then
        ShowMapForm
    End If

    Application.StatusBar = "Page refreshed."
    Application.Cursor = xlDefault
    Exit Sub

ErrHandler:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Could not refresh the page:" & vbCrLf & Err.Description, vbExclamation
End Sub

Sub ShowDashboardInBrowser()
    ShowDashboard (True)
End Sub

Sub ShowDashboardInExcel()
    ShowDashboard (False)
End Sub

Sub ShowDashboard(Optional showInBrowser As Boolean = True)
    Application.Cursor = xlWait
    Application.StatusBar = "Rebuilding dashboard..."

    On Error GoTo ErrHandler

    Call RunMapGenerator("build_dashboard", showInBrowser)
    If showInBrowser = False Then
        ShowMapForm
    End If

    Application.StatusBar = "Dashboard refreshed."
    Application.Cursor = xlDefault
    Exit Sub

ErrHandler:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    MsgBox "Could not refresh the dashboard:" & vbCrLf & Err.Description, vbExclamation
End Sub


