VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MapForm 
   Caption         =   "UserForm1"
   ClientHeight    =   7320
   ClientLeft      =   180
   ClientTop       =   710
   ClientWidth     =   10060
   OleObjectBlob   =   "MapForm.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "MapForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------
'Notes on the "Microsoft Web Browser" control used (Browser1):
'- It's the legacy IE-based ActiveX control (SHDocVw.WebBrowser). It's
'  deprecated but still ships with Windows and works fine for local,
'  reasonably modern HTML/JS such as Folium/Leaflet output.
'- If you hit rendering/JS errors, force it into a modern IE document
'  mode via a registry key for FeatureBrowserEmulation (see README.md),
'  or move up to the WebView2 route described in README.md.

Private Const MAP_HTML_PATH As String = "C:\map.html"
' ^ Must match OUTPUT_HTML in map_generator.py. Consider building this
'   path dynamically instead, e.g.:
'   Environ$("USERPROFILE") & "\excel_map_output\map.html"

Private Sub UserForm_Initialize()
    Me.Caption = "Interactive Map"
    Me.Width = 640
    Me.Height = 480
    ResizeBrowser
    ReloadMap
End Sub

Private Sub UserForm_Resize()
    ResizeBrowser
End Sub

Private Sub ResizeBrowser()
    On Error Resume Next
    Browser1.Left = 0
    Browser1.Top = 0
    Browser1.Width = Me.InsideWidth
    Browser1.Height = Me.InsideHeight
    On Error GoTo 0
End Sub

Public Sub ReloadMap()
    Dim htmlPath As String
    'htmlPath = Environ$("USERPROFILE") & "\excel_map_output\map.html"
    htmlPath = MAP_HTML_PATH

    If Dir(htmlPath) = "" Then
        MsgBox "map.html not found yet - click Refresh Map first.", vbInformation
        Exit Sub
    End If

    ' file:// URL, cache-busted with a timestamp so the browser doesn't
    ' show a stale cached copy after regeneration
    Browser1.Navigate "file:///" & Replace(htmlPath, "\", "/") & "?t=" & Format(Now, "yyyymmddhhnnss")
End Sub



