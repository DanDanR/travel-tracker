VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DieseArbeitsmappe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public credentials As Object

Private Sub Workbook_Open()
    Application.OnKey "{F11}", "ConvertCalculation2StaticVal"
    Application.OnKey "{F12}", "ReCalculateCurrentCell"
    Set credentials = LoadCredentials(ThisWorkbook.Path & "\credentials.txt")
    'Call SetPythonInterpreter
End Sub

Private Sub Workbook_Activate()
    Application.OnKey "{F11}", "ConvertCalculation2StaticVal"
    Application.OnKey "{F12}", "ReCalculateCurrentCell"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.OnKey "{F11}"
    Application.OnKey "{F12}"
End Sub

Private Sub Workbook_Deactivate()
    Application.OnKey "{F11}"
    Application.OnKey "{F12}"
End Sub

Sub SetPythonInterpreter()
    Dim wbPath As String
    wbPath = ThisWorkbook.Path
    
    ThisWorkbook.Names.Add Name:="PYTHON_WIN", RefersTo:="=""" & wbPath & "\venv\Scripts\python.exe"""
    ThisWorkbook.Names.Add Name:="INTERPRETER_WIN", RefersTo:="=""" & wbPath & "\venv\Scripts\python.exe"""
    ThisWorkbook.Names.Add Name:="PYTHONPATH", RefersTo:="=""" & wbPath & """"
End Sub

Public Function LoadCredentials(filePath As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If Dir(filePath) = "" Then
        MsgBox "credentials.txt not found at: " & filePath, vbExclamation
        Set LoadCredentials = dict
        Exit Function
    End If

    Dim fileNum As Integer
    Dim lineText As String
    Dim eqPos As Long
    Dim k As String, v As String

    fileNum = FreeFile
    Open filePath For Input As #fileNum

    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim(lineText)

        ' skip blank lines
        If Len(lineText) > 0 Then
            eqPos = InStr(lineText, "=")
            If eqPos > 0 Then
                k = Trim(Left(lineText, eqPos - 1))
                v = Trim(Mid(lineText, eqPos + 1))
                ' strip surrounding quotes if present
                If Left(v, 1) = """" And Right(v, 1) = """" Then
                    v = Mid(v, 2, Len(v) - 2)
                End If
                dict(k) = v
            End If
        End If
    Loop

    Close #fileNum
    Set LoadCredentials = dict
End Function
