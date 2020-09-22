Attribute VB_Name = "ModFunc"
Option Explicit

Public Sub DoButton(iButton As Integer, objToHide As Object, objToShow As Object)
'i use this routine for the 'buttons'
'each button is really two images, one representing it being pressed, the other unpressed
'to simulate the event of pressing a button, i simply hide one image and show the other
If iButton = 1 Then 'if the user clicks the left mouse button then
objToHide.Visible = False 'hide this one
objToShow.Visible = True 'show this one
End If
End Sub

Public Function GetIP(sHTML As String) As String
'this involves basic string manipulation, using string arrays, along with the InStr() and Mid() functions
'this function parses the IP out of the source code to "http://www.ShowMyIP.com/xml"
'to understand how it works, look at the source code and compare it to this routine
GetIP = Empty
Dim sBuff() As String: sBuff() = Split(sHTML, vbNewLine)
Dim iStart As Integer, iEnd As Integer: iStart = 0: iEnd = Empty
iStart = InStr(1, sBuff(2), "<ip>")
iEnd = InStr(iStart + 1, sBuff(2), "<")
GetIP = Mid(sBuff(2), iStart + 4, (iEnd - iStart) - 4)
End Function

Public Function GetHost(sHTML As String) As String
GetHost = Empty
Dim sBuff() As String: sBuff() = Split(sHTML, vbNewLine)
Dim iStart As Integer, iEnd As Integer: iStart = 0: iEnd = 0
iStart = InStr(1, sBuff(3), "<host>")
iEnd = InStr(iStart + 1, sBuff(3), "<")
GetHost = Mid(sBuff(3), iStart + 6, (iEnd - iStart) - 6)
End Function

Public Sub CreateFile(sPath As String)
Open sPath For Binary As #1
Close #1
End Sub

Public Sub KillFile(sPath As String)
On Error Resume Next
Kill sPath
End Sub

Public Sub RemakeFile(sPath As String)
Call KillFile(sPath)
Call CreateFile(sPath)
End Sub

Public Sub WriteToFile(sPath As String, sData As String)
Call RemakeFile(sPath)
Dim FF As Integer: FF = FreeFile
Open sPath For Binary Access Write As #FF
Put #FF, , sData
Close #FF
End Sub
