Attribute VB_Name = "MyFunctions"
Public Enum JSCommands: Parse = 1: GetKeys = 2: End Enum


Public Function WebScrape(UrlAddress As String, Optional httpObject As Object) As Object

Dim attempts&: attempts = -1

  If IsMissing(httpObject) _
  Or httpObject Is Nothing _
Then Set httpObject = CreateObject("WinHTTP.WinHTTPRequest.5.1")

RetryConnection:
   attempts = attempts + 1
If attempts > 10 Then MsgBox "Maximum connection attempts reached!" & Chr(10) & _
                             "Check the site to see if it is up and running." & Chr(10) & _
                             "Check the given URL and any argument spelling / syntax" & _
                             "for possible grammatical errors.": Exit Function

On Error Resume Next
With httpObject
    .SetAutoLogonPolicy 0
    .SetTimeouts 0, 0, 0, 0
    .Open "GET", UrlAddress
    .send
    .WaitForResponse
 If .Status <> 200 Then GoTo RetryConnection
End With
Set WebScrape = httpObject
On Error GoTo 0

End Function

Public Function ImportCSV(httpObject As Object, TopLeftCell As Range)

Dim ado As Object: Set ado = CreateObject("ADODB.Stream")
Dim alt As Boolean: alt = Application.DisplayAlerts
 
'Checking httpobject type, if not the correct object type- throws error
If (TypeName(httpObject) <> "WinHttpRequest") Then MsgBox "Invalid object type!": Exit Function

On Error Resume Next: Kill Environ("TEMP") & "\tempFile.csv": On Error GoTo 0

With ado
    .Open
    .Type = 1
    .Write httpObject.responseBody
    .SaveToFile Environ("TEMP") & "\tempFile.csv"
    .Close
End With

'Turning off alerts to avoid automation interruption
Application.DisplayAlerts = False
With Workbooks.Open(Environ("TEMP") & "\tempFile.csv")
    .ActiveSheet.UsedRange.Copy: TopLeftCell.PasteSpecial xlPasteValues
    .Close False
End With

'Reverting alert settings to previous condition
Application.DisplayAlerts = alt
Kill Environ("TEMP") & "\tempFile.csv"

'Clean up
Set ado = Nothing

End Function

Public Function ImportHTML(httpObject As Object, TableNum As Double, TopLeftCell As Range)

Dim doc As Object: Set doc = CreateObject("HTMLFile")
Dim frm As Object: Set frm = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")

'Checking httpobject type, if not the correct object type- throws error
If (TypeName(httpObject) <> "WinHttpRequest") Then MsgBox "Invalid object type!": Exit Function

doc.body.innerHTML = httpObject.responseText

With frm
    .SetText doc.getElementsByTagName("table")(TableNum).outerHTML
    .PutInClipboard
End With

TopLeftCell.PasteSpecial

Set doc = Nothing
With TopLeftCell.Worksheet
    For Each doc In .Shapes
        doc.Delete
    Next doc
End With

Application.CutCopyMode = False

'Clean up
Set doc = Nothing
Set frm = Nothing
End Function

Public Function JSON(JSFunction As JSCommands, DataObject As Variant) As Variant

Dim motw$

motw = _
"<!doctype html>" & _
"<!-- saved from url=(0014)about:internet -->" & _
"<html>" & _
"<head><title>Created By Michael Klink</title></head>" & _
"<body>       <p>mklink.dev@gmail.com</p>    </body>" & _
"</html>"

With CreateObject("HTMLFile")
    .Write motw
    With .parentWindow
    Select Case JSFunction
        Case 1
            .execScript "var JSONParse=" & DataObject, "JScript"
 Set JSON = .JSONParse
        Case 2
            .execScript "function GetKeys(O) { var k = new Array(); for (var x in O) { k.push(x); } return k; } ", "JScript"
 Set JSON = .GetKeys(DataObject)
     JSON = Split(JSON, ",")
    End Select
    End With
End With

End Function




