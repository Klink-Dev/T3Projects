Attribute VB_Name = "Toaster"
Sub badgesearch()

'Created by Michael Klink
'

Application.DisplayAlerts = False

Sheets("Search_By_Badge_Number_Bulk").Range("D4:D1048576").Clear

Sheets("REF").Visible = True
Sheets("FCLM").Visible = True
Sheets("FLEX").Visible = True
Sheets("Onsite").Visible = True
Sheets("Filtered").Visible = True
Sheets("Backup").Visible = True
Sheets("Roster").Visible = True

Call Pull_Data
Call jobSplit

Sheets("Search_By_Badge_Number_Bulk").Select
Range("D4").Select

Sheets("REF").Visible = False
Sheets("FCLM").Visible = False
Sheets("FLEX").Visible = False
Sheets("Onsite").Visible = False
Sheets("Filtered").Visible = False
Sheets("Backup").Visible = False
Sheets("Roster").Visible = False

Application.DisplayAlerts = True

End Sub


Sub jobSplit()
'
' jobSplit Macro
' Created by Michael Klink
'

Application.DisplayAlerts = False

    Sheets("FLEX").Activate
    Columns("AA:AM").Select
    Selection.ClearContents
    Columns("Q:Q").Select
    Selection.TextToColumns Destination:=Range("AA1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1)), _
        TrailingMinusNumbers:=True
        


End Sub

Sub filterOnsite()
'
'Created By Michael Klink

'
shifttype = Sheets("Search_By_Job").Range("C9")

Sheets("Filtered").Range("A:H").ClearContents

Sheets("Onsite").Activate

'remove blanks
    Range("A1").Select
    Application.CutCopyMode = False
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$H$3100").AutoFilter Field:=1, Criteria1:="<>"
    
    'filter data according to shift
    If shifttype = Sheets("REF").Range("B2") Then
        ActiveSheet.Range("$A$1:$H$3100").AutoFilter Field:=3, Criteria1:="=D"
    ElseIf shifttype = Sheets("REF").Range("B3") Then
        ActiveSheet.Range("$A$1:$H$3100").AutoFilter Field:=3, Criteria1:="=N"
    ElseIf shifttype = Sheets("REF").Range("B4") Then
        ActiveSheet.Range("$A$1:$H$3100").AutoFilter Field:=3, Criteria1:="=M"
    Else: MsgBox "Something went wrong filtering the shifts. Please try again.", _
        , "Shift Filter"
    
    End If
    
    ActiveCell.Range("A1:H3100").Select
    Selection.Copy
    Sheets("Filtered").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Call createlist
    
    
End Sub

Sub isOnsite()
'
'
'Created by Michael Klink
'
'

x = Sheets("REF").Range("I2")

Dim AAonsite As Boolean
Dim loginRange As Range
Dim EMPrownum As Integer

AAonsite = False

With Sheets("FLEX").Range("A:AG")
    Set loginRange = .Find(x, LookIn:=xlValues)
    If loginRange Is Nothing Then
        MsgBox "Sorry, but it looks like " & x & " is not on the job list." & vbCrLf _
        & "Please try again."
    
    Else
        EMPrownum = loginRange.Row
        If Range("AG" & EMPrownum) > 0 Then
            AAonsite = True
        Else
            MsgBox "Sorry, it looks like " & x & " is not onsite today." & vbCrLf _
            & "Please try another login", , "Failed"
        End If
        
        If AAonsite Then
            MsgBox x & " is onsite and trained in:" & vbCrLf _
            & Range("AA" & EMPrownum).Value & vbCrLf _
            & Range("AB" & EMPrownum).Value & vbCrLf _
            & Range("AC" & EMPrownum).Value & vbCrLf _
            & Range("AD" & EMPrownum).Value & vbCrLf _
            & Range("AE" & EMPrownum).Value & vbCrLf _
            & Range("AF" & EMPrownum).Value, , "Success!"
        End If
    End If
End With



End Sub

Sub createlist()

'
'Created by Michael Klink

'
Dim rownum As Integer
Dim job1 As String
Dim logintemp As String
Dim jobtemp As String
Dim rowcounter As Integer
Dim rng As Object
Dim x As Integer

Sheets("Backup").Select
Range("A:B").Select
Selection.Clear

Sheets("Filtered").Select
x = 1  ' counter
lrow = Cells(Sheets("Filtered").Rows.Count, 1).End(xlUp).Row  'last row variable

'dynamically select cell range
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Set rng = Application.Selection

'loop through cells in selected range checking if they match job field
'when it reaches the last row, send successful message and exit loop
    For Each cell In rng
        If (cell.Row) = lrow Then
            MsgBox "Compilation Complete", , "SUCCESS!"
            Exit For
        End If
        If cell = Sheets("Search_By_Job").Range("$C$11") Then
            logintemp = Sheets("Filtered").Range("B" & (cell.Row))
            job1 = cell
            jobtemp = job1
            Sheets("Backup").Range("A" & x) = logintemp
            Sheets("Backup").Range("B" & x) = job1
            x = x + 1
        End If
    Next cell
    
    'copy cells from backup sheet to paste on Search_By_Job page
    Sheets("Backup").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy

    Sheets("Search_By_Job").Select
    Range("E3").Select
    ActiveSheet.Paste
    
    'Format main table with borders
    Range("E2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

End Sub


Sub NotTrainedOnsite()

'
'Created by Michael Klink

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets("REF").Visible = True
Sheets("FCLM").Visible = True
Sheets("FLEX").Visible = True
Sheets("Onsite").Visible = True
Sheets("Filtered").Visible = True
Sheets("Backup").Visible = True
Sheets("Roster").Visible = True
Sheets("Trained").Visible = True

Sheets("Search_By_Module").Activate
Range("F7:F8").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Clear


Call Pull_Data_Mod

Dim rownum As Integer
Dim EIDtemp As String
Dim logintemp As String
Dim rowcounter As Integer
Dim rng1 As Object
Dim x As Integer

Sheets("Backup").Select
Range("A:Z").Select
Selection.Clear

Sheets("Roster").Select
x = 1  ' counter
lrow = Cells(Sheets("Roster").Rows.Count, 1).End(xlUp).Row  'last row variable

'dynamically select cell range
With Sheets("Roster")
    Range("AA1").Select
    'Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set rng1 = Selection
End With

'loop through cells in selected range checking if they match job field
'when it reaches the last row, send successful message and exit loop
    For Each cell In rng1
        If (cell.Row) = lrow Then
            MsgBox "Compilation Complete", , "SUCCESS!"
            Exit For
        End If
        If cell = Sheets("REF").Range("$K$1") And Sheets("Roster").Range("AB" & (cell.Row)) = Sheets("REF").Range("$K$2") Then
            EIDtemp = Sheets("Roster").Range("A" & (cell.Row))
            logintemp = Sheets("Roster").Range("B" & (cell.Row))
            Sheets("Backup").Range("B" & x) = EIDtemp
            Sheets("Backup").Range("A" & x) = logintemp
            x = x + 1
        End If
    Next cell
    
    'copy cells from backup sheet to paste on Search_By_Job page
    Sheets("Backup").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy

    Sheets("Search_By_Module").Select
    Range("F7").Select
    ActiveSheet.Paste
    
    Sheets("Search_By_Module").Select
    ActiveSheet.Range("A1").Select

Sheets("REF").Visible = False
Sheets("FCLM").Visible = False
Sheets("FLEX").Visible = False
Sheets("Onsite").Visible = False
Sheets("Filtered").Visible = False
Sheets("Backup").Visible = False
Sheets("Roster").Visible = False
Sheets("Trained").Visible = False


Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

Sub TrainedOnsite()

'
'Created by Michael Klink

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheets("Search_By_Module").Activate
Range("F7:F8").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Clear

Sheets("REF").Visible = True
Sheets("FCLM").Visible = True
Sheets("FLEX").Visible = True
Sheets("Onsite").Visible = True
Sheets("Filtered").Visible = True
Sheets("Backup").Visible = True
Sheets("Roster").Visible = True
Sheets("Trained").Visible = True

Call Pull_Data_Mod

Dim rownum As Integer
Dim EIDtemp As String
Dim logintemp As String
Dim rowcounter As Integer
Dim rng1 As Object
Dim x As Integer

Sheets("Backup").Select
Range("A:Z").Select
Selection.Clear

Sheets("Roster").Select
x = 1  ' counter
lrow = Cells(Sheets("Roster").Rows.Count, 1).End(xlUp).Row  'last row variable

'dynamically select cell range
With Sheets("Roster")
    Range("AA1").Select
    'Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Set rng1 = Selection
End With

'loop through cells in selected range checking if they match job field
'when it reaches the last row, send successful message and exit loop
    For Each cell In rng1
        If (cell.Row) = lrow Then
            MsgBox "Compilation Complete", , "SUCCESS!"
            Exit For
        End If
        If cell = Sheets("REF").Range("$K$1") And Sheets("Roster").Range("AB" & (cell.Row)) = Sheets("REF").Range("$K$1") Then
            EIDtemp = Sheets("Roster").Range("A" & (cell.Row))
            logintemp = Sheets("Roster").Range("B" & (cell.Row))
            Sheets("Backup").Range("B" & x) = EIDtemp
            Sheets("Backup").Range("A" & x) = logintemp
            x = x + 1
        End If
    Next cell
    
    'copy cells from backup sheet to paste on Search_By_Job page
    Sheets("Backup").Select
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy

    Sheets("Search_By_Module").Select
    Range("F7").Select
    ActiveSheet.Paste
    
Sheets("Search_By_Module").Select
ActiveSheet.Range("A1").Select

Sheets("REF").Visible = False
Sheets("FCLM").Visible = False
Sheets("FLEX").Visible = False
Sheets("Onsite").Visible = False
Sheets("Filtered").Visible = False
Sheets("Backup").Visible = False
Sheets("Roster").Visible = False
Sheets("Trained").Visible = False


Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub



