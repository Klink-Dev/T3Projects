Attribute VB_Name = "Bread"

'Created by Michael Klink
'github.com/Klink-Dev

Sub jobsearch()
Application.ScreenUpdating = False

Sheets("REF").Visible = True
Sheets("FCLM").Visible = True
Sheets("FLEX").Visible = True
Sheets("Onsite").Visible = True
Sheets("Filtered").Visible = True
Sheets("Backup").Visible = True


Sheets("Search_By_Job").Select
Range("E3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    Range("E2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

Call Pull_Data
Call jobSplit
Call filterOnsite

Sheets("Search_By_Job").Activate
Range("B2").Select

Sheets("REF").Visible = False
Sheets("FCLM").Visible = False
Sheets("FLEX").Visible = False
Sheets("Onsite").Visible = False
Sheets("Filtered").Visible = False
Sheets("Backup").Visible = False

Application.ScreenUpdating = True

End Sub
Sub searchlogin()

InputBox "Please enter a login", "Employee job search"


Application.ScreenUpdating = False

Sheets("REF").Visible = True
Sheets("FCLM").Visible = True
Sheets("FLEX").Visible = True
Sheets("Onsite").Visible = True
Sheets("Filtered").Visible = True
Sheets("Backup").Visible = True

'placeholder
'placeholder


Sheets("Search_By_Job").Activate
Range("B2").Select

Sheets("REF").Visible = False
Sheets("FCLM").Visible = False
Sheets("FLEX").Visible = False
Sheets("Onsite").Visible = False
Sheets("Filtered").Visible = False
Sheets("Backup").Visible = False

Application.ScreenUpdating = True

End Sub
