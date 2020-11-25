Attribute VB_Name = "Scraper"
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Sub Pull_Data()

Application.ScreenUpdating = False



' <--- Adding a " ' " before a string of text will flag it as a comment rather than part of the VBA code.
Dim winhttp As Object
Dim w As Object: Set w = CreateObject("WinHTTP.WinHTTPRequest.5.1")
Dim httpRequest As XMLHTTP
Dim DataObj As New MSForms.DataObject
Dim URL$
Dim fclmURL As String
Dim flexURL As String

'Pull Data Variable Components:
'   Use these to make the URL dynamic/variable
strtDte = Format(Sheets("Search_By_Job").Range("C5"), "yyyy/mm/dd")
endDte = Format(Sheets("Search_By_Job").Range("C6"), "yyyy/mm/dd")
strthr = Sheets("Search_By_Job").Range("C7")
endHr = Sheets("Search_By_Job").Range("C8")
fc = Sheets("Search_By_Job").Range("C4")

'Insert URL Here inside of the quotes
'Example: URL = "https://www.inside.amazon.com"
'Replace the text in the URL that references above Variable Components with assigned call name (__ = )
'Add in a " & __ & " to replace the URL portion
'Example: https://fclm-portal.amazon.com/reports/employeeRoster?reportFormat=CSV&warehouseId=LGB3 --> https://fclm-portal.amazon.com/reports/employeeRoster?reportFormat=CSV&warehouseId=" & FC & "


fclmURL = "https://fclm-portal.amazon.com/reports/timeOnTask?reportFormat=CSV&warehouseId=" & fc & "&maxIntradayDays=30&spanType=Intraday&startDateIntraday=" & strtDte & "&startHourIntraday=" & strthr & "&startMinuteIntraday=0&endDateIntraday=" & endDte & "&endHourIntraday=" & endHr & "&endMinuteIntraday=0"
flexURL = "https://flex.corp.amazon.com/PSP1/exports?file=" & fc & "%2FALL_JOBS.csv.gz"

Call ImportCSV(WebScrape(fclmURL, winhttp), Sheets("FCLM").Range("A1"))
Call ImportCSV(WebScrape(flexURL, winhttp), Sheets("FLEX").Range("A1"))



End Sub

