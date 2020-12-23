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
Dim RosterURL As String

Sheets("FCLM").Activate
Range("A:F").Clear
Sheets("FLEX").Activate
Range("A:Q").Clear
Sheets("Roster").Activate
Range("A:H").Clear

'Pull Data Variable Components:
'   Use these to make the URL dynamic/variable
strtDte = Format(Sheets("Search_By_Job").Range("C5"), "yyyy/mm/dd")
endDte = Format(Sheets("Search_By_Job").Range("C6"), "yyyy/mm/dd")
strthr = Sheets("Search_By_Job").Range("C7")
endHr = Sheets("Search_By_Job").Range("C8")
FC = Sheets("Search_By_Job").Range("C4")

'Insert URL Here inside of the quotes
'Example: URL = "https://www.inside.amazon.com"
'Replace the text in the URL that references above Variable Components with assigned call name (__ = )
'Add in a " & __ & " to replace the URL portion
'Example: https://fclm-portal.amazon.com/reports/employeeRoster?reportFormat=CSV&warehouseId=LGB3 --> https://fclm-portal.amazon.com/reports/employeeRoster?reportFormat=CSV&warehouseId=" & FC & "


fclmURL = "https://fclm-portal.amazon.com/reports/timeOnTask?reportFormat=CSV&warehouseId=" & FC & "&maxIntradayDays=30&spanType=Intraday&startDateIntraday=" & strtDte & "&startHourIntraday=" & strthr & "&startMinuteIntraday=0&endDateIntraday=" & endDte & "&endHourIntraday=" & endHr & "&endMinuteIntraday=0"
flexURL = "https://flex.corp.amazon.com/PSP1/exports?file=" & FC & "%2FALL_JOBS.csv.gz"
RosterURL = "https://fclm-portal.amazon.com/employee/employeeRoster?reportFormat=CSV&warehouseId=" & FC & "&employeeStatusActive=true&_employeeStatusActive=on&_employeeStatusLeaveOfAbsence=on&_employeeStatusExempt=on&employeeTypeAmzn=true&_employeeTypeAmzn=on&employeeTypeTemp=true&_employeeTypeTemp=on&employeeType3Pty=true&_employeeType3Pty=on&Employee+ID=Employee+ID&User+ID=User+ID&Employee+Name=Employee+Name&Badge+Barcode+ID=Badge+Barcode+ID&Department+ID=Department+ID&Employment+Start+Date=Employment+Start+Date&Manager+Name=Manager+Name&Job+Title=Job+Title&hideColumns=Photo%2CEmployment+Type%2CEmployee+Status%2CTemp+Agency+Code%2CManagement+Area+ID%2CShift+Pattern%2CBadge+RFID%2CExempt&submit=true"
Call ImportCSV(WebScrape(fclmURL, winhttp), Sheets("FCLM").Range("A1"))
Call ImportCSV(WebScrape(flexURL, winhttp), Sheets("FLEX").Range("A1"))
Call ImportCSV(WebScrape(RosterURL, winhttp), Sheets("Roster").Range("A1"))


End Sub


Sub Pull_Data_Mod()

Application.DisplayAlerts = False
Application.ScreenUpdating = False

' <--- Adding a " ' " before a string of text will flag it as a comment rather than part of the VBA code.
Dim winhttp As Object
Dim w As Object: Set w = CreateObject("WinHTTP.WinHTTPRequest.5.1")
Dim httpRequest As XMLHTTP
Dim DataObj As New MSForms.DataObject
Dim URL$
Dim SecDte As String

'Pull Data Variable Components:
'   Use these to make the URL dynamic/variable
strtDte = Format(Sheets("Search_By_Job").Range("C5"), "yyyy/mm/dd")
endDte = Format(Sheets("Search_By_Job").Range("C6"), "yyyy/mm/dd")
SecDte = Sheets("Search_By_Module").Range("C19")
StdDte = Format(Sheets("Search_By_Module").Range("C20"), "yyyy/mm/dd")
FC = Sheets("Search_By_Module").Range("C18")
Module = Sheets("Search_By_Module").Range("C21")
strthr = Sheets("Search_By_Job").Range("C7")
endHr = Sheets("Search_By_Job").Range("C8")

Sheets("FCLM").Activate
Range("A:F").Clear
Sheets("FLEX").Activate
Range("A:Q").Clear
Sheets("Roster").Activate
Range("A:H").Clear

'Insert URL Here inside of the quotes
'Example: URL = "https://www.inside.amazon.com"
'Replace the text in the URL that references above Variable Components with assigned call name (__ = )
'Add in a " & __ & " to replace the URL portion
'Example: https://fclm-portal.amazon.com/reports/employeeRoster?reportFormat=CSV&warehouseId=LGB3 --> https://fclm-portal.amazon.com/reports/employeeRoster?reportFormat=CSV&warehouseId=" & FC & "
Dim TrainedURL As String
Dim RosterURL As String
Dim TOTURL As String

Sheets("Backup").Select
Range("A:B").Select
Selection.Clear

TrainedURL = "https://fclearning.amazon.com/fcl/summaryreport?trained=1&download=true&selectedWarehouseIds=" & FC & "&managers=&modules=" & Module & "&moduleType=All&startDateInSecs=" & SecDte & "&endDateInSecs=" & SecDte
RosterURL = "https://fclm-portal.amazon.com/employee/employeeRoster?reportFormat=CSV&warehouseId=" & FC & "&employeeStatusActive=true&_employeeStatusActive=on&_employeeStatusLeaveOfAbsence=on&_employeeStatusExempt=on&employeeTypeAmzn=true&_employeeTypeAmzn=on&employeeTypeTemp=true&_employeeTypeTemp=on&employeeType3Pty=true&_employeeType3Pty=on&Employee+ID=Employee+ID&User+ID=User+ID&Employee+Name=Employee+Name&Badge+Barcode+ID=Badge+Barcode+ID&Department+ID=Department+ID&Manager+Name=Manager+Name&hideColumns=Photo%2CEmployment+Start+Date%2CEmployment+Type%2CEmployee+Status%2CTemp+Agency+Code%2CJob+Title%2CManagement+Area+ID%2CShift+Pattern%2CBadge+RFID%2CExempt&submit=true"
TOTURL = "https://fclm-portal.amazon.com/reports/timeOnTask?reportFormat=CSV&warehouseId=" & FC & "&maxIntradayDays=30&spanType=Intraday&startDateIntraday=" & strtDte & "&startHourIntraday=" & strthr & "&startMinuteIntraday=0&endDateIntraday=" & endDte & "&endHourIntraday=" & endHr & "&endMinuteIntraday=0"

Call ImportCSV(WebScrape(TrainedURL, winhttp), Sheets("Trained").Range("A1"))
Call ImportCSV(WebScrape(RosterURL, winhttp), Sheets("Roster").Range("A1"))
Call ImportCSV(WebScrape(TOTURL, winhttp), Sheets("FCLM").Range("A1"))

Application.DisplayAlerts = True
Application.ScreenUpdating = True

End Sub

