# Report-CalendarItems.ps1
# Generate a report of calendar items for a mailbox
# V1.0 2-Apr-2025

# GitHub link: https://github.com/12Knocksinna/Office365itpros/blob/master/Report-CalendarItems.PS1

Connect-MgGraph -NoWelcome -Scopes Calendars.ReadBasic, User.ReadBasic.All

$StartDate = (Get-Date).AddDays(-180)
$EndDate = (Get-Date).AddDays(1)

$Start = Get-Date($StartDate) -format s
$End = Get-Date($EndDate) -format s

$User = Get-MgUser -UserId (Get-MgContext).Account

Write-Host ("Fetching calendar information from {0} to {1}..." -f $StartDate, $EndDate)
[array]$Data = Get-MgUserCalendarView -StartDate $Start -EndDateTime $End -UserId $User.Id -All
If ($Data) {
    Write-Host ("Found {0} calendar items" -f $Data.Count)
} Else {
    Write-Host "No calendar items found."
    Break
}
# Find meetings arranged by the user
[array]$Meetings = $Data | Where-Object {$_.Attendees.count -gt 1 -and $_.Organizer.emailaddress.Name -eq $User.displayName}
Write-Host ("{0} meetings found arranged by {1}" -f $Meetings.Count, $User.displayName)

$CalendarInfo = [System.Collections.Generic.List[Object]]::new() 

ForEach ($event in $Meetings) {
    [datetime]$MeetingStart =  Get-Date($Event.start.datetime) 
    [datetime]$MeetingEnd   = Get-Date($Event.end.datetime)

    # Calculate meeting duration in minutes. If it's an all-day event, use 480 minutes
    If ($Event.IsAllDay -eq $False) {
        $Duration =  ($MeetingEnd - $MeetingStart).TotalMinutes 
    }  Else { 
        $Duration = 480 
    }
     
    $OnlineMeetingProvider = $null
    [array]$AllAttendees = ($Event.Attendees | Where-Object {$_.Type -ne "resource"} )
    [array]$RequiredAttendees = ($Event.Attendees | Where-Object {$_.Type -eq "required"}) 
    [array]$OptionalAttendees = ($Event.Attendees | Where-Object {$_.Type -eq "optional"})
    # Create output line - add one to the total attendees to account for the organizer
    If ($Event.onlineMeetingProvider -ne "unknown") {
        $OnlineMeetingProvider = $Event.onlineMeetingProvider
    }
    $DataLine = [PSCustomObject] @{
        Type              = $Event.type
        Organizer         = $Event.organizer.emailaddress.name
        OrganizerEmail    = $Event.organizer.emailaddress.address
        Created           = Get-Date($Event.createdDateTime) -format 'dd-MMM-yyyy HH:mm'
        Modified          = Get-Date($Event.lastModifiedDateTime) -format 'dd-MMM-yyyy HH:mm'
        TimeZone          = $Event.originalStartTimeZone
        Subject           = $Event.Subject
        AllDay            = $Event.IsAllDay
        Online            = $Event.isOnlineMeeting
        OnlineProvider    = $OnlineMeetingProvider
        Start             = Get-Date($MeetingStart) -format 'dd-MMM-yyyy HH:mm'
        End               = Get-Date($MeetingEnd) -format 'dd-MMM-yyyy HH:mm'
        Day               = (Get-Date($MeetingStart)).DayOfWeek
        Duration          = $Duration
        Location          = $event.location.displayname
        RequiredAttendees = $RequiredAttendees.emailaddress.name -join ", "
        OptionalAttendees = $OptionalAttendees.emailaddress.name -join ", "
        TotalAttendees    = $AllAttendees.Count
        Required          = $RequiredAttendees.Count
        Optional          = $OptionalAttendees.Count
        TotalAtEvent      = $AllAttendees.Count + 1
    }
    $CalendarInfo.Add($DataLine)
}

$CalendarInfo = $CalendarInfo | Sort-Object {$_.Start -as [datetime]}
$CalendarInfo | Select-Object Start, End, Subject, Organizer, RequiredAttendees, OnlineProvider | Out-GridView -Title ("Calendar items for {0}" -f $User.displayName) 

Write-Host "Generating report..."
If (Get-Module ImportExcel -ListAvailable) {
    $ExcelGenerated = $True
    Import-Module ImportExcel -ErrorAction SilentlyContinue
    $ExcelOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Calendar Items.xlsx"
    If (Test-Path $ExcelOutputFile) {
        Remove-Item $ExcelOutputFile -ErrorAction SilentlyContinue
    }
    $CalendarInfo | Export-Excel -Path $ExcelOutputFile -WorksheetName "Calendar Items" `
        -Title ("Calendar Items {0}" -f (Get-Date -format 'dd-MMM-yyyy')) -TitleBold -TableName "CalendarItems"
} Else {
    $CSVOutputFile = ((New-Object -ComObject Shell.Application).Namespace('shell:Downloads').Self.Path) + "\Calendar Items.CSV"
    $CalendarInfo | Export-Csv -Path $CSVOutputFile -NoTypeInformation -Encoding Utf8
}
 
If ($ExcelGenerated) {
    Write-Host ("An Excel report of calendar items is available in {0}" -f $ExcelOutputFile)
} Else {    
    Write-Host ("A CSV report of calendar items is available in {0}" -f $CSVOutputFile)
}  

Write-Host "All done..."

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment. 