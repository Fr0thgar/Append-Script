# Import the necessary modules
Import-Module SharePointPnPPowerShellOnline
Import-Module ImportExcel

# Define Variables
$siteUrl = # Your sharepoint url here
$calendarName = # Your Calendar name here
$excelFilePath = # Your excel file path here
$appendText = # Text to append to the title if needed

# Prompt for credentials
$credentials = Get-Credential

# Connect to SharePoint Online site
Connect-PnPOnline -Url $siteUrl -Credentials $credentials

# Read the Excel file
$events = Import-Excel -Path $excelFilePath

# Loop through each row in the Excel sheet
foreach ($event in $events) {
    # Extract event details from the current row
    $eventTitle = $event.Title

    try {
        # Find the existing event based on Title only
        $existingEvent = Get-PnPListItem -List $calendarName -Query "<View><Query><Where><Eq><FieldRef Name='Title' /><Value Type='Text'>$eventTitle</Value></Eq></Where></Query></View>"

        if ($existingEvent) {
            # Append text to the existing title
            $updatedTitle = $existingEvent["Title"] + $appendText

            # Update the Title field of the existing event
            Set-PnPListItem -List $calendarName -Identity $existingEvent.Id -Values @{Title = $updatedTitle}
            Write-Host "Updated event title: $eventTitle to new title: $updatedTitle"
        } else {
            Write-Warning "Event not found: $eventTitle"
        }
    } catch {
        Write-Error "Failed to update event: $($event.Title). Error: $_"
    }
}

# Disconnect the PnP session
Disconnect-PnPOnline

Write-Host "All events processed."