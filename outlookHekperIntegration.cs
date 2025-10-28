// Existing: your currentEvent is ready
var currentEvent = new CalendarEvent
{
    Summary = "Test HR Data Mesh Event",
    Description = "Calendar entry for the Event.",
    DtStart = new CalDateTime(2025, 11, 1, 10, 59, 0, "UTC"),
    DtEnd = new CalDateTime(2025, 11, 1, 11, 59, 0, "UTC"),
    Location = "Online"
};

// New: Add directly to Outlook calendar
OutlookHelper.InitializeGraph("YOUR_APP_CLIENT_ID", "YOUR_TENANT_ID", "YOUR_APP_SECRET");
await OutlookHelper.AddEventAsync(currentEvent, "youremail@yourdomain.com");

logger.LogInformation("✅ Event successfully added to Outlook Calendar");