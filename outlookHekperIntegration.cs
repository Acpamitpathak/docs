// Initialize Graph (once per app run)
await OutlookHelper.InitializeGraphAsync(
    clientId: "YOUR_APP_CLIENT_ID",
    tenantId: "YOUR_TENANT_ID",
    clientSecret: "YOUR_APP_SECRET"
);

// Create event directly in Outlook
await OutlookHelper.AddEventAsync(currentEvent, "youremail@yourdomain.com");

Console.WriteLine("✅ Event added to Outlook successfully!");