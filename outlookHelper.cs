using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Ical.Net.CalendarComponents;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace YourNamespace.Helpers
{
    public static class OutlookHelper
    {
        private static GraphServiceClient _graphClient;

        /// <summary>
        /// Initialize Microsoft Graph connection (call once, e.g., during startup)
        /// </summary>
        public static void InitializeGraph(string clientId, string tenantId, string clientSecret)
        {
            if (_graphClient != null) return;

            var app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();

            var authProvider = new Microsoft.Graph.Auth.ClientCredentialProvider(app);
            _graphClient = new GraphServiceClient(authProvider);
        }

        /// <summary>
        /// Adds an iCal.Net CalendarEvent to Outlook calendar via Microsoft Graph.
        /// </summary>
        public static async Task AddEventAsync(CalendarEvent currentEvent, string targetUserEmail)
        {
            if (_graphClient == null)
                throw new InvalidOperationException("Graph client not initialized. Call InitializeGraph() first.");

            if (currentEvent == null)
                throw new ArgumentNullException(nameof(currentEvent));

            // 1️⃣ Convert iCal.Net event → Graph event
            var startUtc = currentEvent.DtStart?.AsUtc ?? DateTime.UtcNow;
            var endUtc = currentEvent.DtEnd?.AsUtc ??
                         (currentEvent.Duration != TimeSpan.Zero
                            ? startUtc + currentEvent.Duration
                            : startUtc.AddHours(1));

            var attendees = new List<Attendee>();
            if (currentEvent.Attendees != null)
            {
                foreach (var a in currentEvent.Attendees)
                {
                    var email = a?.Value?.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase) == true
                        ? a.Value.Substring("mailto:".Length)
                        : a?.Value;

                    if (!string.IsNullOrWhiteSpace(email))
                    {
                        attendees.Add(new Attendee
                        {
                            EmailAddress = new EmailAddress
                            {
                                Address = email,
                                Name = string.IsNullOrWhiteSpace(a.CommonName) ? email : a.CommonName
                            },
                            Type = AttendeeType.Required
                        });
                    }
                }
            }

            // 2️⃣ Create Graph event
            var graphEvent = new Event
            {
                Subject = currentEvent.Summary ?? "(No Subject)",
                Body = new ItemBody
                {
                    ContentType = BodyType.Text,
                    Content = currentEvent.Description ?? string.Empty
                },
                Start = new DateTimeTimeZone
                {
                    DateTime = startUtc.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "UTC"
                },
                End = new DateTimeTimeZone
                {
                    DateTime = endUtc.ToString("yyyy-MM-ddTHH:mm:ss"),
                    TimeZone = "UTC"
                },
                Location = new Location { DisplayName = currentEvent.Location ?? string.Empty },
                Attendees = attendees,
                IsOnlineMeeting = false
            };

            // 3️⃣ Add to Outlook Calendar
            await _graphClient.Users[targetUserEmail].Events.Request().AddAsync(graphEvent);
        }
    }
}