using System;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Generic;
using Ical.Net.CalendarComponents;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;

namespace YourNamespace.Helpers
{
    public static class OutlookHelper
    {
        private static GraphServiceClient _graphClient;

        // ✅ Initializes the Graph client using token-based authentication (works in .NET 8)
        public static async Task InitializeGraphAsync(string clientId, string tenantId, string clientSecret)
        {
            // 1️⃣ Acquire a token from Azure AD using MSAL
            var app = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .Build();

            string[] scopes = new[] { "https://graph.microsoft.com/.default" };
            var result = await app.AcquireTokenForClient(scopes).ExecuteAsync();
            string token = result.AccessToken;

            // 2️⃣ Create a GraphServiceClient using that token
            _graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
                (request) =>
                {
                    request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                    return Task.CompletedTask;
                }));
        }

        // ✅ Adds an iCal.Net CalendarEvent directly to Outlook via Microsoft Graph
        public static async Task AddEventAsync(CalendarEvent currentEvent, string targetUserEmail)
        {
            if (_graphClient == null)
                throw new InvalidOperationException("Graph client not initialized. Call InitializeGraphAsync() first.");

            if (currentEvent == null)
                throw new ArgumentNullException(nameof(currentEvent));

            var startUtc = currentEvent.DtStart?.AsUtc ?? DateTime.UtcNow;
            var endUtc = currentEvent.DtEnd?.AsUtc ??
                         (currentEvent.Duration != TimeSpan.Zero ? startUtc + currentEvent.Duration : startUtc.AddHours(1));

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

            await _graphClient.Users[targetUserEmail].Events.Request().AddAsync(graphEvent);
        }
    }
}