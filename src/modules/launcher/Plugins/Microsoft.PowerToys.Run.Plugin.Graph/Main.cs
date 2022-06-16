// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Runtime.Caching;
using System.Windows.Media.Imaging;
using ManagedCommon;
using Microsoft.Graph;
using Microsoft.PowerToys.Run.Plugin.Graph.Properties;
using Wox.Plugin;

namespace Microsoft.PowerToys.Run.Plugin.Graph
{
    /// <summary>
    /// A power launcher plugin to search across time zones.
    /// </summary>
    public class Main : IPlugin, IDelayedExecutionPlugin, IPluginI18n
    {
        /// <summary>
        /// A helper library for accessing the Microsoft Graph.
        /// </summary>
        private readonly GraphServiceClient? _graphClient;

        private readonly FileCache _cache;

        /// <summary>
        /// The initial context for this plugin (contains API and meta-data)
        /// </summary>
        private PluginInitContext? _context;

        /// <summary>
        /// The path to the icon for each result
        /// </summary>
        private string _iconPath;

        /// <summary>
        /// Initializes a new instance of the <see cref="Main"/> class.
        /// </summary>
        public Main()
        {
            UpdateIconPath(Theme.Light);
            _graphClient = GraphHelper.GetGraphClient("d3590ed6-52b3-4102-aeff-aad2292ab01c", new[] { "People.Read" });
            _cache = new FileCache();
        }

        /// <summary>
        /// Gets the localized name.
        /// </summary>yeahyeahyeah
        public string Name => Resources.PluginTitle;

        /// <summary>
        /// Gets the localized description.
        /// </summary>
        public string Description => Resources.PluginDescription;

        /// <summary>
        /// Initialize the plugin with the given <see cref="PluginInitContext"/>
        /// </summary>
        /// <param name="context">The <see cref="PluginInitContext"/> for this plugin</param>
        public void Init(PluginInitContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));

            _context.API.ThemeChanged += OnThemeChanged;
            UpdateIconPath(_context.API.GetCurrentTheme());
        }

        /// <summary>
        /// Return a filtered list, based on the given query
        /// </summary>
        /// <param name="query">The query to filter the list</param>
        /// <returns>A filtered list, can be empty when nothing was found</returns>
        public List<Result> Query(Query query)
        {
            return Query(query, false);
        }

        /// <summary>
        /// Return a filtered list, based on the given query
        /// </summary>
        /// <param name="query">The query to filter the list</param>
        /// <returns>A filtered list, can be empty when nothing was found</returns>
        public List<Result> Query(Query query, bool delayedExecution)
        {
            if (_graphClient is null || query is null || !delayedExecution)
            {
                return new List<Result>(0);
            }

            var peopleResults = GetPeopleResults(query.Search);
            var meetingResults = GetMeetingResults(query.Search);

            var results = peopleResults.Union(meetingResults).ToList();
            return results;
       }

        [Serializable]
        private record PersonRecord(string Id, string DisplayName, string PersonType, string Department, string JobTitle, string OfficeLocation, string? Email, string? UserPrincipalName);

        private IList<PersonRecord> GetPeopleFromGraph(string query)
        {
            var request = _graphClient!.Me.People.Request().Select(p => new
            {
                p.DisplayName,
                p.Department,
                p.UserPrincipalName,
                p.JobTitle,
                p.OfficeLocation,
                p.ScoredEmailAddresses,
                p.PersonType,
            });

            query = string.IsNullOrWhiteSpace(query) ? string.Empty : $"\"{query}\"";
            request.QueryOptions.Add(new QueryOption("search", query));

            var people = request.GetAsync().Result;

            return people.Select(p =>
            {
                var email = p.UserPrincipalName ?? p.ScoredEmailAddresses.FirstOrDefault()?.Address;
                return new PersonRecord(
                p.Id,
                p.DisplayName,
                p.PersonType.Class,
                p.Department,
                p.JobTitle,
                p.OfficeLocation,
                email,
                p.UserPrincipalName);
            })
            .ToList();
        }

        private IList<Result> GetPeopleResults(string query)
        {
            var people = GetFromCache(
                $"people-{query}",
                () => GetPeopleFromGraph(query),
                TimeSpan.FromDays(1));

            return people.Select(p =>
            {
                var result = new Result
                {
                    Title = p.DisplayName,
                    QueryTextDisplay = p.DisplayName,
                    Action = (_) =>
                    {
                        string launchUri = p.PersonType == "Person" ? $"MSTeams:/l/chat/0/0?users={p.UserPrincipalName}" : $"mailto:{p.Email}";
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { UseShellExecute = true, FileName = launchUri });

                        return true;
                    },
                    ContextData = p,
                };

                if (p.PersonType == "Person")
                {
                    result.Icon = () => GetPhoto(p.Id);
                    var subTitleItems = new[] { p.JobTitle, p.Department, p.Email, p.OfficeLocation };
                    result.SubTitle = string.Join(" - ", subTitleItems.Where(i => i is not null));
                }
                else
                {
                    result.IcoPath = _iconPath;
                    result.SubTitle = @$"{p.Email}";
                }

                return result;
            }).ToList();
        }

        private IList<Result> GetMeetingResults(string query)
        {
            var meetings = GetFromCache(
                "meetings",
                () => GetMeetingsFromGraph(DateTime.Now.AddDays(-14), DateTime.Now.AddDays(14)),
                TimeSpan.FromDays(1));

            return meetings
                .Where(m => string.IsNullOrWhiteSpace(query)
                    || m.Subject.ContainsIC(query)
                    || m.OrganizerEmail.ContainsIC(query)
                    || m.OrganizerName.ContainsIC(query))
                .Where(m => !(m.IsAllDay ?? false) || (m.IsCancelled ?? false))
                .Select(m => new { Meeting = m, Delta = Math.Min(Math.Abs((DateTime.Now - m.Start).TotalMinutes), Math.Abs((DateTime.Now - m.End).TotalMinutes) * 2) })
                .OrderBy(m => m.Delta)
                .DistinctBy(m => new { m.Meeting.Subject, m.Meeting.JoinUrl })
                .Take(10)
                .Select((meeting, i) =>
                {
                    var m = meeting.Meeting;
                    bool isNow = DateTime.Now > m.Start.AddMinutes(-5) && DateTime.Now < m.End.AddMinutes(5);
                    string link = isNow ? m.JoinUrl ?? m.WebLink : m.WebLink;
                    string title = isNow ? $"IN-PROGRESS: {m.Subject}" : m.Subject;

                    var result = new Result
                    {
                        Title = title,
                        QueryTextDisplay = m.Subject,
                        IcoPath = _iconPath,
                        SubTitle = @$"Organized by {m.OrganizerName}",
                        Action = (_) =>
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { UseShellExecute = true, FileName = link });

                            return true;
                        },
                        ContextData = m,
                    };

                    if (isNow)
                    {
                        result.Score = 500;
                    }
                    else
                    {
                        result.Score = (int)Math.Round(500.0 / meeting.Delta);
                    }

                    return result;
                })
                .ToList();
        }

        private IList<CalendarEvent> GetMeetingsFromGraph(DateTime start, DateTime end)
        {
            var viewOptions = new List<QueryOption>
            {
                new QueryOption("startDateTime", start.ToString("o")),
                new QueryOption("endDateTime", end.ToString("o")),
            };

            var events = _graphClient!.Me
                    .CalendarView
                    .Request(viewOptions)

                    // Send user time zone in request so date/time in
                    // response will be in preferred time zone
                    .Header("Prefer", $"outlook.timezone=\"{TimeZoneInfo.Local.StandardName}\"")
                    .Top(100)

                    // Only return fields app will use
                    .Select(e => new
                    {
                        e.Subject,
                        e.Organizer,
                        e.Start,
                        e.End,
                        e.IsAllDay,
                        e.IsCancelled,
                        e.OnlineMeeting,
                        e.WebLink,
                    })

                    // Order results chronologically
                    .OrderBy("start/dateTime")
                    .GetAsync()
                    .Result;

            return events.CurrentPage
                .Select(e => new CalendarEvent(
                    e.Subject,
                    e.Organizer.EmailAddress.Name,
                    e.Organizer.EmailAddress.Address,
                    e.IsAllDay,
                    e.IsCancelled,
                    DateTime.Parse(e.Start.DateTime, CultureInfo.InvariantCulture),
                    DateTime.Parse(e.End.DateTime, CultureInfo.InvariantCulture),
                    e.OnlineMeeting?.JoinUrl.Replace("https://teams.microsoft.com", "MSTeams:"),
                    e.WebLink))
                .ToList();
        }

        private static BitmapImage StreamToBitmapImage(Stream stream)
        {
            var bitmapImage = new BitmapImage();
            bitmapImage.BeginInit();
            bitmapImage.StreamSource = stream;
            bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
            bitmapImage.EndInit();
            bitmapImage.Freeze();

            return bitmapImage;
        }

        private BitmapImage GetPhoto(string id)
        {
            byte[] bytes = GetFromCache(
                id,
                () =>
                {
                    var request = _graphClient!.Users[id].Photos["48x48"].Content.Request();
                    using var stream = request.GetAsync().Result as MemoryStream;
                    var bytes = stream!.ToArray();
                    return bytes;
                },
                TimeSpan.FromDays(7));

            using MemoryStream stream = new(bytes);
            return StreamToBitmapImage(stream);
        }

        private T GetFromCache<T>(string key, Func<T> valueFactory, TimeSpan expiration)
        {
            var value = (T)_cache[key];
            if (value is null)
            {
                value = valueFactory();
                _cache.Add(key, value, DateTimeOffset.Now.Add(expiration));
            }

            return value;
        }

        /// <summary>
        /// Return the translated plugin title.
        /// </summary>
        /// <returns>A translated plugin title.</returns>
        public string GetTranslatedPluginTitle() => Resources.PluginTitle;

        /// <summary>
        /// Return the translated plugin description.
        /// </summary>
        /// <returns>A translated plugin description.</returns>
        public string GetTranslatedPluginDescription() => Resources.PluginDescription;

        private void OnThemeChanged(Theme currentTheme, Theme newTheme)
        {
            UpdateIconPath(newTheme);
        }

        [MemberNotNull(nameof(_iconPath))]
        private void UpdateIconPath(Theme theme)
        {
            _iconPath = theme == Theme.Light || theme == Theme.HighContrastWhite ? "Images/graph.light.png" : "Images/graph.dark.png";
        }

        [Serializable]
        public record CalendarEvent(string Subject, string OrganizerName, string OrganizerEmail, bool? IsAllDay, bool? IsCancelled, DateTime Start, DateTime End, string? JoinUrl, string WebLink);
    }
}
