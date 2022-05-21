// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System.Diagnostics.CodeAnalysis;
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
            if (_graphClient is null || query is null || !delayedExecution || string.IsNullOrWhiteSpace(query.Search))
            {
                return new List<Result>(0);
            }

            var request = _graphClient.Me.People.Request().Select(p => new
            {
                p.DisplayName,
                p.Department,
                p.UserPrincipalName,
                p.JobTitle,
                p.OfficeLocation,
                p.ScoredEmailAddresses,
                p.PersonType,
            });

            request.QueryOptions.Add(new QueryOption("search", query.Search));

            var people = request.GetAsync().Result;
            return people.Select(p =>
            {
                var email = p.UserPrincipalName ?? p.ScoredEmailAddresses.FirstOrDefault()?.Address;

                var result = new Result
                {
                    Title = p.DisplayName,
                    QueryTextDisplay = p.DisplayName,
                    Action = (_) =>
                    {
                        string launchUri;
                        if (p.PersonType.Class == "Person")
                        {
                            launchUri = $"MSTeams:/l/chat/0/0?users={p.UserPrincipalName}";
                        }
                        else
                        {
                            launchUri = $"mailto:{email}";
                        }

                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { UseShellExecute = true, FileName = launchUri });

                        return true;
                    },
                    ContextData = p,
                };

                if (p.PersonType.Class == "Person")
                {
                    result.Icon = () => GetPhoto(p.Id);
                    result.SubTitle = @$"{p.JobTitle} - {p.Department} - {email} - {p.OfficeLocation}";
                }
                else
                {
                    result.IcoPath = _iconPath;
                    result.SubTitle = @$"{email}";
                }

                return result;
            }).ToList();
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
            var request = _graphClient!.Users[id].Photos["48x48"].Content.Request();
            using var stream = request.GetAsync().Result;
            return StreamToBitmapImage(stream);
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
    }
}
