using ParquetViewer.Analytics;
using ParquetViewer.Helpers;
using ParquetViewer.Resources;
using System;
using System.Globalization;
using System.IO;
using System.Windows;

namespace ParquetViewer
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);

            if (AppSettings.UserSelectedCulture is not null)
                CultureInfo.CurrentUICulture = AppSettings.UserSelectedCulture;

            ApplyTheme(AppSettings.DarkMode);
            AppSettings.DarkModeChanged += ApplyTheme;

            if (!System.Diagnostics.Debugger.IsAttached)
            {
                AppDomain.CurrentDomain.UnhandledException += (_, ev) =>
                    HandleException((Exception)ev.ExceptionObject);
                DispatcherUnhandledException += (_, ev) =>
                {
                    HandleException(ev.Exception);
                    ev.Handled = true;
                };
            }

            string? pathToOpen = e.Args.Length > 0
                && (File.Exists(e.Args[0]) || Directory.Exists(e.Args[0]))
                ? e.Args[0]
                : null;

            var mainWindow = new MainWindow(pathToOpen);
            MainWindow = mainWindow;
            mainWindow.Show();
        }

        public static void ApplyTheme(bool isDark)
        {
            var source = isDark
                ? new Uri("/ParquetViewer.Support;component/Themes/DarkTheme.xaml", UriKind.Relative)
                : new Uri("/ParquetViewer.Support;component/Themes/LightTheme.xaml", UriKind.Relative);
            Current.Resources.MergedDictionaries[0] = new ResourceDictionary { Source = source };
        }

        private static void HandleException(Exception ex)
        {
            ExceptionEvent.FireAndForget(ex);
            MessageBox.Show(
                $"{Errors.GenericErrorMessage} {Errors.CopyErrorMessageText}:{Environment.NewLine}{Environment.NewLine}{ex}",
                ex.Message, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static void GetUserConsentToGatherAnalytics()
        {
            if (AppSettings.ConsentLastAskedOnVersion is null || AppSettings.ConsentLastAskedOnVersion < Env.AssemblyVersion)
            {
                if (AppSettings.AnalyticsDataGatheringConsent)
                {
                    AppSettings.ConsentLastAskedOnVersion = Env.AssemblyVersion;
                    return;
                }

                bool isFirstLaunch = AppSettings.ConsentLastAskedOnVersion is null;
                if (isFirstLaunch)
                {
                    AppSettings.ConsentLastAskedOnVersion = new SemanticVersion(0, 0, 0, DateTime.Now.Day);
                }
                else if (AppSettings.ConsentLastAskedOnVersion != new SemanticVersion(0, 0, 0, DateTime.Now.Day))
                {
                    AppSettings.ConsentLastAskedOnVersion = Env.AssemblyVersion;
                    if (MessageBox.Show(
                        Strings.AnalyticsConsentPromptMessage,
                        Strings.AnalyticsConsentPromptTitle,
                        MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        AppSettings.AnalyticsDataGatheringConsent = true;
                    }
                }
            }
        }

        public static void AskUserIfTheyWantToSwitchToDarkMode()
        {
            if (!AppSettings.DarkMode
                && (AppSettings.OpenedFileCount == 30 || AppSettings.OpenedFileCount == 300)
                && (Env.AppsUseDarkTheme == true || Env.SystemUsesDarkTheme == true))
            {
                if (MessageBox.Show(
                    Strings.DarkModePromptMessage,
                    Strings.DarkModePromptTitle,
                    MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    AppSettings.DarkMode = true;
                }
            }
        }
    }
}
