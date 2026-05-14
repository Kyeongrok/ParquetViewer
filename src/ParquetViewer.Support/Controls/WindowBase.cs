using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;

namespace ParquetViewer.Controls
{
    public class WindowBase : Window
    {
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1 = 19;
        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        protected WindowBase()
        {
            Loaded += OnLoaded;
            AppSettings.DarkModeChanged += OnDarkModeChanged;
            Closed += (_, _) => AppSettings.DarkModeChanged -= OnDarkModeChanged;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ApplyTitleBarTheme(AppSettings.DarkMode);
        }

        private void OnDarkModeChanged(bool isDark)
        {
            ApplyTitleBarTheme(isDark);
        }

        private void ApplyTitleBarTheme(bool isDark)
        {
            var handle = new WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero) return;
            UseImmersiveDarkMode(handle, isDark);
        }

        private static bool UseImmersiveDarkMode(IntPtr handle, bool enabled)
        {
            if (!IsWindows10OrGreater(17763)) return false;

            var attribute = IsWindows10OrGreater(18985)
                ? DWMWA_USE_IMMERSIVE_DARK_MODE
                : DWMWA_USE_IMMERSIVE_DARK_MODE_BEFORE_20H1;

            int useImmersiveDarkMode = enabled ? 1 : 0;
            return DwmSetWindowAttribute(handle, attribute, ref useImmersiveDarkMode, sizeof(int)) == 0;
        }

        private static bool IsWindows10OrGreater(int build = -1)
            => Environment.OSVersion.Version.Major >= 10 && Environment.OSVersion.Version.Build >= build;
    }
}
