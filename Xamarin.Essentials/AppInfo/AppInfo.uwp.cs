using System.Diagnostics;
using System.Globalization;
using Windows.ApplicationModel;
using Windows.Graphics.Display;

namespace Xamarin.Essentials
{
    public static partial class AppInfo
    {
        static BrightnessOverride @override;

        static string PlatformGetPackageName() => Package.Current.Id.Name;

        static Brightness PlatformBrightness => new Brightness(BrightnessOverride.GetForCurrentView().BrightnessLevel);

        static string PlatformGetName() => Package.Current.DisplayName;

        static string PlatformGetVersionString()
        {
            var version = Package.Current.Id.Version;
            return $"{version.Major}.{version.Minor}.{version.Build}.{version.Revision}";
        }

        static string PlatformGetBuild() =>
            Package.Current.Id.Version.Build.ToString(CultureInfo.InvariantCulture);

        static void PlatformOpenSettings() =>
            Windows.System.Launcher.LaunchUriAsync(new System.Uri("ms-settings:appsfeatures-app")).WatchForError();

        static void PlatformSetBrightness(Brightness brightness)
        {
            @override = BrightnessOverride.GetForCurrentView();
            @override.SetBrightnessLevel(brightness.Value, DisplayBrightnessOverrideOptions.None);
            @override.StartOverride();
        }
    }
}
