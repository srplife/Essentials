using System;

namespace Xamarin.Essentials
{
    public static partial class AppInfo
    {
        public static string PackageName => PlatformGetPackageName();

        public static string Name => PlatformGetName();

        public static string VersionString => PlatformGetVersionString();

        public static Version Version => Utils.ParseVersion(VersionString);

        public static string BuildString => PlatformGetBuild();

        public static void OpenSettings() => PlatformOpenSettings();

        public static Brightness Brightness => PlatformBrightness;

        public static void SetBrightness(double value) => SetBrightness(new Brightness(value));

        public static void SetBrightness(Brightness brightness) => PlatformSetBrightness(brightness);
    }
}
