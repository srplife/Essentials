using System;
using Windows.Foundation.Metadata;
using Windows.Graphics.Display;
using Windows.Security.ExchangeActiveSyncProvisioning;
using Windows.Storage.Streams;
using Windows.System.Profile;
using Windows.UI.ViewManagement;

namespace Xamarin.Essentials
{
    public static partial class DeviceInfo
    {
        static readonly EasClientDeviceInformation deviceInfo;

        static DeviceInfo()
        {
            deviceInfo = new EasClientDeviceInformation();
        }

        static string GetID()
        {
            if (ApiInformation.IsTypePresent("Windows.System.Profile.SystemIdentification"))
            {
                var info = SystemIdentification.GetSystemIdForPublisher();
                if (info.Source != SystemIdentificationSource.None)
                {
                    var dataReader = DataReader.FromBuffer(info.Id);

                    var bytes = new byte[info.Id.Length];
                    dataReader.ReadBytes(bytes);

                    return Convert.ToBase64String(bytes);
                }
            }
            if (ApiInformation.IsTypePresent("Windows.System.Profile.HardwareIdentification"))
            {
                var hardwareToken = HardwareIdentification.GetPackageSpecificToken(null);
                var dataReader = DataReader.FromBuffer(hardwareToken.Id);

                var bytes = new byte[hardwareToken.Id.Length];
                dataReader.ReadBytes(bytes);

                return Convert.ToBase64String(bytes);
            }
            return string.Empty;
        }

        static string GetModel() => deviceInfo.SystemProductName;

        static string GetManufacturer() => deviceInfo.SystemManufacturer;

        static string GetDeviceName() => deviceInfo.FriendlyName;

        static string GetVersionString()
        {
            var version = AnalyticsInfo.VersionInfo.DeviceFamilyVersion;

            if (ulong.TryParse(version, out var v))
            {
                var v1 = (v & 0xFFFF000000000000L) >> 48;
                var v2 = (v & 0x0000FFFF00000000L) >> 32;
                var v3 = (v & 0x00000000FFFF0000L) >> 16;
                var v4 = v & 0x000000000000FFFFL;
                return $"{v1}.{v2}.{v3}.{v4}";
            }

            return version;
        }

        static string GetPlatform() => Platforms.UWP;

        static string GetIdiom()
        {
            switch (AnalyticsInfo.VersionInfo.DeviceFamily)
            {
                case "Windows.Mobile":
                    return Idioms.Phone;
                case "Windows.Universal":
                case "Windows.Desktop":
                    {
                        var uiMode = UIViewSettings.GetForCurrentView().UserInteractionMode;
                        return uiMode == UserInteractionMode.Mouse ? Idioms.Desktop : Idioms.Tablet;
                    }
                case "Windows.Xbox":
                case "Windows.Team":
                    return Idioms.TV;
                case "Windows.IoT":
                    return Idioms.Unsupported;
            }

            return Idioms.Unsupported;
        }

        static DeviceType GetDeviceType()
        {
            var isVirtual = deviceInfo.SystemProductName == "Virtual";

            if (isVirtual)
                return DeviceType.Virtual;

            return DeviceType.Physical;
        }
    }
}
