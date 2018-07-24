using CoreLocation;

namespace Xamarin.Essentials
{
    public static partial class Photometer
    {
        internal static bool IsSupported =>
            CLLocationManager.HeadingAvailable;

        static CLLocationManager locationManager;

        internal static void PlatformStart(SensorSpeed sensorSpeed)
        {
            IOKit.
            locationManager.UpdatedHeading += LocationManagerUpdatedHeading;
            locationManager.StartUpdatingHeading();
        }

        static bool LocationManagerShouldDisplayHeadingCalibration(CLLocationManager manager) => ShouldDisplayHeadingCalibration;

        static void LocationManagerUpdatedHeading(object sender, CLHeadingUpdatedEventArgs e)
        {
            var data = new CompassData(e.NewHeading.MagneticHeading);
            OnChanged(data);
        }

        internal static void PlatformStop()
        {
            if (locationManager == null)
                return;

            locationManager.ShouldDisplayHeadingCalibration -= LocationManagerShouldDisplayHeadingCalibration;
            locationManager.UpdatedHeading -= LocationManagerUpdatedHeading;
            locationManager.StopUpdatingHeading();
            locationManager.Dispose();
            locationManager = null;
        }
    }
}
