using Windows.Devices.Sensors;

using WindowsCompass = Windows.Devices.Sensors.Compass;

namespace Xamarin.Essentials
{
    public static partial class Photometer
    {
        // Magic numbers from https://docs.microsoft.com/en-us/uwp/api/windows.devices.sensors.compass.reportinterval#Windows_Devices_Sensors_Compass_ReportInterval
        internal const uint FastestInterval = 8;
        internal const uint GameInterval = 22;
        internal const uint NormalInterval = 33;

        static LightSensor sensor;

        internal static LightSensor DefaultLightSensor =>
            LightSensor.GetDefault();

        internal static bool IsSupported =>
            DefaultLightSensor != null;

        internal static void PlatformStart(SensorSpeed sensorSpeed)
        {
            sensor = sensor ?? DefaultLightSensor;

            var interval = NormalInterval;
            switch (sensorSpeed)
            {
                case SensorSpeed.Fastest:
                    interval = FastestInterval;
                    break;
                case SensorSpeed.Game:
                    interval = GameInterval;
                    break;
            }

            sensor.ReportInterval = sensor.MinimumReportInterval >= interval ? sensor.MinimumReportInterval : interval;

            sensor.ReadingChanged += LightSensorReportedInterval;
        }

        static void LightSensorReportedInterval(object sender, LightSensorReadingChangedEventArgs e)
        {
            var data = new LightData(e.Reading.IlluminanceInLux);
            OnChanged(data);
        }

        internal static void PlatformStop()
        {
            sensor.ReadingChanged -= LightSensorReportedInterval;
        }
    }
}
