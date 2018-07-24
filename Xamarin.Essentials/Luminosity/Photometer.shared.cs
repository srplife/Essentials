using System;

namespace Xamarin.Essentials
{
    public static partial class Photometer
    {
        static bool useSyncContext;

        public static event EventHandler<LightChangedEventArgs> IntensityChanged;

        public static bool IsMonitoring { get; private set; }

        public static bool ApplyLowPassFilter { get; set; }

        public static void Start(SensorSpeed sensorSpeed)
        {
            if (!IsSupported)
                throw new FeatureNotSupportedException();

            if (IsMonitoring)
                return;

            IsMonitoring = true;
            useSyncContext = sensorSpeed == SensorSpeed.Normal || sensorSpeed == SensorSpeed.UI;

            try
            {
                PlatformStart(sensorSpeed);
            }
            catch
            {
                IsMonitoring = false;
                throw;
            }
        }

        public static void Stop()
        {
            if (!IsSupported)
                throw new FeatureNotSupportedException();

            if (!IsMonitoring)
                return;

            IsMonitoring = false;

            try
            {
                PlatformStop();
            }
            catch
            {
                IsMonitoring = true;
                throw;
            }
        }

        internal static void OnChanged(LightData reading)
            => OnChanged(new LightChangedEventArgs(reading));

        internal static void OnChanged(LightChangedEventArgs e)
        {
            var handler = IntensityChanged;
            if (handler == null)
                return;

            if (useSyncContext)
                MainThread.BeginInvokeOnMainThread(() => handler?.Invoke(null, e));
            else
                handler?.Invoke(null, e);
        }
    }

    public class LightChangedEventArgs : EventArgs
    {
        internal LightChangedEventArgs(LightData intensity) => Intensity = intensity;

        public LightData Intensity { get; }
    }

    public readonly struct LightData
    {
        internal LightData(double lx) =>
            Illuminance = lx;

        public double Illuminance { get; }
    }
}
