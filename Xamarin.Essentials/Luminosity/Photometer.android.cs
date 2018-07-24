using System;
using Android.Hardware;
using Android.Runtime;
using Xamarin.Essentials.Sensors;

namespace Xamarin.Essentials
{
    public static partial class Photometer
    {
        internal static bool IsSupported =>
            Platform.SensorManager?.GetDefaultSensor(SensorType.Light) != null;

        static ISensorEventListener listener;
        static Sensor photometer;

        internal static void PlatformStart(SensorSpeed sensorSpeed)
        {
            var delay = sensorSpeed.ToNative();
            photometer = Platform.SensorManager.GetDefaultSensor(SensorType.Light);
            listener = new LightSensorListener();
            Platform.SensorManager.RegisterListener(listener, photometer, delay);
        }

        internal static void PlatformStop()
        {
            if (listener == null)
                return;

            Platform.SensorManager.UnregisterListener(listener, photometer);
            listener.Dispose();
            listener = null;
        }
    }

    class LightSensorListener : Java.Lang.Object, ISensorEventListener, IDisposable
    {
        void ISensorEventListener.OnAccuracyChanged(Sensor sensor, [GeneratedEnum] SensorStatus accuracy)
        {
        }

        void ISensorEventListener.OnSensorChanged(SensorEvent e)
        {
            Photometer.OnChanged(new LightData(e.Values[0]));
        }
    }
}
