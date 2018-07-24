using System;
using Android.Hardware;

namespace Xamarin.Essentials.Sensors
{
    internal static class SensorExtensions
    {
        internal static SensorDelay ToNative(this SensorSpeed speed)
        {
            switch (speed)
            {
                case SensorSpeed.Normal:
                    return SensorDelay.Normal;
                case SensorSpeed.Fastest:
                    return SensorDelay.Fastest;
                case SensorSpeed.Game:
                    return SensorDelay.Game;
                case SensorSpeed.UI:
                    return SensorDelay.Ui;
                default:
                    throw new InvalidOperationException();
            }
        }
    }
}
