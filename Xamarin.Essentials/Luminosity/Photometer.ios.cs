using System;
using System.Runtime.InteropServices;
using CoreFoundation;
using Foundation;
using ObjCRuntime;

namespace Xamarin.Essentials
{
    enum IOHIDEventType
    {
    kIOHIDEventTypeNULL,
    kIOHIDEventTypeVendorDefined,
    kIOHIDEventTypeButton,
    kIOHIDEventTypeKeyboard,
    kIOHIDEventTypeTranslation,
    kIOHIDEventTypeRotation,
    kIOHIDEventTypeScroll,
    kIOHIDEventTypeScale,
    kIOHIDEventTypeZoom,
    kIOHIDEventTypeVelocity,
    kIOHIDEventTypeOrientation,
    kIOHIDEventTypeDigitizer,
    kIOHIDEventTypeAmbientLightSensor,
    kIOHIDEventTypeAccelerometer,
    kIOHIDEventTypeProximity,
    kIOHIDEventTypeTemperature,
    kIOHIDEventTypeSwipe,
    kIOHIDEventTypeMouse,
    kIOHIDEventTypeProgress,
    kIOHIDEventTypeCount
    };

    enum IOHIDEventTypeValue
    {
    kIOHIDEventFieldAmbientLightSensorLevel = 11 << 16,
    kIOHIDEventFieldAmbientLightSensorRawChannel0,
    kIOHIDEventFieldAmbientLightSensorRawChannel1,
    kIOHIDEventFieldAmbientLightDisplayBrightnessChanged
    };

    public static partial class Photometer
    {
        static readonly int SInt32Type = 3;

        static readonly int kIOHIDEventTypeAmbientLightSensor = 12;

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IntPtr IOHIDEventSystemCreate(IntPtr nullPtr);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IntPtr CFStringCreateWithCString(IntPtr zero, string matching, IntPtr zero2);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IntPtr CFNumberCreate(IntPtr zero, IntPtr tpye, IntPtr matching);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IntPtr IOHIDEventSystemClose(IntPtr handle, IntPtr nullPtr);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern void CFRelease(IntPtr handle);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IOHIDEventType IOHIDEventGetType(IntPtr handle);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern int IOHIDEventGetIntegerValue(IntPtr eventHandle, int kIOHIDEEventFieldConstant);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IntPtr IOHIDEventSystemOpen(IntPtr handle, IOHIDECallback callback, IntPtr zero, IntPtr zero2, IntPtr zero3);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IntPtr CFDictionaryCreate(IntPtr zero, IntPtr[] keys, IntPtr[] nums, int length, IntPtr kCFTypeKeyCallbacks, IntPtr kCFTypeDictionaryValueCallBacks);

        [DllImport("/System/Library/Frameworks/IOKit.framework/IOKit")]
        static extern IntPtr IOHIDEventSystemCopyMatchingServices(IntPtr handle, IntPtr dict, IntPtr null1, IntPtr null2, IntPtr null3, IntPtr null4);

        static IntPtr systemHandle = IntPtr.Zero;
        static IntPtr serviceHandle = IntPtr.Zero;
        static IntPtr KeyCallbacks = IntPtr.Zero;
        static IntPtr ValueCallbacks = IntPtr.Zero;

        delegate void IOHIDECallback(IntPtr target, IntPtr refcon, IntPtr service, IntPtr @event);

        static void OnSensorChanged(IntPtr target, IntPtr refcon, IntPtr service, IntPtr @event)
        {
            if (IOHIDEventGetType(@event) == IOHIDEventType.kIOHIDEventTypeAmbientLightSensor)
            {
                var luxValue = IOHIDEventGetIntegerValue(@event, (int) IOHIDEventTypeValue.kIOHIDEventFieldAmbientLightSensorLevel); // lux Event Field
                var channel0 = IOHIDEventGetIntegerValue(@event, (int) IOHIDEventTypeValue.kIOHIDEventFieldAmbientLightSensorRawChannel0); // ch0 Event Field
                var channel1 = IOHIDEventGetIntegerValue(@event, (int) IOHIDEventTypeValue.kIOHIDEventFieldAmbientLightSensorRawChannel1); // ch1 Event Field
            }
            Photometer.OnChanged(new LightData(e.Values[0]));
        }

        static Photometer()
        {
            var lib = Libraries.CoreFoundation.Handle;
            KeyCallbacks = Dlfcn.GetIndirect(lib, "kCFTypeDictionaryKeyCallBacks");
            ValueCallbacks = Dlfcn.GetIndirect(lib, "kCFTypeDictionaryValueCallBacks");
        }

        internal static void PlatformStart(SensorSpeed sensorSpeed)
        {
            if (systemHandle != IntPtr.Zero)
                throw new InvalidOperationException();
            systemHandle = IOHIDEventSystemCreate(IntPtr.Zero);
            var page = 0xff00;
            var usage = 4;
            var dictionary = new NSDictionary<NSString, NSNumber>();
            dictionary.Add(new CFString("PrimaryUsagePage"), new NSNumber(page));
            dictionary.Add(new CFString("PrimaryUsage"), new NSNumber(page));
            var services = (CFArray)ObjCRuntime.Runtime.GetNSObject(IOHIDEventSystemCopyMatchingServices(systemHandle, dictionary.Handle, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
            serviceHandle = services[0];
            IOHIDServiceSetProperty(serviceHandle, new CFString("ReportInterval"), new NSNumber(500).Handle);
            IOHIDEventSystemOpen(systemHandle, OnSensorChanged, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);
        }

        internal static void PlatformStop()
        {
            var defaultInterval = new NSNumber(5428500);
            IOHIDServiceSetProperty((IOHIDServiceRef)serviceHandle, new CFString("ReportInterval"), defaultInterval.Handle);

            IOHIDEventSystemClose(systemHandle, NULL);
            CFRelease(systemHandle);
            systemHandle = IntPtr.Zero;
        }
    }
}
