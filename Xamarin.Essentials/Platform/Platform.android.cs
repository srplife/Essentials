using System;
using System.IO;
using System.Linq;
using Android.App;
using Android.Content;
using Android.Content.PM;
using Android.Hardware;
using Android.Hardware.Camera2;
using Android.Locations;
using Android.Net;
using Android.Net.Wifi;
using Android.OS;
using Android.Support.V4.Content;
using Java.Util;
using AndroidEnvironment = Android.OS.Environment;
using AndroidUri = Android.Net.Uri;

namespace Xamarin.Essentials
{
    public static partial class Platform
    {
        static readonly string fileprovider = "xamarin_essentials_fileprovider_name";

        static ActivityLifecycleContextListener lifecycleListener;

        internal static Context AppContext =>
            Application.Context;

        internal static Activity GetCurrentActivity(bool throwOnNull)
        {
            var activity = lifecycleListener?.Activity;
            if (throwOnNull && activity == null)
                throw new NullReferenceException("The current Activity can not be detected. Ensure that you have called Init in your Activity or Application class.");

            return activity;
        }

        internal static AndroidUri GetUriForFile(string filename)
        {
            var stringId = AppContext.Resources.GetIdentifier(fileprovider, "string", AppContext.PackageName);

            // if (stringId <= 0)
            //     throw new MissingResourceException("Missing fileprovider registered in client app", "Platform", fileprovider);
            var fileProviderName = AppContext.Resources.GetString(stringId);
            var dir = AndroidEnvironment.GetExternalStoragePublicDirectory(AndroidEnvironment.DirectoryDownloads);
            var path = Path.Combine(dir.AbsolutePath, "SamplePNGImage_100kbmb.png");
            var uri = FileProvider.GetUriForFile(AppContext, AppContext.PackageName, new Java.IO.File(path));
            return uri;
        }

        public static void Init(Application application)
        {
            lifecycleListener = new ActivityLifecycleContextListener();
            application.RegisterActivityLifecycleCallbacks(lifecycleListener);
        }

        public static void Init(Activity activity, Bundle bundle)
        {
            Init(activity.Application);
            lifecycleListener.Activity = activity;
        }

        public static void OnRequestPermissionsResult(int requestCode, string[] permissions, Permission[] grantResults) =>
            Permissions.OnRequestPermissionsResult(requestCode, permissions, grantResults);

        internal static bool HasSystemFeature(string systemFeature)
        {
            var packageManager = AppContext.PackageManager;
            foreach (var feature in packageManager.GetSystemAvailableFeatures())
            {
                if (feature.Name.Equals(systemFeature, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }

        internal static bool IsIntentSupported(Intent intent)
        {
            var manager = AppContext.PackageManager;
            var activities = manager.QueryIntentActivities(intent, PackageInfoFlags.MatchDefaultOnly);
            return activities.Any();
        }

        internal static bool HasApiLevel(BuildVersionCodes versionCode) =>
            (int)Build.VERSION.SdkInt >= (int)versionCode;

        internal static CameraManager CameraManager =>
            AppContext.GetSystemService(Context.CameraService) as CameraManager;

        internal static ConnectivityManager ConnectivityManager =>
            AppContext.GetSystemService(Context.ConnectivityService) as ConnectivityManager;

        internal static Vibrator Vibrator =>
            AppContext.GetSystemService(Context.VibratorService) as Vibrator;

        internal static WifiManager WifiManager =>
            AppContext.GetSystemService(Context.WifiService) as WifiManager;

        internal static SensorManager SensorManager =>
            AppContext.GetSystemService(Context.SensorService) as SensorManager;

        internal static ClipboardManager ClipboardManager =>
            AppContext.GetSystemService(Context.ClipboardService) as ClipboardManager;

        internal static LocationManager LocationManager =>
            AppContext.GetSystemService(Context.LocationService) as LocationManager;

        internal static PowerManager PowerManager =>
            AppContext.GetSystemService(Context.PowerService) as PowerManager;
    }

    class ActivityLifecycleContextListener : Java.Lang.Object, Application.IActivityLifecycleCallbacks
    {
        WeakReference<Activity> currentActivity = new WeakReference<Activity>(null);

        internal Context Context =>
            Activity ?? Application.Context;

        internal Activity Activity
        {
            get => currentActivity.TryGetTarget(out var a) ? a : null;
            set => currentActivity.SetTarget(value);
        }

        void Application.IActivityLifecycleCallbacks.OnActivityCreated(Activity activity, Bundle savedInstanceState) =>
            Activity = activity;

        void Application.IActivityLifecycleCallbacks.OnActivityDestroyed(Activity activity)
        {
        }

        void Application.IActivityLifecycleCallbacks.OnActivityPaused(Activity activity) =>
            Activity = activity;

        void Application.IActivityLifecycleCallbacks.OnActivityResumed(Activity activity) =>
            Activity = activity;

        void Application.IActivityLifecycleCallbacks.OnActivitySaveInstanceState(Activity activity, Bundle outState)
        {
        }

        void Application.IActivityLifecycleCallbacks.OnActivityStarted(Activity activity)
        {
        }

        void Application.IActivityLifecycleCallbacks.OnActivityStopped(Activity activity)
        {
        }
    }
}
