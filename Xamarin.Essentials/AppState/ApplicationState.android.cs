using System;
using Android.App;
using Android.Content;
using Android.Content.Res;
using Android.OS;
using Android.Runtime;
using static Android.App.Application;

namespace Xamarin.Essentials
{
    public static partial class ApplicationState
    {
        static readonly Lazy<AppStateLifecycleListener> appState = new Lazy<AppStateLifecycleListener>(
            () => new AppStateLifecycleListener((targ) => UpdateStateCallback(targ)));

        static AppState PlatformState { get; set; }

        static void StartStateListeners()
        {
            var app = Application.Context.ApplicationContext as Application;
            app.RegisterActivityLifecycleCallbacks(appState.Value);
            app.RegisterComponentCallbacks(appState.Value);
        }

        static void StopStateListeners()
        {
            var app = Application.Context.ApplicationContext as Application;
            app.UnregisterActivityLifecycleCallbacks(appState.Value);
            app.UnregisterComponentCallbacks(appState.Value);
        }

        static void UpdateStateCallback(AppState state)
        {
            PlatformState = state;
        }
    }

    public class AppStateLifecycleListener : Java.Lang.Object, IActivityLifecycleCallbacks, IComponentCallbacks2
    {
        readonly Action<AppState> callback;

        public AppStateLifecycleListener(Action<AppState> callback)
        {
            this.callback = callback;
        }

        public void OnActivityResumed(Activity activity)
        {
            callback(AppState.Foreground);
            ApplicationState.OnStateChanged(AppState.Foreground);
        }

        public void OnTrimMemory([GeneratedEnum] TrimMemory level)
        {
            if (level == TrimMemory.UiHidden)
            {
                callback(AppState.Background);
                ApplicationState.OnStateChanged(AppState.Background);
            }
        }

        // Unused from here on
        public void OnActivityCreated(Activity activity, Bundle savedInstanceState)
        {
        }

        public void OnActivityDestroyed(Activity activity)
        {
        }

        public void OnActivityPaused(Activity activity)
        {
        }

        public void OnActivitySaveInstanceState(Activity activity, Bundle outState)
        {
        }

        public void OnActivityStarted(Activity activity)
        {
        }

        public void OnActivityStopped(Activity activity)
        {
        }

        public void OnConfigurationChanged(Configuration newConfig)
        {
        }

        public void OnLowMemory()
        {
        }
    }
}
