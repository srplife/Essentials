using System;
using UIKit;

namespace Xamarin.Essentials
{
    public static partial class ApplicationState
    {
        static AppState PlatformState
            => UIApplication.SharedApplication.ApplicationState == UIApplicationState.Active
            ? AppState.Foreground : AppState.Background;

        static IDisposable foregroundObserver;
        static IDisposable backgroundObserver;

        static void StartStateListeners()
        {
            foregroundObserver = UIApplication.Notifications.ObserveWillEnterForeground(
                (t, u) => OnStateChanged(AppState.Foreground));
            backgroundObserver = UIApplication.Notifications.ObserveDidEnterBackground(
                (t, u) => OnStateChanged(AppState.Background));
        }

        static void StopStateListeners()
        {
            foregroundObserver?.Dispose();
            foregroundObserver = null;
            backgroundObserver?.Dispose();
            backgroundObserver = null;
        }
    }
}
