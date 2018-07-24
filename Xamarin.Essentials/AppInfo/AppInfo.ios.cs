﻿using System;
using Foundation;
using UIKit;

namespace Xamarin.Essentials
{
    public static partial class AppInfo
    {
        static string PlatformGetPackageName() => GetBundleValue("CFBundleIdentifier");

        static string PlatformGetName() => GetBundleValue("CFBundleDisplayName") ?? GetBundleValue("CFBundleName");

        static string PlatformGetVersionString() => GetBundleValue("CFBundleShortVersionString");

        static string PlatformGetBuild() => GetBundleValue("CFBundleVersion");

        static string GetBundleValue(string key)
           => NSBundle.MainBundle.ObjectForInfoDictionary(key).ToString();

        static void PlatformOpenSettings() =>
            UIApplication.SharedApplication.OpenUrl(new NSUrl(UIApplication.OpenSettingsUrlString));

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
