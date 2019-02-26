﻿using System;
using System.Threading.Tasks;
using Foundation;
using UIKit;

namespace Xamarin.Essentials
{
    public static partial class Clipboard
    {
        static Task PlatformSetTextAsync(string text)
        {
            UIPasteboard.General.String = text;
            return Task.CompletedTask;
        }

        static NSObject _observer;

        static bool PlatformHasText
            => UIPasteboard.General.HasStrings;

        static Task<string> PlatformGetTextAsync()
            => Task.FromResult(UIPasteboard.General.String);

        static void StartClipboardListeners()
        {
            _observer = NSNotificationCenter.DefaultCenter.AddObserver(
                UIPasteboard.ChangedNotification,
                ClipboardChangedObserver);
        }

        static void StopClipboardListeners()
            => NSNotificationCenter.DefaultCenter.RemoveObserver(_observer);

        static void ClipboardChangedObserver(NSNotification notification)
            => ClipboardChangedInternal();
    }
}
