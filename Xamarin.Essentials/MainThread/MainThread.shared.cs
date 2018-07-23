using System;
using System.Threading.Tasks;

namespace Xamarin.Essentials
{
    public static partial class MainThread
    {
        public static bool IsMainThread =>
            PlatformIsMainThread;

        public static void BeginInvokeOnMainThread(Action action) =>
            PlatformBeginInvokeOnMainThread(action);
    }
}
