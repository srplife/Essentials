using Windows.UI.Core;
using Windows.UI.Xaml;

namespace Xamarin.Essentials
{
    public static partial class ApplicationState
    {
        static AppState PlatformState => Window.Current.Visible ? AppState.Foreground : AppState.Background;

        static void StartStateListeners() => Window.Current.VisibilityChanged += VisibilityChanged;

        static void VisibilityChanged(object sender, VisibilityChangedEventArgs e)
        {
            var state = e.Visible ? AppState.Foreground : AppState.Background;
            MainThread.BeginInvokeOnMainThread(() => OnStateChanged(state));
            e.Handled = true;
        }

        static void StopStateListeners() => Window.Current.VisibilityChanged -= VisibilityChanged;
    }
}
