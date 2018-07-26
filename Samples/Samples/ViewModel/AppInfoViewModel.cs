using System.Windows.Input;
using Xamarin.Essentials;
using Xamarin.Forms;

namespace Samples.ViewModel
{
    public class AppInfoViewModel : BaseViewModel
    {
        public string AppPackageName => AppInfo.PackageName;

        public string AppName => AppInfo.Name;

        public string AppVersion => AppInfo.VersionString;

        public string AppBuild => AppInfo.BuildString;

        public string Brightness => AppInfo.Brightness.Value.ToString();

        double newBrightness;

        public double NewBrightness
        {
            get => newBrightness;
            set
            {
                if (value > 1)
                    SetProperty(ref newBrightness, 1);
                else if (value < 0)
                    SetProperty(ref newBrightness, 0);
                else
                    SetProperty(ref newBrightness, value);
            }
        }

        public string BrightnessString
        {
            get => NewBrightness.ToString();
            set => NewBrightness = double.Parse(string.IsNullOrEmpty(value) ? "0" : value);
        }

        public Command OpenSettingsCommand { get; }

        public ICommand SetBrightnessCommand { get; }

        public AppInfoViewModel()
        {
            OpenSettingsCommand = new Command(() => AppInfo.OpenSettings());
            SetBrightnessCommand = new Command(() => AppInfo.SetBrightness(NewBrightness));
            NewBrightness = AppInfo.Brightness.Value;
        }
    }
}
