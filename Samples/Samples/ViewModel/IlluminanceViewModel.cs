using Xamarin.Essentials;

namespace Samples.ViewModel
{
    public class IlluminanceViewModel : BaseViewModel
    {
        public string Intensity { get; set; }

        public override void OnAppearing()
        {
            base.OnAppearing();
            Photometer.IntensityChanged += LightIntensityChanged;
        }

        void LightIntensityChanged(object sender, LightChangedEventArgs e)
        {
            Intensity = $"Light intensity: {e.Intensity}lx";
        }

        public override void OnDisappearing()
        {
            base.OnDisappearing();
            Photometer.IntensityChanged -= LightIntensityChanged;
        }
    }
}
