using Android.App;
using Android.Content;
using Android.Support.V4.Content;

namespace Xamarin.Essentials
{
    [ContentProvider(new[] { "${applicationId}.fileProvider" }, Name = "android.support.v4.content.FileProvider", Exported = true, GrantUriPermissions = true)]
    [MetaData("android.support.FILE_PROVIDER_PATHS", Resource = "@xml/essentials_file_paths")]
    class EssentialsFileProvider : FileProvider
    {
    }
}
