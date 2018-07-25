using System.Linq;
using System.Threading.Tasks;
using Android.Content;
using Android.Net;
using Android.OS;
using Android.Text;
using Java.Lang;

namespace Xamarin.Essentials
{
    public static partial class Email
    {
        static readonly EmailMessage testEmail =
            new EmailMessage("Testing Xamarin.Essentials", "This is a test email.", "Xamarin.Essentials@example.org");

        internal static bool IsComposeSupported
            => Platform.IsIntentSupported(CreateIntent(testEmail));

        static Task PlatformComposeAsync(EmailMessage message)
        {
            Permissions.RequestAsync(PermissionType.)
            var intent = CreateIntent(message)
                .SetFlags(ActivityFlags.ClearTop)
                .SetFlags(ActivityFlags.NewTask);

            Platform.AppContext.StartActivity(intent);
            return Task.CompletedTask;
        }

        static Intent CreateIntent(EmailMessage message)
        {
            var intent = new Intent(Intent.ActionSend);
            intent.SetType("message/rfc822");

            if (!string.IsNullOrEmpty(message?.Body))
            {
                if (message?.BodyFormat == EmailBodyFormat.Html)
                {
                    ISpanned html;
                    if (Platform.HasApiLevel(BuildVersionCodes.N))
                    {
                        html = Html.FromHtml(message.Body, FromHtmlOptions.ModeLegacy);
                    }
                    else
                    {
#pragma warning disable CS0618 // Type or member is obsolete
                        html = Html.FromHtml(message.Body);
#pragma warning restore CS0618 // Type or member is obsolete
                    }
                    intent.PutExtra(Intent.ExtraText, html);
                }
                else
                {
                    intent.PutExtra(Intent.ExtraText, message.Body);
                }
            }
            if (!string.IsNullOrEmpty(message?.Subject))
                intent.PutExtra(Intent.ExtraSubject, message.Subject);
            if (message.To?.Count > 0)
                intent.PutExtra(Intent.ExtraEmail, message.To.ToArray());
            if (message.Cc?.Count > 0)
                intent.PutExtra(Intent.ExtraCc, message.Cc.ToArray());
            if (message.Bcc?.Count > 0)
                intent.PutExtra(Intent.ExtraBcc, message.Bcc.ToArray());

            var attachmentUris = message.Attachments.Select(
                x => Uri.Parse($"file://{x.Filepath}")).ToArray();
            if (attachmentUris.Length > 0)
            {
                intent.AddFlags(ActivityFlags.GrantReadUriPermission);
                if (attachmentUris.Length > 1)
                    intent.PutParcelableArrayListExtra(Intent.ExtraStream, attachmentUris);
                else
                    intent.PutExtra(Intent.ExtraStream, attachmentUris.Single());
            }

            return intent;
        }
    }
}
