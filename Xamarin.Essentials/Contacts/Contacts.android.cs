using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Xamarin.Essentials
{
    public static partial class Contacts
    {
        static void PlatformContacts()
        {
            Permissions.EnsureDeclared(PermissionType.Contact); 
        }

        static async Task<Contact> PlatformPickAsync()
        {
            PermissionStatus status;
            if (!MainThread.IsMainThread)
                status = await MainThread.InvokeOnMainThread(async () => await Permissions.RequestAsync(PermissionType.Contacts));
            status = await Permissions.RequestAsync(PermissionType.Contacts)

            if (status != PermissionStatus.Granted)
                throw new InvalidOperationException(nameof(Permissions));
        }
    }

    internal class ContactActivity : Activity
    {
        Func<Contact> SuccessCallback;
        Action FailureCallback;

        internal ContactActivity(Func<Contact> successCallback, Action failureCallback)
        {
            if (successCallback == null)
                throw new ArgumentNullException(nameof(successCallback));
            if (failureCallback == null)
                throw new ArgumentNullException(nameof(failureCallback));

            Callback = successCallback;
            FailureCallback = failureCallback;
        }

        protected override void OnCreate(Bundle savedInstanceState)
        {
            base.OnCreate(savedInstanceState);
        }

        protected override void OnActivityResult(int requestCode, [GeneratedEnum] Result resultCode, Content.Intent intent)
        {
            if (resultCode == Result.Cancelled)
            {
                FailureCallback.Invoke();
                Finish();
            }

            if (requestCode == Activity.PICK_CONTACT_REQUEST)
            {
                FailureCallback.Invoke();
                Finish();
            }

            var record = intent.Data;

            var contact = new Contact()
        }
    }
}
