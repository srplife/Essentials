using Contacts;
using System.Linq;
namespace Xamarin.Essentials
{
    public static partial class Contacts
    {
        static readonly Lazy<NSString> _contactKeys = new Lazy<NSString>(
        () => new[] { CNContactKey.GivenName, CNContactKey.EmailAddresses, CNContactKey.PhoneNumbers, CNContactKey.Birthday, CNContactKey.PhoneNumbers, CNContactKey.Type, CNContactKey.PostalAddresses });

        static Task<Contact> PlatformPickAsync()
        {

        }

        static Task<IReadOnlyList<Contact>> PlatformGetContacts()
        {
            using (var store = new CNContactStore())
            {
                var containerId = store.DefaultContainerIdentifier;
                using (var predicate = CNContact.GetPredicateForContactsInContainer(containerId))
                {
                    var contactList = store.GetUnifiedContacts(predicate, _contactKeys, out
                             var error);
                    if(error != null)
                        throw new PlatformException
                    return contactList.Select(x => new Contact(x.GivenName, x.PhoneNumbers.Select(y => y.Value.ToString()), x.Emails, x.Birthday, x.PostalAddresses, FromPlatform(x.Type).AsReadOnly();
                }
            }
        }

        static ContactType FromPlatform(CNContactType type)
        {
            switch (type)
            {
                case CNContactType.Person:
                    return ContactType.Private;
                case CNContactType.Organization:
                    return ContactType.Company;
                default:
                    return ContactType.Unknow;
            }
        }
    }
}
