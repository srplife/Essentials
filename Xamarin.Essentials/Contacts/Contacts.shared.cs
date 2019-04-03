namespace Xamarin.Essentials
{
    public static partial class Contacts
    {
        static Contacts()
        {
            PlatformContacts(); 
        }

        public static Task<Contact> PickContactAsync()
        {
            return PlatformPickAsync();
        }
    }

    public enum ContactType
    {
        Default,
        Private,
        Company 
    }

    public class Contact
    {
        public string Name { get; }

        public IReadOnlyList<string> Numbers { get; }

        public IReadOnlyList<string> Emails { get; }

        public string Birthday { get; }

        public string Address { get; }

        public ContactType Type { get; }

        internal PhoneContact(string name, IReadOnlyList<string> numbers, IReadOnlyList<string> emails, string birthday, string address, ContactType type)
        {
            Name = name;
            Birthday = birthday;
            Numbers = numbers;
            Emails = emails;
            Address = address;
            Type = type;
        }
    }
}
