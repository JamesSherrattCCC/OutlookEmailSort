using EmailHandler.DataTypes;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace EmailToJobDB
{
    static class MailItemToEmail
    {
        public static Email ToEmail(this Outlook.MailItem mailItem)
        {
            return new Email
            {
                Body = mailItem.Body,
                Sender = mailItem.Sender.Address,
                Subject = mailItem.Subject,
                DateRetrieved = mailItem.ReceivedTime,
            };
        }
    }
}
