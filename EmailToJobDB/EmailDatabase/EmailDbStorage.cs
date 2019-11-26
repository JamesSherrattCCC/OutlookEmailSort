using EmailHandler.DataTypes;
using EmailHandler;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Data.Entity;


namespace EmailToJobDB.EmailDatabase
{
    class EmailDbStorage : IEmailStorage
    {

        private EmailContext _dbContext;
        private string _user;

        public EmailDbStorage(string user, EmailContext ctx)
        {
            _user = user;
            _dbContext = ctx;
        }

        public async Task<IEnumerable<Email>> GetEmails()
        {
            return _dbContext.Emails.Where(email => email.User == _user).ToList();
        }

        public async Task<Email> GetLastRetrievedEmail()
        {
            var maxDate = (from Date in _dbContext.Set<Email>()
                        where Date.User == _user
                        group Date by 1 into Dateg
                        select new
                        {
                            MaxDate = Dateg.Max(email => email.DateRetrieved)
                        }).FirstOrDefault();

            return (from Mail in _dbContext.Set<Email>()
                                 where Mail.DateRetrieved == maxDate.MaxDate
                                 select Mail).FirstOrDefault();

        }

        public async Task SaveEmails(IEnumerable<Email> emails)
        {
            foreach (Email email in emails)
            {
                if (! await _dbContext.Emails.AnyAsync(e => e.Id == email.Id))
                {
                    email.User = _user;
                    _dbContext.Emails.Add(email);
                }
            }
            await _dbContext.SaveChangesAsync();
        }
    }
}
