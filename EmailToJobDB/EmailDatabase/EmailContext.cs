using EmailHandler.DataTypes;
using System.Data.Entity;

namespace EmailToJobDB.EmailDatabase
{
    class EmailContext : DbContext
    {
        public DbSet<Email> Emails { get; set; }
        public DbSet<Job> Jobs { get; set; }


        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Job>().HasKey(j => j.Id);
        }
    }
}
