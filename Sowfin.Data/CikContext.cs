using Sowfin.Model;
using Microsoft.EntityFrameworkCore;
using System.Linq;

namespace Sowfin.Data
{
    public class CikContext : DbContext
    {
        private string? ConnectionString { get; set; }

        public CikContext()
        {

        }

        public CikContext(DbContextOptions<CikContext> options) : base(options)
        {
            Database.EnsureCreated();
        }

        public CikContext(string ConnectionString)
        {
            this.ConnectionString = ConnectionString;
        }

        public DbSet<Findata> FinData { get; set; }


        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
           optionsBuilder.UseNpgsql(ConnectionString);
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            foreach (var relationship in modelBuilder.Model.GetEntityTypes().SelectMany(e => e.GetForeignKeys()))
            {
                relationship.DeleteBehavior = DeleteBehavior.Restrict;
            }

            ConfigureModelBuilderForFindata(modelBuilder);

        }

        void ConfigureModelBuilderForFindata(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Findata>().ToTable("findata");

        }

    }
}
