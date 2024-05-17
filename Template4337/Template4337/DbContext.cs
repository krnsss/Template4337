using System.Data.Entity;

namespace Template4337
{
    public class ServiceContext : DbContext
    {
        private const string ConnectionString = "Data Source=(localdb)\\mssqllocaldb;Initial Catalog=Isrpo3;Integrated Security=True;";

        public ServiceContext() : base(ConnectionString)
        {
        }

        public DbSet<Service> Services { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            base.OnModelCreating(modelBuilder);

            modelBuilder.Entity<Service>()
                .Property(s => s.Name)
                .IsRequired()
                .HasMaxLength(100);

            modelBuilder.Entity<Service>()
                .Property(s => s.Type)
                .IsRequired()
                .HasMaxLength(50);

            modelBuilder.Entity<Service>()
                .Property(s => s.Code)
                .IsRequired()
                .HasMaxLength(10);
        }
    }

    public class Service
    {
        public int Id { get; set; }

        public string Name { get; set; }

        public string Type { get; set; }

        public string Code { get; set; }

        public int Price { get; set; }

        public Service()
        {
        }
    }
}