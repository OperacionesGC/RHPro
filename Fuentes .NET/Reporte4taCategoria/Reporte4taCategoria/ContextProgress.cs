using System;
using RHPro.Repository.Providers;
using System.Data.Entity;
using RHPro.Shared.Data;
using RHPro.Shared.Data.Mapping;
using RHPro.Shared.Interface.Data;
using RHPro.Shared.Interface.Data.Mapping;
using RHPro.Shared.Ganancias.Data;
using RHPro.ADP.Data.Document;
using RHPro.ADP.Data.Document.Mapping;

namespace Reporte4taCategoria
{
    public partial class ContextProgress : DataContext, IDataContext
    {
        public string stringConnection;
        public ContextProgress(string nameOrConnectionString) : base(nameOrConnectionString)
        {
            stringConnection = nameOrConnectionString;
            Database.SetInitializer<ContextReporte4taCategoria>(null);
        }

        public DbSet<sistema> sistema { get; set; }
        public DbSet<batch_proceso> batch_Procesos { get; set; }
        public DbSet<batch_empleado> batch_empleado { get; set; }
       

        // Sobreescibimos OnModelCreating de DdContext
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            this.Configuration.LazyLoadingEnabled = true;
            this.Configuration.AutoDetectChangesEnabled = false;

            modelBuilder.Configurations.Add(new sistemaMap(""));
            modelBuilder.Configurations.Add(new batch_empleadoMap(""));
            modelBuilder.Configurations.Add(new batch_procesoMap(""));

        }
    }
}
