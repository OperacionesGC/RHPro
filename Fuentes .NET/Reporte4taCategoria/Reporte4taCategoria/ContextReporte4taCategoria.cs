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
    public partial class ContextReporte4taCategoria : DataContext, IDataContext
    {
        public string stringConnection;
        public ContextReporte4taCategoria(string nameOrConnectionString) : base(nameOrConnectionString)
        {
            stringConnection = nameOrConnectionString;
            Database.SetInitializer<ContextReporte4taCategoria>(null);
        }

        public DbSet<sistema> sistema { get; set; }
        public DbSet<batch_proceso> batch_Procesos { get; set; }
        public DbSet<batch_empleado> batch_empleado { get; set; }
        public DbSet<tercero> tercero { get; set; }
        public DbSet<empleado> empleado { get; set; }
        public DbSet<estructura> estructura { get; set; }
        public DbSet<his_estructura> his_estructura { get; set; }
        public DbSet<traza_gan_item_top> traza_gan_item_top { get; set; }
        public DbSet<traza_gan> traza_gan { get; set; }
        public DbSet<ficharet> ficharet { get; set; }
        public DbSet<empresa> empresa { get; set; }
        public DbSet<ter_doc> ter_doc { get; set; }
        public DbSet<Gan4taCab> Gan4taCab { get; set; }
        public DbSet<Gan4taDet> Gan4taDet { get; set; }
        public DbSet<confrep> confrep { get; set; }

        // Sobreescibimos OnModelCreating de DdContext
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            this.Configuration.LazyLoadingEnabled = true;
            this.Configuration.AutoDetectChangesEnabled = false;

            modelBuilder.Configurations.Add(new empleadoMap(""));
            modelBuilder.Configurations.Add(new estructuraMap(""));
            modelBuilder.Configurations.Add(new his_estructuraMap(""));
            modelBuilder.Configurations.Add(new terceroMap(""));
            modelBuilder.Configurations.Add(new sistemaMap(""));
            modelBuilder.Configurations.Add(new batch_empleadoMap(""));
            modelBuilder.Configurations.Add(new traza_gan_item_topMap("").HasKey(obj => new { obj.itenro, obj.ternro, obj.pronro }));
            modelBuilder.Configurations.Add(new traza_ganMap("").HasKey(obj => new {obj.ternro, obj.pronro }));
			modelBuilder.Configurations.Add(new ficharetMap("").HasKey(obj => new { obj.empleado, obj.pronro,obj.fecha  }));
			modelBuilder.Configurations.Add(new empresaMap(""));
            modelBuilder.Configurations.Add(new ter_docMap(""));
            modelBuilder.Configurations.Add(new confrepMap(""));
            modelBuilder.Configurations.Add(new Gan4taCabMap(""));
            modelBuilder.Configurations.Add(new Gan4taDetMap("").HasKey(obj => new { obj.Gan4tanro, obj.orden })); 
        }
    }
}
