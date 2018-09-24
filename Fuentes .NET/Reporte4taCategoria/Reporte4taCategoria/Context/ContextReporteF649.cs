using System;
using RHPro.Repository.Providers;
using System.Data.Entity;
using RHPro.Shared.Data;
using RHPro.Shared.Data.Mapping;
using RHPro.Documents.Data;
using RHPro.Documents.Data.Mapping;

namespace Reporte_F649.Context
{
    public partial class ContextReporteF649 : DataContext, IDataContext
    {
        public string stringConnection;
        public ContextReporteF649(string nameOrConnectionString) : base(nameOrConnectionString)
        {
            stringConnection = nameOrConnectionString;
            Database.SetInitializer<ContextReporteF649>(null);
        }

        public DbSet<sistema> sistema { get; set; }
        public DbSet<batch_proceso> batch_Procesos { get; set; }
        public DbSet<batch_empleado> batch_empleado { get; set; }
        public DbSet<document> document { get; set; }
        public DbSet<tercero> tercero { get; set; }
        public DbSet<empleado> empleado { get; set; }
        public DbSet<estructura> estructura { get; set; }
        public DbSet<his_estructura> his_estructura { get; set; }
        public DbSet<documentTag> documentTag { get; set; }
        public DbSet<documentEstr> documentEstr  { get; set; }
        public DbSet<documentGrp> documentGrp { get; set; }
        public DbSet<documentGrpDet> documentGrpDet { get; set; }
        public DbSet<documentHist> documentHist { get; set; }
        public DbSet<documentState> documentState { get; set; }
        public DbSet<documentOrigin> documentOrigin { get; set; }
        public DbSet<documentType> documentType { get; set; }
        //public DbSet<empresa> empresa { get; set; }




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
            modelBuilder.Configurations.Add(new documentTagMap(""));
            modelBuilder.Configurations.Add(new batch_empleadoMap(""));
            //modelBuilder.Configurations.Add(new empresaMap(""));
            modelBuilder.Configurations.Add(new documentGrpMap(""));
            modelBuilder.Configurations.Add(new documentGrpDetMap("").HasKey(obj => new { obj.dgrpnro, obj.docnro }));

            modelBuilder.Configurations.Add(new documentEstrMap("").HasKey(obj => new {obj.docnro, obj.tenro, obj.estrnro }));
            modelBuilder.Configurations.Add(new documentHistMap(""));
            modelBuilder.Configurations.Add(new documentStateMap(""));
            modelBuilder.Configurations.Add(new documentOriginMap(""));
            modelBuilder.Configurations.Add(new documentTypeMap(""));
        }
    }
}
