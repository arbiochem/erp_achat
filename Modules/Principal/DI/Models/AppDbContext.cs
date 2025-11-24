using arbioApp.Models;
using arbioApp.Modules.Principal.DI.Models;
using arbioApp.Utils.Connection;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity;

namespace arbioApp.Models
{
    public partial class AppDbContext : DbContext
    {
        private static string connectionString = "";
        public AppDbContext(): base(Db.GetConnectionString()) {
            Database.SetInitializer<AppDbContext>(null);
        }
        
        //    FANAOVANA UPDATE DATABASE
       /* public AppDbContext()
    :   base("Data Source=SRV-ARB;Initial Catalog=ARBIOCHEM_ACHAT;Persist Security Info=True;User ID=DEV;Password=1234;TrustServerCertificate=True")
        {
        }*/


        public DbSet<F_FRET> F_FRETS { get; set; }
        public DbSet<F_DEVISE> F_DEVISES { get; set; }
        public virtual DbSet<P_PREFERENCES> P_PREFERENCES { get; set; }
        public virtual DbSet<F_DOCREGL> F_DOCREGL { get; set; }
        public virtual DbSet<F_CATALOGUE> F_CATALOGUE { get; set; }
        public virtual DbSet<F_LIGNEARCHIVE> F_LIGNEARCHIVE { get; set; }
        public virtual DbSet<F_REGLEARCHIVE> F_REGLEARCHIVE { get; set; }
        public virtual DbSet<P_ANALYTIQUE> P_ANALYTIQUE { get; set; }
        public virtual DbSet<F_DOCCURRENTPIECE> F_DOCCURRENTPIECE { get; set; }
        public virtual DbSet<F_FAMILLE> F_FAMILLE { get; set; }
        public virtual DbSet<F_COMPTET> F_COMPTET { get; set; }
        public virtual DbSet<F_COMPTEG> F_COMPTEG { get; set; }
        public virtual DbSet<F_COMPTEA> F_COMPTEA { get; set; }
        public virtual DbSet<F_CAISSE> F_CAISSE { get; set; }
        public virtual DbSet<F_DEPOT> F_DEPOT { get; set; }
        public virtual DbSet<F_BILLETPIECE> F_BILLETPIECE { get; set; }
        public virtual DbSet<F_TICKETARCHIVE> F_TICKETARCHIVE { get; set; }
        public virtual DbSet<F_DOCENTETE> F_DOCENTETE { get; set; }
        public virtual DbSet<F_DOCLIGNE> F_DOCLIGNE { get; set; }
        public virtual DbSet<F_JOURNAUX> F_JOURNAUX { get; set; }
        public virtual DbSet<F_COLLABORATEUR> F_COLLABORATEUR { get; set; }
        public virtual DbSet<F_ARTICLE> F_ARTICLE { get; set; }
        public virtual DbSet<P_DEVISE> P_DEVISE { get; set; }
        public virtual DbSet<P_SOUCHEVENTE> P_SOUCHEVENTE { get; set; }
        public virtual DbSet<F_CREGLEMENT> F_CREGLEMENT { get; set; }
        public virtual DbSet<F_REGLECH> F_REGLECH { get; set; }
        public virtual DbSet<F_TAXE> F_TAXE { get; set; }
        public virtual DbSet<F_ARTCOMPTA> F_ARTCOMPTA { get; set; }
        public virtual DbSet<P_REGLEMENT> P_REGLEMENT { get; set; }
        public virtual DbSet<P_UNITE> P_UNITE { get; set; }
        public virtual DbSet<F_ARTCLIENT> F_ARTCLIENT { get; set; }
        public virtual DbSet<P_CATTARIF> P_CATTARIF { get; set; }
        public virtual DbSet<F_PAYS> F_PAYS { get; set; }
        public virtual DbSet<F_GLOSSAIRE> F_GLOSSAIRE { get; set; }
        public virtual DbSet<F_ARTGLOSS> F_ARTGLOSS { get; set; }
        public virtual DbSet<F_ARTICLEMEDIA> F_ARTICLEMEDIA { get; set; }
        public virtual DbSet<F_MODELE> F_MODELE { get; set; }
        public virtual DbSet<F_ARTMODELE> F_ARTMODELE { get; set; }
        public virtual DbSet<F_DEPOTEMPL> F_DEPOTEMPL { get; set; }
        public virtual DbSet<F_CONDITION> F_CONDITION { get; set; }
        public virtual DbSet<P_GAMME> P_GAMME { get; set; }
        public virtual DbSet<F_ARTGAMME> F_ARTGAMME { get; set; }
        public virtual DbSet<F_ARTENUMREF> F_ARTENUMREF { get; set; }
        public virtual DbSet<F_ENUMGAMME> F_ENUMGAMME { get; set; }
        public virtual DbSet<F_ARTSTOCK> F_ARTSTOCK { get; set; }
        public virtual DbSet<F_ARTSTOCKEMPL> F_ARTSTOCKEMPL { get; set; }
        public virtual DbSet<F_GAMSTOCK> F_GAMSTOCK { get; set; }
        public virtual DbSet<F_GAMSTOCKEMPL> F_GAMSTOCKEMPL { get; set; }
        public virtual DbSet<P_PARAMETRECIAL> P_PARAMETRECIAL { get; set; }
        public virtual DbSet<P_EXPEDITION> P_EXPEDITION { get; set; }
        public virtual DbSet<F_LIVRAISON> F_LIVRAISON { get; set; }
        public virtual DbSet<P_CONDLIVR> P_CONDLIVR { get; set; }
        public virtual DbSet<P_CATCOMPTA> P_CATCOMPTA { get; set; }
        public virtual DbSet<F_DOCLIGNEEMPL> F_DOCLIGNEEMPL { get; set; }
        public virtual DbSet<F_REGLEMENTT> F_REGLEMENTT { get; set; }
        public virtual DbSet<F_ARTFOURNISS> F_ARTFOURNISS { get; set; }
        public virtual DbSet<P_COLREGLEMENT> P_COLREGLEMENT { get; set; }
        public virtual DbSet<F_AGENDA> F_AGENDA { get; set; }
        public virtual DbSet<F_ARTPRIX> F_ARTPRIX { get; set; }
        public virtual DbSet<F_DOCFRAISIMPORT> F_DOCFRAISIMPORT { get; set; }
        public virtual DbSet<F_DOCCOUTREVIENT> F_DOCCOUTREVIENT { get; set; }
        public virtual DbSet<P_TYPEFRAIS> P_TYPEFRAIS { get; set; }
        public virtual DbSet<P_TYPEREPARTITION> P_TYPEREPARTITION { get; set; }

        public virtual DbSet<P_CRISQUE> P_CRISQUE { get; set; }
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<F_DOCFRAISIMPORT>()
                .HasRequired(f => f.TypeFrais) // Replace HasOne with HasRequired for EF6 compatibility
                .WithMany(t => t.FraisImports)
                .HasForeignKey(f => f.FI_TypeFraisId)
                .WillCascadeOnDelete(false); // Adjust delete behavior for EF6

            modelBuilder.Entity<F_DOCFRAISIMPORT>()
                .HasRequired(f => f.Repartition) // Replace HasOne with HasRequired for EF6 compatibility
                .WithMany(r => r.FraisImports)
                .HasForeignKey(f => f.FI_RepartitionId)
                .WillCascadeOnDelete(false); // Adjust delete behavior for EF6
            modelBuilder.Entity<F_ARTICLE>()
            .HasOptional(a => a.artStock)
            .WithRequired(s => s.FArticle);

            modelBuilder.Entity<F_ARTSTOCK>()
            .HasKey(e => e.cbMarq);

            modelBuilder.Entity<F_ARTSTOCK>()
                .Property(e => e.cbMarq)
                .HasDatabaseGeneratedOption(DatabaseGeneratedOption.Identity);


        }
        //public string GetConnectionInfo()
        //{
        //    var conn = this.Database.Connection; // EF6
        //    return $"Base = {conn.Database}, Serveur = {conn.DataSource}";
        //}



    }
}
