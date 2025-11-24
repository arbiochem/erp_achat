using System;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace arbioApp.Models
{
    

    public partial class F_DOCENTETE
    {
        public short? DO_Domaine { get; set; }
        //public string DO_Intitule { get; set; }
        public short? DO_Type { get; set; }

        [StringLength(13)]
        public string DO_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbDO_Piece { get; set; }

        public DateTime? DO_Date { get; set; }

        [StringLength(17)]
        public string DO_Ref { get; set; }

        [StringLength(17)]
        public string DO_Tiers { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbDO_Tiers { get; set; }

        public int? CO_No { get; set; }

        public int? cbCO_No { get; set; }

        public short? DO_Period { get; set; }

        public short? DO_Devise { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_Cours { get; set; }

        public int? DE_No { get; set; }

        public int? cbDE_No { get; set; }

        public int? LI_No { get; set; }

        public int? cbLI_No { get; set; }

        [StringLength(17)]
        public string CT_NumPayeur { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_NumPayeur { get; set; }

        public short? DO_Expedit { get; set; }

        public short? DO_NbFacture { get; set; }

        public short? DO_BLFact { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_TxEscompte { get; set; }

        public short? DO_Reliquat { get; set; }

        public short? DO_Imprim { get; set; }

        [StringLength(13)]
        public string CA_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCA_Num { get; set; }

        [StringLength(25)]
        public string DO_Coord01 { get; set; }

        [StringLength(25)]
        public string DO_Coord02 { get; set; }

        [StringLength(25)]
        public string DO_Coord03 { get; set; }

        [StringLength(25)]
        public string DO_Coord04 { get; set; }

        public short? DO_Souche { get; set; }

        public DateTime? DO_DateLivr { get; set; }

        public short? DO_Condition { get; set; }

        public short? DO_Tarif { get; set; }

        public short? DO_Colisage { get; set; }

        public short? DO_TypeColis { get; set; }

        public short? DO_Transaction { get; set; }

        public short? DO_Langue { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_Ecart { get; set; }

        public short? DO_Regime { get; set; }

        public short? N_CatCompta { get; set; }

        public short? DO_Ventile { get; set; }

        public int? AB_No { get; set; }

        public DateTime? DO_DebutAbo { get; set; }

        public DateTime? DO_FinAbo { get; set; }

        public DateTime? DO_DebutPeriod { get; set; }

        public DateTime? DO_FinPeriod { get; set; }

        [StringLength(13)]
        public string CG_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCG_Num { get; set; }

        public short? DO_Statut { get; set; }

        [StringLength(9)]
        public string DO_Heure { get; set; }

        public int? CA_No { get; set; }

        public int? cbCA_No { get; set; }

        public int? CO_NoCaissier { get; set; }

        public int? cbCO_NoCaissier { get; set; }

        public short? DO_Transfere { get; set; }

        public short? DO_Cloture { get; set; }

        [StringLength(17)]
        public string DO_NoWeb { get; set; }

        public short? DO_Attente { get; set; }

        public short? DO_Provenance { get; set; }

        [StringLength(13)]
        public string CA_NumIFRS { get; set; }

        public int? MR_No { get; set; }

        public short? DO_TypeFrais { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_ValFrais { get; set; }

        public short? DO_TypeLigneFrais { get; set; }

        public short? DO_TypeFranco { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_ValFranco { get; set; }

        public short? DO_TypeLigneFranco { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_Taxe1 { get; set; }

        public short? DO_TypeTaux1 { get; set; }

        public short? DO_TypeTaxe1 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_Taxe2 { get; set; }

        public short? DO_TypeTaux2 { get; set; }

        public short? DO_TypeTaxe2 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_Taxe3 { get; set; }

        public short? DO_TypeTaux3 { get; set; }

        public short? DO_TypeTaxe3 { get; set; }

        public short? DO_MajCpta { get; set; }

        [StringLength(69)]
        public string DO_Motif { get; set; }

        [StringLength(17)]
        public string CT_NumCentrale { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_NumCentrale { get; set; }

        [StringLength(35)]
        public string DO_Contact { get; set; }

        public short? DO_FactureElec { get; set; }

        public short? DO_TypeTransac { get; set; }

        public DateTime? DO_DateLivrRealisee { get; set; }

        public DateTime? DO_DateExpedition { get; set; }

        [StringLength(35)]
        public string DO_FactureFrs { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbDO_FactureFrs { get; set; }

        [StringLength(13)]
        public string DO_PieceOrig { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbDO_PieceOrig { get; set; }

        public Guid? DO_GUID { get; set; }

        public short? DO_EStatut { get; set; }

        public short? DO_DemandeRegul { get; set; }

        public int? ET_No { get; set; }

        public int? cbET_No { get; set; }

        public short? DO_Valide { get; set; }

        public short? DO_Coffre { get; set; }

        [StringLength(5)]
        public string DO_CodeTaxe1 { get; set; }

        [StringLength(5)]
        public string DO_CodeTaxe2 { get; set; }

        [StringLength(5)]
        public string DO_CodeTaxe3 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_TotalHT { get; set; }

        public short? DO_StatutBAP { get; set; }

        public short? DO_Escompte { get; set; }

        public short? DO_DocType { get; set; }

        public short? DO_TypeCalcul { get; set; }

        public Guid? DO_FactureFile { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_TotalHTNet { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_TotalTTC { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_NetAPayer { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DO_MontantRegle { get; set; }

        public Guid? DO_RefPaiement { get; set; }

        [StringLength(255)]
        public string DO_AdressePaiement { get; set; }

        public short? DO_PaiementLigne { get; set; }

        public short? DO_MotifDevis { get; set; }

        public short? cbProt { get; set; }

        [Key]
        public int cbMarq { get; set; }

        [StringLength(4)]
        public string cbCreateur { get; set; }

        public DateTime? cbModification { get; set; }

        public int? cbReplication { get; set; }

        public short? cbFlag { get; set; }

        public DateTime? cbCreation { get; set; }

        public Guid? cbCreationUser { get; set; }

        [MaxLength(32)]
        public byte[] cbHash { get; set; }

        public short? cbHashVersion { get; set; }

        public DateTime? cbHashDate { get; set; }

        public int? cbHashOrder { get; set; }

        //[StringLength(69)]
        //public string Commentaires { get; set; }

        //[StringLength(69)]
        //public string Divers { get; set; }

        //public short? DO_Conversion { get; set; }
    }
}
