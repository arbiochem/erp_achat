namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_LIGNEARCHIVE
    {
        [Required]
        [StringLength(13)]
        public string TA_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbTA_Piece { get; set; }

        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        [StringLength(69)]
        public string LA_Design { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_PrixUnitaire { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_PUTTC { get; set; }

        public short? LA_TTC { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_Qte { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_Remise01REM_Valeur { get; set; }

        public short? LA_Remise01REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_Remise02REM_Valeur { get; set; }

        public short? LA_Remise02REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_Remise03REM_Valeur { get; set; }

        public short? LA_Remise03REM_Type { get; set; }

        public int? LA_Ligne { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_Taxe1 { get; set; }

        public short? LA_TypeTaux1 { get; set; }

        public short? LA_TypeTaxe1 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_Taxe2 { get; set; }

        public short? LA_TypeTaux2 { get; set; }

        public short? LA_TypeTaxe2 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_Taxe3 { get; set; }

        public short? LA_TypeTaux3 { get; set; }

        public short? LA_TypeTaxe3 { get; set; }

        public int? AG_No1 { get; set; }

        public int? AG_No2 { get; set; }

        [StringLength(31)]
        public string LS_NoSerie { get; set; }

        public int? CO_No { get; set; }

        public int? cbCO_No { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_PoidsNet { get; set; }

        public short? LA_TRemExep { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_PrixRU { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? LA_CMUP { get; set; }

        [StringLength(35)]
        public string EU_Enumere { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? EU_Qte { get; set; }

        public short? LA_FactPoids { get; set; }

        public short? LA_Escompte { get; set; }

        public short? LA_Valorise { get; set; }

        [StringLength(31)]
        public string LS_Complement { get; set; }

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
    }
}
