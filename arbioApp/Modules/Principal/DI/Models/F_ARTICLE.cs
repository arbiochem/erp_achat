namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_ARTICLE
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        [StringLength(69)]
        public string AR_Design { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(70)]
        public byte[] cbAR_Design { get; set; }

        [Required]
        [StringLength(11)]
        public string FA_CodeFamille { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(12)]
        public byte[] cbFA_CodeFamille { get; set; }

        [StringLength(19)]
        public string AR_Substitut { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Substitut { get; set; }

        [StringLength(7)]
        public string AR_Raccourci { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(8)]
        public byte[] cbAR_Raccourci { get; set; }

        public short? AR_Garantie { get; set; }

        public short? AR_UnitePoids { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_PoidsNet { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_PoidsBrut { get; set; }

        public short? AR_UniteVen { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_PrixAch { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Coef { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_PrixVen { get; set; }

        public short? AR_PrixTTC { get; set; }

        public short? AR_Gamme1 { get; set; }

        public short? AR_Gamme2 { get; set; }

        public short? AR_SuiviStock { get; set; }

        public short? AR_Nomencl { get; set; }

        [StringLength(21)]
        public string AR_Stat01 { get; set; }

        [StringLength(21)]
        public string AR_Stat02 { get; set; }

        [StringLength(21)]
        public string AR_Stat03 { get; set; }

        [StringLength(21)]
        public string AR_Stat04 { get; set; }

        [StringLength(21)]
        public string AR_Stat05 { get; set; }

        public short? AR_Escompte { get; set; }

        public short? AR_Delai { get; set; }

        public short? AR_HorsStat { get; set; }

        public short? AR_VteDebit { get; set; }

        public short? AR_NotImp { get; set; }

        public short? AR_Sommeil { get; set; }

        [StringLength(69)]
        public string AR_Langue1 { get; set; }

        [StringLength(69)]
        public string AR_Langue2 { get; set; }

        [StringLength(45)]
        public string AR_EdiCode { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(46)]
        public byte[] cbAR_EdiCode { get; set; }

        [StringLength(19)]
        public string AR_CodeBarre { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_CodeBarre { get; set; }

        [StringLength(25)]
        public string AR_CodeFiscal { get; set; }

        [StringLength(35)]
        public string AR_Pays { get; set; }

        [StringLength(21)]
        public string AR_Frais01FR_Denomination { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais01FR_Rem01REM_Valeur { get; set; }

        public short? AR_Frais01FR_Rem01REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais01FR_Rem02REM_Valeur { get; set; }

        public short? AR_Frais01FR_Rem02REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais01FR_Rem03REM_Valeur { get; set; }

        public short? AR_Frais01FR_Rem03REM_Type { get; set; }

        [StringLength(21)]
        public string AR_Frais02FR_Denomination { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais02FR_Rem01REM_Valeur { get; set; }

        public short? AR_Frais02FR_Rem01REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais02FR_Rem02REM_Valeur { get; set; }

        public short? AR_Frais02FR_Rem02REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais02FR_Rem03REM_Valeur { get; set; }

        public short? AR_Frais02FR_Rem03REM_Type { get; set; }

        [StringLength(21)]
        public string AR_Frais03FR_Denomination { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais03FR_Rem01REM_Valeur { get; set; }

        public short? AR_Frais03FR_Rem01REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais03FR_Rem02REM_Valeur { get; set; }

        public short? AR_Frais03FR_Rem02REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_Frais03FR_Rem03REM_Valeur { get; set; }

        public short? AR_Frais03FR_Rem03REM_Type { get; set; }

        public short? AR_Condition { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_PUNet { get; set; }

        public short? AR_Contremarque { get; set; }

        public short? AR_FactPoids { get; set; }

        public short? AR_FactForfait { get; set; }

        public short? AR_SaisieVar { get; set; }

        public short? AR_Transfere { get; set; }

        public short? AR_Publie { get; set; }

        public DateTime? AR_DateModif { get; set; }

        [StringLength(259)]
        public string AR_Photo { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_PrixAchNouv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_CoefNouv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_PrixVenNouv { get; set; }

        public DateTime? AR_DateApplication { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_CoutStd { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_QteComp { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AR_QteOperatoire { get; set; }

        public int? CO_No { get; set; }

        public int? cbCO_No { get; set; }

        public short? AR_Prevision { get; set; }

        public int? CL_No1 { get; set; }

        public int? cbCL_No1 { get; set; }

        public int? CL_No2 { get; set; }

        public int? cbCL_No2 { get; set; }

        public int? CL_No3 { get; set; }

        public int? cbCL_No3 { get; set; }

        public int? CL_No4 { get; set; }

        public int? cbCL_No4 { get; set; }

        public short? AR_Type { get; set; }

        [StringLength(11)]
        public string RP_CodeDefaut { get; set; }

        public short? AR_Nature { get; set; }

        public short? AR_DelaiFabrication { get; set; }

        public short? AR_NbColis { get; set; }

        public short? AR_DelaiPeremption { get; set; }

        public short? AR_DelaiSecurite { get; set; }

        public short? AR_Fictif { get; set; }

        public short? AR_SousTraitance { get; set; }

        public short? AR_TypeLancement { get; set; }

        public short? AR_Cycle { get; set; }

        public short? AR_Criticite { get; set; }

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

        public virtual F_ARTSTOCK artStock { get; set; }
    }
}
