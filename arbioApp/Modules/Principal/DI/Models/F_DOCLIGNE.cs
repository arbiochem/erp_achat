namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_DOCLIGNE
    {
        [NotMapped]
        public string CT_Intitule { get; set; }
        public short? DO_Domaine { get; set; }

        public short DO_Type { get; set; }

        [StringLength(17)]
        public string CT_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_Num { get; set; }

        [Required]
        [StringLength(9)]
        public string DO_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDO_Piece { get; set; }

        [StringLength(9)]
        public string DL_PieceBC { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDL_PieceBC { get; set; }

        [StringLength(9)]
        public string DL_PieceBL { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDL_PieceBL { get; set; }

        public DateTime? DO_Date { get; set; }

        public DateTime? DL_DateBC { get; set; }

        public DateTime? DL_DateBL { get; set; }

        public int? DL_Ligne { get; set; }

        [StringLength(17)]
        public string DO_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbDO_Ref { get; set; }

        public short? DL_TNomencl { get; set; }

        public short? DL_TRemPied { get; set; }

        public short? DL_TRemExep { get; set; }

        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        [StringLength(69)]
        public string DL_Design { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Qte { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_QteBC { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_QteBL { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_PoidsNet { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_PoidsBrut { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Remise01REM_Valeur { get; set; }

        public short? DL_Remise01REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Remise02REM_Valeur { get; set; }

        public short? DL_Remise02REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Remise03REM_Valeur { get; set; }

        public short? DL_Remise03REM_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_PrixUnitaire { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_PUBC { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Taxe1 { get; set; }

        public short? DL_TypeTaux1 { get; set; }

        public short? DL_TypeTaxe1 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Taxe2 { get; set; }

        public short? DL_TypeTaux2 { get; set; }

        public short? DL_TypeTaxe2 { get; set; }

        public int? CO_No { get; set; }

        public int? cbCO_No { get; set; }

        public int? AG_No1 { get; set; }

        public int? AG_No2 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_PrixRU { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_CMUP { get; set; }

        public short? DL_MvtStock { get; set; }

        public int? DT_No { get; set; }

        public int? cbDT_No { get; set; }

        [StringLength(19)]
        public string AF_RefFourniss { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAF_RefFourniss { get; set; }

        [StringLength(21)]
        public string EU_Enumere { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? EU_Qte { get; set; }

        public short? DL_TTC { get; set; }

        public int? DE_No { get; set; }

        public int cbDE_No { get; set; }

        public short? DL_NoRef { get; set; }

        public short? DL_TypePL { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_PUDevise { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_PUTTC { get; set; }

        public int? DL_No { get; set; }

        public DateTime? DO_DateLivr { get; set; }

        [StringLength(13)]
        public string CA_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCA_Num { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Taxe3 { get; set; }

        public short? DL_TypeTaux3 { get; set; }

        public short? DL_TypeTaxe3 { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Frais { get; set; }

        public short? DL_Valorise { get; set; }

        [StringLength(19)]
        public string AR_RefCompose { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_RefCompose { get; set; }

        public short? DL_NonLivre { get; set; }

        [StringLength(19)]
        public string AC_RefClient { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_MontantHT { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_MontantTTC { get; set; }

        public short? DL_FactPoids { get; set; }

        public short? DL_Escompte { get; set; }

        [StringLength(9)]
        public string DL_PiecePL { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDL_PiecePL { get; set; }

        public DateTime? DL_DatePL { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_QtePL { get; set; }
        public decimal? DL_MontantRegle { get; set; }

        [StringLength(19)]
        public string DL_NoColis { get; set; }

        public int? DL_NoLink { get; set; }

        public int? cbDL_NoLink { get; set; }

        [StringLength(11)]
        public string RP_Code { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(12)]
        public byte[] cbRP_Code { get; set; }

        public int? DL_QteRessource { get; set; }

        public DateTime? DL_DateAvancement { get; set; }

        [Required]
        [StringLength(9)]
        public string PF_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbPF_Num { get; set; }

        [StringLength(5)]
        public string DL_CodeTaxe1 { get; set; }

        [StringLength(5)]
        public string DL_CodeTaxe2 { get; set; }

        [StringLength(5)]
        public string DL_CodeTaxe3 { get; set; }

        public int? DL_PieceOFProd { get; set; }

        [StringLength(9)]
        public string DL_PieceDE { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDL_PieceDE { get; set; }

        public DateTime? DL_DateDE { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_QteDE { get; set; }

        [StringLength(11)]
        public string DL_Operation { get; set; }

        public int? DL_NoSousTotal { get; set; }

        public short? cbProt { get; set; }

        [Key]
        public int cbMarq { get; set; }

        [StringLength(4)]
        public string cbCreateur { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? cbModification { get; set; }

        public int? cbReplication { get; set; }

        public short? cbFlag { get; set; }

        //[Column(TypeName = "numeric")]
        //public decimal? Colisage { get; set; }

        //[Column("Unité de colisage")]
        //[StringLength(21)]
        //public string Unité_de_colisage { get; set; }

        //[StringLength(69)]
        //public string Commentaires { get; set; }

        public int? CA_No { get; set; }

        public int? cbCA_No { get; set; }

        [MaxLength(32)]
        public byte[] cbHash { get; set; }

        public short? cbHashVersion { get; set; }

        public DateTime? cbHashDate { get; set; }

        public int? cbHashOrder { get; set; }
        public bool Retenu { get; set; }
        public string? DL_PieceFourniss { get; set; }
        public DateTime? DL_DatePieceFourniss { get; set; }
        public Guid cbCreationUser { get; set; }

    }
}
