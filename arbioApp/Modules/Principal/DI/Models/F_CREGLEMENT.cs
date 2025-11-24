namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_CREGLEMENT
    {
        public int? RG_No { get; set; }

        [StringLength(17)]
        public string CT_NumPayeur { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_NumPayeur { get; set; }

        public DateTime? RG_Date { get; set; }

        [StringLength(17)]
        public string RG_Reference { get; set; }

        [StringLength(35)]
        public string RG_Libelle { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RG_Montant { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RG_MontantDev { get; set; }

        public short? N_Reglement { get; set; }

        public short? RG_Impute { get; set; }

        public short? RG_Compta { get; set; }

        public int? EC_No { get; set; }

        public int? cbEC_No { get; set; }

        public short? RG_Type { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RG_Cours { get; set; }

        public short? N_Devise { get; set; }

        [Required]
        [StringLength(7)]
        public string JO_Num { get; set; }

        [StringLength(13)]
        public string CG_NumCont { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCG_NumCont { get; set; }

        public DateTime? RG_Impaye { get; set; }

        [StringLength(13)]
        public string CG_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCG_Num { get; set; }

        public short? RG_TypeReg { get; set; }

        [StringLength(9)]
        public string RG_Heure { get; set; }

        [StringLength(13)]
        public string RG_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbRG_Piece { get; set; }

        public int? CA_No { get; set; }

        public int? cbCA_No { get; set; }

        public int? CO_NoCaissier { get; set; }

        public int? cbCO_NoCaissier { get; set; }

        public short? RG_Banque { get; set; }

        public short? RG_Transfere { get; set; }

        public short RG_Cloture { get; set; }

        public short? RG_Ticket { get; set; }

        public short? RG_Souche { get; set; }

        [StringLength(17)]
        public string CT_NumPayeurOrig { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_NumPayeurOrig { get; set; }

        public DateTime? RG_DateEchCont { get; set; }

        [StringLength(13)]
        public string CG_NumEcart { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCG_NumEcart { get; set; }

        [StringLength(7)]
        public string JO_NumEcart { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RG_MontantEcart { get; set; }

        public int? RG_NoBonAchat { get; set; }

        public short? RG_Valide { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RG_Anterieur { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RG_MontantCommission { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RG_MontantNet { get; set; }

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


        // 🔗 Navigation properties
        //public virtual F_COMPTET CT_Num { get; set; }
        //public virtual P_REGLEMENT R_Intitule { get; set; }
        //public virtual P_DEVISE D_Intitule { get; set; }

    }
}
