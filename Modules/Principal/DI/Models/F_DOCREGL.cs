namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_DOCREGL
    {
        public int? DR_No { get; set; }

        public short? DO_Domaine { get; set; }

        public short? DO_Type { get; set; }

        [StringLength(13)]
        public string DO_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbDO_Piece { get; set; }

        public short? DR_TypeRegl { get; set; }

        public DateTime? DR_Date { get; set; }

        [StringLength(35)]
        public string DR_Libelle { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DR_Pourcent { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DR_Montant { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DR_MontantDev { get; set; }

        public short? DR_Equil { get; set; }

        public int? EC_No { get; set; }

        public int? cbEC_No { get; set; }

        public short? DR_Regle { get; set; }

        public short? N_Reglement { get; set; }

        public int? CA_No { get; set; }

        public int? cbCA_No { get; set; }

        public short? DO_DocType { get; set; }

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

        public Guid? DR_RefPaiement { get; set; }

        [StringLength(255)]
        public string DR_AdressePaiement { get; set; }
    }
}
