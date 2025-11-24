namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_TAXE
    {
        [StringLength(35)]
        public string TA_Intitule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbTA_Intitule { get; set; }

        public short? TA_TTaux { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? TA_Taux { get; set; }

        public short? TA_Type { get; set; }

        [Required]
        [StringLength(13)]
        public string CG_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbCG_Num { get; set; }

        public int TA_No { get; set; }

        [Required]
        [StringLength(5)]
        public string TA_Code { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(6)]
        public byte[] cbTA_Code { get; set; }

        public short? TA_NP { get; set; }

        public short? TA_Sens { get; set; }

        public short? TA_Provenance { get; set; }

        [StringLength(5)]
        public string TA_Regroup { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(6)]
        public byte[] cbTA_Regroup { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? TA_Assujet { get; set; }

        [StringLength(3)]
        public string TA_GrilleBase { get; set; }

        [StringLength(3)]
        public string TA_GrilleTaxe { get; set; }

        [StringLength(3)]
        public string TA_EdiCode { get; set; }

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
