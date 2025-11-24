namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_CONDITION
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        public int? CO_No { get; set; }

        [StringLength(35)]
        public string EC_Enumere { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbEC_Enumere { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? EC_Quantite { get; set; }

        [StringLength(19)]
        public string CO_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbCO_Ref { get; set; }

        [StringLength(19)]
        public string CO_CodeBarre { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbCO_CodeBarre { get; set; }

        public short? CO_Principal { get; set; }

        [StringLength(3)]
        public string CO_EdiCode { get; set; }

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
