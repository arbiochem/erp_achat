namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_GAMSTOCKEMPL
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        public int? AG_No1 { get; set; }

        public int? AG_No2 { get; set; }

        public int DE_No { get; set; }

        public int DP_No { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GE_QteSto { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GE_QtePrepa { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GE_QteAControler { get; set; }

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
