namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_GAMSTOCK
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

        [Column(TypeName = "numeric")]
        public decimal? GS_MontSto { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteSto { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteRes { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteCom { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteResCM { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteComCM { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteMini { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteMaxi { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QtePrepa { get; set; }

        public int? DP_NoPrincipal { get; set; }

        public int? cbDP_NoPrincipal { get; set; }

        public int? DP_NoControle { get; set; }

        public int? cbDP_NoControle { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? GS_QteAControler { get; set; }

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
