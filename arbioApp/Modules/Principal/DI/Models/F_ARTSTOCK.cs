namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_ARTSTOCK
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        public int DE_No { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteMini { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteMaxi { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_MontSto { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteSto { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteRes { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteCom { get; set; }

        public short? AS_Principal { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteResCM { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteComCM { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QtePrepa { get; set; }

        public int? DP_NoPrincipal { get; set; }

        public int? cbDP_NoPrincipal { get; set; }

        public int? DP_NoControle { get; set; }

        public int? cbDP_NoControle { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AS_QteAControler { get; set; }

        public short? AS_Mouvemente { get; set; }

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

        public virtual F_ARTICLE FArticle { get; set; }
    }
}
