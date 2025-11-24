namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_DOCLIGNEEMPL
    {
        public int DL_No { get; set; }

        public int DP_No { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_Qte { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? DL_QteAControler { get; set; }

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
