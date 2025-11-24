namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_BILLETPIECE
    {
        public short? N_Devise { get; set; }

        [StringLength(35)]
        public string BI_Intitule { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? BI_Valeur { get; set; }

        public short? cbProt { get; set; }

        [Key]
        public int cbMarq { get; set; }

        [StringLength(4)]
        public string cbCreateur { get; set; }

        [Column(TypeName = "smalldatetime")]
        public DateTime? cbModification { get; set; }

        public int? cbReplication { get; set; }

        public short? cbFlag { get; set; }
    }
}
