namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_REGLECH
    {
        public int RG_No { get; set; }

        public int DR_No { get; set; }

        public short? DO_Domaine { get; set; }

        public short? DO_Type { get; set; }

        [StringLength(9)]
        public string DO_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDO_Piece { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RC_Montant { get; set; }

        public short? RG_TypeReg { get; set; }

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
