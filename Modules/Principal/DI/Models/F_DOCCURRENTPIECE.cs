namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_DOCCURRENTPIECE
    {
        public short? DC_Domaine { get; set; }

        public short? DC_IdCol { get; set; }

        public short? DC_Souche { get; set; }

        [StringLength(9)]
        public string DC_Piece { get; set; }

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
