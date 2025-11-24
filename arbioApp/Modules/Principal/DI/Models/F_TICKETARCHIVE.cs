namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_TICKETARCHIVE
    {
        public int CA_No { get; set; }

        public int? CO_NoCaissier { get; set; }

        public int? cbCO_NoCaissier { get; set; }

        [Required]
        [StringLength(9)]
        public string TA_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbTA_Piece { get; set; }

        public DateTime? TA_Date { get; set; }

        [StringLength(9)]
        public string TA_Heure { get; set; }

        [Required]
        [StringLength(17)]
        public string CT_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_Num { get; set; }

        [StringLength(9)]
        public string DO_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(10)]
        public byte[] cbDO_Piece { get; set; }

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
