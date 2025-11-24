namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_DEPOTEMPL
    {
        public int DE_No { get; set; }

        public int? DP_No { get; set; }

        [StringLength(13)]
        public string DP_Code { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbDP_Code { get; set; }

        [StringLength(35)]
        public string DP_Intitule { get; set; }

        public short? DP_Zone { get; set; }

        public short? DP_Type { get; set; }

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
