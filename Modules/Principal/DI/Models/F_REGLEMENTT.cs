namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_REGLEMENTT
    {
        [Required]
        [StringLength(17)]
        public string CT_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_Num { get; set; }

        public short? N_Reglement { get; set; }

        public short? RT_Condition { get; set; }

        public short? RT_NbJour { get; set; }

        public short? RT_JourTb01 { get; set; }

        public short? RT_JourTb02 { get; set; }

        public short? RT_JourTb03 { get; set; }

        public short? RT_JourTb04 { get; set; }

        public short? RT_JourTb05 { get; set; }

        public short? RT_JourTb06 { get; set; }

        public short? RT_TRepart { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RT_VRepart { get; set; }

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
