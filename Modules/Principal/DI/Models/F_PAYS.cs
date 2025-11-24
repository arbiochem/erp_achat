namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_PAYS
    {
        [StringLength(35)]
        public string PA_Intitule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbPA_Intitule { get; set; }

        [StringLength(3)]
        public string PA_Code { get; set; }

        [StringLength(3)]
        public string PA_CodeEdi { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? PA_Assurance { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? PA_Transport { get; set; }

        [StringLength(3)]
        public string PA_CodeISO2 { get; set; }

        public short? PA_SEPA { get; set; }

        public int? PA_No { get; set; }

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
