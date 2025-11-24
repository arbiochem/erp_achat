namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_ENUMGAMME
    {
        public short? EG_Champ { get; set; }

        public short? EG_Ligne { get; set; }

        [StringLength(35)]
        public string EG_Enumere { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbEG_Enumere { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? EG_BorneSup { get; set; }

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
