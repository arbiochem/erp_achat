namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_REGLEARCHIVE
    {
        [Required]
        [StringLength(13)]
        public string TA_Piece { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(14)]
        public byte[] cbTA_Piece { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RA_Montant { get; set; }

        public short? N_Devise { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? RA_MontantDev { get; set; }

        public short? N_Reglement { get; set; }

        public DateTime? RA_Date { get; set; }

        public int? CA_No { get; set; }

        public int? cbCA_No { get; set; }

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
