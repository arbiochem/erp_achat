namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_ARTICLEMEDIA
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        [StringLength(69)]
        public string ME_Commentaire { get; set; }

        [StringLength(259)]
        public string ME_Fichier { get; set; }

        [StringLength(255)]
        public string ME_TypeMIME { get; set; }

        [StringLength(255)]
        public string ME_Origine { get; set; }

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

        [StringLength(255)]
        public string ME_GedId { get; set; }
    }
}
