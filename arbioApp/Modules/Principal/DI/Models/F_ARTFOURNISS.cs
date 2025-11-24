namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_ARTFOURNISS
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        [Required]
        [StringLength(17)]
        public string CT_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_Num { get; set; }

        [StringLength(19)]
        public string AF_RefFourniss { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAF_RefFourniss { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_PrixAch { get; set; }

        public short? AF_Unite { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_Conversion { get; set; }

        public short? AF_DelaiAppro { get; set; }

        public short? AF_Garantie { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_Colisage { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_QteMini { get; set; }

        public short? AF_QteMont { get; set; }

        public short? EG_Champ { get; set; }

        public short? AF_Principal { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_PrixDev { get; set; }

        public short? AF_Devise { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_Remise { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_ConvDiv { get; set; }

        public short? AF_TypeRem { get; set; }

        [StringLength(19)]
        public string AF_CodeBarre { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAF_CodeBarre { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_PrixAchNouv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_PrixDevNouv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AF_RemiseNouv { get; set; }

        public DateTime? AF_DateApplication { get; set; }

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


        [NotMapped] // ne sera pas mappée en base
        public string CT_Intitule { get; set; }
    }
}
