namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_ARTCLIENT
    {
        [Required]
        [StringLength(19)]
        public string AR_Ref { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAR_Ref { get; set; }

        public short? AC_Categorie { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_PrixVen { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_Coef { get; set; }

        public short? AC_PrixTTC { get; set; }

        public short? AC_Arrondi { get; set; }

        public short? AC_QteMont { get; set; }

        public short? EG_Champ { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_PrixDev { get; set; }

        public short? AC_Devise { get; set; }

        [StringLength(17)]
        public string CT_Num { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(18)]
        public byte[] cbCT_Num { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_Remise { get; set; }

        public short? AC_Calcul { get; set; }

        public short? AC_TypeRem { get; set; }

        [StringLength(19)]
        public string AC_RefClient { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAC_RefClient { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_CoefNouv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_PrixVenNouv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_PrixDevNouv { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? AC_RemiseNouv { get; set; }

        public DateTime? AC_DateApplication { get; set; }

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
