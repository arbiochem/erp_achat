namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class P_DEVISE
    {
        [StringLength(35)]
        public string D_Intitule { get; set; }

        [StringLength(31)]
        public string D_Format { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? D_Cours { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? D_CoursP { get; set; }

        [StringLength(21)]
        public string D_Monnaie { get; set; }

        [StringLength(21)]
        public string D_SousMonnaie { get; set; }

        [StringLength(3)]
        public string D_CodeISO { get; set; }

        [StringLength(5)]
        public string D_Sigle { get; set; }

        public short? D_Mode { get; set; }

        public short? N_DeviseCot { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? D_CoursClot { get; set; }

        public DateTime? D_AncDate { get; set; }

        [Column(TypeName = "numeric")]
        public decimal? D_AncCours { get; set; }

        public short? D_AncMode { get; set; }

        public short? N_DeviseAncCot { get; set; }

        public short? D_CodeRemise { get; set; }

        public short? D_Euro { get; set; }

        [StringLength(5)]
        public string D_CodeISONum { get; set; }

        public DateTime? D_UpdateDate { get; set; }

        [StringLength(9)]
        public string D_UpdateTime { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
