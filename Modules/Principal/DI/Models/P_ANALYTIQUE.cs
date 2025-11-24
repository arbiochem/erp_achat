namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class P_ANALYTIQUE
    {
        [StringLength(35)]
        public string A_Intitule { get; set; }

        [StringLength(21)]
        public string A_Rupture01A_Nom { get; set; }

        public short? A_Rupture01A_Lg { get; set; }

        public short? A_Rupture01A_Type { get; set; }

        [StringLength(21)]
        public string A_Rupture02A_Nom { get; set; }

        public short? A_Rupture02A_Lg { get; set; }

        public short? A_Rupture02A_Type { get; set; }

        [StringLength(21)]
        public string A_Rupture03A_Nom { get; set; }

        public short? A_Rupture03A_Lg { get; set; }

        public short? A_Rupture03A_Type { get; set; }

        [StringLength(21)]
        public string A_Rupture04A_Nom { get; set; }

        public short? A_Rupture04A_Lg { get; set; }

        public short? A_Rupture04A_Type { get; set; }

        [StringLength(21)]
        public string A_Rupture05A_Nom { get; set; }

        public short? A_Rupture05A_Lg { get; set; }

        public short? A_Rupture05A_Type { get; set; }

        [StringLength(21)]
        public string A_Rupture06A_Nom { get; set; }

        public short? A_Rupture06A_Lg { get; set; }

        public short? A_Rupture06A_Type { get; set; }

        [StringLength(13)]
        public string CA_Num { get; set; }

        public short? A_Colonne { get; set; }

        public short? A_Imputation { get; set; }

        public short? A_Obligatoire { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
