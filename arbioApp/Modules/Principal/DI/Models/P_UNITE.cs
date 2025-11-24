namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class P_UNITE
    {
        [StringLength(21)]
        public string U_Intitule { get; set; }

        public short? U_Correspondance { get; set; }

        public short? U_NbUnite { get; set; }

        public short? U_UniteTemps { get; set; }

        [StringLength(3)]
        public string U_EdiCode { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
