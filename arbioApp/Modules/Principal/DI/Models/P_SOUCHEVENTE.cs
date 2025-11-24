namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class P_SOUCHEVENTE
    {
        [StringLength(35)]
        public string S_Intitule { get; set; }

        public short? S_Valide { get; set; }

        [StringLength(7)]
        public string JO_Num { get; set; }

        [StringLength(7)]
        public string JO_NumSituation { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
