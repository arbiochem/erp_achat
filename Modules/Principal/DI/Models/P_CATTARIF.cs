namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class P_CATTARIF
    {
        [StringLength(35)]
        public string CT_Intitule { get; set; }

        public short? CT_PrixTTC { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
