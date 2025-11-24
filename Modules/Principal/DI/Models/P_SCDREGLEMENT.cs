namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class P_SCDREGLEMENT
    {
        public short? SCD_Type { get; set; }

        [StringLength(7)]
        public string JO_NumCli { get; set; }

        [StringLength(7)]
        public string JO_NumFour { get; set; }

        public short? cbIndice { get; set; }

        [Key]
        public int cbMarq { get; set; }

        public short? SCD_Arrondir { get; set; }
    }
}
