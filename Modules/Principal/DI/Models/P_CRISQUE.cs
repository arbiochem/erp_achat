using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Data.Entity.Spatial;

namespace arbioApp.Modules.Principal.DI.Models
{
    public partial class P_CRISQUE
    {
        public string R_Intitule { get; set; }
        public Nullable<short> R_Type { get; set; }
        public Nullable<decimal> R_Min { get; set; }
        public Nullable<decimal> R_Max { get; set; }
        public Nullable<short> cbIndice { get; set; }
        [Key]
        public int cbMarq { get; set; }
    }
}
