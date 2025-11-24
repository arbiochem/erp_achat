namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;
    using System.Runtime.CompilerServices;

    public partial class F_FRET
    {
        public int Id { get; set; }
        [StringLength(100)]
        public string DO_PIECE { get; set; }

        public decimal DO_PRIX { get; set; }
        public decimal DO_POIDS { get; set; }
        public decimal DO_MONTANT { get; set; }


    }
}
