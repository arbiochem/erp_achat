using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Models
{
    public partial class F_COORDINATES
    {
        public int Id { get; set; }

        
        public string DO_Piece { get; set; }

        
        public float? Latitude { get; set; }

        public float? Longitude { get; set; }
        public DateTime? DatePos { get; set; }
    }
}
