using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Models
{
    public class P_TYPEREPARTITION
    {
        [Key]
        public int RP_Num { get; set; }

        [StringLength(69)]
        public string RP_Intitule { get; set; }

        // Navigation inverse (facultatif)
        public virtual ICollection<F_DOCFRAISIMPORT> FraisImports { get; set; }
    }

}
