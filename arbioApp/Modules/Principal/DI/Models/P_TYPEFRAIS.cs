using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Models
{
    public class P_TYPEFRAIS
    {
        [Key]
        public int FR_Num { get; set; }

        [Required]
        [StringLength(69)]
        public string FR_Intitule { get; set; }

        public int? FR_ParentId { get; set; }

        [ForeignKey("FR_ParentId")]
        public virtual P_TYPEFRAIS Parent { get; set; }

        public virtual ICollection<P_TYPEFRAIS> Enfants { get; set; }

        public virtual ICollection<F_DOCFRAISIMPORT> FraisImports { get; set; }
    }





}
