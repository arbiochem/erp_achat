using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Models
{
    public class F_DOCFRAISIMPORT
    {
        [Key]
        public int FI_ID { get; set; }

        [Required]
        [StringLength(50)]
        public string DO_Piece { get; set; }

        [Required]
        public int FI_TypeFraisId { get; set; }

        [Required]
        //[Column(TypeName = "decimal(18,2)")]
        public decimal FI_Montant { get; set; }

        [StringLength(10)]
        public string FI_Devise { get; set; } = "USD";

        [Required]
        //[Column(TypeName = "decimal(18,2)")]
        public decimal FI_Montant_AR { get; set; }

        [Required]
        public int FI_RepartitionId { get; set; }


        [Required]
        public string FI_Piece { get; set; }
        public string FI_Observation { get; set; }

        public DateTime cbModification { get; set; }
        public Guid Username { get; set; }

        [StringLength(255)]
        public string FLAG1 { get; set; }

        [StringLength(255)]
        public string FLAG2 { get; set; }

        [StringLength(255)]
        public string FLAG3 { get; set; }

        // 🔗 Navigation properties
        public virtual P_TYPEFRAIS TypeFrais { get; set; }
        public virtual P_TYPEREPARTITION Repartition { get; set; }
    }


}
