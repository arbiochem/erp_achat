using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.Modules.Principal.DI.Models
{
    public class F_DOCCOUTREVIENT
    {
        [Key]
        public int CR_ID { get; set; }

        [Required]
        [StringLength(50)]
        public string DO_Piece { get; set; }

        [Required]
        public int DL_Ligne { get; set; }

        [Required]
        [StringLength(50)]
        public string AR_Ref { get; set; }

        //[Column(TypeName = "decimal(18,2)")]
        public decimal CR_CoutAchat { get; set; }

        //[Column(TypeName = "decimal(18,2)")]
        public decimal CR_FraisImport { get; set; }

        //[Column(TypeName = "decimal(18,2)")]
        public decimal CR_CoutTotal { get; set; }

        //[Column(TypeName = "decimal(18,4)")]
        public decimal CR_CoutUnitaire { get; set; }
    }

}
