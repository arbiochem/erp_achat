using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace arbioApp.DTO
{
    public class ListeEcheancesPourImpressionDocumentsDeVente
    {
        public DateTime? DR_Date { get; set; }
        [Column(TypeName = "numeric")]
        public decimal? A_Payer { get; set; }
    }
}
