namespace arbioApp
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Collaborateur")]
    public partial class Collaborateur
    {
        public int Id { get; set; }

        [Required]
        [StringLength(50)]
        public string Nom_Collab { get; set; }

        [StringLength(50)]
        public string Prenoms_Collab { get; set; }

        public int UserId { get; set; }
    }
}
