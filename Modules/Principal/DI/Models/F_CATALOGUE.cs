namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_CATALOGUE
    {
        public int? CL_No { get; set; }

        [Required]
        [StringLength(35)]
        public string CL_Intitule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbCL_Intitule { get; set; }

        [StringLength(3)]
        public string CL_Code { get; set; }

        public short? CL_Stock { get; set; }

        public int? CL_NoParent { get; set; }

        public int? cbCL_NoParent { get; set; }

        public short? CL_Niveau { get; set; }

        public short? cbProt { get; set; }

        [Key]
        public int cbMarq { get; set; }

        [StringLength(4)]
        public string cbCreateur { get; set; }

        public DateTime? cbModification { get; set; }

        public int? cbReplication { get; set; }

        public short? cbFlag { get; set; }

        public DateTime? cbCreation { get; set; }

        public Guid? cbCreationUser { get; set; }
    }
}
