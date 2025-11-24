namespace arbioApp.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    public partial class F_AGENDA
    {
        public short? AD_Champ { get; set; }

        [StringLength(21)]
        public string AD_Evenem { get; set; }

        public short? AG_Domaine { get; set; }

        [StringLength(19)]
        public string AG_Interes { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(20)]
        public byte[] cbAG_Interes { get; set; }

        public DateTime? AG_Debut { get; set; }

        public DateTime? AG_Fin { get; set; }

        public short? AG_Veille { get; set; }

        [StringLength(69)]
        public string AG_Comment { get; set; }

        public short? AG_Type { get; set; }

        public short? AG_Confirme { get; set; }

        [StringLength(9)]
        public string AG_HeureDebut { get; set; }

        [StringLength(9)]
        public string AG_HeureFin { get; set; }

        public short? AG_Ignorer { get; set; }

        public short? AG_Continue { get; set; }

        public int? DL_No { get; set; }

        public int? cbDL_No { get; set; }

        public int? DE_No { get; set; }

        public int? cbDE_No { get; set; }

        public int? PP_No { get; set; }

        public int? cbPP_No { get; set; }

        public int? AD_No { get; set; }

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
