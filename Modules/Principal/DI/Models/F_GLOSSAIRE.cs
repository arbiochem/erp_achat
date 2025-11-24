namespace arbioApp.Models
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;

    public partial class F_GLOSSAIRE
    {
        public int? GL_No { get; set; }

        public short? GL_Domaine { get; set; }

        [StringLength(35)]
        public string GL_Intitule { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(36)]
        public byte[] cbGL_Intitule { get; set; }

        [StringLength(7)]
        public string GL_Raccourci { get; set; }

        [DatabaseGenerated(DatabaseGeneratedOption.Computed)]
        [MaxLength(8)]
        public byte[] cbGL_Raccourci { get; set; }

        public DateTime? GL_PeriodeDeb { get; set; }

        public DateTime? GL_PeriodeFin { get; set; }

        public string GL_Text { get; set; }

        public string GL_TextLangue1 { get; set; }

        public string GL_TextLangue2 { get; set; }

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
