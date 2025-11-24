namespace arbioApp.Models
{
    using System.ComponentModel.DataAnnotations;

    public partial class P_COLREGLEMENT
    {
        public short? CR_NumPiece01 { get; set; }

        public short? CR_NumPiece02 { get; set; }

        [StringLength(13)]
        public string CR_Numero01 { get; set; }

        [StringLength(13)]
        public string CR_Numero02 { get; set; }

        public short? CR_ColReglement01 { get; set; }

        public short? CR_ColReglement02 { get; set; }

        public short? CR_ColReglement03 { get; set; }

        public short? CR_ColReglement04 { get; set; }

        public short? CR_ColReglement05 { get; set; }

        public short? CR_ColReglement06 { get; set; }

        public short? CR_ColReglement07 { get; set; }

        public short? CR_ColReglement08 { get; set; }

        public short? CR_ColReglement09 { get; set; }

        public short? CR_ColReglement10 { get; set; }

        public short? CR_ColReglement11 { get; set; }

        public short? CR_ColReglement12 { get; set; }

        public short? CR_ColReglement13 { get; set; }

        public short? CR_ColReglement14 { get; set; }

        public short? CR_ColReglement15 { get; set; }

        public short? CR_ColReglement16 { get; set; }

        public short? CR_ColReglement17 { get; set; }

        public short? CR_ColReglement18 { get; set; }

        public short? CR_ColReglement19 { get; set; }

        public short? CR_ColReglement20 { get; set; }

        public short? CR_ColReglement21 { get; set; }

        public short? CR_ColReglement22 { get; set; }

        public short? CR_ColReglement23 { get; set; }

        public short? CR_ColReglement24 { get; set; }

        public short? CR_ColReglement25 { get; set; }

        public short? CR_ColReglement26 { get; set; }

        public short? CR_ColReglement27 { get; set; }

        [Key]
        public int cbMarq { get; set; }
    }
}
