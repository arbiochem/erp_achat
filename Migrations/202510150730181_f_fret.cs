namespace arbioApp.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class f_fret : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.F_FRET",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        DO_PIECE = c.String(maxLength: 100),
                        DO_PRIX = c.Decimal(nullable: false, precision: 18, scale: 2),
                        DO_POIDS = c.Decimal(nullable: false, precision: 18, scale: 2),
                        DO_MONTANT = c.Decimal(nullable: false, precision: 18, scale: 2),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.F_FRET");
        }
    }
}
