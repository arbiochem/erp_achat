namespace arbioApp.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class f_devise : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.F_DEVISE",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        devise = c.String(),
                        valeur = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.F_DEVISE");
        }
    }
}
