namespace arbioApp.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class F_DOCFRAISIMPORT : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.F_DOCFRAISIMPORT", "FLAG1", c => c.String(maxLength: 255));
            AlterColumn("dbo.F_DOCFRAISIMPORT", "FLAG2", c => c.String(maxLength: 255));
            AlterColumn("dbo.F_DOCFRAISIMPORT", "FLAG3", c => c.String(maxLength: 255));
        }
        
        public override void Down()
        {
            AlterColumn("dbo.F_DOCFRAISIMPORT", "FLAG3", c => c.String(nullable: false));
            AlterColumn("dbo.F_DOCFRAISIMPORT", "FLAG2", c => c.String(nullable: false));
            AlterColumn("dbo.F_DOCFRAISIMPORT", "FLAG1", c => c.String(nullable: false));
        }
    }
}
