namespace arbioApp.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class f_devises : DbMigration
    {
        public override void Up()
        {
            AlterColumn("dbo.F_DEVISE", "valeur", c => c.Decimal(nullable: false, precision: 18, scale: 2));
        }
        
        public override void Down()
        {
            AlterColumn("dbo.F_DEVISE", "valeur", c => c.String());
        }
    }
}
