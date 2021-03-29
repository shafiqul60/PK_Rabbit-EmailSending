namespace EmailSending.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class first : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.MailModels",
                c => new
                    {
                        mailModelID = c.Int(nullable: false, identity: true),
                        To = c.String(),
                        Subject = c.String(),
                        Body = c.String(),
                    })
                .PrimaryKey(t => t.mailModelID);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.MailModels");
        }
    }
}
