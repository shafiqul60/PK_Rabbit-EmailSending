namespace EmailSending.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class IntialDb : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.Mails",
                c => new
                    {
                        MailId = c.Int(nullable: false, identity: true),
                        MailTo = c.String(),
                        MailFrom = c.String(),
                        MailSubject = c.String(),
                        MailBody = c.String(),
                        SendingDateTime = c.DateTime(nullable: false),
                    })
                .PrimaryKey(t => t.MailId);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.Mails");
        }
    }
}
