namespace AIS.Data.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class add__table__ATC_AbsenceRequests : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.ATC_AbsenceRequests",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Type = c.Int(nullable: false),
                        StaffId = c.Int(nullable: false),
                        Authoriser1_Id = c.Int(nullable: false),
                        Authoriser2_Id = c.Int(),
                        DateFrom = c.DateTime(nullable: false),
                        DateTo = c.DateTime(nullable: false),
                        Note = c.String(maxLength: 500),
                        isAuthorisedByHr = c.Boolean(nullable: false),
                        isAuthoriser1Approved = c.Boolean(nullable: false),
                        isAuthoriser2Approved = c.Boolean(),
                        isHrApproved = c.Boolean(),
                    })
                .PrimaryKey(t => t.Id)
                .ForeignKey("dbo.ATC_Employees", t => t.Authoriser1_Id)
                .ForeignKey("dbo.ATC_Employees", t => t.Authoriser2_Id)
                .ForeignKey("dbo.ATC_Employees", t => t.StaffId)
                .Index(t => t.StaffId)
                .Index(t => t.Authoriser1_Id)
                .Index(t => t.Authoriser2_Id);
            
        }
        
        public override void Down()
        {
            DropForeignKey("dbo.ATC_AbsenceRequests", "StaffId", "dbo.ATC_Employees");
            DropForeignKey("dbo.ATC_AbsenceRequests", "Authoriser2_Id", "dbo.ATC_Employees");
            DropForeignKey("dbo.ATC_AbsenceRequests", "Authoriser1_Id", "dbo.ATC_Employees");
            DropIndex("dbo.ATC_AbsenceRequests", new[] { "Authoriser2_Id" });
            DropIndex("dbo.ATC_AbsenceRequests", new[] { "Authoriser1_Id" });
            DropIndex("dbo.ATC_AbsenceRequests", new[] { "StaffId" });
            DropTable("dbo.ATC_AbsenceRequests");
        }
    }
}
