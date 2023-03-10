USE [TMS_CM_ANOnline]
GO
/****** Object:  Table [dbo].[ATC_AbsenceRequests]    Script Date: 6/13/2016 6:52:51 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[ATC_AbsenceRequests](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Type] [int] NOT NULL,
	[StaffId] [int] NOT NULL,
	[Authoriser1_Id] [int] NOT NULL,
	[Authoriser2_Id] [int] NULL,
	[DateFrom] [datetime] NOT NULL,
	[DateTo] [datetime] NOT NULL,
	[Note] [nvarchar](500) NULL,
	[isAuthorisedByHr] [bit] NOT NULL,
	[isAuthoriser1Approved] [bit] NOT NULL,
	[isAuthoriser2Approved] [bit] NULL,
	[isHrApproved] [bit] NULL,
	[Status] [int] NOT NULL,
 CONSTRAINT [PK_dbo.ATC_AbsenceRequests] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO
ALTER TABLE [dbo].[ATC_AbsenceRequests]  WITH CHECK ADD  CONSTRAINT [FK_dbo.ATC_AbsenceRequests_dbo.ATC_Employees_Authoriser1_Id] FOREIGN KEY([Authoriser1_Id])
REFERENCES [dbo].[ATC_Employees] ([StaffID])
GO
ALTER TABLE [dbo].[ATC_AbsenceRequests] CHECK CONSTRAINT [FK_dbo.ATC_AbsenceRequests_dbo.ATC_Employees_Authoriser1_Id]
GO
ALTER TABLE [dbo].[ATC_AbsenceRequests]  WITH CHECK ADD  CONSTRAINT [FK_dbo.ATC_AbsenceRequests_dbo.ATC_Employees_Authoriser2_Id] FOREIGN KEY([Authoriser2_Id])
REFERENCES [dbo].[ATC_Employees] ([StaffID])
GO
ALTER TABLE [dbo].[ATC_AbsenceRequests] CHECK CONSTRAINT [FK_dbo.ATC_AbsenceRequests_dbo.ATC_Employees_Authoriser2_Id]
GO
ALTER TABLE [dbo].[ATC_AbsenceRequests]  WITH CHECK ADD  CONSTRAINT [FK_dbo.ATC_AbsenceRequests_dbo.ATC_Employees_StaffId] FOREIGN KEY([StaffId])
REFERENCES [dbo].[ATC_Employees] ([StaffID])
GO
ALTER TABLE [dbo].[ATC_AbsenceRequests] CHECK CONSTRAINT [FK_dbo.ATC_AbsenceRequests_dbo.ATC_Employees_StaffId]
GO

/*Create HR Receive Report View*/
CREATE VIEW [dbo].[HR_ReceiveReport]
AS
SELECT        TOP (100) PERCENT a.UserID, e.FirstName + ' ' + ISNULL(e.LastName, '') + ' ' + ISNULL(e.MiddleName, '') AS Fullname
FROM            dbo.ATC_UserGroup AS a LEFT OUTER JOIN
                         dbo.ATC_Group AS b ON a.GroupID = b.GroupID LEFT OUTER JOIN
                         dbo.ATC_Permissions AS c ON b.GroupID = c.GroupID LEFT OUTER JOIN
                         dbo.ATC_Functions AS d ON c.FunctionID = d.FunctionID LEFT OUTER JOIN
                         dbo.ATC_PersonalInfo AS e ON a.UserID = e.PersonID
WHERE        (d.Description = 'Receive Report') AND (e.fgDelete = 0)
GO

/*Create Session State Table*/
if exists (select * from sysobjects where id = object_id(N'[dbo].[SessionState]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[SessionState]
GO

CREATE TABLE [dbo].[SessionState] (
	[ID] uniqueidentifier  NOT NULL ,
	[Data] [image] NOT NULL ,
	[Last_Accessed] [datetime] NOT NULL 
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [dbo].[SessionState] WITH NOCHECK ADD 
	CONSTRAINT [PK_SessionState] PRIMARY KEY  NONCLUSTERED 
	(
		[ID]
	)  ON [PRIMARY] 
GO

/*Add Note to AbsenceRequest*/
ALTER TABLE[dbo].[ATC_AbsenceRequests]
ADD Authoriser1Note Nvarchar(max)

ALTER TABLE[dbo].[ATC_AbsenceRequests]
ADD Authoriser2Note Nvarchar(max)

ALTER TABLE[dbo].[ATC_AbsenceRequests]
ADD HrNote Nvarchar(max)

/*Email Templates*/
CREATE TABLE [dbo].[ATC_EmailTemplate](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[Type] [nvarchar](255) NULL,	
	[Subject] [nvarchar](255) NULL,
	[Content] [nvarchar](max) NULL,
	[Note] [nvarchar](max) NULL	
 CONSTRAINT [PK_EmailTemplate] PRIMARY KEY CLUSTERED 
(
	[Id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

INSERT INTO [dbo].[ATC_EmailTemplate]
           ([Type]
           ,[Subject]
           ,[Content]
           ,[Note])
     VALUES
           ('Email Request Approval'
           ,'Email Request Approval'
           ,'<p>Dear #Manager#,</p> 
			 <p>I would be grateful if you could review my leave request, please select the following link <a href="#LinkApproveRequestForManager#">click here</a></p>
			 <p>Best regards,</p> 
			 <p>#Requester#</p>'
           ,null)
GO

INSERT INTO [dbo].[ATC_EmailTemplate]
           ([Type]
           ,[Subject]
           ,[Content]
           ,[Note])
     VALUES
           ('Email Inform Approve By Authorizor1'
           ,'Email Inform Approve'
           ,'<p>Dear #Requester#,</p> 
			 <p>Your leave request from #DateFrom# to #DateTo# has been approved by #Manager#</p>
			 <p>#Note#</p>
			 <p>Your leave request is still being progressed; you will be notified by email when it is approved.</p> 
			 <p>Yours faithfully,</p> 
			 <p>#Manager#</p>'
           ,null)
GO

INSERT INTO [dbo].[ATC_EmailTemplate]
           ([Type]
           ,[Subject]
           ,[Content]
           ,[Note])
     VALUES
           ('Email Inform Approve By Authorizor2'
           ,'Email Inform Approve'
           ,'<p>Dear #Requester#,</p> 
			 <p>Your leave request from #DateFrom# to #DateTo# has been approved by #Manager#</p>
			 <p>#Note#</p>
			 <p>Yours faithfully,</p> 
			 <p>#Manager#</p>'
           ,null)
GO

INSERT INTO [dbo].[ATC_EmailTemplate]
           ([Type]
           ,[Subject]
           ,[Content]
           ,[Note])
     VALUES
           ('Email Reject'
           ,'Reject Request'
           ,'<p>Dear #Requester#,</p> 
			 <p>Your leave request from #DateFrom# to #DateTo# has not been approved.</p>
			 <p>#Note#</p>
			 <p>If you have any queries please contact direct to your Manager.</p>
			 <p>Regards,</p> 
			 <p>#Manager#</p>'
           ,null)
GO

/*Add menu link for new pages*/
INSERT INTO ATC_Functions ( [Description] , Form , GroupID , LevelOrder , fgUpdateable , fgAttribute ) VALUES(
	'HR Authorisation','aisnet/HRAuthorisation/HRAuthorisation',58,3,1,0)

INSERT INTO ATC_Functions ( [Description] , Form , GroupID , LevelOrder , fgUpdateable , fgAttribute ) VALUES(
	'Team Calendar','aisnet/Authoriser/Authoriser',58,4,1,0)
GO

INSERT INTO ATC_Permissions (GroupID,FunctionID) 
	SELECT 6,FunctionID FROM ATC_Functions WHERE [Description] IN ('HR Authorisation','Team Calendar')

