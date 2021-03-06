USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_workflows]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_workflows](
	[workflowID] [int] NOT NULL,
	[WorkflowName] [varchar](50) NOT NULL,
	[Comments] [varchar](150) NULL,
	[datetime_stamp] [datetime] NULL,
 CONSTRAINT [PK_inv_workflows] PRIMARY KEY CLUSTERED 
(
	[workflowID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_workflows] ADD  DEFAULT (getdate()) FOR [datetime_stamp]
GO
