USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_item_workflow_sets]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_item_workflow_sets](
	[setRowID] [int] IDENTITY(1,1) NOT NULL,
	[workflowID] [int] NOT NULL,
	[item_id] [int] NOT NULL,
	[Comments] [varchar](200) NULL,
	[datetime_stamp] [datetime] NULL,
 CONSTRAINT [PK_inv_reagent_workflow_sets] PRIMARY KEY CLUSTERED 
(
	[workflowID] ASC,
	[item_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_item_workflow_sets] ADD  CONSTRAINT [DF__inv_reage__datet__7D439ABD]  DEFAULT (getdate()) FOR [datetime_stamp]
GO
