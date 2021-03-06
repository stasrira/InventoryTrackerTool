USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_usage]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_usage](
	[usage_transaction_id] [bigint] IDENTITY(1,1) NOT NULL,
	[item_id] [int] NOT NULL,
	[productCode] [bigint] NOT NULL,
	[itemQty] [decimal](8, 0) NOT NULL,
	[takenBy] [varchar](30) NULL,
	[takenDate] [datetime] NULL,
	[reservedBy] [varchar](30) NULL,
	[reservedDate] [datetime] NULL,
	[canceled] [int] NULL,
	[update_user] [varchar](50) NULL,
	[update_computer] [varchar](50) NULL,
	[datetime_stamp] [datetime] NOT NULL,
 CONSTRAINT [PK_inv_reagent_usage] PRIMARY KEY CLUSTERED 
(
	[usage_transaction_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Index [IX_inv_reagent_usage_reagent_prodCode]    Script Date: 11/9/2018 5:25:41 PM ******/
CREATE NONCLUSTERED INDEX [IX_inv_reagent_usage_reagent_prodCode] ON [dbo].[inv_usage]
(
	[item_id] ASC,
	[productCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_usage] ADD  CONSTRAINT [DF__inv_reage__datet__6EF57B66]  DEFAULT (getdate()) FOR [datetime_stamp]
GO
