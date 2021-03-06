USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_item_orders]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_item_orders](
	[order_id] [bigint] IDENTITY(1001,1) NOT NULL,
	[productCode] [bigint] NOT NULL,
	[dateOrdered] [datetime] NOT NULL,
	[orderedBy] [varchar](30) NOT NULL,
	[orderQty] [int] NOT NULL,
	[order_number] [varchar](50) NULL,
	[comments] [varchar](150) NULL,
	[update_user] [varchar](50) NULL,
	[update_computer] [varchar](50) NULL,
	[datetime_stamp] [datetime] NOT NULL,
 CONSTRAINT [PK_inv_item_orders] PRIMARY KEY CLUSTERED 
(
	[order_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Index [IX_inv_item_orders_productCode]    Script Date: 11/9/2018 5:25:41 PM ******/
CREATE NONCLUSTERED INDEX [IX_inv_item_orders_productCode] ON [dbo].[inv_item_orders]
(
	[productCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_item_orders] ADD  CONSTRAINT [DF__inv_item___datet__6FE99F9F]  DEFAULT (getdate()) FOR [datetime_stamp]
GO
