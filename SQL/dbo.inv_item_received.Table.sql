USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_item_received]    Script Date: 11/9/2018 5:25:40 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_item_received](
	[recieving_id] [bigint] IDENTITY(10001,1) NOT NULL,
	[order_id] [bigint] NOT NULL,
	[productCode] [bigint] NULL,
	[dateReceived] [datetime] NOT NULL,
	[receivedBy] [varchar](30) NULL,
	[productQty] [int] NOT NULL,
	[lotNum] [varchar](30) NULL,
	[expirationDate] [datetime] NULL,
	[StockLocation] [varchar](50) NULL,
	[comments] [varchar](150) NULL,
	[update_user] [varchar](50) NULL,
	[update_computer] [varchar](50) NULL,
	[datetime_stamp] [datetime] NOT NULL,
 CONSTRAINT [PK_inv_items_received] PRIMARY KEY CLUSTERED 
(
	[recieving_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
/****** Object:  Index [IX_inv_items_received_productCode]    Script Date: 11/9/2018 5:25:41 PM ******/
CREATE NONCLUSTERED INDEX [IX_inv_items_received_productCode] ON [dbo].[inv_item_received]
(
	[order_id] ASC,
	[productCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_item_received] ADD  CONSTRAINT [DF_inv_items_received_datetime_stamp]  DEFAULT (getdate()) FOR [datetime_stamp]
GO
