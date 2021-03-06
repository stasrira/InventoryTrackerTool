USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_items]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_items](
	[item_id] [int] NOT NULL,
	[ItemName] [varchar](80) NOT NULL,
	[itemCategory] [varchar](50) NULL,
	[item_units] [varchar](20) NULL,
	[minSampleGroupQty] [int] NULL,
	[minItemQtyPerSampleGroup] [decimal](10, 4) NULL,
	[minStockQty] [decimal](8, 2) NULL,
	[refillFactorOfMinStockQty] [decimal](8, 2) NULL,
	[notifyFactorOfMinStockQty] [decimal](8, 2) NULL,
	[show_order] [decimal](8, 2) NULL,
	[datetime_stamp] [datetime] NULL,
 CONSTRAINT [PK_inv_reagents] PRIMARY KEY CLUSTERED 
(
	[item_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_items] ADD  CONSTRAINT [DF_inv_items_minSampleGroupQty]  DEFAULT ((1)) FOR [minSampleGroupQty]
GO
ALTER TABLE [dbo].[inv_items] ADD  CONSTRAINT [DF_inv_items_minItemQtyPerSampleGroup]  DEFAULT ((1)) FOR [minItemQtyPerSampleGroup]
GO
ALTER TABLE [dbo].[inv_items] ADD  CONSTRAINT [DF__inv_reage__datet__7B5B524B]  DEFAULT (getdate()) FOR [datetime_stamp]
GO
