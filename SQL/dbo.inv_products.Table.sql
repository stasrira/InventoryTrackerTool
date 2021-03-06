USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_products]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_products](
	[productCode] [bigint] NOT NULL,
	[productName] [varchar](80) NOT NULL,
	[productType] [varchar](50) NULL,
	[item_id] [int] NOT NULL,
	[unitsPerProduct] [int] NOT NULL,
	[manufact_code] [varchar](20) NOT NULL,
	[productCatalogNum] [varchar](50) NULL,
	[productURL] [varchar](150) NULL,
	[datetime_stamp] [datetime] NULL,
 CONSTRAINT [PK_inv_products] PRIMARY KEY CLUSTERED 
(
	[productCode] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_inv_products_manufact_code]    Script Date: 11/9/2018 5:25:41 PM ******/
CREATE NONCLUSTERED INDEX [IX_inv_products_manufact_code] ON [dbo].[inv_products]
(
	[manufact_code] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
/****** Object:  Index [IX_inv_products_reagent_id]    Script Date: 11/9/2018 5:25:41 PM ******/
CREATE NONCLUSTERED INDEX [IX_inv_products_reagent_id] ON [dbo].[inv_products]
(
	[item_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_products] ADD  CONSTRAINT [DF__inv_produ__datet__7C4F7684]  DEFAULT (getdate()) FOR [datetime_stamp]
GO
