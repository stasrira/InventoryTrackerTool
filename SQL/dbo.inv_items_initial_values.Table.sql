USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_items_initial_values]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_items_initial_values](
	[item_id] [int] NOT NULL,
	[initial_amount] [decimal](8, 2) NULL,
	[initial_date] [datetime] NULL,
	[datetime_stamp] [datetime] NULL,
 CONSTRAINT [PK_inv_items_initial_values] PRIMARY KEY CLUSTERED 
(
	[item_id] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
ALTER TABLE [dbo].[inv_items_initial_values] ADD  CONSTRAINT [DF__inv_items__datet__3F115E1A]  DEFAULT (getdate()) FOR [datetime_stamp]
GO
