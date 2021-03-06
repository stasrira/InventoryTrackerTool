USE [sl_inventory]
GO
/****** Object:  Table [dbo].[inv_manufacturers]    Script Date: 11/9/2018 5:25:41 PM ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[inv_manufacturers](
	[manufact_code] [varchar](20) NOT NULL,
	[manufact_name] [varchar](100) NOT NULL,
	[manufact_URL] [varchar](150) NULL,
	[manufact_details] [varchar](200) NULL,
	[comments] [varchar](200) NULL,
 CONSTRAINT [PK_inv_manufacturers] PRIMARY KEY CLUSTERED 
(
	[manufact_code] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]
GO
SET ANSI_PADDING ON
GO
/****** Object:  Index [IX_inv_manufacturers_manufat_name]    Script Date: 11/9/2018 5:25:41 PM ******/
CREATE NONCLUSTERED INDEX [IX_inv_manufacturers_manufat_name] ON [dbo].[inv_manufacturers]
(
	[manufact_name] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
GO
