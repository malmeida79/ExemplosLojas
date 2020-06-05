
/*	============================  Tabelas =================================	*/

create table tblEstoque (
	estqId int identity(1,1) primary key not null,
	estqDesc nvarchar(50)
)
GO

create table tblItemEstoque (
	estqItemId int,
	idProd nvarchar(50),
	estqQtde nvarchar(50),
	estqValUnit decimal(10,2)
)
GO

/*	============================  Procedures =================================	*/

create procedure spc_cad_estq @estqDesc nvarchar(50) as
begin
	insert into tblEstoque values (@estqDesc)
end
go

create procedure spc_alt_estq @estqId int, @estqDesc nvarchar(50) as
begin
	update tblEstoque set estqDesc = @estqDesc where estqId = @estqId 
end
go

create procedure spc_del_estq @estqId int as
begin
	delete from tblEstoque where estqId = @estqId 
end
go

create procedure spc_cad_it_estq 
	@estqId int,
	@idProd nvarchar(50),
	@estqQtde nvarchar(50),
	@estqValUnit decimal(10,2)
	as
begin
	insert into tblItemEstoque values (@estqId,@idprod,@estqQtde,@estqValUnit)
end
go

create procedure spc_alt_it_estq 
	@estqItemId int,
	@estqId int,
	@idEstq int,
	@idProd nvarchar(50),
	@estqQtde nvarchar(50),
	@estqValUnit decimal(10,2)
as
begin
	update tblItemEstoque set 
		estqId=@estqId,
		IdProd = @idProd,
		estqQtde = @estqQtde,
		estqValUnit = @estqValUnit
	where 
		estqItemId = @estqItemId 
end
go

create procedure spc_del_it_estq @estqItemId int as
begin
	delete from tblItemEstoque where estqItemId = @estqItemId 
end
go

/*	============================ Relacionamentos =================================	*/

ALTER TABLE [dbo].[tblItemEstoque]  WITH CHECK ADD  CONSTRAINT [FK_itemestoque_grupo] FOREIGN KEY([estqId])
REFERENCES [dbo].[tblEstoque] ([estqId])
ON UPDATE CASCADE
ON DELETE CASCADE
GO 