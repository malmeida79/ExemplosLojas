
/*	============================  Tabelas =================================	*/

	create table cadastro_cliente(
		idCliente int identity(1,1) primary key,
		cliNome nvarchar(50),
		cliFantasia nvarchar(50),
		cliCnpj nvarchar(20),
		idEndereco int,
		cliNumero int
	)
	go

	create table cadastro_endereco(
		idEndereco int identity(1,1) primary key,
		endCep nvarchar(8),
		endEndereco nvarchar(100),
		idSigla int,
		idBairro int,
		idCidade int,
		IdEstado int
	)
	go

	create table cadastro_bairro(
		idBairro int identity(1,1) primary key,
		brrNome nvarchar(30)
	)
	go

	create table cadastro_cidade(
		idCidade int identity(1,1) primary key,
		cidNome nvarchar(30)
	)
	go

	create table cadastro_estado(
		idEstado int identity(1,1) primary key,
		estNome nvarchar(30)
	)
	go

	create table cadastro_sigla(
		idSigla int identity(1,1) primary key,
		sgNome nvarchar(30)
	)
	go

	create table tblUser (
		idUser int identity(1,1) primary key,
		logUser nvarchar(12),
		pwdUser nvarchar(12),
		blkUser bit
	)
	go

/*	============================  Procedures =================================	*/

create procedure spc_cad_cliente 
		@cliNome nvarchar(50),
		@cliFantasia nvarchar(50),
		@cliCnpj nvarchar(20),
		@idEndereco int,
		@cliNumero int
as
begin
	insert into cadastro_cliente values (@cliNome,@cliFantasia,@cliCnpj,@idEndereco,@cliNumero)		
end
go

create procedure spc_alt_cliente 
		@idCliente int,
		@cliNome nvarchar(50),
		@cliFantasia nvarchar(50),
		@cliCnpj nvarchar(20),
		@idEndereco int,
		@cliNumero int
as
begin

		update
			 cadastro_cliente
		set 	
			cliNome = @cliNome,
			cliFantasia = @cliFantasia,
			cliCnpj = @cliCnpj,
			idEndereco = @idEndereco,
			cliNumero = @cliNumero
		where 
			idCliente = @idCliente

end
go

create procedure spc_del_cliente 
		@idCliente int		
as
begin

		delete from
			 cadastro_cliente
		where 
			idCliente = @idCliente

end
go

create procedure spc_cad_endereco
		@endCep nvarchar(8),
		@idSigla int,
		@endEndereco nvarchar(100),
		@idBairro int,
		@idCidade int,
		@IdEstado int,
		@codigo int output
as
begin
	insert into cadastro_endereco values(@endCep,@endEndereco,@idSigla,@idBairro,@idCidade,@IdEstado)
	select @codigo = @@identity
	return @codigo
end
go

create procedure spc_alt_endereco
		@idEndereco int,
		@endCep nvarchar(8),
		@idSigla int,
		@endEndereco nvarchar(100),
		@idBairro int,
		@idCidade int,
		@IdEstado int
as
begin
		update 
			cadastro_endereco
		set
			endCep=@endCep,
			idSigla=@idSigla,
			endEndereco=@endEndereco,
			idBairro=@idBairro,
			idCidade=@idCidade,
			IdEstado=@IdEstado
		where
			idEndereco = @idEndereco
end
go

create procedure spc_del_endereco
		@idEndereco int
as
begin
		delete from
			cadastro_endereco
		where
			idEndereco = @idEndereco
end
go

create procedure spc_cad_bairro 
		@brrNome nvarchar(30)
as
begin
	insert into cadastro_bairro values (@brrNome)		
end
go

create procedure spc_cad_cidade 
		@cidNome nvarchar(30)
as
begin
	insert into cadastro_cidade values (@cidNome)		
end
go

create procedure spc_cad_estado 
		@estNome nvarchar(30)
as
begin
	insert into cadastro_estado values (@estNome)		
end
go

create procedure spc_cad_sigla
		@sgNome nvarchar(30)
as
begin
	insert into cadastro_sigla values (@sgNome)		
end
go

create procedure spc_alt_bairro 
		@idBairro int,
		@brrNome nvarchar(30)
as
begin
	update cadastro_bairro set brrNome=@brrNome where idBairro = @idBairro
end
go

create procedure spc_alt_cidade 
		@idCidade int,
		@cidNome nvarchar(30)
as
begin
	update cadastro_cidade set cidNome=@cidNome where idCidade = @idCidade
end
go

create procedure spc_alt_estado 
		@idEstado int,
		@estNome nvarchar(30)
as
begin
	update cadastro_estado set estNome=@estNome where idestado = @idEstado
end
go

create procedure spc_alt_sigla 
		@idSigla int,
		@sgNome nvarchar(30)
as
begin
	update cadastro_sigla set sgNome=@sgNome where idSigla = @idSigla
end
go

create procedure spc_del_bairro 
		@idBairro int
as
begin
	delete from cadastro_bairro where idBairro = @idBairro
end
go

create procedure spc_del_cidade 
		@idCidade int
as
begin
	delete from cadastro_cidade where idCidade = @idCidade
end
go

create procedure spc_del_estado 
		@idEstado int
as
begin
	delete from cadastro_estado where idEstado = @idEstado
end
go

create procedure spc_del_sigla 
		@idSigla int
as
begin
	delete from cadastro_sigla where idSigla = @idSigla
end
go

/*	============================ Relacionamentos =================================	*/

ALTER TABLE [dbo].[cadastro_endereco]  WITH CHECK ADD  CONSTRAINT [FK_cadastro_endereco_cadastro_bairro] FOREIGN KEY([idBairro])
REFERENCES [dbo].[cadastro_bairro] ([idBairro])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[cadastro_endereco]  WITH CHECK ADD  CONSTRAINT [FK_cadastro_endereco_cadastro_cidade] FOREIGN KEY([idCidade])
REFERENCES [dbo].[cadastro_cidade] ([idCidade])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[cadastro_endereco]  WITH CHECK ADD  CONSTRAINT [FK_cadastro_endereco_cadastro_estado] FOREIGN KEY([IdEstado])
REFERENCES [dbo].[cadastro_estado] ([idEstado])
ON UPDATE CASCADE
ON DELETE CASCADE
GO

ALTER TABLE [dbo].[cadastro_endereco]  WITH CHECK ADD  CONSTRAINT [FK_cadastro_endereco_cadastro_sigla] FOREIGN KEY([idSigla])
REFERENCES [dbo].[cadastro_sigla] ([idSigla])
ON UPDATE CASCADE
ON DELETE CASCADE
go

ALTER TABLE [dbo].[cadastro_cliente]  WITH CHECK ADD  CONSTRAINT [FK_cadastro_cliente_cadastro_endereco] FOREIGN KEY([idEndereco])
REFERENCES [dbo].[cadastro_endereco] ([idEndereco])
ON UPDATE CASCADE
ON DELETE CASCADE
go

/*	============================ Views =================================	*/

CREATE VIEW [dbo].[VW_ENDERECO]
AS
SELECT     dbo.cadastro_endereco.idEndereco, dbo.cadastro_sigla.sgNome, dbo.cadastro_endereco.endEndereco, dbo.cadastro_endereco.endCep, 
                      dbo.cadastro_bairro.brrNome, dbo.cadastro_cidade.cidNome, dbo.cadastro_estado.estNome
FROM         dbo.cadastro_bairro INNER JOIN
                      dbo.cadastro_endereco ON dbo.cadastro_bairro.idBairro = dbo.cadastro_endereco.idBairro INNER JOIN
                      dbo.cadastro_cidade ON dbo.cadastro_endereco.idCidade = dbo.cadastro_cidade.idCidade INNER JOIN
                      dbo.cadastro_estado ON dbo.cadastro_endereco.IdEstado = dbo.cadastro_estado.idEstado INNER JOIN
                      dbo.cadastro_sigla ON dbo.cadastro_endereco.idSigla = dbo.cadastro_sigla.idSigla
go