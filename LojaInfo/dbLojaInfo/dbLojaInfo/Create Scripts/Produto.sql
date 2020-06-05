
/*	============================  Tabelas =================================	*/

create table tblGrpProd (
	grpId int identity(1,1) primary key not null,
	grpDesc nvarchar(50)
)
GO

create table tblSGrpProd (
	sgrpId int identity(1,1) primary key not null,
	sgrpDesc nvarchar(50),
	grpId int
)
GO

create table tblProduto (
	idprod int identity(1,1) primary key,
	descProd nvarchar(50)
)
GO

create table tblItemProduto (
	idItemProd int identity(1,1) primary key,
	idProd int,
	idEstqOrigem int,
	qtdeItemProd int
)
GO

/*	============================  Procedures =================================	*/

create procedure spc_cad_grp @grpDesc nvarchar(50) as
begin
	insert into tblGrpProd values (@grpDesc)
end
go

create procedure spc_alt_grp @grpId int, @grpDesc nvarchar(50) as
begin
	update tblGrpProd set grpDesc = @grpDesc where grpId = @grpId 
end
go

create procedure spc_del_grp @grpId int as
begin
	delete from tblGrpProd where grpId = @grpId 
end
go

create procedure spc_cad_sgrp @sGrpDesc nvarchar(50),@grpId int as
begin
	insert into tblSGrpProd values (@sGrpDesc,@grpId)
end
go

create procedure spc_alt_sgrp @sgrpId int, @grpid int ,@grpDesc nvarchar(50) as
begin
	update tblSGrpProd set sGrpDesc = @grpDesc, grpId = @grpId where sGrpId = @sGrpId 
end
go

create procedure spc_del_sgrp @sGrpId int as
begin
	delete from tblSGrpProd where grpId = @sGrpId 
end
go

create procedure spc_cad_prd @descprod nvarchar(50) as 
begin
	insert into tblProduto values (@descprod)
end
go

create procedure spc_alt_prd @idProd int, @descprod nvarchar(50) as 
begin
	update tblProduto set descprod=@descprod where idprod = @idprod
end
go

create procedure spc_del_prd @idProd int as 
begin
	delete from tblProduto where idprod = @idprod
end
go

create procedure spc_cad_it_prd 
	@idProd int,
	@idSgrp int,
	@idEstqOrigem int,
	@qtdeItemProd int,
	@mensagem nvarchar(100) output
as
begin transaction

	-- declaração de variaveis
	declare @estq int
	
	-- inicialização de variaveis
	set @estq = 0
	set @mensagem = ''

	-- pesquisa no estoque o  item inserido para contar sua quantidade no local especifico.
	set @estq = (select  sum(convert(int,estqQtde)) as [qtde] from tblItemEstoque where idprod = @idprod and tblItemEstoque.estqId = @idEstqOrigem)
	-- se quantridade em estoque maior que zero
	if @estq > 0
		begin
		-- se quanrtidade inserida maior que quantidade em estoque
		if @qtdeItemProd >= @estq
			begin
				-- insere o item na tabela de item produto
				insert into tblItemEstoque values (@idprod,@idSgrp,@idEstqOrigem,@qtdeItemProd)

				-- desconta a quantidade utilizada
				set @estq = @estq - @qtdeItemProd

				-- atualiza a nova quantidade
				update tblItemEstoque set estqQtde = @estq where idprod = @idprod and estqId = @idEstqOrigem

				-- retorna OK
				set @mensagem = 'OK'

				-- fim da trans com sucesso
				commit transaction

				-- fim do comando
				return
			end
		else
			begin
				set @mensagem = 'NOK'
				-- voltar tudo em caso de erros
				rollback transaction
				-- finalizar execução			
				return
			end
		end
	else
		begin
			
			set @mensagem = 'NOK'
			-- voltar tudo em caso de erros
			rollback transaction
			-- finalizar execução
			return
		end		
go

/*	============================ Relacionamentos =================================	*/

ALTER TABLE [dbo].[tblSGrpProd]  WITH CHECK ADD  CONSTRAINT [FK_subgrupo_grupo] FOREIGN KEY([grpId])
REFERENCES [dbo].[tblGrpProd] ([grpId])
ON UPDATE CASCADE
ON DELETE CASCADE
GO