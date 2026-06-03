Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Disponibilidade
    Inherits BasePage

    Private objDisponibilidade As [Lib].Negocio.Disponibilidade
    Private objListDisp As [Lib].Negocio.ListDisponibilidade
    Private Ano As Integer = Now.Year

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("DISPONIBILIDADE", "ACESSAR") Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
                Funcoes.VerificaEmpresa(ddlEmpresa)
                ddl.Carregar(ddlEmpresaConfiguracao, CarregarDDL.Tabela.Empresas, "")
                Funcoes.VerificaEmpresa(ddlEmpresaConfiguracao)
                ddl.Carregar(ddlSafraDe, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(ddlSafraAte, CarregarDDL.Tabela.Safra, "")
                txtDataDeEmissao.Text = Date.Now.ToString("dd/MM/yyyy")
                HID.Value = Guid.NewGuid().ToString
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissã para acessar essa paginá.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
            ddlEmpresaConfiguracao.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objClienteDisp" & HID.Value) IsNot Nothing Then
            Dim cli As Cliente = CType(Session("objClienteDisp" & HID.Value), [Lib].Negocio.Cliente)
            hdfCodigoCliente.Value = cli.Codigo & "-" & cli.CodigoEndereco
            Funcoes.FormatarClienteTXT(txtNomeCliente, cli)
            Session.Remove("objClienteDisp" & HID.Value)
        End If
    End Sub

#Region "Sessão"
    Public Sub SessaoSalvarDisponibilidade()
        Session("objDisponibilidade") = objDisponibilidade
    End Sub

    Public Sub SessaoRecuperarDisponibilidade()
        If Session("objDisponibilidade") Is Nothing Then
            objDisponibilidade = New [Lib].Negocio.Disponibilidade
        Else
            objDisponibilidade = Session("objDisponibilidade")
        End If
    End Sub

    Public Sub SessaoSalvarListaDisponibilidade()
        Session("objListDisp") = objListDisp
    End Sub

    Public Sub SessaoRecuperarListaDisponibilidade()
        objListDisp = Session("objListDisp")
    End Sub
#End Region

    Public Function ValidaRelatorio() As Boolean
        If Not ddlEmpresa.SelectedIndex > 0 Then
            MsgBox(Me.Page, "Informe a empresa")
            Return False
        End If

        If Not IsDate(txtDataDeEmissao.Text) Then
            MsgBox(Me.Page, "Informe a Data da Disponibilidade")
            Return False
        End If

        If Not ddlSafraDe.SelectedIndex > 0 Then
            MsgBox(Me.Page, "Informe a safra Inicial")
            Return False
        End If

        If Not ddlSafraAte.SelectedIndex > 0 Then
            MsgBox(Me.Page, "Informe a safra Final")
            Return False
        End If

        Return True
    End Function

    Private Function getDataSet(ByVal NumNaLista As Integer) As DataSet
        SessaoRecuperarListaDisponibilidade()
        objDisponibilidade = objListDisp(NumNaLista)

        Dim str As String
        Dim pos As New [Lib].Negocio.ListPosicaoDePedido

        Dim listEmpresa As New List(Of String)
        listEmpresa.Add(objDisponibilidade.CodigoEmpresa & objDisponibilidade.EndEmpresa)

        str = pos.SqlPosicao(txtDataDeEmissao.Text, listEmpresa, objDisponibilidade.Consolidado, Nothing, False, "#Base", objDisponibilidade.getSqlProduto(), ddlSafraDe.SelectedValue, ddlSafraAte.SelectedValue)

        str &= "Select '1' as Disparador," & vbCrLf & _
               "       P.Safra," & vbCrLf & _
               "       case" & vbCrLf & _
               "		 when ((OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe  =  '" & eClassesOperacoes.AFIXAR.ToString & "') or so.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "') and teste.id = 1  then '1 - Compras a Fixar'" & vbCrLf & _
               "		 when   OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.AFIXAR.ToString & "')                                     then '2 - Compras'" & vbCrLf & _
               "         when   OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe  =  '" & eClassesOperacoes.AFIXAR.ToString & "'                                                                                   then '3 - Vendas a Fixar'" & vbCrLf & _
               "		 when   OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')  then '4 - Vendas'" & vbCrLf & _
               "	   end as grupo," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		 when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'CIF' then isnull(ent.Cidade,'NAO INFORMADO') + ' - ' + isnull(ent.Estado,'DEPOSITO DE OD')" & vbCrLf & _
               "		 when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'FOB' then b.CidadeEmpresa + ' - ' + b.EstadoEmpresa" & vbCrLf & _
               "		 else ''" & vbCrLf & _
               "	   end Municipio," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 1                                                                    then 'Compras a Fixar'" & vbCrLf & _
               "		 when SO.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "'                                                                                                                                               then 'Compras a Fixar em Deposito'" & vbCrLf & _
               "         When OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 2                                                                    then 'Compras a Fixar a Receber'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(P.Antecipada,0) = 1                                                     then 'Compras Antecipada'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(P.Antecipada,0) = 0 and isnull(P.Troca,0) = 1                           then 'Compras a Receber P/ Troca'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'CIF' and isnull(P.Antecipada,0) = 0 and isnull(P.Troca,0) = 0 then 'Compras a Receber CIF'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'FOB'                                                          then 'Compras a Receber FOB'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')                                             then 'Vendas a Entregar'" & vbCrLf & _
               "         when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'                                                                                     then 'Vendas a Fixar'" & vbCrLf & _
               "	   end as Descricao," & vbCrLf & _
               "	   sum(case" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 1                                                                    then aFixar" & vbCrLf & _
               "			 when SO.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "'                                                                                                            then aFixar" & vbCrLf & _
               "             When OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 2                                                                    then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(p.Antecipada,0) = 1                                                     then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(p.Antecipada,0) = 0 and isnull(p.troca,0) = 1                           then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'CIF' and isnull(p.Antecipada,0) = 0 and isnull(p.troca,0) = 0 then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'FOB'                                                          then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')                                             then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'                                                                                     then aEntregar" & vbCrLf & _
               "		   end) as Valor" & vbCrLf & _
               "  from #base b" & vbCrLf & _
               " inner join Operacoes OP" & vbCrLf & _
               "    on OP.Operacao_id     = b.Operacao" & vbCrLf & _
               " inner join SubOperacoes SO" & vbCrLf & _
               "    on SO.Operacao_id     = b.Operacao" & vbCrLf & _
               "   and SO.SubOperacoes_id = b.SubOperacao" & vbCrLf & _
               " Inner Join Pedidos P" & vbCrLf & _
               "    on P.Empresa_Id    = b.Empresa" & vbCrLf & _
               "   and P.EndEmpresa_Id = b.EndEmpresa" & vbCrLf & _
               "   and P.Pedido_Id     = b.Pedido" & vbCrLf & _
               "  Left Join PedidosXDepositos PxD" & vbCrLf & _
               "    on PxD.Empresa_Id    = P.Empresa_id" & vbCrLf & _
               "   and PxD.EndEmpresa_Id = P.EndEmpresa_id" & vbCrLf & _
               "   and PxD.Pedido_Id     = P.Pedido_id" & vbCrLf & _
               "   and PXD.Tipo          = 'OD'" & vbCrLf & _
               "  Left Join Clientes as Ent" & vbCrLf & _
               "    on pxd.Deposito_Id    = Ent.Cliente_Id" & vbCrLf & _
               "   and pxd.EndDeposito_Id = Ent.Endereco_Id" & vbCrLf & _
               "  Inner join (" & vbCrLf & _
               "               Select 1 ID" & vbCrLf & _
               "                union" & vbCrLf & _
               "               Select 2" & vbCrLf & _
               "              ) teste" & vbCrLf & _
               "     on 1 = 1" & vbCrLf & _
               " Where SO.classe in ('" & eClassesOperacoes.AFIXAR.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
               "   and (Teste.id = 1 or (OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 2))" & vbCrLf & _
               " Group by" & vbCrLf & _
               "       P.Safra," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		 when ((OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "') or so.Classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "') and teste.id = 1  then '1 - Compras a Fixar'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe in ('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.AFIXAR.ToString & "')                                     then '2 - Compras'" & vbCrLf & _
               "         when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'                                                  then '3 - Vendas a Fixar'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')          then '4 - Vendas'" & vbCrLf & _
               "	   end," & vbCrLf & _
               "	   case " & vbCrLf & _
               "		 when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'CIF' then isnull(ent.Cidade,'NAO INFORMADO') + ' - ' + isnull(ent.Estado,'DEPOSITO DE OD')" & vbCrLf & _
               "		 when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'FOB' then b.CidadeEmpresa + ' - ' + b.EstadoEmpresa" & vbCrLf & _
               "		 else ''" & vbCrLf & _
               "	   end," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 1                                                                    then 'Compras a Fixar'" & vbCrLf & _
               "		 when SO.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "'                                                                                                            then 'Compras a Fixar em Deposito'" & vbCrLf & _
               "         When OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 2                                                                    then 'Compras a Fixar a Receber'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(P.Antecipada,0) = 1                                                     then 'Compras Antecipada'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(P.Antecipada,0) = 0 and isnull(P.Troca,0) = 1                           then 'Compras a Receber P/ Troca'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'CIF' and isnull(P.Antecipada,0) = 0 and isnull(P.Troca,0) = 0 then 'Compras a Receber CIF'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'FOB'                                                          then 'Compras a Receber FOB'" & vbCrLf & _
               "		 when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')                                             then 'Vendas a Entregar'" & vbCrLf & _
               "         when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'                                                                                     then 'Vendas a Fixar'" & vbCrLf & _
               "	   end " & vbCrLf & _
               "having sum(case" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 1                                                                    then aFixar" & vbCrLf & _
               "			 when SO.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "'                                                                                                            then aFixar" & vbCrLf & _
               "             When OP.Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "' and teste.id = 2                                                                    then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(p.Antecipada,0) = 1                                                     then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and isnull(p.Antecipada,0) = 0 and isnull(p.troca,0) = 1                           then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'CIF' and isnull(p.Antecipada,0) = 0 and isnull(p.troca,0) = 0 then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.COMPRAS.ToString & "' and SO.classe = '" & eClassesOperacoes.COMPRAS.ToString & "' and b.FreteCifFob = 'FOB'                                                          then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')                                             then aEntregar" & vbCrLf & _
               "			 when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'                                                                                     then aEntregar" & vbCrLf & _
               "		   end) > 0" & vbCrLf


        ''***********************************************************************************************
        ''*****************  EXPLOSAO VENDAS A ENTREGAR *************************************************
        ''***********************************************************************************************
        str &= "Select P.Safra," & vbCrLf & _
               "       b.Cliente as Cliente_id," & vbCrLf & _
               "       b.EndCliente as Endereco_id," & vbCrLf & _
               "       B.NomeCliente as Nome," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		  when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'CIF' then isnull(ent.Cidade,'NAO INFORMADO') + ' - ' + isnull(ent.Estado,'DEPOSITO DE OD')" & vbCrLf & _
               "		  when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'FOB' then b.CidadeEmpresa + ' - ' + b.EstadoEmpresa" & vbCrLf & _
               "		  else ''" & vbCrLf & _
               "	   end Municipio," & vbCrLf & _
               "	   sum(case" & vbCrLf & _
               "			  when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "') then aEntregar" & vbCrLf & _
               "			  else 0" & vbCrLf & _
               "		    end) as Valor," & vbCrLf & _
               "	    b.UnitarioOficial," & vbCrLf & _
               "	    sum(case" & vbCrLf & _
               "	  	      when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "') then aEntregar" & vbCrLf & _
               "			  else 0" & vbCrLf & _
               "		    end) * b.UnitarioOficial as ValorTotal," & vbCrLf & _
               "	    b.DataEntrega" & vbCrLf & _
               "  from #base b" & vbCrLf & _
               " inner join Operacoes OP" & vbCrLf & _
               "    on OP.Operacao_id     = b.Operacao" & vbCrLf & _
               " inner join SubOperacoes SO" & vbCrLf & _
               "    on SO.Operacao_id     = b.Operacao" & vbCrLf & _
               "   and SO.SubOperacoes_id = b.SubOperacao" & vbCrLf & _
               " Inner Join Pedidos P" & vbCrLf & _
               "    on P.Empresa_Id    = b.Empresa" & vbCrLf & _
               "   and P.EndEmpresa_Id = b.EndEmpresa" & vbCrLf & _
               "   and P.Pedido_Id     = b.Pedido" & vbCrLf & _
               "  Left Join PedidosXDepositos PxD" & vbCrLf & _
               "    on PxD.Empresa_Id    = P.Empresa_id" & vbCrLf & _
               "   and PxD.EndEmpresa_Id = P.EndEmpresa_id" & vbCrLf & _
               "   and PxD.Pedido_Id     = P.Pedido_id" & vbCrLf & _
               "   and PXD.Tipo          = 'OD'" & vbCrLf & _
               "  Left Join Clientes as Ent" & vbCrLf & _
               "    on pxd.Deposito_Id    = Ent.Cliente_Id" & vbCrLf & _
               "   and pxd.EndDeposito_Id = Ent.Endereco_Id" & vbCrLf & _
               " Where SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
               " Group by" & vbCrLf & _
               "       P.Safra," & vbCrLf & _
               "       b.Cliente," & vbCrLf & _
               "	   b.EndCliente," & vbCrLf & _
               "	   B.NomeCliente," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		 when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'CIF' then isnull(ent.Cidade,'NAO INFORMADO') + ' - ' + isnull(ent.Estado,'DEPOSITO DE OD')" & vbCrLf & _
               "		  when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'FOB' then b.CidadeEmpresa + ' - ' + b.EstadoEmpresa" & vbCrLf & _
               "		  else ''" & vbCrLf & _
               "	    end," & vbCrLf & _
               "	    b.UnitarioOficial," & vbCrLf & _
               "       b.DataEntrega" & vbCrLf & _
               " having sum(case" & vbCrLf & _
               "			   when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "') then aEntregar" & vbCrLf & _
               "			   else 0" & vbCrLf & _
               "		    end) > 0" & vbCrLf & _
               " order by 5,9" & vbCrLf

        ''**************************************************************************************************
        ''*****************  RESUMO VENDAS SAFRA MUNICIPIO *************************************************
        ''**************************************************************************************************
        str &= "Select P.Safra," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		  when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'CIF' then isnull(ent.Cidade,'NAO INFORMADO') + ' - ' + isnull(ent.Estado,'DEPOSITO DE OD')" & vbCrLf & _
               "		  when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'FOB' then b.CidadeEmpresa + ' - ' + b.EstadoEmpresa" & vbCrLf & _
               "		  else ''" & vbCrLf & _
               "	    end Municipio," & vbCrLf & _
               "	    sum(case" & vbCrLf & _
               "			  when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
               "                then Contratado" & vbCrLf & _
               " 			  when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
               "                then Fixada" & vbCrLf & _
               "			  else 0" & vbCrLf & _
               "		    end) as Qtde," & vbCrLf & _
               "	    sum(case" & vbCrLf & _
               "	  	      when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "')" & vbCrLf & _
               "                then Contratado  * b.UnitarioOficial" & vbCrLf & _
               "              when OP.CLASSE = '" & eClassesOperacoes.VENDAS.ToString & "'  and SO.classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
               "                then ValorFixadoOficial " & vbCrLf & _
               "			  else 0" & vbCrLf & _
               "		    end) as ValorTotal" & vbCrLf & _
               "  from #base b" & vbCrLf & _
               " inner join Operacoes OP" & vbCrLf & _
               "    on OP.Operacao_id     = b.Operacao" & vbCrLf & _
               " inner join SubOperacoes SO" & vbCrLf & _
               "    on SO.Operacao_id     = b.Operacao" & vbCrLf & _
               "   and SO.SubOperacoes_id = b.SubOperacao" & vbCrLf & _
               " Inner Join Pedidos P" & vbCrLf & _
               "    on P.Empresa_Id    = b.Empresa" & vbCrLf & _
               "   and P.EndEmpresa_Id = b.EndEmpresa" & vbCrLf & _
               "   and P.Pedido_Id     = b.Pedido" & vbCrLf & _
               "  Left Join PedidosXDepositos PxD" & vbCrLf & _
               "    on PxD.Empresa_Id    = P.Empresa_id" & vbCrLf & _
               "   and PxD.EndEmpresa_Id = P.EndEmpresa_id" & vbCrLf & _
               "   and PxD.Pedido_Id     = P.Pedido_id" & vbCrLf & _
               "   and PXD.Tipo          = 'OD'" & vbCrLf & _
               "  Left Join Clientes as Ent" & vbCrLf & _
               "    on pxd.Deposito_Id    = Ent.Cliente_Id" & vbCrLf & _
               "   and pxd.EndDeposito_Id = Ent.Endereco_Id" & vbCrLf & _
               " Where SO.classe in ('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "','" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.AFIXAR.ToString & "')" & vbCrLf & _
               " Group by" & vbCrLf & _
               "       P.Safra," & vbCrLf & _
               "	   case" & vbCrLf & _
               "		 when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'CIF' then isnull(ent.Cidade,'NAO INFORMADO') + ' - ' + isnull(ent.Estado,'DEPOSITO DE OD')" & vbCrLf & _
               "		 when OP.classe = '" & eClassesOperacoes.VENDAS.ToString & "' and b.FreteCifFob = 'FOB' then b.CidadeEmpresa + ' - ' + b.EstadoEmpresa" & vbCrLf & _
               "	     else ''" & vbCrLf & _
               "	   end" & vbCrLf & _
               " order by 1,2" & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(str, "Consulta")
        ds.Tables(0).TableName = "PosicaoPedido"
        ds.Tables(1).TableName = "ExplosaoVendas"
        ds.Tables(2).TableName = "ResumoSafraMunicipioVendas"

        ds.Merge(Banco.ConsultaDataSet(sqlDisponibilidade(1, objDisponibilidade.getSqlProduto()), "DisponibilidadeDia"))
        ds.Merge(Banco.ConsultaDataSet(sqlDisponibilidade(2, objDisponibilidade.getSqlProduto()), "DisponibilidadeAno"))

        Dim Achou As Boolean
        For Each rowAno As DataRow In ds.Tables("DisponibilidadeAno").Rows
            Achou = False
            For Each rowDia As DataRow In ds.Tables("DisponibilidadeDia").Rows
                If rowDia("Deposito") = rowAno("Deposito") And rowDia("EndDeposito") = rowAno("EndDeposito") Then
                    Achou = True
                    Exit For
                End If
            Next

            If Not Achou Then
                Dim novaLinhaDia As DataRow = ds.Tables("DisponibilidadeDia").NewRow()
                novaLinhaDia("Disparador") = rowAno("Disparador")
                novaLinhaDia("Empresa") = rowAno("Empresa")
                novaLinhaDia("EndEmpresa") = rowAno("EndEmpresa")
                novaLinhaDia("NomeEmpresa") = rowAno("NomeEmpresa")
                novaLinhaDia("CidadeEmpresa") = rowAno("CidadeEmpresa")
                novaLinhaDia("EstadoEmpresa") = rowAno("EstadoEmpresa")
                novaLinhaDia("Deposito") = rowAno("Deposito")
                novaLinhaDia("EndDeposito") = rowAno("EndDeposito")
                novaLinhaDia("NomeDeposito") = rowAno("NomeDeposito")
                novaLinhaDia("CidadeDeposito") = rowAno("CidadeDeposito")
                novaLinhaDia("EstadoDeposito") = rowAno("EstadoDeposito")
                novaLinhaDia("Inicial") = 0
                novaLinhaDia("LaudoSemNota") = 0
                novaLinhaDia("LaudoSemNotaSaida") = 0
                novaLinhaDia("EntradaProducao") = 0
                novaLinhaDia("SaidaProducao") = 0
                novaLinhaDia("Recebimento") = 0
                novaLinhaDia("Fob") = 0
                novaLinhaDia("CompraPorto") = 0
                novaLinhaDia("CompraEmDeposito") = 0
                novaLinhaDia("EntradaTransferenciaFilial") = 0
                novaLinhaDia("EntradaTransferenciaTitularidade") = 0
                novaLinhaDia("EntradaDevolucaoVenda") = 0
                novaLinhaDia("EntradaDeposito") = 0
                novaLinhaDia("RetornoSimbolicoFormacaoDeLote") = 0
                novaLinhaDia("RetornoDeDeposito") = 0
                novaLinhaDia("RetiradaPrestacaoServico") = 0
                novaLinhaDia("Retirada") = 0
                novaLinhaDia("RetiradaFob") = 0
                novaLinhaDia("RetiradaFobTer") = 0
                novaLinhaDia("EmbarquePorto") = 0
                novaLinhaDia("SaidaTransferenciaFilial") = 0
                novaLinhaDia("SaidaDeposito") = 0
                novaLinhaDia("SaidaTransferenciaTitularidade") = 0
                novaLinhaDia("DevParaCompra") = 0
                novaLinhaDia("DevDeCompra") = 0
                novaLinhaDia("DevDeposito") = 0
                novaLinhaDia("AcobertamentoFiscal") = 0
                novaLinhaDia("DevAcobertamentoFiscal") = 0
                novaLinhaDia("Quebras") = 0
                novaLinhaDia("Sobras") = 0
                novaLinhaDia("SobrasCliFor") = 0
                novaLinhaDia("EstoqueAtual") = 0
                If rowAno("EndDeposito") = 99 Then
                    ds.Tables("DisponibilidadeDia").Rows.InsertAt(novaLinhaDia, 0)
                Else
                    ds.Tables("DisponibilidadeDia").Rows.Add(novaLinhaDia)
                End If

            End If
        Next

        Return ds
    End Function

    Public Sub Disponibilidade(ByVal pdf As Boolean, ByVal NumNaLista As Integer)
        Try
            If Not ValidaRelatorio() Then Exit Sub

            Dim pEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim Empresa As New [Lib].Negocio.Cliente(pEmpresa(0), pEmpresa(1))
            Dim ds As DataSet = getDataSet(NumNaLista)

            Dim estoque As Decimal
            For Each row As DataRow In ds.Tables("DisponibilidadeAno").Rows
                If row("EndDeposito") <> 99 Then
                    estoque += row("EstoqueAtual")
                End If
            Next

            Dim Parametros As String = ""
            Parametros = "Disponibilidade do dia " & txtDataDeEmissao.Text & vbCrLf & _
                         "Safra de: " & ddlSafraDe.SelectedItem.Text & " a " & ddlSafraAte.SelectedItem.Text & vbCrLf & _
                         "Produto(s): " & objDisponibilidade.getSqlProduto()

            Dim param As New Dictionary(Of String, Object)
            param.Add("Nome", Empresa.Nome)
            param.Add("Cidade", Empresa.Cidade & " - " & Empresa.CodigoEstado & " - CNPJ: " & Funcoes.FormatarCpfCnpj(Empresa.Codigo))
            param.Add("Parametros", Parametros)
            param.Add("Estoque", estoque)

            Funcoes.BindReport(Me.Page, ds, "Cr_Disponibilidade", IIf(pdf, eExportType.PDF, eExportType.ExcelCrystal), param, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try


    End Sub

    Public Function sqlDisponibilidade(ByVal tipo As Integer, ByVal Produtos As String) As String
        Dim sql As String = ""

        sql = "	Declare " & vbCrLf & _
              "	@data nvarchar(10)," & vbCrLf & _
              "	@dataCargaPatio nvarchar(10)," & vbCrLf & _
              " @Empresa nvarchar(14)," & vbCrLf & _
              " @EndEmpresa int," & vbCrLf & _
              " @Ano int" & vbCrLf & _
              "	set @data           ='" & CDate(txtDataDeEmissao.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "	set @Ano            = " & Ano & vbCrLf & _
              "	set @dataCargaPatio ='" & objDisponibilidade.DataInicialCargaPatio.ToString("yyyy-MM-dd") & "'" & vbCrLf


        If objDisponibilidade.Consolidado Then
            sql &= " set @Empresa    ='" & objDisponibilidade.CodigoEmpresa.Substring(0, 8) & "'" & vbCrLf & _
                   " set @EndEmpresa = " & objDisponibilidade.EndEmpresa & vbCrLf
        Else
            sql &= " set @Empresa    ='" & objDisponibilidade.CodigoEmpresa & "'" & vbCrLf & _
                   " set @EndEmpresa = " & objDisponibilidade.EndEmpresa & vbCrLf
        End If

        sql &= "	CREATE TABLE #Disponibilidade " & vbCrLf & _
              "	( " & vbCrLf & _
              "		Empresa                        nvarchar(18) default '', " & vbCrLf & _
              "		EndEmpresa                     int default 0," & vbCrLf & _
              "		Deposito                       nvarchar(18) default '', " & vbCrLf & _
              "		EndDeposito                    int default 0," & vbCrLf & _
              "		LaudoSemNota                   numeric(18,0) default 0," & vbCrLf & _
              "     LaudoSemNotaSaida              numeric(18,0) default 0," & vbCrLf & _
              "		Inicial                        numeric(18,0) default 0," & vbCrLf & _
              "		EntradaProducao                numeric(18,0) default 0," & vbCrLf & _
              "		Recebimento                    numeric(18,0) default 0," & vbCrLf & _
              "		CompraEmDeposito               numeric(18,0) default 0," & vbCrLf & _
              "		Fob                            numeric(18,0) default 0," & vbCrLf & _
              "		CompraPorto                    numeric(18,0) default 0," & vbCrLf & _
              "		EntradaTransferenciaFilial     numeric(18,0) default 0," & vbCrLf & _
              "     EntradaTransferenciaTitularidade numeric(18,0) default 0," & vbCrLf & _
              "		EntradaDevolucaoVenda          numeric(18,0) default 0," & vbCrLf & _
              "		EntradaDeposito                numeric(18,0) default 0," & vbCrLf & _
              "		RetornoSimbolicoFormacaoDeLote numeric(18,0) default 0," & vbCrLf & _
              "     RetornoDeDeposito              numeric(18,0) default 0," & vbCrLf & _
              "		SaidaProducao                  numeric(18,0) default 0," & vbCrLf & _
              "		Retirada                       numeric(18,0) default 0," & vbCrLf & _
              "		RetiradaFob                    numeric(18,0) default 0," & vbCrLf & _
              "     RetiradaFobTer                 numeric(18,0) default 0," & vbCrLf & _
              "		RetiradaPrestacaoServico       numeric(18,0) default 0," & vbCrLf & _
              "		EmbarquePorto                  numeric(18,0) default 0," & vbCrLf & _
              "		SaidaTransferenciaFilial       numeric(18,0) default 0," & vbCrLf & _
              "		SaidaDeposito                  numeric(18,0) default 0," & vbCrLf & _
              "		SaidaTransferenciaTitularidade numeric(18,0) default 0," & vbCrLf & _
              "		DevParaCompra                  numeric(18,0) default 0," & vbCrLf & _
              "		DevDeCompra                    numeric(18,0) default 0," & vbCrLf & _
              "		DevDeposito                    numeric(18,0) default 0," & vbCrLf & _
              "		AcobertamentoFiscal            numeric(18,0) default 0," & vbCrLf & _
              "		DevAcobertamentoFiscal         numeric(18,0) default 0," & vbCrLf & _
              "		Quebras                        numeric(18,0) default 0," & vbCrLf & _
              "		Sobras                         numeric(18,0) default 0," & vbCrLf & _
              "     SobrasCliFor                   numeric(18,0) default 0 " & vbCrLf & _
              "	);" & vbCrLf

        '/*****************************************************************************/ 
        '/******************* Entradas ************************************************/ 
        '/*****************************************************************************/ 
        sql &= " insert into #Disponibilidade(Empresa, EndEmpresa, Deposito, EndDeposito, Recebimento,EntradaDeposito, CompraEmDeposito, Fob,CompraPorto,EntradaTransferenciaFilial,EntradaDevolucaoVenda,RetornoSimbolicoFormacaoDeLote,RetornoDeDeposito, Quebras, Sobras, EntradaTransferenciaTitularidade)" & vbCrLf & _
               "	 (" & vbCrLf & _
               "	   SELECT NF.Empresa_Id," & vbCrLf & _
               "			  NF.EndEmpresa_Id," & vbCrLf & _
               "			  NFxI.deposito," & vbCrLf & _
               "			  NFxI.EndDeposito," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and NF.Finalidade not in (17,23)                and RxP.Pesagem_id is not null                                                                                                                     THEN R.PesoLiquido ELSE 0 END),0) AS Recebimento," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and NF.Finalidade = 23                          and RxP.Pesagem_id is not null                                                                                                                     THEN R.PesoLiquido ELSE 0 END),0) AS EntradaDeposito," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and NF.Finalidade not in (17,23)                and RxP.Pesagem_id is null     and isnull(nf.troca,0) = 0                                                                                          THEN R.PesoLiquido ELSE 0 END),0) AS CompraEmDeposito," & vbCrLf & _
               "              isnull(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and SO.Deposito = 'N'                         and isnull(nf.troca,0) = 1                                                                                          THEN R.PesoLiquido ELSE 0 END),0) as Fob," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and NF.Finalidade  = 21                                                                                                                                                                            THEN R.PesoLiquido ELSE 0 END),0) AS CompraPorto," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe  = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and isnull(nf.troca,0) = 0  and (left(NF.Empresa_id,8) = left(NF.Cliente_id,8) and Nf.Empresa_id <> NF.Cliente_Id)  THEN R.PesoLiquido ELSE 0 END),0) AS EntradaTransferenciaFilial," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.Classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'      /*and isnull(nf.troca,0) = 0*/                                                                                      THEN R.PesoLiquido ELSE 0 END),0) AS EntradaDevolucaoVenda," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      and NF.Finalidade  = 7           and SO.deposito = 'S'                                                                                                                                             THEN R.PesoLiquido ELSE 0 END),0) AS RetornoSimbolicoFormacaoDeLote," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      and NF.Finalidade  in (9,25)     and SO.deposito = 'S'                                                                                                                                             THEN R.PesoLiquido ELSE 0 END),0) AS RetornoDeDeposito," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.Classe in ('" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "')            and NF.Finalidade  = 16                                                                                                                                                                            THEN R.PesoLiquido ELSE 0 END),0) AS Quebras," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe in ('" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.AFIXAR.ToString & "')  and NF.Finalidade  = 17                                                                                                                                                                            THEN R.PesoLiquido ELSE 0 END),0) AS Sobras," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N' and SO.Classe in ('" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.AFIXAR.ToString & "')  and Nf.finalidade  = 23                         and RxP.Pesagem_id is null                                                                                                                         THEN R.PesoLiquido ELSE 0 END),0) AS EntradaTransferenciaTitularida" & vbCrLf & _
               "		 FROM NotasFiscais NF" & vbCrLf & _
               "		Inner Join NotasFiscaisxItens NFxI" & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
               "		Inner JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxR.Empresa_Id" & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxR.Serie_Id" & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxR.Nota_Id" & vbCrLf & _
               "		INNER JOIN Romaneios R" & vbCrLf & _
               "		   ON R.Empresa_Id    = NFxR.Empresa_Id " & vbCrLf & _
               "		  AND R.EndEmpresa_Id = NFxR.EndEmpresa_Id " & vbCrLf & _
               "		  AND R.Romaneio_Id   = NFxR.Romaneio_Id" & vbCrLf & _
               "         LEFT join RomaneiosxPesagens RxP" & vbCrLf & _
               "           ON RxP.Empresa_Id    = NFxR.Empresa_Id" & vbCrLf & _
               "          AND RxP.EndEmpresa_Id = NFxR.EndEmpresa_Id" & vbCrLf & _
               "          AND RxP.Romaneio_Id   = NFxR.Romaneio_Id" & vbCrLf & _
               "		INNER JOIN SubOperacoes SO" & vbCrLf & _
               "		   ON SO.Operacao_Id     = R.Operacao " & vbCrLf & _
               "		  AND SO.SubOperacoes_Id = R.SubOperacao " & vbCrLf & _
               "		WHERE SO.EstoqueFisico      = 'S' " & vbCrLf & _
               "		  AND SO.EntradaSaida       = 'E'" & vbCrLf & _
               "		  AND NF.Situacao           = 1 " & vbCrLf & _
               "		  AND NF.Serie_id          not in ('501','502','101','102','103','104') " & vbCrLf & _
               "          And NF.Finalidade         <> 5" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(R.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND R.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND R.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If


        If tipo = 1 Then
            sql &= "		  AND NF.movimento         = @data" & vbCrLf
        Else
            sql &= "		  AND year(NF.movimento)   >= @Ano" & vbCrLf & _
                   "		  AND NF.movimento         <= @data" & vbCrLf & _
                   "          AND NF.movimento         >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If


        sql &= "		  AND NFxI.produto_id in (" & Produtos & ")" & vbCrLf & _
               "		GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NFxI.deposito, NFxI.EndDeposito" & vbCrLf & _
               "	  );" & vbCrLf

        '	/*********************************************************************************************/ 
        '	/********************* Contra Partida Entrada ************************************************/ 
        '	/*********************************************************************************************/ 

        sql &= "	insert into #Disponibilidade(Empresa, EndEmpresa, Deposito, EndDeposito," & vbCrLf & _
               "								 EmbarquePorto, " & vbCrLf & _
               "								 DevDeposito, DevAcobertamentoFiscal)" & vbCrLf & _
               "	 (" & vbCrLf & _
               "	   SELECT NF.Empresa_Id, " & vbCrLf & _
               "			  NF.EndEmpresa_Id, " & vbCrLf & _
               "			  NF.Destino, " & vbCrLf & _
               "			  NF.EndDestino,  " & vbCrLf & _
               "			  ISNULL(SUM(CASE WHEN SOD.Devolucao = 'N' and SOD.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'                       and NF.Finalidade  = 7                    THEN R.PesoLiquido ELSE 0 END),0) AS EmbarquePorto," & vbCrLf & _
               "			  ISNULL(SUM(CASE WHEN SOD.Devolucao = 'S' and SOD.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "' and SOD.CobraServico = 'N' and NF.Finalidade not in (16,20,22,23,25) THEN R.PesoLiquido ELSE 0 END),0) " & vbCrLf & _
               "			+ ISNULL(SUM(CASE WHEN SO.Devolucao  = 'S' and SO.Classe   = '" & eClassesOperacoes.DEPOSITOS.ToString & "' and SO.deposito      = 'S' and NF.Finalidade  in (16, 25)            THEN R.PesoLiquido ELSE 0 END),0) AS DevDeposito," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao  = 'S' and SO.Classe   = '" & eClassesOperacoes.DEPOSITOS.ToString & "' and SO.deposito      = 'S' and Nf.finalidade  = 25                   THEN R.PesoLiquido ELSE 0 END),0) AS DevAcobertamentoFiscal" & vbCrLf & _
               "		 FROM NotasFiscais NF" & vbCrLf & _
               "		Inner Join NotasFiscaisxItens NFxI" & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
               "		Inner JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxR.Empresa_Id" & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxR.Serie_Id" & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxR.Nota_Id" & vbCrLf & _
               "		INNER JOIN Romaneios R" & vbCrLf & _
               "		   ON R.Empresa_Id    = NFxR.Empresa_Id " & vbCrLf & _
               "		  AND R.EndEmpresa_Id = NFxR.EndEmpresa_Id " & vbCrLf & _
               "		  AND R.Romaneio_Id   = NFxR.Romaneio_Id" & vbCrLf & _
               "		INNER JOIN SubOperacoes SO" & vbCrLf & _
               "		   ON SO.Operacao_Id     = R.Operacao " & vbCrLf & _
               "		  AND SO.SubOperacoes_Id = R.SubOperacao " & vbCrLf & _
               "		INNER JOIN SubOperacoes SOD" & vbCrLf & _
               "		   ON SOD.Operacao_Id     = SO.OperacaoDestino " & vbCrLf & _
               "		  AND SOD.SubOperacoes_Id = SO.SubOperacaoDestino " & vbCrLf & _
               "		INNER JOIN Pedidos P" & vbCrLf & _
               "		   ON P.Empresa_id    = Nf.Empresa_id" & vbCrLf & _
               "		  AND P.EndEmpresa_id = Nf.EndEmpresa_id" & vbCrLf & _
               "		  AND P.Pedido_id     = Nf.Pedido" & vbCrLf & _
               "		WHERE (SO.Deposito          = 'S' or  left(NFxI.deposito,8) = left(NF.destino,8))" & vbCrLf & _
               "		  AND SO.EstoqueFisico      = 'S' " & vbCrLf & _
               "		  AND SO.EntradaSaida       = 'E' " & vbCrLf & _
               "		  AND NF.Situacao           = 1 " & vbCrLf & _
               "		  AND NF.Serie_id          not in ('501','502','101','102','103','104') " & vbCrLf & _
               "          And NF.Finalidade         <> 5" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(R.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND R.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND R.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If


        If tipo = 1 Then
            sql &= "		  AND NF.movimento          = @data " & vbCrLf
        Else
            sql &= "		  AND year(NF.movimento)     >= @Ano" & vbCrLf & _
                   "		  AND NF.movimento          <= @data " & vbCrLf & _
                   "          AND NF.movimento          >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'"
        End If

        sql &= "		  AND NFxI.produto_id      in (" & Produtos & ")" & vbCrLf & _
               "		GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.destino, NF.EndDestino" & vbCrLf & _
               "	  );" & vbCrLf

        '	/*****************************************************************************/ 
        '	/********************* Saidas ************************************************/ 
        '	/*****************************************************************************/ 
        sql &= "	insert into #Disponibilidade(Empresa, EndEmpresa, Deposito, EndDeposito, RetiradaPrestacaoServico, Retirada, RetiradaFob, RetiradaFobTer, SobrasCliFor, EmbarquePorto, SaidaTransferenciaFilial, SaidaDeposito, SaidaTransferenciaTitularidade, DevParaCompra, DevDeCompra, DevDeposito, Quebras)" & vbCrLf & _
               "	 (" & vbCrLf & _
               "	   SELECT NF.Empresa_Id," & vbCrLf & _
               "			  NF.EndEmpresa_Id," & vbCrLf & _
               "			  NFxI.deposito," & vbCrLf & _
               "			  NFxI.EndDeposito," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.CobraServico = 'S'                                                                                                                                         THEN R.PesoLiquido ELSE 0 END),0) AS RetiradaPrestacaoServico," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N'                           and SO.Classe not in ('" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "')     and NF.Finalidade not in (7,17) and RxP.Pesagem_id is not null                           THEN R.PesoLiquido ELSE 0 END),0) AS Retirada," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N'                           and SO.Classe not in ('" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "')     and NF.Finalidade not in (7,17) and RxP.Pesagem_id is null     and NN.existe is not null THEN R.PesoLiquido ELSE 0 END),0) AS RetiradaFob," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N'                           and SO.Classe not in ('" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "')     and NF.Finalidade not in (7,17) and RxP.Pesagem_id is null     and NN.existe is     null THEN R.PesoLiquido ELSE 0 END),0) AS RetiradaFobTer," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N'                           and NF.Finalidade = 17 and NN.existe is null                                                                                            THEN R.PesoLiquido ELSE 0 END),0) AS SobrasCliFor," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N'                           and SO.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'                       and NF.Finalidade  = 7                                                          THEN R.PesoLiquido ELSE 0 END),0) AS EmbarquePorto," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N'                           and SO.Classe  = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'                                                                                                       THEN R.PesoLiquido ELSE 0 END),0) AS SaidaTransferenciaFilial," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'N'                           and SO.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'                            and NF.Finalidade not in (7,24)                                                 THEN R.PesoLiquido ELSE 0 END),0) AS SaidaDeposito," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.CobraServico = 'N' and SO.Classe  in ('" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.AFIXAR.ToString & "')      and NF.Finalidade = 22                                                          THEN R.PesoLiquido ELSE 0 END),0) AS SaidaTransferenciaTitularidade," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.CobraServico = 'N' and SO.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'                            and NF.Finalidade = 20                                                          THEN R.PesoLiquido ELSE 0 END),0) AS DevParaCompra," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.CobraServico = 'N' and SO.Classe <> '" & eClassesOperacoes.DEPOSITOS.ToString & "'                            and NF.Finalidade <> 22                                                         THEN R.PesoLiquido ELSE 0 END),0) AS DevDeCompra," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.CobraServico = 'N' and SO.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'                            and NF.Finalidade not in (16,20,22,23)                                          THEN R.PesoLiquido ELSE 0 END),0) AS DevDeposito," & vbCrLf & _
               "              ISNULL(SUM(CASE WHEN SO.Devolucao = 'S' and SO.CobraServico = 'N' and SO.Classe  in ('" & eClassesOperacoes.DEPOSITOS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "')                and NF.Finalidade = 16                                                          THEN R.PesoLiquido ELSE 0 END),0) AS Quebras" & vbCrLf & _
               "		 FROM NotasFiscais NF" & vbCrLf & _
               "		Inner Join NotasFiscaisxItens NFxI" & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
               "		Inner JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxR.Empresa_Id " & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxR.Serie_Id" & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxR.Nota_Id" & vbCrLf & _
               "		INNER JOIN Romaneios R" & vbCrLf & _
               "		   ON R.Empresa_Id    = NFxR.Empresa_Id " & vbCrLf & _
               "		  AND R.EndEmpresa_Id = NFxR.EndEmpresa_Id " & vbCrLf & _
               "		  AND R.Romaneio_Id   = NFxR.Romaneio_Id" & vbCrLf & _
               "         LEFT join RomaneiosxPesagens RxP" & vbCrLf & _
               "           ON RxP.Empresa_Id    = NFxR.Empresa_Id" & vbCrLf & _
               "          AND RxP.EndEmpresa_Id = NFxR.EndEmpresa_Id" & vbCrLf & _
               "          AND RxP.Romaneio_Id   = NFxR.Romaneio_Id" & vbCrLf & _
               "		INNER JOIN SubOperacoes SO" & vbCrLf & _
               "		   ON SO.Operacao_Id     = R.Operacao " & vbCrLf & _
               "		  AND SO.SubOperacoes_Id = R.SubOperacao " & vbCrLf & _
               "		INNER JOIN Pedidos P" & vbCrLf & _
               "		   ON P.Empresa_id    = Nf.Empresa_id" & vbCrLf & _
               "		  AND P.EndEmpresa_id = Nf.EndEmpresa_id" & vbCrLf & _
               "		  AND P.Pedido_id     = Nf.Pedido" & vbCrLf & _
               "         Left Join (Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, 1 as Existe" & vbCrLf & _
               "                      from NotasXNotas" & vbCrLf & _
               "                     group by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id " & vbCrLf & _
               "                   ) as NN" & vbCrLf & _
               "           on NN.Empresa_Id      = NF.Empresa_Id" & vbCrLf & _
               "          and NN.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf & _
               "          and NN.Cliente_Id      = NF.Cliente_Id" & vbCrLf & _
               "          and NN.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf & _
               "          and NN.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf & _
               "          and NN.Serie_Id        = NF.Serie_Id" & vbCrLf & _
               "          and NN.Nota_Id         = NF.Nota_Id" & vbCrLf & _
               "		WHERE SO.EstoqueFisico      = 'S'" & vbCrLf & _
               "		  AND SO.EntradaSaida       = 'S'" & vbCrLf & _
               "		  AND NF.Situacao           = 1 " & vbCrLf & _
               "		  AND NF.Serie_id          not in ('501','502','101','102','103','104')" & vbCrLf & _
               "          And NF.Finalidade         <> 5" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(R.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND R.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND R.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        If tipo = 1 Then
            sql &= "		  AND NF.movimento          = @data " & vbCrLf
        Else
            sql &= "          AND Year(NF.Movimento)    >= @Ano" & vbCrLf & _
                   "		  AND NF.movimento         >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "		  AND NF.movimento         <= @data " & vbCrLf
        End If

        sql &= "		  AND NFxI.produto_id      in (" & Produtos & ")" & vbCrLf & _
               "		GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NFxI.deposito, NFxI.EndDeposito" & vbCrLf & _
               "	  );" & vbCrLf
        '	/*****************************************************************************************/ 
        '	/******************* Contra Partida Saida ************************************************/ 
        '	/*****************************************************************************************/ 
        sql &= "	insert into #Disponibilidade(Empresa, EndEmpresa, Deposito, EndDeposito, " & vbCrLf & _
               "								 EntradaDeposito, " & vbCrLf & _
               "								 RetornoSimbolicoFormacaoDeLote,RetornoDeDeposito, AcobertamentoFiscal)" & vbCrLf & _
               "	 (" & vbCrLf & _
               "	  SELECT NF.Empresa_Id, " & vbCrLf & _
               "			 NF.EndEmpresa_Id, " & vbCrLf & _
               "			 NF.destino, " & vbCrLf & _
               "			 NF.EndDestino,  " & vbCrLf & _
               "			 ISNULL(SUM(CASE WHEN SOD.Devolucao = 'N' and SOD.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      and NF.Finalidade not in (23,24)                            THEN R.PesoLiquido ELSE 0 END),0) AS EntradaDeposito," & vbCrLf & _
               "			 ISNULL(SUM(CASE WHEN SOD.Devolucao = 'S' and SOD.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      and NF.Finalidade = 7                                       THEN R.PesoLiquido ELSE 0 END),0) AS RetornoSimbolicoFormacaoDeLote," & vbCrLf & _
               "			 ISNULL(SUM(CASE WHEN SOD.Devolucao = 'S' and SOD.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      and NF.Finalidade = 9                                       THEN R.PesoLiquido ELSE 0 END),0) AS RetornoDeDeposito," & vbCrLf & _
               "		     ISNULL(SUM(CASE WHEN SOD.Devolucao = 'N' and SOD.Classe  = '" & eClassesOperacoes.DEPOSITOS.ToString & "'      and NF.Finalidade = 24                                      THEN R.PesoLiquido ELSE 0 END),0) AS AcobertamentoFiscal" & vbCrLf & _
               "		 FROM NotasFiscais NF " & vbCrLf & _
               "		Inner Join NotasFiscaisxItens NFxI" & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxI.Empresa_Id  " & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id  " & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxI.Cliente_Id  " & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxI.EndCliente_Id  " & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id  " & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxI.Serie_Id  " & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxI.Nota_Id  " & vbCrLf & _
               "		Inner JOIN NotasFiscaisXRomaneios NFxR " & vbCrLf & _
               "		   ON NF.Empresa_Id      = NFxR.Empresa_Id  " & vbCrLf & _
               "		  AND NF.EndEmpresa_Id   = NFxR.EndEmpresa_Id  " & vbCrLf & _
               "		  AND NF.Cliente_Id      = NFxR.Cliente_Id  " & vbCrLf & _
               "		  AND NF.EndCliente_Id   = NFxR.EndCliente_Id  " & vbCrLf & _
               "		  AND NF.EntradaSaida_Id = NFxR.EntradaSaida_Id  " & vbCrLf & _
               "		  AND NF.Serie_Id        = NFxR.Serie_Id  " & vbCrLf & _
               "		  AND NF.Nota_Id         = NFxR.Nota_Id  " & vbCrLf & _
               "		INNER JOIN Romaneios R" & vbCrLf & _
               "		   ON R.Empresa_Id    = NFxR.Empresa_Id " & vbCrLf & _
               "		  AND R.EndEmpresa_Id = NFxR.EndEmpresa_Id " & vbCrLf & _
               "		  AND R.Romaneio_Id   = NFxR.Romaneio_Id" & vbCrLf & _
               "		INNER JOIN SubOperacoes SO" & vbCrLf & _
               "		   ON SO.Operacao_Id     = R.Operacao " & vbCrLf & _
               "		  AND SO.SubOperacoes_Id = R.SubOperacao  " & vbCrLf & _
               "		INNER JOIN SubOperacoes SOD" & vbCrLf & _
               "		   ON SOD.Operacao_Id     = SO.OperacaoDestino" & vbCrLf & _
               "		  AND SOD.SubOperacoes_Id = SO.SubOperacaoDestino " & vbCrLf & _
               "		WHERE (SO.Deposito           = 'S' or  left(NFxI.deposito,8) = left(NF.destino,8))" & vbCrLf & _
               "		  And SO.EstoqueFisico      = 'S' " & vbCrLf & _
               "		  AND SO.EntradaSaida       = 'S'" & vbCrLf & _
               "		  AND NF.Situacao           = 1 " & vbCrLf & _
               "		  AND NF.Serie_id          not in ('501','502','101','102','103','104') " & vbCrLf & _
               "          And NF.Finalidade         <> 5" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(R.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND R.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND R.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        If tipo = 1 Then
            sql &= "		  AND NF.movimento          = @data" & vbCrLf
        Else
            sql &= "		  AND year(NF.movimento)   >= @Ano" & vbCrLf & _
                   "		  AND NF.movimento         >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "		  AND NF.movimento         <= @data" & vbCrLf
        End If

        sql &= "		  AND NFxI.produto_id       in (" & Produtos & ")" & vbCrLf & _
               "		GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.destino, NF.EndDestino" & vbCrLf & _
               "	  );" & vbCrLf
        '	/***********************************************************************************/   
        '	/******************************* Lancamentos Producao  *****************************/   
        '	/***********************************************************************************/  
        sql &= "	insert into #Disponibilidade(Empresa, EndEmpresa, Deposito, EndDeposito,EntradaProducao,SaidaProducao)" & vbCrLf & _
               "	   (" & vbCrLf & _
               "		SELECT Producao.Empresa_Id, " & vbCrLf & _
               "			  Producao.EndEmpresa_Id, " & vbCrLf & _
               "			  Producao.Deposito_Id As Deposito, " & vbCrLf & _
               "			  Producao.EndDeposito_Id as EndDeposito, " & vbCrLf & _
               "			  ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Entradas END), 0) AS Entradas, " & vbCrLf & _
               "			  ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Saidas   END), 0) AS Saidas" & vbCrLf & _
               "		 FROM Producao  " & vbCrLf & _
               "		INNER JOIN SubOperacoes" & vbCrLf & _
               "		   ON SubOperacoes.Operacao_Id     = Producao.Operacao_Id " & vbCrLf & _
               "		  AND SubOperacoes.SubOperacoes_Id = Producao.SubOperacao_Id" & vbCrLf & _
               "		WHERE SubOperacoes.EstoqueFisico   = 'S' " & vbCrLf & _
               "		  AND Producao.FisicoFiscal_Id     = 1" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(Producao.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND Producao.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND Producao.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        If tipo = 1 Then
            sql &= "		  AND Producao.Movimento_Id        = @data" & vbCrLf
        Else
            sql &= "          And year(Producao.Movimento_Id)  >= @Ano" & vbCrLf & _
                   "          And Producao.Movimento_Id       >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "		  AND Producao.Movimento_Id       <= @data" & vbCrLf
        End If

        sql &= "		  AND Producao.Produto_Id in (" & Produtos & ")" & vbCrLf & _
               "		GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id	" & vbCrLf & _
               "	   ) " & vbCrLf
        '	/***********************************************************************************/   
        '	/******************************* Agrupa Lancamentos  *******************************/   
        '	/***********************************************************************************/  
        sql &= "	   SELECT Empresa, " & vbCrLf & _
               "			  EndEmpresa, " & vbCrLf & _
               "			  deposito, " & vbCrLf & _
               "			  EndDeposito," & vbCrLf & _
               "			  convert(numeric(18,0), 0) as Inicial," & vbCrLf & _
               "			  convert(numeric(18,0), 0) as LaudoSemNota," & vbCrLf & _
               "			  convert(numeric(18,0), 0) as LaudoSemNotaSaida," & vbCrLf & _
               "			  SUM(EntradaProducao) as EntradaProducao," & vbCrLf & _
               "			  SUM(SaidaProducao) as SaidaProducao," & vbCrLf & _
               "			  SUM(Recebimento) as Recebimento," & vbCrLf & _
               "			  SUM(Fob) as Fob," & vbCrLf & _
               "			  SUM(CompraPorto) as CompraPorto," & vbCrLf & _
               "			  SUM(CompraEmDeposito) as CompraEmDeposito," & vbCrLf & _
               "			  SUM(EntradaTransferenciaFilial) as EntradaTransferenciaFilial," & vbCrLf & _
               "              SUM(EntradaTransferenciaTitularidade) as EntradaTransferenciaTitularidade," & vbCrLf & _
               "			  SUM(EntradaDevolucaoVenda) as EntradaDevolucaoVenda," & vbCrLf & _
               "			  SUM(EntradaDeposito) as EntradaDeposito," & vbCrLf & _
               "			  SUM(RetornoSimbolicoFormacaoDeLote) as RetornoSimbolicoFormacaoDeLote," & vbCrLf & _
               "              SUM(RetornoDeDeposito) as RetornoDeDeposito," & vbCrLf & _
               "			  SUM(RetiradaPrestacaoServico) as RetiradaPrestacaoServico," & vbCrLf & _
               "			  SUM(Retirada) as Retirada," & vbCrLf & _
               "			  SUM(RetiradaFob) as RetiradaFob," & vbCrLf & _
               "              SUM(RetiradaFobTer) as RetiradaFobTer," & vbCrLf & _
               "			  SUM(EmbarquePorto) as EmbarquePorto," & vbCrLf & _
               "			  SUM(SaidaTransferenciaFilial) as SaidaTransferenciaFilial," & vbCrLf & _
               "			  SUM(SaidaDeposito) as SaidaDeposito," & vbCrLf & _
               "			  SUM(SaidaTransferenciaTitularidade) as SaidaTransferenciaTitularidade," & vbCrLf & _
               "			  SUM(DevParaCompra) as DevParaCompra," & vbCrLf & _
               "			  SUM(DevDeCompra) as DevDeCompra," & vbCrLf & _
               "			  SUM(DevDeposito) as DevDeposito," & vbCrLf & _
               "			  SUM(AcobertamentoFiscal) as AcobertamentoFiscal," & vbCrLf & _
               "			  SUM(DevAcobertamentoFiscal) as DevAcobertamentoFiscal," & vbCrLf & _
               "			  SUM(Quebras) as Quebras," & vbCrLf & _
               "			  SUM(Sobras) as Sobras," & vbCrLf & _
               "              SUM(SobrasCliFor) as SobrasCliFor" & vbCrLf & _
               "		 into #Disponibilidade2" & vbCrLf & _
               "		 from #Disponibilidade" & vbCrLf & _
               "		Group by Empresa, EndEmpresa, deposito, EndDeposito;" & vbCrLf

        '	/***********************************************************************************/   
        '	/****************************** Apura Saldo Inicial  *******************************/   
        '	/***********************************************************************************/  
        sql &= "	SELECT Consulta.Empresa, " & vbCrLf & _
               "		   Consulta.EndEmpresa," & vbCrLf & _
               "		   Consulta.Deposito, " & vbCrLf & _
               "		   Consulta.EndDeposito," & vbCrLf & _
               "		   SUM(Consulta.Inicial) AS Inicial" & vbCrLf & _
               "	 Into #SaldoInicial" & vbCrLf & _
               "	 FROM (Select Empresa_Id    as Empresa," & vbCrLf & _
               "				  EndEmpresa_Id as EndEmpresa," & vbCrLf & _
               "				  Deposito," & vbCrLf & _
               "				  EndDeposito," & vbCrLf & _
               "				  Produto_Id as Produto," & vbCrLf & _
               "				  isnull(SUM(CASE WHEN Movimento < @Data   THEN Entradas - Saidas END),0) As Inicial" & vbCrLf & _
               "			 From (" & vbCrLf & _
               "				   SELECT Producao.Empresa_Id, " & vbCrLf & _
               "						  Producao.EndEmpresa_Id, " & vbCrLf & _
               "						  Producao.Deposito_Id As Deposito," & vbCrLf & _
               "						  Producao.EndDeposito_Id as EndDeposito," & vbCrLf & _
               "						  Producao.Produto_Id, " & vbCrLf & _
               "						  Producao.Movimento_Id as Movimento," & vbCrLf & _
               "						  ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Entradas END), 0) AS Entradas, " & vbCrLf & _
               "						  ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Saidas   END), 0) AS Saidas" & vbCrLf & _
               "					 FROM Producao  " & vbCrLf & _
               "					INNER JOIN SubOperacoes" & vbCrLf & _
               "					   ON SubOperacoes.Operacao_Id     = Producao.Operacao_Id " & vbCrLf & _
               "					  AND SubOperacoes.SubOperacoes_Id = Producao.SubOperacao_Id" & vbCrLf & _
               "					WHERE SubOperacoes.EstoqueFisico   = 'S' " & vbCrLf & _
               "					  AND Producao.FisicoFiscal_Id     = 1" & vbCrLf & _
               "                      AND Producao.Produto_Id in (" & Produtos & ")" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(Producao.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND Producao.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND Producao.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        If tipo = 1 Then
            sql &= "					  AND Producao.Movimento_Id        < @Data " & vbCrLf
        Else
            sql &= "                      AND Year(Producao.Movimento_Id)  < @Ano" & vbCrLf
        End If

        sql &= "                      AND Producao.Movimento_Id >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
               "					GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Movimento_Id" & vbCrLf & _
               "				   ) as Consulta" & vbCrLf & _
               "			Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto_Id" & vbCrLf & _
               "		    UNION" & vbCrLf & _
               "		   Select Empresa_Id as Empresa, " & vbCrLf & _
               "				  EndEmpresa_Id as EndEmpresa, " & vbCrLf & _
               "				  Deposito, " & vbCrLf & _
               "				  EndDeposito, " & vbCrLf & _
               "				  Produto, " & vbCrLf & _
               "				  isnull(SUM(CASE  WHEN Movimento < @Data  THEN Entradas - Saidas END),0) As Inicial" & vbCrLf & _
               "			 From (" & vbCrLf & _
               "				   SELECT Romaneios.Empresa_Id, " & vbCrLf & _
               "						  Romaneios.EndEmpresa_Id, " & vbCrLf & _
               "						  NotasFiscaisXitens.Deposito, " & vbCrLf & _
               "						  NotasFiscaisXitens.EndDeposito, " & vbCrLf & _
               "						  Romaneios.Produto, " & vbCrLf & _
               "						  NotasFiscais.Movimento," & vbCrLf & _
               "						  ISNULL(SUM(CASE" & vbCrLf & _
               "									   WHEN SubOperacoes.EntradaSaida = 'E'" & vbCrLf & _
               "										 THEN" & vbCrLf & _
               "										   case" & vbCrLf & _
               "											  when Notasfiscais.Finalidade = 16" & vbCrLf & _
               "												Then 0--Romaneios.PesoLiquido * -1" & vbCrLf & _
               "												Else Romaneios.PesoLiquido" & vbCrLf & _
               "										   end" & vbCrLf & _
               "										 ELSE 0" & vbCrLf & _
               "									 END), 0) AS Entradas," & vbCrLf & _
               "						  ISNULL(SUM(CASE" & vbCrLf & _
               "									   WHEN SubOperacoes.EntradaSaida = 'S'" & vbCrLf & _
               "										 THEN Romaneios.PesoLiquido" & vbCrLf & _
               "										 ELSE 0" & vbCrLf & _
               "									 END), 0) AS Saidas" & vbCrLf & _
               "					 FROM SubOperacoes " & vbCrLf & _
               "					INNER JOIN Romaneios" & vbCrLf & _
               "					   ON SubOperacoes.Operacao_Id     = Romaneios.Operacao " & vbCrLf & _
               "					  AND SubOperacoes.SubOperacoes_Id = Romaneios.SubOperacao" & vbCrLf & _
               "					 Left JOIN NotasFiscaisXRomaneios nfxr " & vbCrLf & _
               "					   ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf & _
               "					  AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
               "					  AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id" & vbCrLf & _
               "					INNER JOIN NotasFiscais " & vbCrLf & _
               "					   ON nfxr.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf & _
               "					  AND nfxr.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
               "					  AND nfxr.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
               "					  AND nfxr.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
               "					  AND nfxr.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
               "					  AND nfxr.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
               "					  AND nfxr.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf & _
               "					INNER Join NotasFiscaisXitens" & vbCrLf & _
               "					   ON NotasFiscais.Empresa_Id      = NotasFiscaisXitens.Empresa_Id  " & vbCrLf & _
               "					  AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXitens.EndEmpresa_Id  " & vbCrLf & _
               "					  AND NotasFiscais.Cliente_Id      = NotasFiscaisXitens.Cliente_Id  " & vbCrLf & _
               "					  AND NotasFiscais.EndCliente_Id   = NotasFiscaisXitens.EndCliente_Id  " & vbCrLf & _
               "					  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXitens.EntradaSaida_Id  " & vbCrLf & _
               "					  AND NotasFiscais.Serie_Id        = NotasFiscaisXitens.Serie_Id  " & vbCrLf & _
               "					  AND NotasFiscais.Nota_Id         = NotasFiscaisXitens.Nota_Id " & vbCrLf & _
               "					WHERE SubOperacoes.EstoqueFisico = 'S' " & vbCrLf & _
               "					  AND NotasFiscais.Situacao      = 1 " & vbCrLf & _
               "					  AND NotasFiscais.Serie_id    not in('501','502','101','102','103','104') " & vbCrLf & _
               "                      AND NotasFiscais.Finalidade  not in (5, 24, 25)" & vbCrLf & _
               "                      AND not (notasfiscais.entradasaida_id = 'S' and NotasFiscais.Finalidade = 17)" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(Romaneios.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND Romaneios.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND Romaneios.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        If tipo = 1 Then
            sql &= "					  AND NotasFiscais.Movimento        < @Data " & vbCrLf
        Else
            sql &= "					  AND year(NotasFiscais.Movimento)  < @Ano" & vbCrLf
        End If
        sql &= "					  AND NotasFiscais.Movimento >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf

        sql &= "					GROUP BY Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, NotasFiscaisXitens.Deposito, NotasFiscaisXitens.EndDeposito, Romaneios.Produto, NotasFiscais.Movimento " & vbCrLf & _
               "				    Union ALL" & vbCrLf & _
               "				   SELECT Romaneios.Empresa_Id, " & vbCrLf & _
               "						  Romaneios.EndEmpresa_Id, " & vbCrLf & _
               "						  NotasFiscais.Destino, " & vbCrLf & _
               "						  NotasFiscais.EndDestino, " & vbCrLf & _
               "						  Romaneios.Produto, " & vbCrLf & _
               "						  NotasFiscais.Movimento," & vbCrLf & _
               "						  ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'E' THEN Romaneios.PesoLiquido END), 0) AS Entradas," & vbCrLf & _
               "						  ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'S' THEN Romaneios.PesoLiquido END), 0) AS Saidas" & vbCrLf & _
               "					 FROM Romaneios  " & vbCrLf & _
               "					INNER JOIN SubOperacoes" & vbCrLf & _
               "					   ON SubOperacoes.Operacao_Id     = Romaneios.Operacao " & vbCrLf & _
               "					  AND SubOperacoes.SubOperacoes_Id = Romaneios.SubOperacao" & vbCrLf & _
               "					INNER JOIN SubOperacoes SOD" & vbCrLf & _
               "					   ON SOD.Operacao_Id     = SubOperacoes.OperacaoDestino " & vbCrLf & _
               "					  AND SOD.SubOperacoes_Id = SubOperacoes.SubOperacaoDestino" & vbCrLf & _
               "					 Left JOIN NotasFiscaisXRomaneios nfxr " & vbCrLf & _
               "					   ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf & _
               "					  AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
               "					  AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id" & vbCrLf & _
               "					INNER JOIN NotasFiscais " & vbCrLf & _
               "					   ON nfxr.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf & _
               "					  AND nfxr.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf & _
               "					  AND nfxr.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf & _
               "					  AND nfxr.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf & _
               "					  AND nfxr.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf & _
               "					  AND nfxr.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
               "					  AND nfxr.Nota_Id         = NotasFiscais.Nota_Id " & vbCrLf & _
               "					INNER Join NotasFiscaisXitens" & vbCrLf & _
               "					   ON NotasFiscais.Empresa_Id      = NotasFiscaisXitens.Empresa_Id  " & vbCrLf & _
               "					  AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXitens.EndEmpresa_Id  " & vbCrLf & _
               "					  AND NotasFiscais.Cliente_Id      = NotasFiscaisXitens.Cliente_Id  " & vbCrLf & _
               "					  AND NotasFiscais.EndCliente_Id   = NotasFiscaisXitens.EndCliente_Id  " & vbCrLf & _
               "					  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXitens.EntradaSaida_Id  " & vbCrLf & _
               "					  AND NotasFiscais.Serie_Id        = NotasFiscaisXitens.Serie_Id  " & vbCrLf & _
               "					  AND NotasFiscais.Nota_Id         = NotasFiscaisXitens.Nota_Id  " & vbCrLf & _
               "					WHERE (SubOperacoes.Deposito = 'S')"  'or left(NotasFiscais.Destino,8) = left(NotasFiscaisXitens.deposito,8))" & vbCrLf
        ' "                      and NotasFiscais.Empresa_id = NotasFiscais.Cliente_Id" & vbCrLf & _
        sql &= "					  AND SubOperacoes.EstoqueFisico   = 'S' " & vbCrLf & _
               "					  AND NotasFiscais.Situacao        = 1 " & vbCrLf & _
               "					  AND NotasFiscais.Serie_id      not in('501','502','101','102','103','104') " & vbCrLf & _
               "                      AND NotasFiscais.Finalidade    not in (5, 24, 25)" & vbCrLf & _
               "                      and not (notasfiscais.entradasaida_id = 'S' and NotasFiscais.Finalidade = 17)" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(Romaneios.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND Romaneios.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND Romaneios.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        If tipo = 1 Then
            sql &= "					  AND NotasFiscais.Movimento       < @Data" & vbCrLf & _
                   "					  AND NotasFiscais.Movimento       >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf
        Else
            sql &= "                      and YEAR(NotasFiscais.Movimento)  < @Ano" & vbCrLf & _
                   "					  and NotasFiscais.Movimento       >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf

        End If

        sql &= "					GROUP BY Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, NotasFiscais.Destino, NotasFiscais.EndDestino, Romaneios.Produto, NotasFiscais.Movimento" & vbCrLf & _
               "				  ) as Consulta" & vbCrLf & _
               "             where Produto in (" & Produtos & ")" & vbCrLf & _
               "			 Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto" & vbCrLf & _
               "             UNION" & vbCrLf
        '*************************************************************************************************
        '***************** Saldo Inicial Informado para Disponibilidade **********************************
        '*************************************************************************************************
        sql &= "            Select D.Empresa," & vbCrLf & _
               "                   D.EndEmpresa," & vbCrLf & _
               "                   SI.Deposito_Id," & vbCrLf & _
               "                   SI.EndDeposito_Id," & vbCrLf & _
               "                   '' as Produto," & vbCrLf & _
               "                   SaldoInicialFiscal" & vbCrLf & _
               "              from SaldoInicialDisponibilidade D" & vbCrLf & _
               "             inner join SaldoInicialDisponibilidadeDeposito SI" & vbCrLf & _
               "                on D.Disponibilidade_Id = SI.Disponibilidade_Id" & vbCrLf & _
               "             where D.Disponibilidade_Id = " & objDisponibilidade.CodigoDisponibilidade & vbCrLf
        '*************************************************************************************************


        sql &= "		  )AS Consulta " & vbCrLf & _
               "	 GROUP BY Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito" & vbCrLf & _
               "	 having SUM(Consulta.Inicial) <> 0; " & vbCrLf


        '	/***********************************************************************************/   
        '	/****************** Atualiza Saldo Inicial na Disponibilidade  *********************/   
        '	/***********************************************************************************/ 
        sql &= "	Update #Disponibilidade2 set" & vbCrLf & _
               "	  inicial = #SaldoInicial.Inicial" & vbCrLf & _
               "	 from #Disponibilidade2" & vbCrLf & _
               "	Inner Join #SaldoInicial" & vbCrLf & _
               "	   on #Disponibilidade2.Empresa     = #SaldoInicial.Empresa" & vbCrLf & _
               "	  and #Disponibilidade2.EndEmpresa  = #SaldoInicial.EndEmpresa" & vbCrLf & _
               "	  and #Disponibilidade2.Deposito    = #SaldoInicial.Deposito" & vbCrLf & _
               "	  and #Disponibilidade2.EndDeposito = #SaldoInicial.EndDeposito; " & vbCrLf


        '	/***********************************************************************************/   
        '	/***** Inseri os Deposito somente com  Saldo Inicial sem movimentacao no ano  ******/   
        '	/***********************************************************************************/
        sql &= "	  insert into #Disponibilidade2(Empresa, EndEmpresa, deposito, EndDeposito, Inicial, LaudoSemNota, LaudoSemNotaSaida, EntradaProducao, SaidaProducao, Recebimento, Fob, CompraPorto, CompraEmDeposito, EntradaTransferenciaFilial, EntradaTransferenciaTitularidade, EntradaDevolucaoVenda, EntradaDeposito, RetornoSimbolicoFormacaoDeLote, RetornoDeDeposito, RetiradaPrestacaoServico, Retirada, RetiradaFob, RetiradaFobTer, EmbarquePorto, SaidaTransferenciaFilial, SaidaDeposito, SaidaTransferenciaTitularidade, DevParaCompra, DevDeCompra, DevDeposito,AcobertamentoFiscal,DevAcobertamentoFiscal, Quebras, Sobras,SobrasCliFor)" & vbCrLf & _
               "			  (select #SaldoInicial.Empresa, #SaldoInicial.EndEmpresa, #SaldoInicial.Deposito, #SaldoInicial.EndDeposito, #SaldoInicial.Inicial,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0" & vbCrLf & _
               "				 from #SaldoInicial" & vbCrLf & _
               "				where not exists(Select 1" & vbCrLf & _
               "								   from #Disponibilidade2 D" & vbCrLf & _
               "								  Where D.Empresa     = #SaldoInicial.Empresa" & vbCrLf & _
               "									and D.EndEmpresa  = #SaldoInicial.EndEmpresa" & vbCrLf & _
               "									and D.Deposito    = #SaldoInicial.Deposito" & vbCrLf & _
               "									and D.EndDeposito = #SaldoInicial.EndDeposito" & vbCrLf & _
               "								 )" & vbCrLf & _
               "			  ); " & vbCrLf
        '	/********************************************************************************************/   
        '	/*****************************  INSERI LAUDO SEM NOTA  ENTRADA *****************************/     
        '	/*******************************************************************************************/
        sql &= "	Update #Disponibilidade2 set" & vbCrLf & _
               "	   LaudoSemNota = sb.LaudosSemNota" & vbCrLf & _
               "	  From #Disponibilidade2" & vbCrLf & _
               "	 Inner Join (" & vbCrLf & _
               "					SELECT Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, Romaneios.Deposito, Romaneios.EndDeposito, sum(Romaneios.PesoLiquido) as LaudosSemNota" & vbCrLf & _
               "					  FROM Romaneios " & vbCrLf & _
               "					 INNER JOIN RomaneiosXPesagens " & vbCrLf & _
               "						ON Romaneios.Empresa_Id    = RomaneiosXPesagens.Empresa_Id " & vbCrLf & _
               "					   AND Romaneios.EndEmpresa_Id = RomaneiosXPesagens.EndEmpresa_Id " & vbCrLf & _
               "					   AND Romaneios.Romaneio_Id   = RomaneiosXPesagens.Romaneio_Id " & vbCrLf & _
               "					 Inner join Pesagem" & vbCrLf & _
               "                        ON RomaneiosXPesagens.Empresa_Id    = Pesagem.Empresa_Id " & vbCrLf & _
               "                       AND RomaneiosXPesagens.EndEmpresa_Id = Pesagem.EndEmpresa_Id " & vbCrLf & _
               "                       AND RomaneiosXPesagens.Pesagem_id    = Pesagem.Pesagem_id " & vbCrLf & _
               "                       AND Pesagem.Situacao                 = 1" & vbCrLf & _
               "					  LEFT JOIN Pedidos " & vbCrLf & _
               "						ON Romaneios.Empresa_Id    = Pedidos.Empresa_Id " & vbCrLf & _
               "					   AND Romaneios.EndEmpresa_Id = Pedidos.EndEmpresa_Id " & vbCrLf & _
               "					   AND Romaneios.Pedido        = Pedidos.Pedido_Id " & vbCrLf & _
               "					  LEFT OUTER JOIN Clientes " & vbCrLf & _
               "						ON Pedidos.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
               "					   AND Pedidos.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
               "					  LEFT OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf & _
               "						ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf & _
               "					   AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
               "					   AND Pedidos.Cliente         = nfxr.Cliente_id" & vbCrLf & _
               "					   AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
               "                      LEFT JOIN NotasFiscais" & vbCrLf & _
               "                        ON NotasFiscais.Empresa_Id      = nfxr.Empresa_Id " & vbCrLf & _
               "					   AND NotasFiscais.EndEmpresa_Id   = nfxr.EndEmpresa_Id " & vbCrLf & _
               "					   AND NotasFiscais.Cliente_Id      = nfxr.Cliente_Id  " & vbCrLf & _
               "					   AND NotasFiscais.EndCliente_Id   = nfxr.EndCliente_Id  " & vbCrLf & _
               "					   AND NotasFiscais.EntradaSaida_Id = nfxr.EntradaSaida_Id " & vbCrLf & _
               "					   AND NotasFiscais.Serie_Id        = nfxr.Serie_Id  " & vbCrLf & _
               "					   AND NotasFiscais.Nota_Id         = nfxr.Nota_Id  " & vbCrLf & _
               "                       AND NotasFiscais.Movimento      <= @Data " & vbCrLf & _
               "                       And NotasFiscais.Finalidade     <> 5" & vbCrLf

        If tipo = 1 Then
            sql &= "					 WHERE Romaneios.Movimento         = @Data  " & vbCrLf & _
                   "					   and Romaneios.Movimento        >= @dataCargaPatio"
            '"					   and Romaneios.Movimento       >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf
        Else
            sql &= "					 WHERE Romaneios.Movimento       <= @Data  " & vbCrLf & _
                   "					   and Romaneios.Movimento       >= @dataCargaPatio"
            '"					   and Romaneios.Movimento       >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(Romaneios.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND Romaneios.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND Romaneios.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        sql &= "					   AND (NotasFiscais.Nota_Id IS NULL) " & vbCrLf & _
               "					   and Romaneios.EntradaSaida  = 'E'" & vbCrLf & _
               "					   and Romaneios.Produto in (" & Produtos & ")" & vbCrLf & _
               "                       and Pesagem.Situacao        = 1 " & vbCrLf & _
               "					group by Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, Romaneios.Deposito, Romaneios.EndDeposito" & vbCrLf & _
               "				) sb" & vbCrLf & _
               "	  on #Disponibilidade2.Empresa     = sb.Empresa_Id" & vbCrLf & _
               "	 and #Disponibilidade2.EndEmpresa  = sb.EndEmpresa_Id" & vbCrLf & _
               "	 and #Disponibilidade2.Deposito    = sb.Deposito" & vbCrLf & _
               "	 and #Disponibilidade2.EndDeposito = sb.EndDeposito;" & vbCrLf

        '	/********************************************************************************************/   
        '	/*****************************  INSERI LAUDO SEM NOTA  Saida *******************************/     
        '	/*******************************************************************************************/
        sql &= "	Update #Disponibilidade2 set" & vbCrLf & _
               "	   LaudoSemNotaSaida = sb.LaudosSemNota" & vbCrLf & _
               "	  From #Disponibilidade2" & vbCrLf & _
               "	 Inner Join (" & vbCrLf & _
               "					SELECT Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, Romaneios.Deposito, Romaneios.EndDeposito, sum(Romaneios.PesoLiquido) as LaudosSemNota" & vbCrLf & _
               "					  FROM Romaneios " & vbCrLf & _
               "					 INNER JOIN RomaneiosXPesagens " & vbCrLf & _
               "						ON Romaneios.Empresa_Id    = RomaneiosXPesagens.Empresa_Id " & vbCrLf & _
               "					   AND Romaneios.EndEmpresa_Id = RomaneiosXPesagens.EndEmpresa_Id " & vbCrLf & _
               "					   AND Romaneios.Romaneio_Id   = RomaneiosXPesagens.Romaneio_Id " & vbCrLf & _
               "					 Inner join Pesagem" & vbCrLf & _
               "                        ON RomaneiosXPesagens.Empresa_Id    = Pesagem.Empresa_Id " & vbCrLf & _
               "                       AND RomaneiosXPesagens.EndEmpresa_Id = Pesagem.EndEmpresa_Id " & vbCrLf & _
               "                       AND RomaneiosXPesagens.Pesagem_id    = Pesagem.Pesagem_id " & vbCrLf & _
               "                       AND Pesagem.Situacao                 = 1" & vbCrLf & _
               "					  LEFT JOIN Pedidos " & vbCrLf & _
               "						ON Romaneios.Empresa_Id    = Pedidos.Empresa_Id " & vbCrLf & _
               "					   AND Romaneios.EndEmpresa_Id = Pedidos.EndEmpresa_Id " & vbCrLf & _
               "					   AND Romaneios.Pedido        = Pedidos.Pedido_Id " & vbCrLf & _
               "					  LEFT OUTER JOIN Clientes " & vbCrLf & _
               "						ON Pedidos.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
               "					   AND Pedidos.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
               "					  LEFT OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf & _
               "						ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf & _
               "					   AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
               "					   AND Pedidos.Cliente         = nfxr.Cliente_id" & vbCrLf & _
               "					   AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
               "                      LEFT JOIN NotasFiscais" & vbCrLf & _
               "                        ON NotasFiscais.Empresa_Id      = nfxr.Empresa_Id " & vbCrLf & _
               "					   AND NotasFiscais.EndEmpresa_Id   = nfxr.EndEmpresa_Id " & vbCrLf & _
               "					   AND NotasFiscais.Cliente_Id      = nfxr.Cliente_Id  " & vbCrLf & _
               "					   AND NotasFiscais.EndCliente_Id   = nfxr.EndCliente_Id  " & vbCrLf & _
               "					   AND NotasFiscais.EntradaSaida_Id = nfxr.EntradaSaida_Id " & vbCrLf & _
               "					   AND NotasFiscais.Serie_Id        = nfxr.Serie_Id  " & vbCrLf & _
               "					   AND NotasFiscais.Nota_Id         = nfxr.Nota_Id  " & vbCrLf & _
               "                       AND NotasFiscais.Movimento       <=  @Data " & vbCrLf & _
               "                       And NotasFiscais.Finalidade      <> 5" & vbCrLf

        If tipo = 1 Then
            sql &= "					 WHERE Romaneios.Movimento         = @Data  " & vbCrLf & _
                   "					   and Romaneios.Movimento        >= @dataCargaPatio"
            '"					   and Romaneios.Movimento       >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf
        Else
            sql &= "					 WHERE Romaneios.Movimento       <= @Data  " & vbCrLf & _
                   "					   and Romaneios.Movimento       >= @dataCargaPatio"
            '"					   and Romaneios.Movimento       >='" & objDisponibilidade.DataInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If objDisponibilidade.Consolidado Then
            sql &= "		  AND left(Romaneios.Empresa_Id,8)  = @Empresa " & vbCrLf
        Else
            sql &= "		  AND Romaneios.Empresa_Id          = @Empresa " & vbCrLf & _
                   "		  AND Romaneios.EndEmpresa_Id       = @EndEmpresa " & vbCrLf
        End If

        sql &= "					   AND (NotasFiscais.Nota_Id IS NULL) " & vbCrLf & _
               "					   and Romaneios.EntradaSaida = 'S'" & vbCrLf & _
               "					   and Romaneios.Produto in (" & Produtos & ")" & vbCrLf & _
               "                       and Pesagem.Situacao        = 1 " & vbCrLf & _
               "					group by Romaneios.Empresa_Id, Romaneios.EndEmpresa_Id, Romaneios.Deposito, Romaneios.EndDeposito" & vbCrLf & _
               "				) sb" & vbCrLf & _
               "	  on #Disponibilidade2.Empresa     = sb.Empresa_Id" & vbCrLf & _
               "	 and #Disponibilidade2.EndEmpresa  = sb.EndEmpresa_Id" & vbCrLf & _
               "	 and #Disponibilidade2.Deposito    = sb.Deposito" & vbCrLf & _
               "	 and #Disponibilidade2.EndDeposito = sb.EndDeposito;" & vbCrLf
        '	/***********************************************************************************/     
        '	/***********************************************************************************/          
        '	/***********************************************************************************/   




        ' sql &= "	select row_number()over(ORDER BY D.Deposito DESC) as Disparador," & vbCrLf & _
        sql &= " Select 0 as Disparador, D.Empresa, " & vbCrLf & _
               "		   D.EndEmpresa," & vbCrLf & _
               "		   Emp.Nome as NomeEmpresa," & vbCrLf & _
               "		   Emp.Cidade as CidadeEmpresa," & vbCrLf & _
               "		   Emp.Estado as EstadoEmpresa," & vbCrLf & _
               "		   D.Deposito," & vbCrLf & _
               "		   D.EndDeposito," & vbCrLf & _
               "		   Dep.Nome as NomeDeposito," & vbCrLf & _
               "		   Dep.Cidade as CidadeDeposito," & vbCrLf & _
               "		   Dep.Estado as EstadoDeposito," & vbCrLf & _
               "		   D.Inicial," & vbCrLf & _
               "		   D.LaudoSemNota," & vbCrLf & _
               "           D.LaudoSemNotaSaida," & vbCrLf & _
               "		   D.EntradaProducao," & vbCrLf & _
               "		   D.SaidaProducao," & vbCrLf & _
               "		   D.Recebimento," & vbCrLf & _
               "		   D.Fob," & vbCrLf & _
               "		   D.CompraPorto," & vbCrLf & _
               "		   D.CompraEmDeposito," & vbCrLf & _
               "		   D.EntradaTransferenciaFilial," & vbCrLf & _
               "           D.EntradaTransferenciaTitularidade, " & vbCrLf & _
               "		   D.EntradaDevolucaoVenda," & vbCrLf & _
               "		   D.EntradaDeposito," & vbCrLf & _
               "		   D.RetornoSimbolicoFormacaoDeLote," & vbCrLf & _
               "           D.RetornoDeDeposito," & vbCrLf & _
               "		   D.RetiradaPrestacaoServico," & vbCrLf & _
               "		   D.Retirada," & vbCrLf & _
               "		   D.RetiradaFob," & vbCrLf & _
               "           D.RetiradaFobTer," & vbCrLf & _
               "		   D.EmbarquePorto," & vbCrLf & _
               "		   D.SaidaTransferenciaFilial," & vbCrLf & _
               "		   D.SaidaDeposito," & vbCrLf & _
               "		   D.SaidaTransferenciaTitularidade," & vbCrLf & _
               "		   D.DevParaCompra," & vbCrLf & _
               "		   D.DevDeCompra," & vbCrLf & _
               "		   D.DevDeposito," & vbCrLf & _
               "           D.AcobertamentoFiscal," & vbCrLf & _
               "           D.DevAcobertamentoFiscal," & vbCrLf & _
               "           D.Quebras," & vbCrLf & _
               "           D.Sobras," & vbCrLf & _
               "           D.SobrasCliFor," & vbCrLf & _
               "		   D.EntradaProducao + D.Inicial + D.Recebimento + D.Fob + D.CompraPorto + D.CompraEmDeposito + D.EntradaTransferenciaFilial + D.EntradaTransferenciaTitularidade + D.EntradaDevolucaoVenda" & vbCrLf & _
               "           + D.EntradaDeposito -  D.RetiradaPrestacaoServico + RetornoSimbolicoFormacaoDeLote + D.RetornoDeDeposito -D.SaidaProducao  -  D.Retirada - D.RetiradaFob - D.RetiradaFobTer - D.EmbarquePorto " & vbCrLf & _
               "		   - D.SaidaTransferenciaFilial - D.SaidaTransferenciaTitularidade - D.SaidaDeposito - DevDeposito - D.DevParaCompra - D.DevDeCompra - Quebras + Sobras as EstoqueAtual" & vbCrLf

        If objDisponibilidade.Consolidado Then
            sql &= "   into #Consolidado"
        End If

        sql &= "	  from #Disponibilidade2 D" & vbCrLf & _
               "	 inner Join Clientes Emp" & vbCrLf & _
               "		on D.Empresa    = Emp.Cliente_id" & vbCrLf & _
               "	   and D.EndEmpresa = Emp.Endereco_id" & vbCrLf & _
               "	 inner Join Clientes Dep" & vbCrLf & _
               "		on D.Deposito    = Dep.Cliente_id" & vbCrLf & _
               "	   and D.EndDeposito = Dep.Endereco_id" & vbCrLf

        If Not objDisponibilidade.Consolidado Then
            sql &= "	order by case when left(D.Empresa,8) = left(D.Deposito,8) and D.EndDeposito = 99 then 1 " & vbCrLf & _
                   "	              when left(D.Empresa,8) = left(D.Deposito,8)                        then 2 " & vbCrLf & _
                   "	              else 3 " & vbCrLf & _
                   "	         end, Dep.reduzido, D.Deposito; " & vbCrLf
        End If

        If objDisponibilidade.Consolidado Then
            'sql &= "Select row_number()over(ORDER BY CO.Deposito DESC) as Disparador," & vbCrLf & _
            sql &= " Select 0 as disparador, c.Cliente_id as Empresa," & vbCrLf & _
                   "       c.Endereco_id as EndEmpresa," & vbCrLf & _
                   "       c.Nome as NomeEmpresa," & vbCrLf & _
                   "       c.Cidade as CidadeEmpresa," & vbCrLf & _
                   "       c.Estado as EstadoEmpresa," & vbCrLf & _
                   "       CO.Deposito," & vbCrLf & _
                   "       CO.EndDeposito," & vbCrLf & _
                   "       CO.NomeDeposito," & vbCrLf & _
                   "       CO.CidadeDeposito," & vbCrLf & _
                   "       CO.EstadoDeposito," & vbCrLf & _
                   "       sum(CO.Inicial) as inicial," & vbCrLf & _
                   "	   sum(CO.LaudoSemNota) as LaudoSemNota," & vbCrLf & _
                   "       sum(CO.LaudoSemNotaSaida) as LaudoSemNotaSaida," & vbCrLf & _
                   "	   sum(CO.EntradaProducao) as EntradaProducao," & vbCrLf & _
                   "	   sum(CO.SaidaProducao) as SaidaProducao," & vbCrLf & _
                   "	   sum(CO.Recebimento) as Recebimento," & vbCrLf & _
                   "	   sum(CO.Fob) as Fob," & vbCrLf & _
                   "	   sum(CO.CompraPorto) as CompraPorto," & vbCrLf & _
                   "	   sum(CO.CompraEmDeposito) as CompraEmDeposito," & vbCrLf & _
                   "	   sum(CO.EntradaTransferenciaFilial) as EntradaTransferenciaFilial," & vbCrLf & _
                   "       sum(CO.EntradaTransferenciaTitularidade) as EntradaTransferenciaTitularidade," & vbCrLf & _
                   "	   sum(CO.EntradaDevolucaoVenda) as EntradaDevolucaoVenda," & vbCrLf & _
                   "	   sum(CO.EntradaDeposito) as EntradaDeposito," & vbCrLf & _
                   "	   sum(CO.RetornoSimbolicoFormacaoDeLote) as RetornoSimbolicoFormacaoDeLote," & vbCrLf & _
                   "       sum(CO.RetornoDeDeposito) as RetornoDeDeposito," & vbCrLf & _
                   "	   sum(CO.RetiradaPrestacaoServico) as RetiradaPrestacaoServico," & vbCrLf & _
                   "	   sum(CO.Retirada) as Retirada," & vbCrLf & _
                   "	   sum(CO.RetiradaFob) as RetiradaFob," & vbCrLf & _
                   "       sum(CO.RetiradaFobTer) as RetiradaFobTer," & vbCrLf & _
                   "	   sum(CO.EmbarquePorto) as EmbarquePorto," & vbCrLf & _
                   "	   sum(CO.SaidaTransferenciaFilial) as SaidaTransferenciaFilial," & vbCrLf & _
                   "	   sum(CO.SaidaDeposito) as SaidaDeposito," & vbCrLf & _
                   "	   sum(CO.SaidaTransferenciaTitularidade) as SaidaTransferenciaTitularidade," & vbCrLf & _
                   "	   sum(CO.DevParaCompra) as DevParaCompra," & vbCrLf & _
                   "	   sum(CO.DevDeCompra) as DevDeCompra," & vbCrLf & _
                   "	   sum(CO.DevDeposito) as DevDeposito," & vbCrLf & _
                   "       sum(CO.AcobertamentoFiscal) as AcobertamentoFiscal," & vbCrLf & _
                   "       sum(CO.DevAcobertamentoFiscal) as DevAcobertamentoFiscal," & vbCrLf & _
                   "       sum(CO.Quebras) as Quebras," & vbCrLf & _
                   "       sum(CO.Sobras)  as Sobras," & vbCrLf & _
                   "       sum(CO.SobrasCliFor) as SobrasCliFor," & vbCrLf & _
                   "       sum(CO.EstoqueAtual) as EstoqueAtual" & vbCrLf & _
                   "  from #Consolidado CO" & vbCrLf & _
                   " inner join ClientesxEmpresas CxE" & vbCrLf & _
                   "    on left(CO.Empresa,8) = left(CxE.Empresa_Id,8)" & vbCrLf & _
                   "   and CxE.Matriz             = 'S'" & vbCrLf & _
                   " inner Join Clientes C" & vbCrLf & _
                   "    on CxE.Empresa_id    = C.Cliente_id" & vbCrLf & _
                   "   and CxE.EndEmpresa_id = C.Endereco_Id" & vbCrLf & _
                   " where left(CxE.Empresa_Id,8) = @Empresa" & vbCrLf & _
                   " group by c.Cliente_id," & vbCrLf & _
                   "		  c.Endereco_id," & vbCrLf & _
                   "		  c.Nome," & vbCrLf & _
                   "		  c.Cidade," & vbCrLf & _
                   "		  c.Estado," & vbCrLf & _
                   "		  CO.Deposito," & vbCrLf & _
                   "		  CO.EndDeposito," & vbCrLf & _
                   "		  CO.NomeDeposito," & vbCrLf & _
                   "		  CO.CidadeDeposito," & vbCrLf & _
                   "		  CO.EstadoDeposito" & vbCrLf & _
                   " order by case when left(C.Cliente_id,8) = left(CO.Deposito,8) and CO.EndDeposito = 99 then 1 " & vbCrLf & _
                   "	           when left(C.Cliente_id,8) = left(CO.Deposito,8)                        then 2 " & vbCrLf & _
                   "	           else 3 " & vbCrLf & _
                   "	       end, CO.Deposito; " & vbCrLf & _
                   "	drop table #Consolidado  "
        End If

        sql &= "	drop table #Disponibilidade;" & vbCrLf & _
               "	drop table #Disponibilidade2;" & vbCrLf & _
               "	drop table #SaldoInicial;" & vbCrLf
        Return sql
    End Function

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        If ddlEmpresa.SelectedIndex > 0 Then
            Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
            objListDisp = New ListDisponibilidade(Emp(0), Emp(1))
            gridDisponibilidade.DataSource = objListDisp.ToArray
            gridDisponibilidade.DataBind()
            SessaoSalvarListaDisponibilidade()
        Else
            gridDisponibilidade.DataSource = Nothing
            gridDisponibilidade.DataBind()
        End If

    End Sub

    Protected Sub btnDeposito_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDeposito.Click
        ucConsultaCliente.Limpar()
        Dim txtNome As TextBox = CType(ucConsultaCliente.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeClientes(Me, "objClienteDisp" & HID.Value, txtNome.ClientID, True, 500)
    End Sub

    Protected Sub imgAdicionarDeposito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgAdicionarDeposito.Click
        If txtNomeCliente.Text.Length = 0 Then Exit Sub
        SessaoRecuperarDisponibilidade()
        'If objDisponibilidade.IUD <> "I" Then
        '    MsgBox(Me.Page, "Inicie uma nova configuracao de disponibilidade para incluir os saldos dos depositos")
        '    Exit Sub
        'End If

        Dim dep As String() = hdfCodigoCliente.Value.Split("-")
        Dim dispdep As New DisponibilidadeDeposito(objDisponibilidade)
        dispdep.IUD = "I"
        dispdep.CodigoDeposito = dep(0)
        dispdep.EndDeposito = dep(1)
        dispdep.SaldoInicialFiscal = txtQtdeFiscal.Text
        dispdep.SaldoinicialFisico = txtQtdeFisica.Text
        objDisponibilidade.Depositos.Add(dispdep)

        SessaoSalvarDisponibilidade()

        gridDeposito.DataSource = objDisponibilidade.Depositos.ToArray
        gridDeposito.DataBind()
        limparCamposDeposito()
    End Sub

    Public Function ValidaDisponibilidade() As Boolean
        If ddlEmpresaConfiguracao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione uma empresa")
            Return False
        End If

        If Not ucSelecaoProduto.TemSelecionado Then
            MsgBox(Me.Page, "Selecione ao Menos 1 Produto para Disponibilidade")
            Return False
        End If

        If Not IsDate(txtDataInicioDisponibilidade.Text) Then
            MsgBox(Me.Page, "Informe uma Data Valida para o Inicio da Apuracao da Disponibilidade")
            Return False
        End If


        If Not IsDate(txtDataCargaPatio.Text) Then
            MsgBox(Me.Page, "Informe a Data de Carga Patio")
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim imgbtn As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(imgbtn.NamingContainer, GridViewRow)
            Ano = CDate(row.Cells(7).Text).Year
            Disponibilidade(True, row.RowIndex)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim imgbtn As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(imgbtn.NamingContainer, GridViewRow)
            Ano = CDate(row.Cells(7).Text).Year
            Disponibilidade(False, row.RowIndex)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridDisponibilidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles gridDisponibilidade.SelectedIndexChanged
        SessaoRecuperarListaDisponibilidade()
        objDisponibilidade = objListDisp(gridDisponibilidade.SelectedIndex)

        ddlEmpresaConfiguracao.SelectedValue = objDisponibilidade.CodigoEmpresa & "-" & objDisponibilidade.EndEmpresa
        chkConsolidarEmpresa.Checked = objDisponibilidade.Consolidado
        txtDataCargaPatio.Text = objDisponibilidade.DataInicialCargaPatio.ToString("dd-MM-yyyy")
        txtDataInicioDisponibilidade.Text = objDisponibilidade.DataInicial.ToString("dd-MM-yyyy")

        Dim Produtos As String
        Produtos = ";"
        For Each obj In objDisponibilidade.Produtos
            Produtos &= obj.CodigoProduto & ","
        Next

        ucSelecaoProduto.PreenheGrid(Produtos)

        gridDeposito.DataSource = objDisponibilidade.Depositos.ToArray
        gridDeposito.DataBind()

        SessaoSalvarDisponibilidade()
        tcDisponibilidade.ActiveTabIndex = 1
    End Sub

    Protected Sub lnkExcluirDisponibilidade_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim imgbtn As LinkButton = CType(sender, LinkButton)
        Dim row As GridViewRow = CType(imgbtn.NamingContainer, GridViewRow)

        SessaoRecuperarListaDisponibilidade()
        objDisponibilidade = objListDisp(row.RowIndex)
        objDisponibilidade.IUD = "D"
        If objDisponibilidade.Salvar Then
            MsgBox(Me.Page, "Excluído com Sucesso.", eTitulo.Sucess)
            objListDisp.RemoveAt(row.RowIndex)
            gridDisponibilidade.DataSource = objListDisp.ToArray
            gridDisponibilidade.DataBind()
            SessaoSalvarListaDisponibilidade()
        Else
            MsgBox(Me.Page, "Falha ao Excluir o Registro!")
        End If
    End Sub

    Protected Sub lnkExcluirDeposito_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim imgbtn As LinkButton = CType(sender, LinkButton)
        Dim row As GridViewRow = CType(imgbtn.NamingContainer, GridViewRow)

        SessaoRecuperarDisponibilidade()
        objDisponibilidade.Depositos(row.RowIndex).IUD = "D"

        If objDisponibilidade.Depositos(row.RowIndex).Salvar Then
            MsgBox(Me.Page, "Excluído com Sucesso.", eTitulo.Sucess)
            objDisponibilidade.Depositos.RemoveAt(row.RowIndex)
            gridDeposito.DataSource = objDisponibilidade.Depositos.ToArray
            gridDeposito.DataBind()
            SessaoSalvarDisponibilidade()
        Else
            MsgBox(Me.Page, "Falha ao Excluir o Registro!")
        End If
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If ValidaDisponibilidade() Then
            SessaoRecuperarDisponibilidade()
            Dim emp As String() = ddlEmpresaConfiguracao.SelectedValue.Split("-")
            objDisponibilidade.IUD = "I"
            objDisponibilidade.CodigoEmpresa = emp(0)
            objDisponibilidade.EndEmpresa = emp(1)
            objDisponibilidade.Consolidado = chkConsolidarEmpresa.Checked
            objDisponibilidade.DataInicial = CDate(txtDataInicioDisponibilidade.Text)
            objDisponibilidade.DataInicialCargaPatio = CDate(txtDataCargaPatio.Text)


            Dim prd As ArrayList = ucSelecaoProduto.GetSqlEParametrosRelatorio("Grupo", "Produto_Id")
            objDisponibilidade.Produtos = New ListDisponibilidadeProduto(objDisponibilidade, prd(0))

            If objDisponibilidade.Salvar Then
                MsgBox(Me.Page, "Configuracao Salva com Sucesso.", eTitulo.Sucess)
                objDisponibilidade.IUD = ""
                ddlEmpresa.SelectedIndex = ddlEmpresaConfiguracao.SelectedIndex
                objListDisp = New ListDisponibilidade(emp(0), emp(1))
                gridDisponibilidade.DataSource = objListDisp.ToArray
                gridDisponibilidade.DataBind()
                SessaoSalvarListaDisponibilidade()
                tcDisponibilidade.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Erro ao Salvar a configuracao da disponibilidade")
            End If
        End If
    End Sub

    Protected Sub limparCamposDeposito()
        txtNomeCliente.Text = ""
        txtQtdeFiscal.Text = ""
        txtQtdeFisica.Text = ""
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Disponibilidade")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaConf_Click(sender As Object, e As EventArgs) Handles lnkAjudaConf.Click
        Try
            Funcoes.Ajuda(Me.Page, "Disponibilidade")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class