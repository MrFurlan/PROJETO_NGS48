Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaNotaVendaAOrdem
    Inherits BaseUserControl

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            txtAno.Text = Today.Year.ToString()
        End If
    End Sub

    Protected Sub gridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim notaFiscal As New [Lib].Negocio.NotaFiscal
        notaFiscal.CodigoEmpresa = gridNotas.SelectedRow.Cells(2).Text
        notaFiscal.EnderecoEmpresa = gridNotas.SelectedRow.Cells(3).Text
        notaFiscal.CodigoCliente = gridNotas.SelectedRow.Cells(7).Text
        notaFiscal.EnderecoCliente = gridNotas.SelectedRow.Cells(8).Text
        notaFiscal.Codigo = gridNotas.SelectedRow.Cells(13).Text
        notaFiscal.Serie = gridNotas.SelectedRow.Cells(14).Text
        notaFiscal.EntradaSaida = IIf(gridNotas.SelectedRow.Cells(12).Text = "S", [Lib].Negocio.eEntradaSaida.Saida, [Lib].Negocio.eEntradaSaida.Entrada)

        Dim Nota As New [Lib].Negocio.NotaFiscal(notaFiscal)

        'Nao atende a Revenda 12/02/2016
        'Nota.Itens(0).SaldoPedidoFiscal = CDec(gridNotas.SelectedRow.Cells(17).Text) - CDec(gridNotas.SelectedRow.Cells(18).Text)

        Popup.CloseDialog(Me.Page, "divConsultaNotaVendaAOrdem")
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objVendaAOrdemNXI")) Then
            Session("objVendaAOrdemNXI" & HID.Value.ToString) = Nota
            If TypeOf Me.Page Is NotaFiscalXItens Then
                CType(Me.Page, NotaFiscalXItens).CarregarVendaAOrdem()
            End If
        End If
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaNotaVendaAOrdem")
        If TypeOf Me.Page Is NotaFiscalXItens Then
            CType(Me.Page, NotaFiscalXItens).CarregarVendaAOrdem()
        End If
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), NotaFiscal)
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Sub BindGridView()
        SessaoRecuperaNotaFiscal()
        Dim Prod As String = objNotaFiscal.Itens(0).CodigoProduto
        Dim Emp As String = objNotaFiscal.CodigoEmpresa
        Dim EndEmp As String = objNotaFiscal.EnderecoEmpresa
        Dim ds As DataSet
        Dim sql As String
        Dim prd As New [Lib].Negocio.Produto(Prod)

        'sql = "select CONVERT(varchar, NF.Movimento, 103) as Movimento, NF.Empresa_Id, NF.EndEmpresa_Id,  " & vbCrLf & _
        '        "		Empresa.Nome As EmpresaNome, Empresa.Cidade as EmpresaCidade, Empresa.Estado as EmpresaEstado, " & vbCrLf & _
        '        "		NF.Cliente_Id, NF.EndCliente_Id,  " & vbCrLf & _
        '        "		Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf & _
        '        "		NF.EntradaSaida_Id, " & vbCrLf & _
        '        "		NF.Serie_Id, " & vbCrLf & _
        '        "		NF.Nota_Id, " & vbCrLf & _
        '        "		NFxI.Produto_id, " & vbCrLf & _
        '        "		P.Nome as NomeProduto, " & vbCrLf & _
        '        "		P.Unidade, " & vbCrLf & _
        '        "       NF.Movimento as Data, " & vbCrLf & _
        '        "		max(case " & vbCrLf & _
        '        "		       when NFxI.QuantidadeFisica > 0 " & vbCrLf & _
        '        "		          then NFxI.QuantidadeFisica " & vbCrLf & _
        '        "		          else NFxI.QuantidadeFiscal " & vbCrLf & _
        '        "		       end) AS QtdeNota, " & vbCrLf & _
        '        "		max(isnull(NOri.EntregueNota,0)) AS Entregue " & vbCrLf & _
        '        "	from NotasFiscais NF " & vbCrLf & _
        '        "			Inner Join NotasFiscaisxItens as NFxI " & vbCrLf & _
        '        "					ON NF.Empresa_Id      = NFxI.Empresa_Id  " & vbCrLf & _
        '        "					AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id  " & vbCrLf & _
        '        "					AND NF.Cliente_Id      = NFxI.Cliente_Id  " & vbCrLf & _
        '        "					AND NF.EndCliente_Id   = NFxI.EndCliente_Id  " & vbCrLf & _
        '        "					AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id  " & vbCrLf & _
        '        "					AND NF.Serie_Id        = NFxI.Serie_Id  " & vbCrLf & _
        '        "					AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
        '        "			left Join (Select nXn.OrigemEmpresa_id, nXn.OrigemEndEmpresa_id, nXn.OrigemCliente_id, nXn.OrigemEndCliente_id, " & vbCrLf & _
        '        "								nXn.OrigemEntradaSaida_id, nXn.OrigemSerie_id, nXn.OrigemNota_id,  " & vbCrLf & _
        '        "								sum(case " & vbCrLf & _
        '        "								when so.Devolucao = 'N' " & vbCrLf & _
        '        "								then nXi.QuantidadeFisica " & vbCrLf & _
        '        "								else nXi.QuantidadeFisica * -1 " & vbCrLf & _
        '        "								end) AS EntregueNota " & vbCrLf & _
        '        "						From NotasFiscaisXItens nXi " & vbCrLf & _
        '        "								inner join NotasFiscais as n " & vbCrLf & _
        '        "										ON n.Empresa_Id      = nXi.Empresa_Id  " & vbCrLf & _
        '        "										AND n.EndEmpresa_Id   = nXi.EndEmpresa_Id  " & vbCrLf & _
        '        "										AND n.Cliente_Id      = nXi.Cliente_Id  " & vbCrLf & _
        '        "										AND n.EndCliente_Id   = nXi.EndCliente_Id  " & vbCrLf & _
        '        "										AND n.EntradaSaida_Id = nXi.EntradaSaida_Id  " & vbCrLf & _
        '        "										AND n.Serie_Id        = nXi.Serie_Id  " & vbCrLf & _
        '        "										AND n.Nota_Id         = nXi.Nota_Id " & vbCrLf & _
        '        "								inner join NotasXNotas nXn  " & vbCrLf & _
        '        "										on nXn.Empresa_id       = nXi.Empresa_id " & vbCrLf & _
        '        "										and nXn.EndEmpresa_id   = nXi.EndEmpresa_id " & vbCrLf & _
        '        "										and nXn.Cliente_id      = nXi.Cliente_id " & vbCrLf & _
        '        "										and nXn.EndCliente_id   = nXi.EndCliente_id " & vbCrLf & _
        '        "										and nXn.EntradaSaida_id = nXi.EntradaSaida_id " & vbCrLf & _
        '        "										and nXn.Serie_id        = nXi.Serie_id " & vbCrLf & _
        '        "										and nXn.Nota_id         = nXi.Nota_id " & vbCrLf & _
        '        "								inner join SubOperacoes so  " & vbCrLf & _
        '        "									on so.Operacao_Id      = nXi.Operacao " & vbCrLf & _
        '        "									and so.SubOperacoes_Id = nXi.SubOperacao " & vbCrLf & _
        '        "								Inner Join Produtos Prd " & vbCrLf & _
        '        "									on Prd.Produto_id = nXi.Produto_id " & vbCrLf & _
        '        "						where n.Situacao        = 1  " & vbCrLf & _
        '        "						  and n.TipoDeDocumento = 1  " & vbCrLf & _
        '        "						  and n.Empresa_Id      = '" & Emp & "' " & vbCrLf & _
        '        "						  and n.EndEmpresa_Id   = " & EndEmp & vbCrLf & _
        '        "						  and Prd.Grupo         = '" & prd.CodigoGrupo & "' " & vbCrLf & _
        '        "						  and so.Classe = '" & eClassesOperacoes.CONTAEORDEM.ToString & "' " & vbCrLf & _
        '        "						Group By nXn.OrigemEmpresa_id, nXn.OrigemEndEmpresa_id, nXn.OrigemCliente_id, nXn.OrigemEndCliente_id, " & vbCrLf & _
        '        "								nXn.OrigemEntradaSaida_id, nXn.OrigemSerie_id, nXn.OrigemNota_id) AS NOri " & vbCrLf & _
        '        "					on NOri.OrigemEmpresa_Id = NF.Empresa_Id " & vbCrLf & _
        '        "					AND NOri.OrigemEndEmpresa_Id   = NF.EndEmpresa_Id  " & vbCrLf & _
        '        "					AND NOri.OrigemCliente_Id      = NF.Cliente_Id  " & vbCrLf & _
        '        "					AND NOri.OrigemEndCliente_Id   = NF.EndCliente_Id  " & vbCrLf & _
        '        "					AND NOri.OrigemEntradaSaida_Id = NF.EntradaSaida_Id  " & vbCrLf & _
        '        "					AND NOri.OrigemSerie_Id        = NF.Serie_Id  " & vbCrLf & _
        '        "					AND NOri.OrigemNota_Id         = NF.Nota_Id " & vbCrLf & _
        '        "			Inner Join Produtos P " & vbCrLf & _
        '        "					on P.Produto_id = NFxI.Produto_id " & vbCrLf & _
        '        "			Inner Join Clientes Empresa " & vbCrLf & _
        '        "					on Empresa.Cliente_id  = NF.Empresa_id " & vbCrLf & _
        '        "					and Empresa.Endereco_id = NF.EndEmpresa_Id " & vbCrLf & _
        '        "			Inner Join Clientes  " & vbCrLf & _
        '        "					on Clientes.Cliente_id  = NF.Cliente_Id " & vbCrLf & _
        '        "					and Clientes.Endereco_id = NF.EndCliente_Id " & vbCrLf & _
        '        "			Inner Join SubOperacoes  " & vbCrLf & _
        '        "					on SubOperacoes.Operacao_Id     = NF.Operacao " & vbCrLf & _
        '        "					and SubOperacoes.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
        '        "where NF.Situacao        = 1  " & vbCrLf & _
        '        "  and NF.TipoDeDocumento = 1  " & vbCrLf & _
        '        "  and NF.Empresa_Id      = '" & Emp & "' " & vbCrLf & _
        '        "  and NF.EndEmpresa_Id   = " & EndEmp & vbCrLf

        sql = "select CONVERT(varchar, NF.Movimento, 103) as Movimento, NF.Empresa_Id, NF.EndEmpresa_Id,  " & vbCrLf & _
               "		Empresa.Nome As EmpresaNome, Empresa.Cidade as EmpresaCidade, Empresa.Estado as EmpresaEstado, " & vbCrLf & _
               "		NF.Cliente_Id, NF.EndCliente_Id,  " & vbCrLf & _
               "		Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf & _
               "		NF.EntradaSaida_Id, " & vbCrLf & _
               "		NF.Serie_Id, " & vbCrLf & _
               "		NF.Nota_Id, " & vbCrLf & _
               "		NFxI.Produto_id, " & vbCrLf & _
               "		P.Nome as NomeProduto, " & vbCrLf & _
               "		P.Unidade, " & vbCrLf & _
               "       NF.Movimento as Data, " & vbCrLf & _
               "		max(NFxI.QuantidadeFiscal) AS QtdeNota, " & vbCrLf & _
               "		max(isnull(NOri.EntregueNota,0)) AS Entregue " & vbCrLf & _
               "	from NotasFiscais NF " & vbCrLf & _
               "			Inner Join NotasFiscaisxItens as NFxI " & vbCrLf & _
               "					ON NF.Empresa_Id      = NFxI.Empresa_Id  " & vbCrLf & _
               "					AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id  " & vbCrLf & _
               "					AND NF.Cliente_Id      = NFxI.Cliente_Id  " & vbCrLf & _
               "					AND NF.EndCliente_Id   = NFxI.EndCliente_Id  " & vbCrLf & _
               "					AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id  " & vbCrLf & _
               "					AND NF.Serie_Id        = NFxI.Serie_Id  " & vbCrLf & _
               "					AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
               "			left Join (Select nXn.OrigemEmpresa_id, nXn.OrigemEndEmpresa_id, nXn.OrigemCliente_id, nXn.OrigemEndCliente_id, " & vbCrLf & _
               "								nXn.OrigemEntradaSaida_id, nXn.OrigemSerie_id, nXn.OrigemNota_id,  " & vbCrLf & _
               "								sum(case " & vbCrLf & _
               "								when so.Devolucao = 'N' " & vbCrLf & _
               "								then nXi.QuantidadeFiscal " & vbCrLf & _
               "								else nXi.QuantidadeFiscal * -1 " & vbCrLf & _
               "								end) AS EntregueNota " & vbCrLf & _
               "						From NotasFiscaisXItens nXi " & vbCrLf & _
               "								inner join NotasFiscais as n " & vbCrLf & _
               "										ON n.Empresa_Id      = nXi.Empresa_Id  " & vbCrLf & _
               "										AND n.EndEmpresa_Id   = nXi.EndEmpresa_Id  " & vbCrLf & _
               "										AND n.Cliente_Id      = nXi.Cliente_Id  " & vbCrLf & _
               "										AND n.EndCliente_Id   = nXi.EndCliente_Id  " & vbCrLf & _
               "										AND n.EntradaSaida_Id = nXi.EntradaSaida_Id  " & vbCrLf & _
               "										AND n.Serie_Id        = nXi.Serie_Id  " & vbCrLf & _
               "										AND n.Nota_Id         = nXi.Nota_Id " & vbCrLf & _
               "								inner join NotasXNotas nXn  " & vbCrLf & _
               "										on nXn.Empresa_id       = nXi.Empresa_id " & vbCrLf & _
               "										and nXn.EndEmpresa_id   = nXi.EndEmpresa_id " & vbCrLf & _
               "										and nXn.Cliente_id      = nXi.Cliente_id " & vbCrLf & _
               "										and nXn.EndCliente_id   = nXi.EndCliente_id " & vbCrLf & _
               "										and nXn.EntradaSaida_id = nXi.EntradaSaida_id " & vbCrLf & _
               "										and nXn.Serie_id        = nXi.Serie_id " & vbCrLf & _
               "										and nXn.Nota_id         = nXi.Nota_id " & vbCrLf & _
               "								inner join SubOperacoes so  " & vbCrLf & _
               "									on so.Operacao_Id      = nXi.Operacao " & vbCrLf & _
               "									and so.SubOperacoes_Id = nXi.SubOperacao " & vbCrLf & _
               "								Inner Join Produtos Prd " & vbCrLf & _
               "									on Prd.Produto_id = nXi.Produto_id " & vbCrLf & _
               "						where n.Situacao        = 1  " & vbCrLf & _
               "						  and n.TipoDeDocumento = 1  " & vbCrLf & _
               "						  and n.Empresa_Id      = '" & Emp & "' " & vbCrLf & _
               "						  and n.EndEmpresa_Id   = " & EndEmp & vbCrLf & _
               "						  and Prd.Grupo         = '" & prd.CodigoGrupo & "' " & vbCrLf & _
               "						  and so.Classe = '" & eClassesOperacoes.CONTAEORDEM.ToString & "' " & vbCrLf & _
               "						Group By nXn.OrigemEmpresa_id, nXn.OrigemEndEmpresa_id, nXn.OrigemCliente_id, nXn.OrigemEndCliente_id, " & vbCrLf & _
               "								nXn.OrigemEntradaSaida_id, nXn.OrigemSerie_id, nXn.OrigemNota_id) AS NOri " & vbCrLf & _
               "					on NOri.OrigemEmpresa_Id = NF.Empresa_Id " & vbCrLf & _
               "					AND NOri.OrigemEndEmpresa_Id   = NF.EndEmpresa_Id  " & vbCrLf & _
               "					AND NOri.OrigemCliente_Id      = NF.Cliente_Id  " & vbCrLf & _
               "					AND NOri.OrigemEndCliente_Id   = NF.EndCliente_Id  " & vbCrLf & _
               "					AND NOri.OrigemEntradaSaida_Id = NF.EntradaSaida_Id  " & vbCrLf & _
               "					AND NOri.OrigemSerie_Id        = NF.Serie_Id  " & vbCrLf & _
               "					AND NOri.OrigemNota_Id         = NF.Nota_Id " & vbCrLf & _
               "			Inner Join Produtos P " & vbCrLf & _
               "					on P.Produto_id = NFxI.Produto_id " & vbCrLf & _
               "			Inner Join Clientes Empresa " & vbCrLf & _
               "					on Empresa.Cliente_id  = NF.Empresa_id " & vbCrLf & _
               "					and Empresa.Endereco_id = NF.EndEmpresa_Id " & vbCrLf & _
               "			Inner Join Clientes  " & vbCrLf & _
               "					on Clientes.Cliente_id  = NF.Cliente_Id " & vbCrLf & _
               "					and Clientes.Endereco_id = NF.EndCliente_Id " & vbCrLf & _
               "			Inner Join SubOperacoes  " & vbCrLf & _
               "					on SubOperacoes.Operacao_Id     = NF.Operacao " & vbCrLf & _
               "					and SubOperacoes.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
               "where NF.Situacao        = 1  " & vbCrLf & _
               "  and NF.TipoDeDocumento = 1  " & vbCrLf & _
               "  and NF.Empresa_Id      = '" & Emp & "' " & vbCrLf & _
               "  and NF.EndEmpresa_Id   = " & EndEmp & vbCrLf

        If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
            sql &= "  and NF.EntradaSaida_Id = 'E' " & vbCrLf
        Else
            sql &= "  and NF.EntradaSaida_Id = 'S' " & vbCrLf
        End If

        'sql &= "  and year(NF.Movimento) = '" & txtAno.Text & "' " & vbCrLf & _
        '        "  and P.Grupo            = '" & prd.CodigoGrupo & "' " & vbCrLf & _
        '        "  and SubOperacoes.Classe IN ('" & eClassesOperacoes.COMPRASAORDEM.ToString & "','" & eClassesOperacoes.VENDASAORDEM.ToString & "','" & eClassesOperacoes.INDUSTRIALIZACAO.ToString & "') " & vbCrLf & _
        '        "Group by CONVERT(varchar, NF.Movimento, 103), NF.Empresa_Id, NF.EndEmpresa_Id,  " & vbCrLf & _
        '        "			Empresa.Nome, Empresa.Cidade, Empresa.Estado, NF.Cliente_Id, NF.EndCliente_Id,  " & vbCrLf & _
        '        "			Clientes.Nome, Clientes.Cidade, Clientes.Estado,NF.EntradaSaida_Id, " & vbCrLf & _
        '        "			NF.Serie_Id, NF.Nota_Id, NFxI.Produto_id, P.Nome, P.Unidade, NF.Movimento " & vbCrLf & _
        '        "having (max(case " & vbCrLf & _
        '        "				when NFxI.QuantidadeFisica > 0 " & vbCrLf & _
        '        "				   then NFxI.QuantidadeFisica " & vbCrLf & _
        '        "				   else NFxI.QuantidadeFiscal " & vbCrLf & _
        '        "			   end) - max(isnull(NOri.EntregueNota,0))) > 0 " & vbCrLf & _
        '        "Order By NF.Movimento Desc "

        sql &= "  and year(NF.Movimento) = '" & txtAno.Text & "' " & vbCrLf & _
                "  and P.Grupo            = '" & prd.CodigoGrupo & "' " & vbCrLf & _
                "  and SubOperacoes.Classe IN ('" & eClassesOperacoes.COMPRASAORDEM.ToString & "','" & eClassesOperacoes.VENDASAORDEM.ToString & "','" & eClassesOperacoes.INDUSTRIALIZACAO.ToString & "') " & vbCrLf & _
                "Group by CONVERT(varchar, NF.Movimento, 103), NF.Empresa_Id, NF.EndEmpresa_Id,  " & vbCrLf & _
                "			Empresa.Nome, Empresa.Cidade, Empresa.Estado, NF.Cliente_Id, NF.EndCliente_Id,  " & vbCrLf & _
                "			Clientes.Nome, Clientes.Cidade, Clientes.Estado,NF.EntradaSaida_Id, " & vbCrLf & _
                "			NF.Serie_Id, NF.Nota_Id, NFxI.Produto_id, P.Nome, P.Unidade, NF.Movimento " & vbCrLf & _
                "having (max(NFxI.QuantidadeFiscal) - max(isnull(NOri.EntregueNota,0))) > 0 " & vbCrLf & _
                "Order By NF.Movimento Desc "

        ds = Banco.ConsultaDataSet(sql, "Notas")
        gridNotas.DataSource = ds
        gridNotas.DataBind()

        If ds.Tables(0).Rows.Count = 0 Then
            Session("SemVendaAOrdem" & HID.Value.ToString) = True
            Popup.CloseDialog(Me.Page, "divConsultaNotaVendaAOrdem")
        End If
    End Sub

    Protected Sub btnBuscar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnBuscar.Click
        BindGridView()
    End Sub

End Class