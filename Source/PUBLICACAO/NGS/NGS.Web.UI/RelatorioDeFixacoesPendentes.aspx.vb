Imports System.Data
Imports System.IO
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioDeFixacoesPendentes
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeFixacoesPendentes", "ACESSAR") Then
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
                    ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)

                    ChkPeriodo.Checked = False
                    txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")

                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    ucConsultaPedidos.SetarHID(HID.Value)

                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(sender As Object, e As EventArgs) Handles cmdBuscaCliente.Click
        Try
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteOrigem" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objClienteOrigem" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteOrigem" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtCliente.Text = itemCliente.Text
            HDCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteOrigem" & HID.Value)
        ElseIf Session("objPedido" & HID.Value) IsNot Nothing Then
            Dim objPedido As [Lib].Negocio.Pedido = CType(Session("objPedido" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = objPedido.Codigo
            Session.Remove("objPedido" & HID.Value)
        End If
    End Sub

    Protected Sub cmdBuscaPedido_Click(sender As Object, e As EventArgs) Handles cmdBuscaPedido.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) OrElse String.IsNullOrWhiteSpace(HDCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido")
            Else
                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("pedido", txtPedido.Text)
                parameters.Add("empresa", ddlEmpresa.SelectedValue.Split("-")(0))
                parameters.Add("enderecoEmpresa", ddlEmpresa.SelectedValue.Split("-")(1))
                parameters.Add("cliente", HDCodigoCliente.Value.Split("-")(0))
                parameters.Add("enderecoCliente", HDCodigoCliente.Value.Split("-")(1))
                parameters.Add("CodigoSafra", ddlSafra.SelectedValue)
                Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
                Session("ssTipoRetorno") = "objPedidos" & HID.Value

                Dim numberRows As Integer = ucConsultaPedidos.BindGridView(parameters)
                If numberRows > 1 Then
                    Popup.ConsultaDePedidos(Me.Page, "objPedido" & HID.Value)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        ddlClasse.SelectedIndex = 0
        txtCliente.Text = ""
        HDCodigoCliente.Value = ""
        ddlClasse.SelectedIndex = 0
        ddlSafra.SelectedIndex = 0
        txtPedido.Text = ""
        pnlData.Visible = ChkPeriodo.Checked
        Funcoes.VerificaEmpresa(ddlEmpresa)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Selecione a empresa.")
            Return False
        ElseIf ChkPeriodo.Checked Then
            If String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) OrElse _
               Not IsDate(txtDataInicial.Text) OrElse Not IsDate(txtDataFinal.Text) Then
                MsgBox(Me.Page, "Informe datas válidas.")
                Return False
            ElseIf CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
                MsgBox(Me.Page, "Data inicial não pode ser maior que a data final.")
                Return False
            End If
        End If
        Return True
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidarCampos() Then
                BuscarRegistros()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BuscarRegistros()
        Dim Parametros As String = ""
        Dim ds As New DataSet
        Dim sql As String = ""

        Dim Emp() As String = ddlEmpresa.SelectedValue.Split("-")
        Dim Cli() As String = HDCodigoCliente.Value.Split("-")

        'Moeda exibicao
        If ChkConsEmpresa.Checked Then Parametros &= "Empresas Consolidadas" & vbCrLf
        If ddlEmpresa.SelectedIndex > 0 Then Parametros &= ddlEmpresa.SelectedItem.Text & vbCrLf
        If HDCodigoCliente.Value.Length > 0 Then
            If chkConsCliente.Checked Then Parametros &= "Cliente Consolidado" & vbCrLf
            Parametros &= txtCliente.Text
        End If
        If ddlClasse.SelectedIndex > 0 Then Parametros &= ddlClasse.SelectedItem.Text & vbCrLf
        If ddlSafra.SelectedIndex > 0 Then Parametros &= ddlSafra.SelectedItem.Text & vbCrLf
        If txtPedido.Text.Length > 0 Then Parametros &= "Pedido: " & txtPedido.Text & vbCrLf
        If ChkPeriodo.Checked Then
            Parametros &= "Periodo de: " & txtDataInicial.Text.ToStrDate() & " a " & txtDataFinal.Text.ToStrDate()
        End If

        sql &= "Select PxF.Empresa_Id," & vbCrLf & _
               "       PxF.EndEmpresa_Id," & vbCrLf & _
               "       Empresa.Nome + ' - ' + Empresa.Cidade + '/' + Empresa.Estado as NomeEmpresa," & vbCrLf & _
               "       P.Cliente," & vbCrLf & _
               "       P.EndCliente," & vbCrLf & _
               "       C.Nome + ' - ' + C.Cidade + '/' + C.Estado as NomeCliente," & vbCrLf & _
               "       p.Safra," & vbCrLf & _
               "       pxf.Pedido_Id," & vbCrLf & _
               "       pxf.Fixacao_Id," & vbCrLf & _
               "       pxf.Produto_id," & vbCrLf & _
               "       prd.Nome as NomeProduto," & vbCrLf & _
               "       OPFix.Classe as ClasseOperacao," & vbCrLf & _
               "       pxf.Operacao," & vbCrLf & _
               "       pxf.SubOperacao," & vbCrLf & _
               "       sofix.Descricao as NomeOperacao," & vbCrLf & _
               "       pxf.Movimento," & vbCrLf & _
               "       pxf.Quantidade," & vbCrLf & _
               "       M.Cifrao as Moeda," & vbCrLf & _
               "       case" & vbCrLf & _
               "         when M.Classificacao = 'O'" & vbCrLf & _
               "           then UnitarioOficial" & vbCrLf & _
               "           else UnitarioMoeda" & vbCrLf & _
               "       end as Unitario," & vbCrLf & _
               "       case" & vbCrLf & _
               "         when M.Classificacao = 'O'" & vbCrLf & _
               "           then TotalOficial" & vbCrLf & _
               "           else TotalMoeda" & vbCrLf & _
               "       end as Total" & vbCrLf & _
               "       into #base" & vbCrLf & _
               "  from Pedidos P" & vbCrLf & _
               " inner join PedidosXItensXFixacoes pxf" & vbCrLf & _
               "    on P.Empresa_Id    = pxf.Empresa_Id" & vbCrLf & _
               "   and P.EndEmpresa_Id = pxf.EndEmpresa_Id" & vbCrLf & _
               "   and P.Pedido_Id     = pxf.Pedido_Id" & vbCrLf & _
               " Inner join SubOperacoes SO" & vbCrLf & _
               "    on p.Operacao    = so.Operacao_Id" & vbCrLf & _
               "   and p.SubOperacao = so.SubOperacoes_Id" & vbCrLf & _
               " Inner join Operacoes OPFix" & vbCrLf & _
               "    on p.Operacao    = OPFix.Operacao_Id" & vbCrLf & _
               " Inner join SubOperacoes SOFix" & vbCrLf & _
               "    on p.Operacao    = SOFix.Operacao_Id" & vbCrLf & _
               "   and p.SubOperacao = SOFix.SubOperacoes_Id" & vbCrLf & _
               " Inner Join Clientes Empresa" & vbCrLf & _
               "    on P.Empresa_id    = Empresa.Cliente_Id" & vbCrLf & _
               "   and P.EndEmpresa_id = Empresa.Endereco_id" & vbCrLf & _
               " Inner Join Clientes C" & vbCrLf & _
               "    on P.Cliente    = C.Cliente_Id" & vbCrLf & _
               "   and P.EndCliente = C.Endereco_id" & vbCrLf & _
               " Inner Join Produtos prd" & vbCrLf & _
               "    on Prd.Produto_id = pxf.produto_id" & vbCrLf & _
               " Inner Join Moedas M" & vbCrLf & _
               "    on M.Moeda_Id = P.Moeda" & vbCrLf & _
               "  left join NotasFiscaisxitens nxi" & vbCrLf & _
               "    on pxf.Empresa_Id    = nxi.Empresa_Id" & vbCrLf & _
               "   and pxf.EndEmpresa_Id = nxi.EndEmpresa_Id" & vbCrLf & _
               "   and pxf.Pedido_Id     = nxi.Pedido" & vbCrLf & _
               "   and pxf.Fixacao_Id    = nxi.Fixacao" & vbCrLf & _
               "  where nxi.Empresa_Id is null" & vbCrLf & _
               "    and p.Situacao = 1" & vbCrLf & _
               "    and so.Classe = '" & eClassesOperacoes.AFIXAR.ToString & "'" & vbCrLf & _
               "    and isnull(P.FiscalAberto,0) = 1" & vbCrLf

        If ChkConsEmpresa.Checked Then
            sql &= "    and left(PxF.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
        Else
            sql &= "    and PxF.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                   "    and PxF.EndEmpresa_id = " & Emp(1)
        End If

        If Not String.IsNullOrWhiteSpace(HDCodigoCliente.Value) Then
            If chkConsCliente.Checked Then
                sql &= "    and left(P.Cliente,8) = '" & Cli(0).Substring(0, 8) & "'" & vbCrLf
            Else
                sql &= "    and P.Cliente    ='" & Cli(0) & "'" & vbCrLf & _
                       "    and P.EndCliente = " & Cli(1)
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            sql &= "    and pxf.Pedido_Id =" & txtPedido.Text & vbCrLf
        End If

        If ddlClasse.SelectedIndex > 0 Then
            sql &= "    and OPFix.Classe = '" & ddlClasse.SelectedValue & "'" & vbCrLf
        End If

        If ddlSafra.SelectedIndex > 0 Then
            sql &= "    and p.Safra = '" & ddlSafra.SelectedValue & "'" & vbCrLf
        End If

        If ChkPeriodo.Checked Then
            sql &= "    and pxf.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf
        End If

        If ucSelecaoProdutoVenda.TemSelecionado Then
            Dim ret As New ArrayList
            ret = ucSelecaoProdutoVenda.GetSqlEParametrosRelatorio("prd.Grupo", "pxf.Produto_Id", "", True)
            sql &= " and " & ret(0) & vbCrLf
            Parametros &= ret(1)
        End If
        sql &= "   order by PxF.Empresa_Id, P.Cliente, P.EndCliente, pxf.Pedido_Id, pxf.Movimento " & vbCrLf

        sql &= "   select Empresa_Id, EndEmpresa_Id, NomeEmpresa, Cliente, EndCliente, NomeCliente, Safra,                " & vbCrLf & _
               "           Pedido_Id, Fixacao_Id, Produto_id, NomeProduto, ClasseOperacao, Operacao, SubOperacao, " & vbCrLf & _
               "           NomeOperacao, Movimento, Quantidade, Moeda, Unitario, Total                                    " & vbCrLf & _
               "       from #base                                                                                         " & vbCrLf & _
               "                                                                                                          " & vbCrLf & _
               "   select Empresa_id,                                                                                     " & vbCrLf & _
               "          Endempresa_id,                                                                                  " & vbCrLf & _
               "          cliente,                                                                                        " & vbCrLf & _
               "          Endcliente,                                                                                     " & vbCrLf & _
               "          Moeda,                                                                                          " & vbCrLf & _
               "          SUM(quantidade) as Quantidade,                                                                        " & vbCrLf & _
               "          SUM(total) / SUM(quantidade) as Unitario,                                                           " & vbCrLf & _
               "          SUM(total) as Total                                                                             " & vbCrLf & _
               "     from #base                                                                                           " & vbCrLf & _
               "    group by Empresa_id,                                                                                  " & vbCrLf & _
               "             Endempresa_id,                                                                               " & vbCrLf & _
               "             cliente,                                                                                     " & vbCrLf & _
               "             Endcliente,                                                                                  " & vbCrLf & _
               "             Moeda                                                                                        " & vbCrLf & _
               "                                                                                                          " & vbCrLf & _
               "    select Empresa_id,                                                                                    " & vbCrLf & _
               "          Endempresa_id,                                                                                  " & vbCrLf & _
               "          Moeda,                                                                                          " & vbCrLf & _
               "          SUM(quantidade) as Quantidade,                                                                        " & vbCrLf & _
               "          SUM(total) / SUM(quantidade) as Unitario,                                                           " & vbCrLf & _
               "          SUM(total) as Total                                                                             " & vbCrLf & _
               "     from #base                                                                                           " & vbCrLf & _
               "    group by Empresa_id,                                                                                  " & vbCrLf & _
               "             Endempresa_id,                                                                               " & vbCrLf & _
               "             Moeda                                                                                        " & vbCrLf & _
               "                                                                                                          " & vbCrLf & _
               "   select Moeda,                                                                                          " & vbCrLf & _
               "          SUM(quantidade) as Quantidade,                                                                        " & vbCrLf & _
               "          SUM(total) / SUM(quantidade) as Unitario,                                                           " & vbCrLf & _
               "          SUM(total) as Total                                                                             " & vbCrLf & _
               "     from #base                                                                                           " & vbCrLf & _
               "    group by Moeda                                                                                        " & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "FixacoesPendentes")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count = 4 Then
            ds.Tables(1).TableName = "ResumoCliente"
            ds.Tables(2).TableName = "ResumoEmpresa"
            ds.Tables(3).TableName = "ResumoGeral"

            AlimentaCrptRelatorios(ds, "Cr_FixacoesPendentes", Parametros)
        Else
            MsgBox(Me.Page, "Nenhum resultado encontrado.")
        End If



    End Sub

    Public Sub AlimentaCrptRelatorios(ByVal Ds As DataSet, ByVal Caminho As String, ByVal Parametros As String)
        Try
            Dim objEmpresa As Cliente = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Parametros", Parametros)
            parameters.Add("NomeEmpresa", objEmpresa.Nome)
            parameters.Add("CidadeEmpresa", objEmpresa.Cidade & " - " & objEmpresa.Estado.Descricao)

            Funcoes.BindReport(Me.Page, Ds, Caminho, eExportType.PDF, parameters)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub ChkPeriodo_CheckedChanged(sender As Object, e As EventArgs) Handles ChkPeriodo.CheckedChanged
        Try
            pnlData.Visible = ChkPeriodo.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeFixacoesPendentes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class