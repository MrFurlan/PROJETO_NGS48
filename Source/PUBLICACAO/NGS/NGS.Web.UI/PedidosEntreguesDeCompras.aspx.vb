Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class PedidosEntreguesDeCompras
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim Mensagem As String
    Dim Cliente() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PedidosEntreguesDeCompras", "ACESSAR") Then
                    CargaUnidade()
                    VerificaUnidade()
                    txtDataInicial.Text = String.Format("01/{0}/{1}", Month(DateTime.Now).ToString().PadLeft(2, "0"), Year(DateTime.Now))
                    txtDataFinal.Text = Now.ToString("dd/MM/yyyy")
                    GruposdeEstoques()
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
                    ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
                    ddl.Carregar(ddlGruposCFOP, CarregarDDL.Tabela.CFOPGrupo, "", True)
                    ddl.Carregar(ddlCentroDeCusto, CarregarDDL.Tabela.CentroDeCusto, "", True)
                    Panel3.Visible = False
                    GridProduto.Parent.Visible = False
                    GridCliente.Parent.Visible = False
                    divCentroDeCusto.Visible = False
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf &
              "FROM Clientes C " & vbCrLf &
              "INNER JOIN ClientesXTipos CT " & vbCrLf &
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
              "WHERE CT.Tipo_Id = 050 " & vbCrLf &
              "ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarSubOperacoes(cmbOperacao.SelectedValue)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Private Sub GruposdeEstoques()
        Sql = "  SELECT     distinct GruposDeEstoques.Grupo_Id as Codigo, GruposDeEstoques.Descricao" & vbCrLf & _
              " FROM         GruposDeEstoques INNER JOIN" & vbCrLf & _
              "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo" & vbCrLf & _
              " WHERE     (LEN(GruposDeEstoques.Grupo_Id) >= 5) Order by GruposDeEstoques.Descricao" & vbCrLf
        'Sql &= " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) and Produtos.Agrupar = 'N' order by GruposDeEstoques.Descricao"
        DS = Banco.ConsultaDataSet(Sql, "GruposDeEstoques")
    End Sub

    Private Sub Limpar()
        txtCliente.Text = ""
        txtCodigoCliente.Value = ""

        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()

        ddlGruposCFOP.SelectedIndex = 0
        Panel3.Visible = False

        lstCfop.Items.Clear()
        lstCfopSelecionados.Items.Clear()

        ucSelecaoProduto.Limpar()

        GridProduto.Parent.Visible = False
        GridCliente.Parent.Visible = False
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Unidade: " & ddlUnidade.SelectedItem.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Data: " & txtDataInicial.Text & " "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "à " & txtDataFinal.Text & " - "
        End If

        param &= IIf(RadCliente.Checked, "Visualizar: Por Cliente", "Visualizar: Por Produto")

        Return param
    End Function

    Private Function VerificarGrupoProduto(ByRef Sql As String) As String
        If ucSelecaoProduto.TemSelecionado() Then
            Dim retorno As ArrayList
            retorno = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", " NotasFiscaisXItens.Produto_id")
            Sql &= " AND " & retorno(0)
            Return retorno(1)
        Else
            Return ""
        End If
    End Function

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim NomeArquivo As String = "Manual/PedidosEntreguesDeVendas.mht"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(ddlUnidade.Text) Then
            MsgBox(Me.Page, "Informe a unidade.")
            Return False
        End If
        If String.IsNullOrWhiteSpace(DdlEmpresa.Text) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        End If

        If divCentroDeCusto.Visible AndAlso ddlCentroDeCusto.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Centro de Custo não foi selecionado.")
            Return False
        End If

        Return True
    End Function

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteCxN" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteCxN" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteCxN" & HID.Value)
        End If
    End Sub

    Protected Sub ddlCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCxN" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGruposCFOP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlGruposCFOP.SelectedIndex > 0 Then
                Panel3.Visible = True
                Dim ListaCFOP As New [Lib].Negocio.ListCFOP(True, ddlGruposCFOP.SelectedValue)

                lstCfop.Items.Clear()
                lstCfopSelecionados.Items.Clear()

                Dim j As Integer = 0
                While j < ListaCFOP.Count
                    lstCfop.Items.Add(New ListItem(Format(ListaCFOP(j).Codigo, "0000") & "-" & ListaCFOP(j).Descricao, ListaCFOP(j).Codigo))
                    j += 1
                End While
            Else
                Panel3.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgAdicionar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If lstCfop.SelectedIndex > -1 Then

                Dim temCFOP As Boolean = False

                Dim i As Integer = 0
                While i < lstCfopSelecionados.Items.Count
                    If lstCfopSelecionados.Items(i).Text = lstCfop.SelectedItem.Text Then
                        temCFOP = True
                    End If
                    i += 1
                End While

                If temCFOP Then
                    MsgBox(Me.Page, "CFOP já foi selecionado")
                Else
                    lstCfopSelecionados.Items.Add(New ListItem(lstCfop.SelectedItem.Text))
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemover_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If lstCfopSelecionados.SelectedIndex > -1 Then
                lstCfopSelecionados.Items.RemoveAt(lstCfopSelecionados.SelectedIndex)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            If divCentroDeCusto.Visible Then
                MsgBox(Me.Page, "Relatório não pode ser utilizado para Centro de Custo")
            Else
                EmitirRelatorio(True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            If divCentroDeCusto.Visible Then
                MsgBox(Me.Page, "Relatório não pode ser utilizado para Centro de Custo")
            Else
                EmitirRelatorio(False)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelCC_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcelCC.Click
        Try
            EmitirRelatorioCC()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Validar() Then
                Cliente = txtCodigoCliente.Value.Split("-")
                If RadCliente.Checked = True Then

                    'Sql = " SELECT  NotasFiscais.Nota_Id as Nota, NotasFiscais.Movimento as Data, Clientes.Nome AS Cliente, Clientes.Bairro, Clientes.Cidade AS Cidade, " & vbCrLf & _
                    '      "        Clientes.Estado AS Estado, SUM(NotasFiscaisXItens.QuantidadeFiscal) AS Quantidade, SUM(NotasFiscaisXItens.Valor) AS Valor" & vbCrLf & _
                    '      " FROM   NotasFiscais INNER JOIN" & vbCrLf & _
                    '      "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                    '      "        NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                    '      "        NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                    '      "        NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                    '      "          INNER jOIN SUBOPERACOES SO                              " & vbCrLf & _
                    '      "             ON NOTASFISCAIS.OPERACAO = SO.OPERACAO_ID            " & vbCrLf & _
                    '      "             AND NOTASFISCAIS.SUBOPERACAO = SO.SUBOPERACOES_ID    " & vbCrLf & _
                    '      "        INNER JOIN Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf & _
                    '      "        Pedidos ON NotasFiscais.Empresa_Id = Pedidos.Empresa_Id AND NotasFiscais.EndEmpresa_Id = Pedidos.EndEmpresa_Id AND " & vbCrLf & _
                    '      "        NotasFiscais.Pedido = Pedidos.Pedido_Id" & vbCrLf & _
                    '      " WHERE  NotasFiscais.tipodedocumento = 1 And (NotasFiscais.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "') AND (NotasFiscais.EntradaSaida_Id = 'E') AND " & vbCrLf & _
                    '      "        (NotasFiscais.Situacao = 1)" & vbCrLf

                    Sql = "SELECT  n.Nota_Id as Nota, n.Movimento as Data, c.Nome AS Cliente, c.Bairro, c.Cidade AS Cidade, " & vbCrLf & _
                            "        c.Estado AS Estado, SUM(ni.QuantidadeFiscal) AS Quantidade, SUM(ni.Valor) AS Valor, SUM(ne.Valor) AS ValorLiquido " & vbCrLf & _
                            "FROM   NotasFiscais n" & vbCrLf & _
                            "	INNER JOIN NotasFiscaisXItens ni" & vbCrLf & _
                            "			ON ni.Empresa_Id       = n.Empresa_Id " & vbCrLf & _
                            "			AND ni.EndEmpresa_Id   = n.EndEmpresa_Id " & vbCrLf & _
                            "			AND ni.Cliente_Id      = n.Cliente_Id " & vbCrLf & _
                            "			AND ni.EndCliente_Id   = n.EndCliente_Id " & vbCrLf & _
                            "			AND ni.EntradaSaida_Id = n.EntradaSaida_Id " & vbCrLf & _
                            "			AND ni.Serie_Id        = n.Serie_Id " & vbCrLf & _
                            "			AND ni.Nota_Id         = n.Nota_Id " & vbCrLf & _
                            "	INNER JOIN NotasFiscaisXEncargos ne" & vbCrLf & _
                            "			ON ne.Empresa_Id       = ni.Empresa_Id " & vbCrLf & _
                            "			AND ne.EndEmpresa_Id   = ni.EndEmpresa_Id " & vbCrLf & _
                            "			AND ne.Cliente_Id      = ni.Cliente_Id " & vbCrLf & _
                            "			AND ne.EndCliente_Id   = ni.EndCliente_Id " & vbCrLf & _
                            "			AND ne.EntradaSaida_Id = ni.EntradaSaida_Id " & vbCrLf & _
                            "			AND ne.Serie_Id        = ni.Serie_Id " & vbCrLf & _
                            "			AND ne.Nota_Id         = ni.Nota_Id " & vbCrLf & _
                            "			AND ne.Produto_Id      = ni.Produto_Id " & vbCrLf & _
                            "			AND ne.CFOP_Id         = ni.CFOP_Id" & vbCrLf & _
                            "			AND ne.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
                            "	INNER jOIN SUBOPERACOES so " & vbCrLf & _
                            "			ON so.Operacao_Id      = n.Operacao " & vbCrLf & _
                            "			AND so.SubOperacoes_Id = n.SubOperacao " & vbCrLf & _
                            "	INNER JOIN Clientes c " & vbCrLf & _
                            "			ON c.Cliente_Id   = n.Cliente_Id" & vbCrLf & _
                            "			AND c.Endereco_Id = n.EndCliente_Id " & vbCrLf & _
                            "	INNER JOIN Pedidos p" & vbCrLf & _
                            "			ON p.Empresa_Id     = n.Empresa_Id  " & vbCrLf & _
                            "			AND p.EndEmpresa_Id = n.EndEmpresa_Id " & vbCrLf & _
                            "			AND p.Pedido_Id     = n.Pedido" & vbCrLf & _
                            "WHERE  n.tipodedocumento = 1 And (n.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "') " & vbCrLf & _
                            "AND (n.EntradaSaida_Id = 'E') " & vbCrLf & _
                            "AND (n.Situacao = 1)" & vbCrLf

                    If DdlEmpresa.SelectedIndex > 0 Then
                        Dim strEmpresa As String()
                        strEmpresa = DdlEmpresa.SelectedValue.Split(New Char() {"-"})
                        Sql &= "AND  n.Empresa_Id    ='" & strEmpresa(0) & "' " & vbCrLf & _
                               "AND  n.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
                    End If

                    If txtCliente.Text <> "" Then
                        Sql &= "AND (n.Cliente_Id = '" & Cliente(0) & "' AND n.EndCliente_Id = " & Cliente(1) & ")" & vbCrLf
                    End If

                    If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND n.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf

                    If cmbSubOperacao.SelectedIndex > 0 Then
                        Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
                        Sql &= "AND n.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
                    End If

                    VerificarGrupoProduto(Sql)

                    If lstCfopSelecionados.Items.Count > 0 Then
                        Sql &= "AND ("
                        Dim strOr As String = ""

                        Dim k As Integer = 0
                        While k < lstCfopSelecionados.Items.Count
                            Sql &= strOr & "ni.CFOP_Id = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                            strOr = " OR "
                            k += 1
                        End While

                        Sql &= ") " & vbCrLf
                    End If

                    If ckApenasFinanceiro.Checked Then
                        Sql &= " AND so.Financeiro = 'S' " & vbCrLf
                    End If

                    Sql &= " GROUP  BY n.Cliente_Id, n.EndCliente_Id, n.EntradaSaida_Id, n.Serie_Id, n.Nota_Id, " & vbCrLf & _
                           "        n.Movimento, c.Nome, c.Bairro, c.Cidade, c.Estado, p.PedidoEfetivo" & vbCrLf & _
                           " ORDER  BY c.Nome, n.Nota_Id" & vbCrLf

                    Dim dsCli As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("ConsultaParametros", getParam())

                    Funcoes.BindReport(Me.Page, dsCli, "Cr_PedidosEntreguesDeComprasCli", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                Else

                    Sql = " SELECT  NotasFiscais.Nota_Id AS Nota, " & vbCrLf & _
                          "        NotasFiscais.Movimento AS Data, " & vbCrLf & _
                          "        Clientes.Nome AS Cliente, " & vbCrLf & _
                          "        NotasFiscaisXItens.Produto_Id as Produto, " & vbCrLf & _
                          "        Produtos.Unidade, " & vbCrLf & _
                          "        Produtos.Nome as NomeDoProduto," & vbCrLf & _
                          "        NotasFiscaisXItens.QuantidadeFiscal AS Quantidade, " & vbCrLf & _
                          "        NotasFiscaisXItens.Unitario," & vbCrLf & _
                          "        NotasFiscaisXItens.Valor," & vbCrLf & _
                          "        Produtos.NCM" & vbCrLf & _
                          " FROM   NotasFiscais INNER JOIN" & vbCrLf & _
                          "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                          "        NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                          "        NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                          "        NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                          "          INNER jOIN SUBOPERACOES SO                              " & vbCrLf & _
                          "             ON NOTASFISCAIS.OPERACAO = SO.OPERACAO_ID            " & vbCrLf & _
                          "             AND NOTASFISCAIS.SUBOPERACAO = SO.SUBOPERACOES_ID    " & vbCrLf & _
                          "        INNER JOIN Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf & _
                          "        INNER JOIN Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                          " WHERE  NotasFiscais.tipodedocumento = 1 And (NotasFiscais.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "') AND (NotasFiscais.EntradaSaida_Id = 'E') AND (NotasFiscais.Situacao = 1)" & vbCrLf

                    If DdlEmpresa.SelectedIndex > 0 Then
                        Dim strEmpresa As String()
                        strEmpresa = DdlEmpresa.SelectedValue.Split(New Char() {"-"})
                        Sql &= " AND  NotasFiscaisXItens.Empresa_Id    ='" & strEmpresa(0) & "' " & vbCrLf & _
                               " AND  NotasFiscaisXItens.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
                    End If

                    If txtCliente.Text <> "" Then
                        Sql &= "  And (NotasFiscais.Cliente_Id = '" & Cliente(0) & "') And NotasFiscais.EndCliente_Id = " & Cliente(1) & vbCrLf
                    End If

                    If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND NotasFiscais.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf

                    If cmbSubOperacao.SelectedIndex > 0 Then
                        Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
                        Sql &= "AND NotasFiscais.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
                    End If

                    VerificarGrupoProduto(Sql)

                    If lstCfopSelecionados.Items.Count > 0 Then
                        Sql &= "AND ("
                        Dim strOr As String = ""

                        Dim k As Integer = 0
                        While k < lstCfopSelecionados.Items.Count
                            Sql &= strOr & "NotasFiscaisXItens.CFOP_Id = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                            strOr = " OR "
                            k += 1
                        End While

                        Sql &= ") " & vbCrLf
                    End If

                    If ckApenasFinanceiro.Checked Then
                        Sql &= " AND SO.Financeiro = 'S' " & vbCrLf
                    End If

                    Sql &= " ORDER  BY Cliente, Nota" & vbCrLf

                    Dim dsProd As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("ConsultaParametros", getParam())

                    Funcoes.BindReport(Me.Page, dsProd, "Cr_PedidosEntreguesDeComprasProd", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorioCC()
        Try
            If Validar() Then
                Dim objEmpresa = New Cliente(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1))

                Cliente = txtCodigoCliente.Value.Split("-")


                Sql = " SELECT  NotasFiscais.Nota_Id AS Nota, " & vbCrLf & _
                        "        NotasFiscais.Movimento AS Data, " & vbCrLf & _
                        "        Clientes.Nome AS Cliente, " & vbCrLf & _
                        "        NotasFiscaisXEncargos.CentroDeCusto, " & vbCrLf & _
                        "        NotasFiscaisXItens.Produto_Id as Produto, " & vbCrLf & _
                        "        Produtos.Nome as NomeDoProduto," & vbCrLf & _
                        "        Produtos.Unidade, " & vbCrLf & _
                        "        Produtos.NCM," & vbCrLf & _
                        "        NotasFiscaisXItens.QuantidadeFiscal AS Quantidade, " & vbCrLf & _
                        "        NotasFiscaisXItens.Unitario," & vbCrLf & _
                        "        NotasFiscaisXItens.Valor" & vbCrLf & _
                        " FROM   NotasFiscais INNER JOIN" & vbCrLf & _
                        "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                        "        NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                        "        NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                        "        NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                        "          INNER JOIN NotasFiscaisXEncargos" & vbCrLf & _
                        "                  ON NotasFiscaisXEncargos.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.Produto_Id      = NotasFiscaisXItens.Produto_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.CFOP_Id         = NotasFiscaisXItens.CFOP_Id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.Sequencia_id    = NotasFiscaisXItens.Sequencia_id " & vbCrLf & _
                        "                 AND NotasFiscaisXEncargos.Encargo_id      = 'PRODUTO' " & vbCrLf & _
                        "          INNER jOIN SUBOPERACOES SO                              " & vbCrLf & _
                        "             ON NOTASFISCAIS.OPERACAO = SO.OPERACAO_ID            " & vbCrLf & _
                        "             AND NOTASFISCAIS.SUBOPERACAO = SO.SUBOPERACOES_ID    " & vbCrLf & _
                        "        INNER JOIN Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf & _
                        "        INNER JOIN Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                        " WHERE  NotasFiscais.tipodedocumento = 1 And (NotasFiscais.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "') AND (NotasFiscais.EntradaSaida_Id = 'E') AND (NotasFiscais.Situacao = 1)" & vbCrLf

                If DdlEmpresa.SelectedIndex > 0 Then
                    Sql &= " AND  NotasFiscaisXItens.Empresa_Id    ='" & objEmpresa.Codigo & "' " & vbCrLf & _
                           " AND  NotasFiscaisXItens.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & " " & vbCrLf
                End If

                If txtCliente.Text <> "" Then
                    Sql &= "  And (NotasFiscais.Cliente_Id = '" & Cliente(0) & "') And NotasFiscais.EndCliente_Id = " & Cliente(1) & vbCrLf
                End If

                If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND NotasFiscais.Operacao = " & cmbOperacao.SelectedValue & " " & vbCrLf

                If cmbSubOperacao.SelectedIndex > 0 Then
                    Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
                    Sql &= "AND NotasFiscais.SubOperacao = " & strSubOpe(1) & " " & vbCrLf
                End If

                If ddlCentroDeCusto.SelectedIndex > 0 Then Sql &= "AND NotasFiscaisXEncargos.CentroDeCusto = " & ddlCentroDeCusto.SelectedValue & " " & vbCrLf

                VerificarGrupoProduto(Sql)

                If lstCfopSelecionados.Items.Count > 0 Then
                    Sql &= "AND ("
                    Dim strOr As String = ""

                    Dim k As Integer = 0
                    While k < lstCfopSelecionados.Items.Count
                        Sql &= strOr & "NotasFiscaisXItens.CFOP_Id = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                        strOr = " OR "
                        k += 1
                    End While

                    Sql &= ") " & vbCrLf
                End If

                If ckApenasFinanceiro.Checked Then
                    Sql &= " AND SO.Financeiro = 'S' " & vbCrLf
                End If

                Sql &= " ORDER  BY Cliente, Nota" & vbCrLf

                Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then
                    File.Delete(fileName)
                End If

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando planilha títulos
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("PEDIDOS ENTREGUES")

                        'criando linha com o cabeçalho da planilha
                        Dim rowIndex As Integer = 1
                        Dim columnIndex As Integer = 1

                        'criando linha que informa o nome da empresa e o cnpj
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa a cidade e o estado da empresa
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o título do relatório
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELATÓRIO DE PEDIDOS ENTREGUES")
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A5:F" & rowIndex).AutoFilter = True

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using
                        rowIndex += 1

                        ' criando conteúdo da planilha com os dados do dataset
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                columnIndex += 1
                            Next

                            'formatando células datas
                            worksheet.Cells(String.Format("B{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                            'formatando células numéricas
                            worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "##0.0000_ ;[Red]-##0.0000"
                            worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            'aplicando formatação nas células do conteúdo
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        rowIndex += 1

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(String.Format("F{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                        'criando colunas de totalizadores
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Bold = True
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Formula = String.Format("=SUM(K6:K{0})", rowIndex - 1)
                        worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando quinta linha (cabeçalho)
                        worksheet.View.FreezePanes(6, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Validar() Then

                GridProduto.DataSource = Nothing
                GridProduto.DataBind()

                GridCliente.DataSource = Nothing
                GridCliente.DataBind()
                Cliente = txtCodigoCliente.Value.Split("-")

                If RadCliente.Checked = True Then

                    Sql = " SELECT  NotasFiscais.Nota_Id as Nota, NotasFiscais.Movimento as Data, Clientes.Nome AS Cliente, Clientes.Bairro, Clientes.Cidade AS Cidade, " & vbCrLf & _
                          "        Clientes.Estado AS Estado, SUM(NotasFiscaisXItens.QuantidadeFiscal) AS Quantidade, SUM(NotasFiscaisXItens.Valor) AS Valor" & vbCrLf & _
                          " FROM   NotasFiscais INNER JOIN" & vbCrLf & _
                          "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                          "        NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                          "        NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                          "        NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                          "          INNER jOIN SUBOPERACOES SO                              " & vbCrLf & _
                          "             ON NOTASFISCAIS.OPERACAO = SO.OPERACAO_ID            " & vbCrLf & _
                          "             AND NOTASFISCAIS.SUBOPERACAO = SO.SUBOPERACOES_ID    " & vbCrLf & _
                          "        INNER JOIN Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf & _
                          "        Pedidos ON NotasFiscais.Empresa_Id = Pedidos.Empresa_Id AND NotasFiscais.EndEmpresa_Id = Pedidos.EndEmpresa_Id AND " & vbCrLf & _
                          "        NotasFiscais.Pedido = Pedidos.Pedido_Id" & vbCrLf & _
                          " WHERE  NotasFiscais.tipodedocumento = 1 And (NotasFiscais.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "') AND (NotasFiscais.EntradaSaida_Id = 'E') AND " & vbCrLf & _
                          "        (NotasFiscais.Situacao = 1)" & vbCrLf
                    If DdlEmpresa.SelectedIndex > 0 Then
                        Dim strEmpresa As String()
                        strEmpresa = DdlEmpresa.SelectedValue.Split(New Char() {"-"})
                        Sql &= " AND  NotasFiscaisXItens.Empresa_Id    ='" & strEmpresa(0) & "' " & vbCrLf & _
                               " AND  NotasFiscaisXItens.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
                    End If
                    If txtCliente.Text <> "" Then
                        Sql &= "  And (NotasFiscais.Cliente_Id = '" & Cliente(0) & "') And NotasFiscais.EndCliente_Id = " & Cliente(1) & vbCrLf
                    End If
                    VerificarGrupoProduto(Sql)
                    If lstCfopSelecionados.Items.Count > 0 Then
                        Sql &= "AND ("
                        Dim strOr As String = ""

                        Dim k As Integer = 0
                        While k < lstCfopSelecionados.Items.Count
                            Sql &= strOr & "NotasFiscaisXItens.CFOP_Id = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                            strOr = " OR "
                            k += 1
                        End While

                        Sql &= ") " & vbCrLf
                    End If
                    If ckApenasFinanceiro.Checked Then
                        Sql &= " AND SO.Financeiro = 'S' " & vbCrLf
                    End If
                    Sql &= " GROUP  BY NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, " & vbCrLf & _
                           "        NotasFiscais.Movimento, Clientes.Nome, Clientes.Bairro, Clientes.Cidade, Clientes.Estado, Pedidos.PedidoEfetivo" & vbCrLf & _
                           " ORDER  BY Cliente, NotasFiscais.Nota_Id" & vbCrLf

                    Dim dsCli As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")
                    GridCliente.DataSource = dsCli
                    GridCliente.DataBind()

                    GridProduto.Parent.Visible = False
                    GridCliente.Parent.Visible = True
                Else
                    Sql = " SELECT  NotasFiscais.Nota_Id AS Nota, " & vbCrLf & _
                          "        NotasFiscais.Movimento AS Data, " & vbCrLf & _
                          "        Clientes.Nome AS Cliente, " & vbCrLf & _
                          "        NotasFiscaisXItens.Produto_Id as Produto, " & vbCrLf & _
                          "        Produtos.Unidade, " & vbCrLf & _
                          "        Produtos.Nome as NomeDoProduto," & vbCrLf & _
                          "        NotasFiscaisXItens.QuantidadeFiscal AS Quantidade, " & vbCrLf & _
                          "        NotasFiscaisXItens.Unitario," & vbCrLf & _
                          "        NotasFiscaisXItens.Valor" & vbCrLf & _
                          " FROM   NotasFiscais INNER JOIN" & vbCrLf & _
                          "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                          "        NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                          "        NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
                          "        NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                          "          INNER jOIN SUBOPERACOES SO                              " & vbCrLf & _
                          "             ON NOTASFISCAIS.OPERACAO = SO.OPERACAO_ID            " & vbCrLf & _
                          "             AND NOTASFISCAIS.SUBOPERACAO = SO.SUBOPERACOES_ID    " & vbCrLf & _
                          "        INNER JOIN Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf & _
                          "        Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                          " WHERE  NotasFiscais.tipodedocumento = 1 And (NotasFiscais.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "') AND (NotasFiscais.EntradaSaida_Id = 'E') AND (NotasFiscais.Situacao = 1)" & vbCrLf
                    If DdlEmpresa.SelectedIndex > 0 Then
                        Dim strEmpresa As String()
                        strEmpresa = DdlEmpresa.SelectedValue.Split(New Char() {"-"})
                        Sql &= " AND  NotasFiscaisXItens.Empresa_Id    ='" & strEmpresa(0) & "' " & vbCrLf & _
                               " AND  NotasFiscaisXItens.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
                    End If
                    If txtCliente.Text <> "" Then
                        Sql &= "  And (NotasFiscais.Cliente_Id = '" & Cliente(0) & "') And NotasFiscais.EndCliente_Id = " & Cliente(1) & vbCrLf
                    End If
                    VerificarGrupoProduto(Sql)
                    If lstCfopSelecionados.Items.Count > 0 Then
                        Sql &= "AND ("
                        Dim strOr As String = ""

                        Dim k As Integer = 0
                        While k < lstCfopSelecionados.Items.Count
                            Sql &= strOr & "NotasFiscaisXItens.CFOP_Id = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                            strOr = " OR "
                            k += 1
                        End While

                        Sql &= ") " & vbCrLf
                    End If
                    If ckApenasFinanceiro.Checked Then
                        Sql &= " AND SO.Financeiro = 'S' " & vbCrLf
                    End If
                    Sql &= " ORDER  BY Cliente, Nota" & vbCrLf

                    Dim dsProd As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")
                    GridProduto.DataSource = dsProd
                    GridProduto.DataBind()

                    GridProduto.Parent.Visible = True
                    GridCliente.Parent.Visible = False
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

    Protected Sub lnkAjuda_Click1(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PedidosEntreguesDeCompras")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub RadCliente_CheckedChanged(sender As Object, e As EventArgs) Handles RadCliente.CheckedChanged
        ddlCentroDeCusto.SelectedIndex = 0
        divCentroDeCusto.Visible = False
    End Sub

    Protected Sub RadProduto_CheckedChanged(sender As Object, e As EventArgs) Handles RadProduto.CheckedChanged
        ddlCentroDeCusto.SelectedIndex = 0
        divCentroDeCusto.Visible = False
    End Sub

    Protected Sub RadCentroDeCusto_CheckedChanged(sender As Object, e As EventArgs) Handles RadCentroDeCusto.CheckedChanged
        ddlCentroDeCusto.SelectedIndex = 0
        divCentroDeCusto.Visible = True
    End Sub
End Class