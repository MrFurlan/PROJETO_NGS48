Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ConsistenciaDeMeiosEPecas
    Inherits BasePage

    Dim Sql As String
    Dim Sqla As String
    Dim Row As DataRow

    Dim ds As New DataSet
    Dim dss As New DataSet

    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim Cliente() As String
    Dim SequenciaLote As Integer = 0

    Dim Encargo As String = ""
    Dim EntradaSaida As String = ""
    Dim EntradaSaidaDescricao As String = ""

    Dim Codigo As String = ""
    Dim Descricao As String = ""

    Dim Endereco As Integer = 0
    Dim Mensagem As String = ""

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date

    Dim Processo As Integer = 0
    Dim Livro As Integer = 0
    Dim Folha As Integer = 0
    Dim Opcao As String = ""

    Dim EmpresaNome As String = ""
    Dim EmpresaCNPJ As String = ""
    Dim EmpresaInscricao As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ConsistenciaDeMeiosEPecas", "ACESSAR") Then
                CarregarUnidade()
                VerirficaUnidade()
                CargaOperacao()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUnidade()
        ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, " Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario"))
    End Sub

    Private Sub VerirficaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa)
    End Sub

    Private Sub CargaOperacao()
        ddl.Carregar(ddlOperacao, CarregarDDL.Tabela.Operacao, "", True)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue)
    End Sub

    Private Sub CargaSubOperacao()
        ddlSubOperacao.Items.Clear()

        Dim objSubOperacoes As New [Lib].Negocio.ListSubOperacao()

        If objSubOperacoes.Selecionar(ddlOperacao.SelectedValue) Then
            For Each objSubOperacao As [Lib].Negocio.SubOperacao In objSubOperacoes
                ddlSubOperacao.Items.Add(New ListItem(objSubOperacao.Codigo.ToString("00") & " - " & objSubOperacao.Descricao, _
                                                      objSubOperacao.Codigo))
            Next

            Funcoes.InserirLinhaEmBranco(ddlSubOperacao)
        End If
    End Sub

    Private Function validar()
        If DdlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Unidade é obrigatório.")
            Return False
        ElseIf DdlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf txtDataInicial.Text = "" Then
            MsgBox(Me.Page, "Período inicial é obrigatório.")
            Return False
        ElseIf txtDataFinal.Text = "" Then
            MsgBox(Me.Page, "Período final é obrigatório.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf
        If Not String.IsNullOrWhiteSpace(DdlUnidade.SelectedValue) Then
            param &= "Unidade: " & DdlUnidade.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Período: " & txtDataInicial.Text & " "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "à: " & txtDataFinal.Text & " - "
        End If
        param &= IIf(RadEntradas.Checked, "Notas: Entradas", "Notas: Saidas")

        Return param
    End Function

    Protected Sub DdlUnidade_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidade.SelectedIndexChanged
        CargaEmpresas()
    End Sub

    Private Sub Relatorio()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
        Dim EmpresaCidade As String = ""
        Dim SqlArray As New ArrayList
        Dim Cliente() As String = txtCodigoCliente.Value.Split("-")

        Sql = " SELECT *" & vbCrLf & _
            " FROM   Clientes LEFT OUTER JOIN" & vbCrLf & _
            " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
            " WHERE  Clientes.Cliente_Id = '" & Empresa(0) & "'" & vbCrLf
        ds = Banco.ConsultaDataSet(Sql, "Clientes")

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            EmpresaNome = Dr("Nome")
            EmpresaCidade = Dr("Cidade") & " - " & Dr("Estado")
            Exit For
        Next
        Dim Periodo As String = "Período : " & Format(PeriodoInicial, "dd/MM/yyyy") & " a " & Format(PeriodoFinal, "dd/MM/yyyy") & " - " & EntradaSaidaDescricao

        Sql = " SELECT NotasFiscais.Movimento, NotasFiscais.Nota_Id as Nota, NotasFiscais.Serie_Id as Serie, NotasFiscais.Pedido, NotasFiscais.Operacao, NotasFiscais.SubOperacao, " & vbCrLf & _
            "        Clientes.Cliente_Id as Cliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, NotasFiscaisXItens.Produto_Id as Produto, Produtos.Nome AS NomeDoProduto, " & vbCrLf & _
            "        NotasFiscaisXItens.QuantidadeFiscal as Quantidade, NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor, isnull(NotasFiscaisXItens.NumeroPecas,0) NumeroPecas " & vbCrLf & _
            "  FROM  NotasFiscais INNER JOIN" & vbCrLf & _
            "        Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf & _
            "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
            "        NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
            "        NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND " & vbCrLf & _
            "        NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
            "        Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
            "        Where NotasFiscais.Situacao = 1 And NotasFiscais.Empresa_Id = '" & Empresa(0) & "' " & vbCrLf

        If txtCliente.Text <> "" Then
            Sql &= "    And NotasFiscais.Cliente_Id = '" & Cliente(0) & "' And NotasFiscais.EndCliente_Id = " & Cliente(1)
        End If
        Sql &= "        And NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.tosqldate() & "'" & vbCrLf & _
            "        And NotasFiscais.EntradaSaida_Id = '" & EntradaSaida & "'"

        If ddlOperacao.SelectedIndex > 0 Then
            Sql &= "    And NotasFiscaisXItens.Operacao = " & ddlOperacao.SelectedValue & ""
        End If
        If ddlSubOperacao.SelectedIndex > 0 Then
            Sql &= "    And NotasFiscaisXItens.SubOperacao = " & ddlSubOperacao.SelectedValue & ""
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim RetornoProdutos As ArrayList
            RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NotasFiscaisXItens.Produto_Id", "")
            Sql &= " AND " & RetornoProdutos(0)
        End If

        Sql &= "        Order By NotasFiscais.Movimento"
        ds = Banco.ConsultaDataSet(Sql, "ConsistenciaDeNotas")

        Dim parameters = New Dictionary(Of String, Object)()
        parameters.Add("ConsultaParametros", getParam())

        Funcoes.BindReport(Me.Page, ds, "Cr_ConsistenciaDeMeiosEPecas", eExportType.PDF, parameters)
    End Sub

    Private Sub Limpar()
        VerirficaUnidade()
        ddlOperacao.SelectedIndex = 0
        ddlSubOperacao.Items.Clear()
        ucSelecaoProduto.Limpar()
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCliente.Click
        HttpContext.Current.Session("ssCampo") = "ConsistenciaDeNotas"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCMP" & HID.Value, "txtNome")
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteCMP" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteCMP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteCMP" & HID.Value)
        End If
    End Sub

    Protected Sub ddlOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaSubOperacao()
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("ConsistenciaDeMeiosEPecas", "RELATORIO") Then
                If validar() Then
                    If RadEntradas.Checked = True Then
                        EntradaSaida = "E"
                        EntradaSaidaDescricao = "Entradas"
                    Else
                        EntradaSaida = "S"
                        EntradaSaidaDescricao = "Saidas"
                    End If

                    Relatorio()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConsistenciaDeMeiosEPecas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class