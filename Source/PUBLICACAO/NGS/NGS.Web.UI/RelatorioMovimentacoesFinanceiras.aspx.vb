Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioMovimentacoesFinanceiras
    Inherits BasePage

#Region "Locais"
    Dim ds As DataSet
    Dim Sql As String
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioMovimentacoesFinanceiras", "ACESSAR") Then
                    HttpContext.Current.Session("ssRegistros") = ""
                    HttpContext.Current.Session("ssObservacoes") = ""

                    txtPeriodoInicialConsultaTitulos.Text = "01/01/" & Today.Year.ToString
                    txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                    ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "")
                    Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)

                    RadAtivos.Checked = True
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeConsultaTitulos.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            Dim Campo As String() = Nothing
            Dim Carteiras As String = ""
            Dim UnidadeDeNegocio As String = ""
            Dim NomeUnidadeDeNegocio As String = "Consolidado"
            Dim Empresa As String = ""
            Dim EndEmpresa As String = ""
            Dim NomeEmpresa As String = ""
            Dim CidadeEmpresa As String = ""
            Dim EstadoEmpresa As String = ""
            Dim Cliente As String = ""
            Dim EndCliente As String = ""
            Dim Ativos As Integer

            If Funcoes.VerificaPermissao("RelatorioMovimentacoesFinanceiras", "RELATORIO") Then
                If DdlUnidadeConsultaTitulos.Text <> "" Then
                    UnidadeDeNegocio = DdlUnidadeConsultaTitulos.SelectedValue
                Else
                    UnidadeDeNegocio = ""                   'UnidadeDeNegocio
                End If

                If DdlEmpresaConsultaTitulos.Text <> "" Then
                    Cliente = DdlEmpresaConsultaTitulos.SelectedValue
                    Campo = Cliente.Split("-")
                    Empresa = Campo(0)                      'EmpresaCliente
                    EndEmpresa = Campo(1)                   'Endereco Empresa Cliente
                Else
                    Empresa = ""                            'Empresa Cliente
                    EndEmpresa = 0                          'Endereco Empresa Cliente
                End If

                If CodigoCliente.Value.Length > 0 Then
                    Campo = CodigoCliente.Value.Split("-")
                    Cliente = Campo(0).Trim                    'Cliente
                    EndCliente = Campo(1)                   'Endereco Cliente
                Else
                    Cliente = ""                            'Cliente
                    EndCliente = 0                          'Endereco Cliente
                End If

                If DdlCarteiras.Text <> "" Then
                    Carteiras = DdlCarteiras.SelectedValue
                Else
                    Carteiras = ""                            'Cliente
                End If

                If RadAtivos.Checked Then
                    Ativos = 2
                ElseIf RadBaixados.Checked Then
                    Ativos = 1
                End If

                If Not IsDate(txtPeriodoInicialConsultaTitulos.Text) Or Not IsDate(txtPeriodoFinalConsultaTitulos.Text) Then
                    MsgBox(Me.Page, "Periodo Invalido.")
                    Exit Sub
                End If

                Dim DataInicial As String = Format(CDate(txtPeriodoInicialConsultaTitulos.Text), "yyyy/MM/dd")
                Dim DataFinal As String = Format(CDate(txtPeriodoFinalConsultaTitulos.Text), "yyyy/MM/dd")

                Dim DataInicialD As String = Format(CDate(txtPeriodoInicialConsultaTitulos.Text), "dd/MM/yyyy")
                Dim DataFinalD As String = Format(CDate(txtPeriodoFinalConsultaTitulos.Text), "dd/MM/yyyy")

                Sql = "  SELECT Clientes.Cliente_Id AS Cliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, 'CONSOLIDADO' as Descricao"
                Sql &= " FROM Clientes "
                Sql &= " WHERE Clientes.Cliente_Id = '" & HttpContext.Current.Session("ssEmpresa") & "'"
                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        NomeEmpresa = dr("Nome")
                        CidadeEmpresa = dr("Cidade")
                        EstadoEmpresa = dr("Estado")
                        'UnidadeDeNegocio = dr("Descricao")
                        Exit For
                    Next
                End If

                Sql = "  SELECT Clientes.Nome, Clientes.Cidade, Clientes.Estado"
                Sql &= " FROM Clientes "
                Sql &= " WHERE Clientes.Cliente_Id = '" & UnidadeDeNegocio & "'"
                ds = Banco.ConsultaDataSet(Sql, "Clientes")
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        NomeUnidadeDeNegocio = dr("Nome")
                        Exit For
                    Next
                End If

                If RadDebitosECreditos.Checked = False Then 'se true relatorio de debitos e creditos
                    Sql = "  SELECT Empresas.Reduzido AS ReduzidoEmpresa, MovimentacoesFinanceiras.Empresa, MovimentacoesFinanceiras.EndEmpresa, Empresas.Nome AS NomeEmpresa, "
                    'Sql &= " Empresas.Cidade AS CidadeEmpresa, Empresas.Estado AS EstadoEmpresa, MovimentacoesFinanceiras.Registro_Id AS Registro, MovimentacoesFinanceiras.Cliente, "
                    Sql &= " Empresas.Cidade AS CidadeEmpresa, Empresas.Estado AS EstadoEmpresa, MovimentacoesFinanceiras.Registro_Id AS Registro, MovimentacoesFinanceiras.EmpresaPagadora , "
                    'Sql &= " MovimentacoesFinanceiras.EmpresaPagadora as ClienteCodigo, MovimentacoesFinanceiras.EndEmpresaPagadora as ClienteCodigoEnd, MovimentacoesFinanceiras.Historico as Historico , MovimentacoesFinanceiras.ValorLiquido as ValorBaixado, Empresas.Nome as EmpresaNome , Empresas.Estado as EmpresaEstado, "
                    Sql &= " Clientes.Nome AS NomeCliente, MovimentacoesFinanceiras.Movimento, MovimentacoesFinanceiras.Vencimento, MovimentacoesFinanceiras.Baixa, "
                    Sql &= " MovimentacoesFinanceiras.ValorDoDocumento AS ValorDocumento, MovimentacoesFinanceiras.ValorLiquido, MovimentacoesFinanceiras.Provisao, MovimentacoesFinanceiras.Carteira, "
                    Sql &= " Carteiras.Descricao AS NomeCarteira, MovimentacoesFinanceiras.Historico, MovimentacoesFinanceiras.solicitacao"

                    Sql &= " FROM  MovimentacoesFinanceiras INNER JOIN"
                    Sql &= "       Clientes AS Empresas ON MovimentacoesFinanceiras.Empresa = Empresas.Cliente_Id AND "
                    Sql &= "       MovimentacoesFinanceiras.EndEmpresa = Empresas.Endereco_Id INNER JOIN"
                    Sql &= "       ComprasXProdutos AS Carteiras ON MovimentacoesFinanceiras.Carteira = Carteiras.Produto_Id LEFT OUTER JOIN"
                    Sql &= "       Clientes ON MovimentacoesFinanceiras.EmpresaPagadora = Clientes.Cliente_ID"

                    ''Sql &= " WHERE MovimentacoesFinanceiras.Situacao = 1  "
                    '20/01/2009 - Incluido para desconsiderar registros agrupados.
                    ''Sql &= " WHERE MovimentacoesFinanceiras.Situacao = 1  AND MovimentacoesFinanceiras.RegistroMestre is NULL"
                    '11/02/2010 - Incluido para desconsiderar registros agrupados.
                    Sql &= " WHERE MovimentacoesFinanceiras.Situacao = 1  AND MovimentacoesFinanceiras.Grupado = 'M'"

                    If Ativos = 1 Then
                        Sql &= " and MovimentacoesFinanceiras.Provisao = 1"
                    End If

                    If Ativos = 2 Then
                        Sql &= " and MovimentacoesFinanceiras.Provisao = 2"
                    End If

                    If UnidadeDeNegocio <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.UnidadeDeNegocio = '" & UnidadeDeNegocio & "' "
                    End If

                    If Empresa <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Empresa = '" & Empresa & "'"      'Empresa
                        Sql &= " and MovimentacoesFinanceiras.EndEmpresa = " & EndEmpresa       'Endereco da Empresa
                    End If

                    If Cliente <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Cliente = '" & Cliente & "'"      'Cliente
                        Sql &= " and MovimentacoesFinanceiras.EndCliente = " & EndCliente       'Endereco da Empresa
                    End If

                    If DataInicial <> "" And DataFinal <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Vencimento between '" & DataInicial & "' and '" & DataFinal & "'"
                    End If

                    'Sql &= " order by MovimentacoesFinanceiras.Empresa , MovimentacoesFinanceiras.Vencimento , MovimentacoesFinanceiras.Cliente , MovimentacoesFinanceiras.EndCliente "
                    ds = Banco.ConsultaDataSet(Sql, "Titulos")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Nome", NomeEmpresa)
                    parameters.Add("Cidade", CidadeEmpresa)
                    parameters.Add("Estado", EstadoEmpresa)
                    parameters.Add("UnidadeDeNegocio", NomeUnidadeDeNegocio)
                    parameters.Add("DataInicial", DataInicialD)
                    parameters.Add("DataFinal", DataFinalD)

                    Funcoes.BindReport(Me.Page, ds, "Cr_MovimentacoesFinanceiras", eExportType.PDF)

                Else 'relatorio de debitos e creditos

                    Sql = " select Reduzido as ReduzidoEmpresa, Empresa, EndEmpresa, Nome as NomeEmpresa, Credito, Debito, Carteira,NomeCarteira,Historico, " & vbCrLf & _
                          " CidadeEmpresa, EstadoEmpresa, Registro,DescConta, " & vbCrLf & _
                          " Movimento, Vencimento, Baixa, Sequencia " & vbCrLf & _
                          "  from (" & vbCrLf & _
                          " SELECT Empresas.Reduzido, MovimentacoesFinanceiras.EmpresaPagadora as Empresa, " & vbCrLf & _
                          "       MovimentacoesFinanceiras.EndEmpresaPagadora as EndEmpresa,Empresas.Nome as Nome, " & vbCrLf & _
                          " MovimentacoesFinanceiras.ValorLiquido as Credito," & vbCrLf & _
                          " convert(decimal(18,4),0.00) as Debito, " & vbCrLf & _
                          " MovimentacoesFinanceiras.Carteira," & vbCrLf & _
                          " Carteiras.Descricao AS NomeCarteira, MovimentacoesFinanceiras.Historico,PlanoDeContas.Titulo as DescConta, " & vbCrLf & _
                          " Empresas.Cidade AS CidadeEmpresa, Empresas.Estado AS EstadoEmpresa, MovimentacoesFinanceiras.Registro_Id AS Registro, " & vbCrLf & _
                          " MovimentacoesFinanceiras.Movimento, MovimentacoesFinanceiras.Vencimento, MovimentacoesFinanceiras.Baixa, MovimentacoesFinanceiras.Sequencia_Id AS Sequencia " & vbCrLf & _
                          " FROM MovimentacoesFinanceiras " & vbCrLf & _
                          " INNER JOIN Clientes AS Empresas " & vbCrLf & _
                          " ON MovimentacoesFinanceiras.EmpresaPagadora    = Empresas.Cliente_Id " & vbCrLf & _
                          " AND MovimentacoesFinanceiras.EndEmpresaPagadora = Empresas.Endereco_Id " & vbCrLf & _
                          " left JOIN ComprasXProdutos AS Carteiras " & vbCrLf & _
                          " ON MovimentacoesFinanceiras.Carteira = Carteiras.Produto_Id " & vbCrLf & _
                          " Left JOIN  PlanoDeContas ON '99999999999999' = PlanoDeContas.Empresa_Id " & vbCrLf & _
                          " AND 0 = PlanoDeContas.EndEmpresa_Id " & vbCrLf & _
                          " AND MovimentacoesFinanceiras.ContaContabilPagadora = PlanoDeContas.Conta_Id " & vbCrLf & _
                          "where MovimentacoesFinanceiras.Situacao = 1 " & vbCrLf
                    If Empresa <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Empresa = '" & Empresa & "'"      'Empresa
                        Sql &= " and MovimentacoesFinanceiras.EndEmpresa = " & EndEmpresa       'Endereco da Empresa
                    End If

                    If Cliente <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Cliente = '" & Cliente & "'"      'Cliente
                        Sql &= " and MovimentacoesFinanceiras.EndCliente = " & EndCliente       'Endereco da Empresa
                    End If

                    If DataInicial <> "" And DataFinal <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Vencimento between '" & DataInicial & "' and '" & DataFinal & "'"
                    End If
                    Sql &= " union " & vbCrLf & _
                           " SELECT Clientes.Reduzido, MovimentacoesFinanceiras.Cliente as Empresa,  " & vbCrLf & _
                           " MovimentacoesFinanceiras.EndCliente as EndEmpresa, Clientes.Nome as nome,   " & vbCrLf & _
                           " convert(decimal(18,4),0.00) as Credito, " & vbCrLf & _
                           " MovimentacoesFinanceiras.ValorLiquido as Debito, " & vbCrLf & _
                           " MovimentacoesFinanceiras.Carteira," & vbCrLf & _
                           " Carteiras.Descricao AS NomeCarteira, MovimentacoesFinanceiras.Historico,PlanoDeContas.Titulo as DescConta, " & vbCrLf & _
                           " Clientes.Cidade AS CidadeEmpresa, Clientes.Estado AS EstadoEmpresa, MovimentacoesFinanceiras.Registro_Id AS Registro, " & vbCrLf & _
                           " MovimentacoesFinanceiras.Movimento, MovimentacoesFinanceiras.Vencimento, MovimentacoesFinanceiras.Baixa, MovimentacoesFinanceiras.Sequencia_Id AS Sequencia " & vbCrLf & _
                           " FROM MovimentacoesFinanceiras " & vbCrLf & _
                           " INNER JOIN Clientes AS Clientes " & vbCrLf & _
                           " ON MovimentacoesFinanceiras.Cliente = Clientes.Cliente_Id " & vbCrLf & _
                           " AND MovimentacoesFinanceiras.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
                           " left JOIN ComprasXProdutos AS Carteiras " & vbCrLf & _
                           " ON MovimentacoesFinanceiras.Carteira = Carteiras.Produto_Id " & vbCrLf & _
                           " Left JOIN  PlanoDeContas ON '99999999999999' = PlanoDeContas.Empresa_Id " & vbCrLf & _
                           " AND 0 = PlanoDeContas.EndEmpresa_Id " & vbCrLf & _
                           " AND MovimentacoesFinanceiras.ContaContabilCliente = PlanoDeContas.Conta_Id " & vbCrLf & _
                           "where MovimentacoesFinanceiras.Situacao = 1 " & vbCrLf
                    If Empresa <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Empresa = '" & Empresa & "'"      'Empresa
                        Sql &= " and MovimentacoesFinanceiras.EndEmpresa = " & EndEmpresa       'Endereco da Empresa
                    End If
                    If Cliente <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Cliente = '" & Cliente & "'"      'Cliente
                        Sql &= " and MovimentacoesFinanceiras.EndCliente = " & EndCliente       'Endereco da Empresa
                    End If
                    If DataInicial <> "" And DataFinal <> "" Then
                        Sql &= " and MovimentacoesFinanceiras.Vencimento between '" & DataInicial & "' and '" & DataFinal & "'"
                    End If
                    Sql &= " ) as consulta "
                    Sql &= " order by Registro"
                    ds = Banco.ConsultaDataSet(Sql, "Titulos")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Nome", NomeEmpresa)
                    parameters.Add("Cidade", CidadeEmpresa)
                    parameters.Add("Estado", EstadoEmpresa)
                    parameters.Add("UnidadeDeNegocio", NomeUnidadeDeNegocio)
                    parameters.Add("DataInicial", DataInicialD)
                    parameters.Add("DataFinal", "DataFinalD")

                    Funcoes.BindReport(Me.Page, ds, "Cr_MovimentacoesFinanceirasDebitoECredito", eExportType.PDF, parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            DdlCarteiras.SelectedIndex = 0
            txtPeriodoInicialConsultaTitulos.Text = "01/01/" & Today.Year.ToString
            txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
            RadTodos.Checked = True
            txtNomeCliente.Text = ""
            CodigoCliente.Value = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioMovimentacoesFinanceiras")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objClienteMovimentacaoFinanceira" & HID.Value) Is Nothing Then
                Dim cli As [Lib].Negocio.Cliente = CType(Session("objClienteMovimentacaoFinanceira" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(cli)
                With cli
                    txtNomeCliente.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                    CodigoCliente.Value = .Codigo & " - " & .CodigoEndereco
                End With
                Session.Remove("objClienteMovimentacaoFinanceira" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteMovimentacaoFinanceira" & HID.Value.ToString, "txtNomeCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class