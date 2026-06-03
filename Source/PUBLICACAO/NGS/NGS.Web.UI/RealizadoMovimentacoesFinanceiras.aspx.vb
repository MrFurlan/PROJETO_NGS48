Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RealizadoMovimentacoesFinanceiras
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RealizadoMovimentacoesFinanceiras", "ACESSAR") Then
                    txtPeriodoInicialConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                    txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
                    CargaUnidadeDeNegocioEmpresaCliente()
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
                    HID.Value = Guid.NewGuid().ToString
                    ucConsultaClientes.SetarHID(HID.Value)
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
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objCliente" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub RelRealizadoMovimentacoesFinanceiras()
        If Funcoes.VerificaPermissao("RealizadoMovimentacoesFinanceiras", "RELATORIO") Then
            Try

                Dim DataInicial As String = Format(CDate(txtPeriodoInicialConsultaTitulos.Text), "yyyy/MM/dd")
                Dim DataFinal As String = Format(CDate(txtPeriodoFinalConsultaTitulos.Text), "yyyy/MM/dd")

                Dim Sql As String = "SELECT MovimentacoesFinanceiras.Registro_Id as Registro, MovimentacoesFinanceiras.Baixa as Baixa, MovimentacoesFinanceiras.Vencimento as Vencimento, MovimentacoesFinanceiras.Empresa As EmpresaCodigo, MovimentacoesFinanceiras.EndEmpresa EmpresaCodigoEnd," & vbCrLf & _
                                    " MovimentacoesFinanceiras.EmpresaPagadora as ClienteCodigo, MovimentacoesFinanceiras.EndEmpresaPagadora as ClienteCodigoEnd, MovimentacoesFinanceiras.Historico as Historico , MovimentacoesFinanceiras.ValorLiquido as ValorBaixado, Empresas.Nome as EmpresaNome , Empresas.Estado as EmpresaEstado, " & vbCrLf & _
                                    " Empresas.Cidade as EmpresaCidade, Empresas.Reduzido as EmpresaReduzido , Clientes.Nome AS ClienteNome" & vbCrLf & _
                                    " FROM MovimentacoesFinanceiras " & vbCrLf & _
                                    " INNER JOIN Clientes AS Empresas ON MovimentacoesFinanceiras.Empresa = Empresas.Cliente_Id AND MovimentacoesFinanceiras.EndEmpresa = Empresas.Endereco_Id INNER JOIN" & vbCrLf & _
                                    " Clientes ON MovimentacoesFinanceiras.EmpresaPagadora = Clientes.Cliente_Id AND MovimentacoesFinanceiras.EndEmpresaPagadora = Clientes.Endereco_Id " & vbCrLf

                Sql &= " WHERE MovimentacoesFinanceiras.Situacao = 1  and MovimentacoesFinanceiras.Grupado <> 'M'"

                If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
                    Sql &= " and MovimentacoesFinanceiras.UnidadeDeNegocio = '" & ddlUnidade.SelectedValue & "' "
                End If


                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    Sql &= " and MovimentacoesFinanceiras.Empresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                           " and MovimentacoesFinanceiras.EndEmpresa = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                    Sql &= " and MovimentacoesFinanceiras.Cliente = '" & txtCodigoCliente.Value.Split("-")(0) & "'"   'Cliente
                    Sql &= " and MovimentacoesFinanceiras.EndCliente = " & txtCodigoCliente.Value.Split("-")(1)       'Cliente da Empresa
                End If

                If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
                    Sql &= " and Vencimento between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
                End If

                Sql &= " order by  MovimentacoesFinanceiras.Vencimento , MovimentacoesFinanceiras.Empresa , MovimentacoesFinanceiras.EmpresaPagadora , MovimentacoesFinanceiras.EndEmpresaPagadora "

                Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Ds_realizado")

                Dim empresa As String = ""
                If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    empresa = String.Format("{0}-{1}", UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
                Else
                    empresa = String.Format("{0}-{1}", ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                End If

                Dim objEmpresa As New Cliente(empresa.Split("-")(0), empresa.Split("-")(1))

                Dim param As New Dictionary(Of String, Object)
                param.Add("Titulo", "Realizado Movimentacoes Financeiras")
                param.Add("Nome", objEmpresa.Nome)
                param.Add("Cidade", objEmpresa.Cidade & "-" & objEmpresa.CodigoEstado)
                param.Add("DataInicial", DataInicial)
                param.Add("DataFinal", DataFinal)

                Funcoes.BindReport(Me.Page, ds, "Cr_RealizadoMovimentacoesFinanceiras", eExportType.PDF, param)

            Catch ex As Exception
                MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            End Try
        Else
            MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            RelRealizadoMovimentacoesFinanceiras()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
            txtCliente.Text = String.Empty
            txtCodigoCliente.Value = String.Empty
            txtPeriodoInicialConsultaTitulos.Text = ""
            txtPeriodoFinalConsultaTitulos.Text = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RealizadoMovimentacoesFinanceiras")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class