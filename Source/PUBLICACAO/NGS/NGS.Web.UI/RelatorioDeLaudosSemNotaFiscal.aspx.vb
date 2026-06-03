Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioDeLaudosSemNotaFiscal
    Inherits BasePage

    Private Sql As String
    Private SqlArray As New ArrayList
    Private ds As DataSet
    Private dr As DataRow
    Private strm As StreamWriter

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeLaudosSemNotaFiscal", "ACESSAR") Then
                    ucConsultaClientes.SetarHID(HID.Value)

                    txtPeriodoInicial.Text = Today.ToString("dd/MM/yyyy")
                    txtPeriodoFinal.Text = Today.ToString("dd/MM/yyyy")
                    CargaUnidadeDeNegocio()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuários sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocio.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(DdlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocio, DdlEmpresa, True)
    End Sub

    Protected Sub DdlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocio.SelectedValue.ToString(), True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDataset() As DataSet
        Sql = "SELECT CONVERT(varchar, R.Movimento, 103) AS Movimento, RxP.Pesagem_Id, " & vbCrLf & _
              "       R.Romaneio_Id, R.Pedido, CONVERT(VARCHAR,Cli.Cliente_Id) + ' - ' + CONVERT(VARCHAR,Cli.Endereco_Id) AS Cliente_Id, Cli.Endereco_Id, " & vbCrLf & _
              "       Cli.Nome, Cli.Complemento, Pr.Nome AS NomeProduto, R.PesoLiquido, " & vbCrLf & _
              "       ISNULL(R.Processo, '') AS Processo,R.EntradaSaida, P.EntradaBalanca, P.SaidaBalanca, P.Observacoes " & vbCrLf & _
              "  FROM Romaneios R " & vbCrLf & _
              " INNER JOIN RomaneiosXPesagens RxP " & vbCrLf & _
              "    ON R.Empresa_Id    = RxP.Empresa_Id " & vbCrLf & _
              "	  AND R.EndEmpresa_Id = RxP.EndEmpresa_Id " & vbCrLf & _
              "   AND R.Romaneio_Id   = RxP.Romaneio_Id " & vbCrLf & _
              "  LEFT JOIN Produtos Pr" & vbCrLf & _
              "    ON R.Produto = Pr.Produto_Id " & vbCrLf & _
              "  LEFT JOIN Pedidos Ped " & vbCrLf & _
              "    ON R.Empresa_Id    = Ped.Empresa_Id " & vbCrLf & _
              "   AND R.EndEmpresa_Id = Ped.EndEmpresa_Id " & vbCrLf & _
              "   AND R.Pedido        = Ped.Pedido_Id " & vbCrLf & _
              "  LEFT OUTER JOIN Clientes Cli " & vbCrLf & _
              "    ON Ped.Cliente    = Cli.Cliente_Id " & vbCrLf & _
              "   AND Ped.EndCliente = Cli.Endereco_Id " & vbCrLf & _
              "  LEFT OUTER JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf & _
              "    ON R.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf & _
              "   AND R.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf & _
              "   AND R.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf & _
              "  LEFT JOIN NotasFiscais NF" & vbCrLf & _
              "    ON nfxr.Empresa_Id    = NF.Empresa_Id" & vbCrLf & _
              "   AND nfxr.EndEmpresa_Id = NF.EndEmpresa_Id" & vbCrLf & _
              "   AND nfxr.Cliente_Id = NF.Cliente_id" & vbCrLf & _
              "   AND nfxr.EndCliente_Id = NF.EndCliente_Id " & vbCrLf & _
              "   AND nfxr.Nota_Id = NF.Nota_Id" & vbCrLf & _
              "   AND nfxr.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
              "   AND nfxr.Serie_Id = NF.Serie_Id " & vbCrLf & _
              " INNER JOIN Pesagem P " & vbCrLf & _
              "    ON RxP.Empresa_Id    = P.Empresa_Id  " & vbCrLf & _
              "   AND RxP.EndEmpresa_Id = P.EndEmpresa_Id  " & vbCrLf & _
              "   AND RxP.Pesagem_Id    = P.Pesagem_Id " & vbCrLf & _
              "   AND P.Sequencia_Id    = 0 " & vbCrLf & _
              " WHERE Pr.ControlarRomaneio = 1" & vbCrLf & _
              "   AND P.Situacao = 1" & vbCrLf & _
              "   AND R.Movimento BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "' " & vbCrLf

        If DdlEmpresa.SelectedIndex > 0 Then
            Sql &= "   AND R.Empresa_ID = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf & _
                   "   AND R.EndEmpresa_ID = " & DdlEmpresa.SelectedValue.Split("-")(1) & " " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            If chkConsolidarCliente.Checked Then
                Sql &= "   AND left(Cli.Cliente_ID,8) = '" & txtCodigoCliente.Value.Split("-")(0).Substring(0, 8) & "' " & vbCrLf
            Else
                Sql &= "   AND Cli.Cliente_ID    ='" & txtCodigoCliente.Value.Split("-")(0) & "' " & vbCrLf & _
                       "   AND Cli.Endereco_ID = " & txtCodigoCliente.Value.Split("-")(1) & " " & vbCrLf
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoDeposito.Value) Then
            Sql &= "   AND R.Deposito    ='" & txtCodigoDeposito.Value.Split("-")(0) & "'" & vbCrLf & _
                   "   AND R.EndDeposito = " & txtCodigoDeposito.Value.Split("-")(1) & " " & vbCrLf
        End If

        VerificarGrupoProduto(Sql)

        Sql &= "   AND (nfxr.Nota_Id IS NULL OR NF.situacao NOT IN (1,4,7))  " & vbCrLf & _
               "   AND ISNULL(P.RegistroMestre,0) = 0 " & vbCrLf

        Sql &= " GROUP BY CONVERT(varchar, R.Movimento, 103), RxP.Pesagem_Id,  " & vbCrLf & _
               "          R.Romaneio_Id, R.Pedido, Cli.Cliente_Id, Cli.Endereco_Id, " & vbCrLf & _
               "          Cli.Nome, Cli.Complemento, Pr.Nome, R.PesoLiquido, " & vbCrLf & _
               "          ISNULL(R.Processo, ''),R.EntradaSaida, P.EntradaBalanca, P.SaidaBalanca, P.Observacoes " & vbCrLf & _
               " ORDER BY Movimento, R.EntradaSaida "

        Return Banco.ConsultaDataSet(Sql, "RelatorioDeLaudosSemNotaFiscal")
    End Function

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            Dim Parameters As New Dictionary(Of String, Object)
            Dim Par As String = String.Empty
            Dim objEmpresa As [Lib].Negocio.Cliente

            If DdlEmpresa.SelectedIndex > 0 Then
                Dim Empresa() As String
                Empresa = DdlEmpresa.SelectedValue.ToString.Split("-")
                objEmpresa = New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))
            Else
                objEmpresa = New [Lib].Negocio.Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))
            End If

            ds = getDataset()

            Par = String.Empty
            Par &= "Parâmetros:" & vbCrLf
            If DdlUnidadeDeNegocio.SelectedIndex > 0 Then
                Par &= "Unidade de Negócio:" & DdlUnidadeDeNegocio.SelectedItem.Text & vbCrLf
            End If

            Par &= "Empresa: " & objEmpresa.Nome & " - " & objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado & " " & objEmpresa.Codigo & vbCrLf
            If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                Par &= IIf(chkConsolidarCliente.Checked, "Cliente Consolidado:", "Cliente:") & txtCliente.Text & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtDeposito.Text) Then
                Par &= "Depósito: " & txtDeposito.Text & vbCrLf
            End If
            Par &= "Período: " & txtPeriodoInicial.Text & " à " & txtPeriodoFinal.Text & vbCrLf

            If ucSelecaoProduto.TemSelecionado Then
                Par &= ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo_Id", "P.Produto_Id", "", True)(1)
            End If

            Parameters.Add("EmpresaNome", objEmpresa.Nome)
            Parameters.Add("EmpresaCidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)
            Parameters.Add("EmpresaCodigo", Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
            Parameters.Add("Parametros", Par)

            Funcoes.BindReport(Me.Page, ds, "Cr_RelatorioDeLaudosSemNotaFiscal", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), Parameters)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            DdlUnidadeDeNegocio.SelectedIndex = 0
            DdlEmpresa.SelectedIndex = 0
            txtPeriodoInicial.Text = ""
            txtPeriodoFinal.Text = ""

            txtCliente.Text = String.Empty
            txtCodigoCliente.Value = String.Empty

            txtDeposito.Text = String.Empty
            txtCodigoDeposito.Value = String.Empty

            txtPeriodoInicial.Text = Today.ToString("dd/MM/yyyy")
            txtPeriodoFinal.Text = Today.ToString("dd/MM/yyyy")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeLaudosSemNotaFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(sender As Object, e As EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Session("objCliente" & HID.Value) IsNot Nothing Then
            Dim objCliente As Cliente = Session("objCliente" & HID.Value)
            txtCliente.Text = Funcoes.FormatarListItemCliente(objCliente).Text
            txtCodigoCliente.Value = Funcoes.FormatarListItemCliente(objCliente).Value
            Session.Remove("objCliente" & HID.Value)
        End If

        If Session("objDeposito" & HID.Value) IsNot Nothing Then
            Dim objDeposito As Cliente = Session("objDeposito" & HID.Value)
            txtDeposito.Text = Funcoes.FormatarListItemCliente(objDeposito).Text
            txtCodigoDeposito.Value = Funcoes.FormatarListItemCliente(objDeposito).Value
            Session.Remove("objDeposito" & HID.Value)
        End If
    End Sub

    Protected Sub btnDeposito_Click(sender As Object, e As EventArgs) Handles btnDeposito.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objDeposito" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function VerificarGrupoProduto(ByRef Sql As String) As String
        If ucSelecaoProduto.TemSelecionado() Then
            Dim retorno As ArrayList
            retorno = ucSelecaoProduto.GetSqlEParametrosRelatorio("Pr.Grupo", "Pr.Produto_id")
            Sql &= " AND " & retorno(0)
            Return retorno(1)
        Else
            Return ""
        End If
    End Function

End Class