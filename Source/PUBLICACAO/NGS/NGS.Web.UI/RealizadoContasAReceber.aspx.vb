Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RealizadoContasAReceber
    Inherits BasePage

    Private Sql As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RealizadoContasAReceber", "ACESSAR") Then
                    CargaUnidadeDeNegocio()
                    CargaCarteiras()
                    Limpar()
                    txtPeriodoInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtPeriodoFinal.Text = Format(Today, "dd/MM/yyyy")
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub CargaCarteiras()
        ddl.Carregar(ddlCarteira, CarregarDDL.Tabela.CarteiraFinanceira, " Classificacao = 'R'", True)
    End Sub

    Private Sub Limpar()
        Session.Remove("objClienteRXP" & HID.Value)
        Session.Remove("objExtratoDeConta" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPlanoDeContas.SetarHID(HID.Value)

        hdfConta.Value = ""
        txtConta.Text = ""

        txtCliente.Text = ""
        txtCodigoCliente.Value = ""

        ddlCarteira.SelectedIndex = 0
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConta_Click(sender As Object, e As EventArgs) Handles btnConta.Click
        HttpContext.Current.Session("ssCampo") = "ExtratoDeConta"
        ucConsultaPlanoDeContas.Limpar()
        ucConsultaPlanoDeContas.BindGridView(True)
        Popup.ConsultaDePlanoDeContas(Me.Page, "objExtratoDeConta" & HID.Value)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteRXP" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteRXP" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteRXP" & HID.Value)
        ElseIf Not Session("objExtratoDeConta" & HID.Value) Is Nothing Then
            Dim objConta As [Lib].Negocio.PlanoDeConta = CType(obj, [Lib].Negocio.PlanoDeConta)
            hdfConta.Value = objConta.Conta
            txtConta.Text = objConta.Conta & "-" & objConta.Titulo
            Session.Remove("objExtratoDeConta" & HID.Value)
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            HttpContext.Current.Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteRXP" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

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
        Dim dsRelatorio As New DataSet
        Dim Cliente As String = ""
        Dim Campo() As String = Nothing

        Try
            Dim strSQL As String = "SELECT CR.Registro_Id AS Registro, CR.Baixa, CR.Prorrogacao as Vencimento, CR.Empresa AS EmpresaCodigo, " & vbCrLf & _
                                   "CR.EndEmpresa AS EmpresaCodigoEnd, CR.Cliente AS ClienteCodigo, CR.EndCliente AS ClienteCodigoEnd, " & vbCrLf & _
                                   "CR.Historico, " & IIf(hdfConta.Value.Length > 0, "CR.ValorDoDocumento", "CR.ValorLiquido") & " AS ValorBaixado, E.Nome AS EmpresaNome, E.Estado as EmpresaEstado, " & vbCrLf & _
                                   "E.Cidade AS EmpresaCidade, E.Reduzido AS EmpresaReduzido, C.Nome AS ClienteNome " & vbCrLf & _
                                   "FROM ContasAReceber CR " & vbCrLf & _
                                   "INNER JOIN Clientes E " & vbCrLf & _
                                   "ON CR.Empresa = E.Cliente_Id " & vbCrLf & _
                                   "AND CR.EndEmpresa = E.Endereco_Id " & vbCrLf & _
                                   "INNER JOIN Clientes C " & vbCrLf & _
                                   "ON CR.Cliente = C.Cliente_Id " & vbCrLf & _
                                   "AND CR.EndCliente = C.Endereco_Id " & vbCrLf & _
                                   "INNER JOIN ComprasXProdutos CxP" & vbCrLf & vbCrLf & _
                                   "ON CxP.Produto_Id = CR.Carteira" & vbCrLf & vbCrLf & _
                                   "WHERE CR.Situacao = 1 " & vbCrLf & _
                                   "AND CR.Provisao = 1 AND (cr.Grupado <> 'M' )" & vbCrLf

            If radMovimento.Checked Then
                strSQL &= "AND CR.Baixa BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "' " & vbCrLf
            ElseIf radVencimento.Checked Then
                strSQL &= "AND CR.Prorrogacao BETWEEN '" & txtPeriodoInicial.Text.ToSqlDate() & "' AND '" & txtPeriodoFinal.Text.ToSqlDate() & "' " & vbCrLf
            End If

            '"AND (CR.ContaContabilCliente not  like '1010101%' and CR.ContaContabilCliente not like '1010102%') " & _
            ' (CR.ContaContabilCliente not  like '1010101%' and CR.ContaContabilCliente not like '1010102%')   'Isolando contas CAIXA e BANCOS C/  MOVIMENTO  INSOL

            If ddlUnidadeDeNegocio.SelectedIndex > 0 Then strSQL &= "AND CR.UnidadeDeNegocio = '" & ddlUnidadeDeNegocio.SelectedValue & "' " & vbCrLf

            If ddlEmpresa.SelectedIndex > 0 Then
                Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
                strSQL &= "AND CR.Empresa = '" & strEmpresa(0) & "' " & vbCrLf
                strSQL &= "AND CR.EndEmpresa = " & strEmpresa(1) & " " & vbCrLf
            End If

            If hdfConta.Value.ToString.Length > 0 Then
                strSQL &= "AND CxP.ContaClientes = '" & hdfConta.Value & "' " & vbCrLf
            End If

            If txtCodigoCliente.Value.ToString.Length > 0 Then
                Dim strCliente As String() = txtCodigoCliente.Value.ToString.Split("-")
                If chkConsolidarCliente.Checked Then
                    strSQL &= "AND left(CR.Cliente,8) = '" & strCliente(0).Substring(0, 8) & "' " & vbCrLf
                Else
                    strSQL &= "AND CR.Cliente = '" & strCliente(0) & "' " & vbCrLf
                    strSQL &= "AND CR.EndCliente = " & strCliente(1) & " " & vbCrLf
                End If
            End If

            If ddlCarteira.SelectedIndex > 0 Then strSQL &= "AND CR.Carteira = '" & ddlCarteira.SelectedValue & "' " & vbCrLf

            strSQL &= "ORDER BY CR.Prorrogacao, CR.Empresa, CR.Cliente, CR.EndCliente " & vbCrLf

            dsRelatorio = Banco.ConsultaDataSet(strSQL, "Ds_realizado")

            Dim empresa As String = ""
            Dim endempresa As String = ""
            If ddlEmpresa.Text <> "" Then
                Cliente = ddlEmpresa.SelectedValue
                Campo = Cliente.Split("-")
                empresa = Campo(0)                      'EmpresaCliente
                endempresa = Campo(1)                   'Endereco Empresa Cliente
            Else
                empresa = ""                            'Empresa Cliente
                endempresa = 0                          'Endereco Empresa Cliente
            End If

            Dim NomeEmpresa As String = ""
            Dim CidadeEmpresa As String = ""
            Dim EstadoEmpresa As String = ""
            Dim DS As DataSet
            Dim dr As DataRow
            Sql = "  SELECT Clientes.Cliente_Id AS Cliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, 'CONSOLIDADO' as Descricao"
            Sql &= " FROM Clientes "
            'Sql &= " WHERE Clientes.Cliente_Id = '" & HttpContext.Current.Session("ssEmpresa") & "'"
            Sql &= " WHERE Clientes.Cliente_Id + convert(nvarchar,Clientes.Endereco_Id) = '" & empresa & endempresa & "' "

            DS = Banco.ConsultaDataSet(Sql, "Clientes")
            If DS.Tables(0).Rows.Count > 0 Then
                For Each dr In DS.Tables(0).Rows
                    NomeEmpresa = dr("Nome")
                    CidadeEmpresa = dr("Cidade")
                    EstadoEmpresa = dr("Estado")
                    'UnidadeDeNegocio = dr("Descricao")
                    Exit For
                Next
            Else
                NomeEmpresa = HttpContext.Current.Session("ssNomeEmpresa")
                CidadeEmpresa = "Geral"
                EstadoEmpresa = ""
            End If

            Dim param As New Dictionary(Of String, Object)
            param.Add("Titulo", "Realizado Contas A Receber - " & txtPeriodoInicial.Text & " à " & txtPeriodoFinal.Text)
            param.Add("Nome", NomeEmpresa)
            param.Add("Cidade", Trim(CidadeEmpresa) & " " & IIf(Trim(EstadoEmpresa) <> "", "-", "") & " " & Trim(EstadoEmpresa))

            Funcoes.BindReport(Me.Page, dsRelatorio, "Cr_RealizadoContasAReceber", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), param)

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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RealizadoContasAReceber")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class