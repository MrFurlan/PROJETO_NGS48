Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucRateio
    Inherits BaseUserControl

    Private Property ObjNotaFiscal() As [Lib].Negocio.NotaFiscal
        Get
            Return CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
        End Get
        Set(ByVal value As [Lib].Negocio.NotaFiscal)
            Session("objNotaFiscal" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack Then
                ddl.Carregar(ddlCentroCusto, CarregarDDL.Tabela.CentroDeCustoDescricao, " Len(centrodecusto_id) = 5")
                ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnEmpresaRateio_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim ucConsultaEmpresas As ucConsultaEmpresas = CType(Me.Page.FindControlRecursive("ucConsultaEmpresas"), ucConsultaEmpresas)
            If ucConsultaEmpresas IsNot Nothing Then
                ucConsultaEmpresas.Limpar()
                ucConsultaEmpresas.MainUserControl = Me
                Popup.ConsultaDeEmpresas(Me.Page, "objRateioNFG" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        hdfIndex.Value = ""
        hdfEdicao.Value = ""
        LimparDados()
    End Sub

    Private Sub LimparDados()
        ddlCentroCusto.SelectedIndex = -1
        txtValorRateado.Text = "0,00"
        txtSaldoRateio.Text = "0,00"
        txtValorRateio.Text = "0,00"
    End Sub

    Public Sub Inicializar(ByVal posicao As Integer)
        hdfIndex.Value = posicao
        lblProduto.Text = String.Format("{0} - {1}", ObjNotaFiscal.Itens(posicao).CodigoProduto, ObjNotaFiscal.Itens(posicao).Produto.Nome)
        Dim valorTotal As Decimal = Decimal.Zero
        If ObjNotaFiscal.Itens(posicao).Rateios IsNot Nothing AndAlso ObjNotaFiscal.Itens(posicao).Rateios.Count > 0 Then
            valorTotal = ObjNotaFiscal.Itens(posicao).Rateios.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Valor)
        End If
        txtSaldoRateio.Text = (ObjNotaFiscal.Itens(posicao).ValorTotal - valorTotal).ToString("N2")
        txtTotal.Text = ObjNotaFiscal.Itens(posicao).ValorTotal.ToString("N2")
        If ObjNotaFiscal.Itens(posicao).Rateios IsNot Nothing AndAlso ObjNotaFiscal.Itens(posicao).Rateios.Count > 0 Then
            txtValorRateado.Text = ObjNotaFiscal.Itens(posicao).Rateios.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Valor).ToString("N2")
            txtValorRateio.Text = txtSaldoRateio.Text
            gridRateio.DataSource = ObjNotaFiscal.Itens(posicao).Rateios.Where(Function(s) s.IUD <> "D").ToList()
            gridRateio.DataBind()
        End If
    End Sub

    Private Sub SalvarDadosRateio()
        If Not String.IsNullOrWhiteSpace(txtValorRateio.Text) AndAlso CDec(txtValorRateio.Text) > 0 Then
            If ObjNotaFiscal.Itens(hdfIndex.Value).Rateios IsNot Nothing AndAlso ObjNotaFiscal.Itens(hdfIndex.Value).Rateios.Count > 0 Then
                For Each item As NotaFiscalxRateio In ObjNotaFiscal.Itens(hdfIndex.Value).Rateios
                    Dim lst = ObjNotaFiscal.Itens(hdfIndex.Value).Rateios.Where(Function(s) s.CodigoCentroDeCusto = ddlCentroCusto.SelectedValue).FirstOrDefault
                    If lst IsNot Nothing Then
                        If lst.IUD <> "D" Then
                            MsgBox(Me.Page, "Não é possível possuir mais de um valor rateado no mesmo centro de custo!")
                            Exit Sub
                        End If

                        If lst.IUD = "D" Then lst.IUD = "U"
                        lst.CodigoUnidadeDeNegocio = ddlUnidadeNegocio.SelectedValue
                        lst.EndUnidadeDeNegocio = 0
                        lst.CodigoEmpresaRateio = ddlEmpresa.SelectedValue.Split("-")(0)
                        lst.EndEmpresaRateio = ddlEmpresa.SelectedValue.Split("-")(1)
                        lst.Valor = txtValorRateio.Text
                        If ObjNotaFiscal.Itens(hdfIndex.Value).Rateios Is Nothing Then
                            ObjNotaFiscal.Itens(hdfIndex.Value).Rateios = New ListNotaFiscalxRateio(ObjNotaFiscal.Itens(hdfIndex.Value))
                        End If
                        Inicializar(hdfIndex.Value)
                        Exit Sub
                    End If
                Next
            End If

            If Not String.IsNullOrWhiteSpace(hdfEdicao.Value) Then
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).CodigoUnidadeDeNegocio = ddlUnidadeNegocio.SelectedValue
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).EndUnidadeDeNegocio = 0
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).CodigoEmpresaRateio = ddlEmpresa.SelectedValue.Split("-")(0)
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).EndEmpresaRateio = ddlEmpresa.SelectedValue.Split("-")(1)
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).CodigoCentroDeCusto = ddlCentroCusto.SelectedValue
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).DescCentroDeCusto = ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).CentroDeCusto.Descricao
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(hdfEdicao.Value).Valor = txtValorRateio.Text
                Inicializar(hdfIndex.Value)
            Else
                Dim objRateio As New NotaFiscalxRateio(ObjNotaFiscal.Itens(hdfIndex.Value))
                objRateio.IUD = "I"
                objRateio.CodigoUnidadeDeNegocio = ddlUnidadeNegocio.SelectedValue
                objRateio.EndUnidadeDeNegocio = 0
                objRateio.CodigoEmpresaRateio = ddlEmpresa.SelectedValue.Split("-")(0)
                objRateio.EndEmpresaRateio = ddlEmpresa.SelectedValue.Split("-")(1)
                objRateio.CodigoCentroDeCusto = ddlCentroCusto.SelectedValue
                objRateio.DescCentroDeCusto = objRateio.CentroDeCusto.Descricao
                objRateio.Valor = txtValorRateio.Text
                If ObjNotaFiscal.Itens(hdfIndex.Value).Rateios Is Nothing Then
                    ObjNotaFiscal.Itens(hdfIndex.Value).Rateios = New ListNotaFiscalxRateio(ObjNotaFiscal.Itens(hdfIndex.Value))
                End If
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios.Add(objRateio)
                Inicializar(hdfIndex.Value)
            End If
        End If
    End Sub

    Private Function ValidarCamposRateio() As Boolean
        If String.IsNullOrWhiteSpace(txtValorRateio.Text) OrElse Not IsNumeric(txtValorRateio.Text) Then txtValorRateio.Text = "0"

        If String.IsNullOrWhiteSpace(ddlUnidadeNegocio.SelectedValue) Then
            MsgBox(Me.Page, "É necessário selecionar uma unidade de negócio!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "É necessário selecionar uma empresa!")
            Return False
        ElseIf ddlCentroCusto.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "É necessário selecionar um centro de custo!")
            Return False
        ElseIf CDec(txtValorRateio.Text) <= 0 Then
            MsgBox(Me.Page, "Informe o valor a ser rateado!")
            Return False
        ElseIf CDec(txtValorRateio.Text) > (CDec(txtSaldoRateio.Text)) Then
            MsgBox(Me.Page, "O valor informado não pode ser maior do que o saldo para o rateio!")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparDados()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Try
            If Not String.IsNullOrWhiteSpace(txtSaldoRateio.Text) Then
                If CDec(txtSaldoRateio.Text) > 0 Then
                    MsgBox(Me.Page, "É necessário finalizar o saldo de rateio para completar o processo!")
                    Exit Sub
                End If
            End If
            If TypeOf Me.Page Is NotasFiscaisGerais Then
                CType(Me.Page, NotasFiscaisGerais).AtualizarItensNoGrid()
            End If
            Popup.CloseDialog(Me.Page, "divRateio")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAdicionar_Click(sender As Object, e As EventArgs) Handles lnkAdicionar.Click
        Try
            If ValidarCamposRateio() Then
                SalvarDadosRateio()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridRateio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridRateio.SelectedIndexChanged
        Try
            Dim Rateio As NotaFiscalxRateio = ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(gridRateio.SelectedRow.RowIndex)
            ddlUnidadeNegocio.SelectedValue = Rateio.CodigoUnidadeDeNegocio
            ddlEmpresa.SelectedValue = Rateio.CodigoEmpresaRateio & "-" & Rateio.EndEmpresaRateio
            ddlCentroCusto.SelectedIndex = ddlCentroCusto.Items.IndexOf(ddlCentroCusto.Items.FindByValue(Rateio.CodigoCentroDeCusto))
            txtSaldoRateio.Text = ObjNotaFiscal.Itens(hdfIndex.Value).Rateios.Sum(Function(s) s.Valor) - Rateio.Valor
            txtValorRateio.Text = Rateio.Valor
            hdfEdicao.Value = gridRateio.SelectedRow.RowIndex
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRemover_Click(sender As Object, e As EventArgs)
        Try
            Dim btn As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)

            Dim pos As Integer = ObjNotaFiscal.Itens(hdfIndex.Value).Rateios.FindIndex(Function(s) s.CodigoCentroDeCusto = gridRateio.DataKeys(row.RowIndex).Value)

            If ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(pos).IUD = "I" Then
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios.RemoveAt(pos)
            Else
                ObjNotaFiscal.Itens(hdfIndex.Value).Rateios(pos).IUD = "D"
            End If

            ObjNotaFiscal.Itens(hdfIndex.Value).IUD = "U"
            Inicializar(hdfIndex.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub SetarEmpresa(ByVal Unidade As String, ByVal Empresa As String)
        ddlUnidadeNegocio.SelectedValue = Unidade
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, Unidade)
        ddlEmpresa.SelectedValue = Empresa
    End Sub
End Class