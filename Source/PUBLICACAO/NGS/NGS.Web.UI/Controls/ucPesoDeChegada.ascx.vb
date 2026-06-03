Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucPesoDeChegada
    Inherits BaseUserControl

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Protected Sub txtBrutoDeChegada_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtBrutoDeChegada.Text.Length = 0 Then txtBrutoDeChegada.Text = 0
        txtDesconto.Text = 0

        If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) Then
            txtLiquido.Text = String.Format("{0:N0}", CDec(txtBrutoDeChegada.Text))
        End If

        Dim vlrBrutoDeSaida As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtBrutoDeSaida.Text) Then
            vlrBrutoDeSaida = CDec(txtBrutoDeSaida.Text)
        End If

        Dim vlrBrutoDeChegada As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) Then
            vlrBrutoDeChegada = CDec(txtBrutoDeChegada.Text)
        End If

        txtSaldo.Text = String.Format("{0:N0}", vlrBrutoDeChegada - vlrBrutoDeSaida)

        If Not Session("objTituloContratoDeFrete") Is Nothing Then
            Dim tolerancia As Decimal = Math.Round(CType(Session("objNFPesoDeChegada"), [Lib].Negocio.NotaFiscal).Itens(0).QuantidadeFiscal * 0.25 / 100, 2, MidpointRounding.AwayFromZero)
            Dim quebra As Decimal = CType(Session("objNFPesoDeChegada"), [Lib].Negocio.NotaFiscal).Itens(0).QuantidadeFiscal - CDec(txtBrutoDeChegada.Text)
            If quebra > tolerancia Then
                CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).Descontos = Math.Round(CType(Session("objNFPesoDeChegada"), [Lib].Negocio.NotaFiscal).Itens(0).Unitario * quebra, 2, MidpointRounding.AwayFromZero)
                'txtValorQuebra.Text = CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).Descontos.ToString("N2")
                'txtValorLiquido.Text = CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).ValorLiquido.ToString("N2")
            End If
        End If
    End Sub

    Protected Sub txtDesconto_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim vlrBrutoDeChegada As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) Then
            vlrBrutoDeChegada = CDec(txtBrutoDeChegada.Text)
        End If

        Dim vlrDesconto As Decimal = 0
        If Not String.IsNullOrWhiteSpace(txtDesconto.Text) Then
            vlrDesconto = CDec(txtDesconto.Text)
        End If

        txtLiquido.Text = String.Format("{0:N2}", vlrBrutoDeChegada - vlrDesconto)
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub BindGridView(ByVal Produto As String, ByVal Peso As Decimal)
        txtBrutoDeSaida.Text = String.Format("{0:N0}", Peso)
        txtSaldo.Text = String.Format("{0:N0}", Decimal.Zero)
        CarregarClassificacao(Produto)
    End Sub

    Public Sub LimparCampos()
        txtDataDeChegada.Text = ""
        txtBrutoDeChegada.Text = ""
        txtDesconto.Text = ""
        txtLiquido.Text = ""
        txtTarifaFrete.Text = 0
        txtValorDoFrete.Text = 0
        GridDescontos.DataSource = Nothing
        GridDescontos.DataBind()
    End Sub

    Private Sub LimparSessoes()
        Session.Remove("objNFPesoDeChegada")
        Session.Remove("objContratoDeFrete")
        Session.Remove("objTituloContratoDeFrete")
    End Sub

    Private Sub CarregarClassificacao(ByVal Produto As String)
        Sql = "SELECT Analises.Analise_Id as Codigo, Analises.Descricao, 0.00 as Percentual, 0 as Indice, 0 as Desconto " & vbCrLf & _
              "  FROM ProdutosXAnalises INNER JOIN Analises ON ProdutosXAnalises.Analise_Id = Analises.Analise_Id " & vbCrLf & _
              " WHERE ProdutosXAnalises.Produto_Id = '" & Produto & "' AND Analises.Analise_Id < 100 " & vbCrLf & _
              " ORDER BY Analises.Analise_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            GridDescontos.DataSource = ds
            GridDescontos.DataBind()
            txtDesconto.Enabled = False
        Else
            txtDesconto.Enabled = True
        End If

    End Sub

    Private Sub InclusaoPesoDeChegada()
        If Funcoes.VerificaPermissao("PesosDeChegada", "GRAVAR") Then
            If Session("objNFPesoDeChegada") Is Nothing Then
                MsgBox(Me.Page, "É necessário selecionar uma nota fiscal!")
                Exit Sub
            End If

            If String.IsNullOrWhiteSpace(txtDataDeChegada.Text) OrElse String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) OrElse String.IsNullOrWhiteSpace(txtLiquido.Text) Then
                MsgBox(Me.Page, "É necessário informar os campos data de chegada, bruto de chegada e líquido!")
                Exit Sub
            End If

            If Not String.IsNullOrWhiteSpace(txtBrutoDeChegada.Text) AndAlso Not CInt(txtBrutoDeChegada.Text) > 0 Then
                MsgBox(Me.Page, "É necessário informar bruto de chegada maior que zero!")
                Exit Sub
            End If

            If Not String.IsNullOrWhiteSpace(txtLiquido.Text) AndAlso Not CInt(txtLiquido.Text) > 0 Then
                MsgBox(Me.Page, "É necessário informar líquido de chegada maior que zero!")
                Exit Sub
            End If

            Dim objNFDestino As New [Lib].Negocio.NotaFiscalXDestino(CType(Session("objNFPesoDeChegada"), [Lib].Negocio.NotaFiscal))
            objNFDestino.IUD = "I"
            objNFDestino.Movimento = txtDataDeChegada.Text
            objNFDestino.PesoBruto = IIf(txtBrutoDeChegada.Text.Length = 0, 0, txtBrutoDeChegada.Text)
            objNFDestino.Desconto = IIf(txtDesconto.Text.Length = 0, 0, txtDesconto.Text)
            objNFDestino.PesoLiquido = IIf(txtLiquido.Text.Length = 0, 0, txtLiquido.Text)
            objNFDestino.Sinistro = chkSinistro.Checked
            objNFDestino.TarifaFrete = CDec(txtTarifaFrete.Text)

            Dim i As Integer
            For i = 0 To GridDescontos.Rows.Count - 1
                If Not String.IsNullOrWhiteSpace(CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text) AndAlso CDec(CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text) > 0 Then
                    Dim objDesc As New [Lib].Negocio.NotaFiscalXDestinoXDescontos(objNFDestino.NF.Itens(0))
                    objDesc.IUD = "I"
                    objDesc.CodigoAnalise = GridDescontos.Rows(i).Cells(0).Text()
                    objDesc.Percentual = IIf(CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text.Length = 0, 0, CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text)
                    objDesc.Indice = IIf(CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text.Length = 0, 0, CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text)
                    objDesc.Desconto = IIf(CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text.Length = 0, 0, CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text)
                    objNFDestino.NF.Itens(0).DescontosPesoDeChegada.Add(objDesc)
                End If
            Next

            Dim sqls As New ArrayList
            objNFDestino.SalvarSql(sqls)
            objNFDestino.NF.Itens(0).DescontosPesoDeChegada.SalvarSql(sqls)

            If Session("objTituloContratoDeFrete") IsNot Nothing Then
                CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).IUD = "U"
                CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).SalvarSql(sqls)
            End If

            If Banco.GravaBanco(sqls) Then
                LimparCampos()
                LimparSessoes()
                If TypeOf Me.Page Is PesosDeChegada Then
                    CType(Me.Page, PesosDeChegada).CarregarNotasFiscais()
                End If
                Popup.CloseDialog(Me.Page, "divPesoDeChegada")
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
        End If
    End Sub

    'Protected Sub txtOutrosDescontos_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    'CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).Deducoes = IIf(txtOutrosDescontos.Text.Length = 0, 0, CDec(txtOutrosDescontos.Text))
    '    'txtValorLiquido.Text = CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).ValorLiquido.ToString("N2")
    'End Sub

    'Protected Sub txtEstadia_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    'CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).Juros = IIf(txtEstadia.Text.Length = 0, 0, CDec(txtEstadia.Text))
    '    'txtValorLiquido.Text = CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).ValorLiquido.ToString("N2")
    'End Sub

    'Protected Sub txtOutrosAcrescimos_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
    '    'CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).Juros = IIf(txtOutrosAcrescimos.Text.Length = 0, 0, CDec(txtOutrosAcrescimos.Text))
    '    'txtValorLiquido.Text = CType(Session("objTituloContratoDeFrete"), [Lib].Negocio.Titulo).ValorLiquido.ToString("N2")
    'End Sub

    'Protected Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
    '    Try
    '        InclusaoPesoDeChegada()
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    'Protected Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
    '    Popup.CloseDialog(Me.Page, "divPesoDeChegada")
    'End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divPesoDeChegada")
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            InclusaoPesoDeChegada()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class