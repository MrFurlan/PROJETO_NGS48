Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucNotaFiscalXClassificacao
    Inherits BaseUserControl

    Dim sql As String = String.Empty


#Region "Session"
    Dim objNotaFiscal As NotaFiscal
    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), NotaFiscal)
    End Sub
    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub
#End Region

#Region "Property"
    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Private ReadOnly Property SessaoDsXML() As DataSet
        Get
            Return CType(Session("dsXML" & HID.Value), DataSet)
        End Get
    End Property
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        Else
            chkPesagens.Checked = False
            chkPesagens_CheckedChanged(chkPesagens, Nothing)
        End If
    End Sub

#Region "Eventos de componentes"
    Protected Sub btnCalcular_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCalcular.Click
        Calcular()
    End Sub

    Protected Sub btnOK_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnOK.Click
        finalizarPesagem()
    End Sub

    Private Sub finalizarPesagem()
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.Romaneio.Observacoes = txtObservacao.Text
        SessaoSalvaNotaFiscal()

        Session(Session("ssTipoRetorno")) = "S"

        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objClassificacaoALT")) Then
            CType(Me.Page, AlterarNotaFiscal).CarregarClassificacao()
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objClassificacaoNXI")) Then
            CType(Me.Page, NotaFiscalXItens).CarregarClassificacao()
        End If

        Popup.CloseDialog(Me.Page, "divConsultaNotaFiscalXClassificacao")
    End Sub

    Protected Sub btnAjustar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAjustar.Click
        Try

            If Funcoes.VerificaPermissao("AjustarRomaneio", "ACESSAR") Then
                SessaoRecuperaNotaFiscal()

                Dim Sqls As New ArrayList

                objNotaFiscal.Romaneio.IUD = "U"
                objNotaFiscal.Romaneio.SalvarSql(Sqls)

                sql = "Update NotasFiscaisXItens set " & vbCrLf &
                      "  QuantidadeFisica     = " & objNotaFiscal.Romaneio.PesoLiquido & vbCrLf &
                      " Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                      "   and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                      "   and Produto_Id      ='" & objNotaFiscal.Romaneio.CodigoProduto & "'"
                Sqls.Add(sql)

                Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Ajuste Romaneio em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                Else
                    obs = "Ajuste Romaneio em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                End If

                objNotaFiscal.ObservacoesControleInterno = obs

                sql = "Update NotasFiscais set " & vbCrLf &
                      "  ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                      "	   ,UsuarioAlteracao        ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                      "	   ,UsuarioAlteracaoData    = getdate() " & vbCrLf &
                      " Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                      "   and Nota_Id         = " & objNotaFiscal.Codigo

                Sqls.Add(sql)

                If Sqls.Count > 0 Then

                    If Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, "Classificações do Romaneio atualizadas com sucesso. Não é necessário atualizar a Nota Fiscal.")

                        CType(Me.Page, NotaFiscalXItens).CarregarNotaClassificacao()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Else
                    MsgBox(Me.Page, "Houve algum problema na Reclassificação, entre em contato com o Suporte.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para ajustar Romaneio!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnCriarRomaneio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCriarRomaneio.Click
        Try
            SessaoRecuperaNotaFiscal()

            Dim Sqls As New ArrayList

            objNotaFiscal.Romaneio.IUD = "I"
            objNotaFiscal.Romaneio.Movimento = objNotaFiscal.Movimento
            objNotaFiscal.Romaneio.SalvarSql(Sqls)
            objNotaFiscal.IUD = "I"
            objNotaFiscal.CodigoRomaneio = objNotaFiscal.Romaneio.Codigo
            objNotaFiscal.SalvarNotasXRomaneio(Sqls)

            If Sqls.Count > 0 Then
                If Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, "Romaneio criado com sucesso. Não é necessário atualizar a Nota Fiscal.")

                    If TypeOf Me.Page Is NotaFiscalXItens Then
                        CType(Me.Page, NotaFiscalXItens).LimparCampos("", "", False, objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa)
                    End If

                    Popup.CloseDialog(Me.Page, "divConsultaNotaFiscalXClassificacao")

                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Houve algum problema na criação do Romaneio, entre em contato com o Suporte.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    Protected Sub btnLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnLimpar.Click
        LimparCampos()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        SessaoRecuperaNotaFiscal()
        If objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso objNotaFiscal.CodigoRomaneio <= 0 Then
            MsgBox(Me.Page, "Faça a Classificacao antes de sair.")
            Exit Sub
        End If
        Popup.CloseDialog(Me.Page, "divConsultaNotaFiscalXClassificacao")
    End Sub

    Protected Sub chkPesagens_CheckedChanged(sender As Object, e As EventArgs) Handles chkPesagens.CheckedChanged
        If chkPesagens.Checked Then
            lblPesagem.Text = "1º Pesagem"
            txtSegundaPesagem.Parent.Visible = True
        Else
            lblPesagem.Text = "Pesagem"
            txtSegundaPesagem.Parent.Visible = False
        End If

        LimparCampos()
    End Sub
#End Region


    Public Sub LimparCampos()
        'chkPesagens.Checked = False
        btnOK.Visible = False
        btnAjustar.Visible = False
        btnCriarRomaneio.Visible = False
        txtPrimeiraPesagem.Text = String.Empty
        txtSegundaPesagem.Text = String.Empty
        txtPesoBruto.Text = String.Empty
        txtDesconto.Text = String.Empty
        txtLiquido.Text = String.Empty
        txtPrimeiraPesagem.Focus()

        For Each row As GridViewRow In GridDescontos.Rows
            CType(row.FindControl("txtPercentual"), TextBox).Enabled = True
            CType(row.FindControl("txtIndice"), TextBox).Enabled = False
            CType(row.FindControl("txtDesconto"), TextBox).Enabled = False

            CType(row.FindControl("txtPercentual"), TextBox).Text = ""
            CType(row.FindControl("txtIndice"), TextBox).Text = ""
            CType(row.FindControl("txtDesconto"), TextBox).Text = ""
        Next
    End Sub

    Public Sub CarregarPercentualPadraoClassificacao(quant As Integer, produto As String, ByVal bClose As Boolean)

        Dim indexColAnalise As Integer

        For Each coluna As DataControlField In GridDescontos.Columns

            If coluna.HeaderText = "Código" Then
                indexColAnalise = GridDescontos.Columns.IndexOf(coluna)
                Exit For
            End If

        Next

        SessaoRecuperaNotaFiscal()

        For Each row As GridViewRow In GridDescontos.Rows

            'Sugerir valor padrão de classificação para Nutri e Baxi
            If Left(objNotaFiscal.CodigoEmpresa, 8) = "05366261" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "40938762" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "49673784" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "62747840" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "62780383" OrElse Left(objNotaFiscal.CodigoEmpresa, 8) = "63358210" Then

                If Left(produto, 5) = "10101" And (CType(row.FindControl("txtPercentual"), TextBox).Text = "" Or CType(row.FindControl("txtPercentual"), TextBox).Text = "0") Then

                    If row.Cells(indexColAnalise).Text = 1 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "14,00"
                    ElseIf row.Cells(indexColAnalise).Text = 2 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "1,00"
                    ElseIf row.Cells(indexColAnalise).Text = 3 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "8,00"
                    ElseIf row.Cells(indexColAnalise).Text > 3 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "0,00"
                    End If

                ElseIf Left(produto, 5) = "10102" And (CType(row.FindControl("txtPercentual"), TextBox).Text = "" Or CType(row.FindControl("txtPercentual"), TextBox).Text = "0") Then

                    If row.Cells(indexColAnalise).Text = 1 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "14,00"
                    ElseIf row.Cells(indexColAnalise).Text = 2 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "1,00"
                    ElseIf row.Cells(indexColAnalise).Text = 3 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "5,00"
                    ElseIf row.Cells(indexColAnalise).Text > 3 Then
                        CType(row.FindControl("txtPercentual"), TextBox).Text = "0,00"
                    End If

                End If

            End If

        Next

        'Calcular()

        If bClose Then
            finalizarPesagem()
        End If

    End Sub

    Public Sub BindGridView()
        If Not IsPostBack And IsConnect Then
            LimparCampos()
        Else
            If Not Session("ssTipoRetorno") Is Nothing AndAlso
                (Session("ssTipoRetorno").ToString.Contains("objClassificacaoNXI") Or
                 Session("ssTipoRetorno").ToString.Contains("objClassificacaoALT")) Then
                If Session("ProcurandoClassificacao" & HID.Value) = True Then
                    SessaoRecuperaNotaFiscal()

                    txtProduto.Value = objNotaFiscal.Itens(0).CodigoProduto
                    txtEntradaSaida.Value = objNotaFiscal.Itens(0).SubOperacao.EntradaSaida
                    If txtEntradaSaida.Value.Length > 0 Then lblEntSai.Text = IIf(objNotaFiscal.Itens(0).SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA")

                    If SessaoDsXML IsNot Nothing AndAlso objNotaFiscal.Itens.Count = 1 Then

                        txtPrimeiraPesagem.Text = objNotaFiscal.Itens(0).QuantidadeFiscal.ToString("N0")

                        txtPesoBruto.Text = objNotaFiscal.Itens(0).QuantidadeFiscal.ToString("N0")
                        txtDesconto.Text = 0
                        txtLiquido.Text = objNotaFiscal.Itens(0).QuantidadeFiscal.ToString("N0")
                        txtObservacao.Text = ""

                        Dim Rom As New Romaneio(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.CodigoRomaneio)
                        Rom.Codigo = -1
                        Rom.CodigoPedido = objNotaFiscal.CodigoPedido
                        Rom.Pedido = objNotaFiscal.Pedido
                        Rom.CodigoProduto = objNotaFiscal.Itens(0).CodigoProduto
                        Rom.EntradaSaida = Left(objNotaFiscal.SubOperacao.EntradaSaida.ToString, 1)

                        objNotaFiscal.Romaneio = Rom
                        objNotaFiscal.CodigoRomaneio = -1

                        GridDescontos.DataSource = Rom.DescontosAnalises
                        GridDescontos.DataBind()
                        SessaoSalvaNotaFiscal()

                    Else
                        If objNotaFiscal.CodigoRomaneio <> 0 Then
                            txtPesoBruto.Text = objNotaFiscal.Romaneio.PesoBruto
                            txtDesconto.Text = objNotaFiscal.Romaneio.Desconto
                            txtLiquido.Text = objNotaFiscal.Romaneio.PesoLiquido
                            txtObservacao.Text = objNotaFiscal.Romaneio.Observacoes

                            If objNotaFiscal.Romaneio.PrimeiraPesagem > 0 Then txtPrimeiraPesagem.Text = objNotaFiscal.Romaneio.PrimeiraPesagem

                            If objNotaFiscal.Romaneio.SegundaPesagem > 0 Then
                                chkPesagens.Checked = True
                                txtSegundaPesagem.Parent.Visible = chkPesagens.Checked
                                txtSegundaPesagem.Text = objNotaFiscal.Romaneio.SegundaPesagem
                            End If

                            GridDescontos.DataSource = objNotaFiscal.Romaneio.DescontosAnalises
                            GridDescontos.DataBind()
                        Else
                            Dim Rom As New Romaneio(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.CodigoRomaneio)
                            Rom.Codigo = -1
                            Rom.CodigoPedido = objNotaFiscal.CodigoPedido
                            Rom.Pedido = objNotaFiscal.Pedido
                            Rom.CodigoProduto = objNotaFiscal.Itens(0).CodigoProduto
                            Rom.EntradaSaida = Left(objNotaFiscal.SubOperacao.EntradaSaida.ToString, 1)

                            If chkPesagens.Checked Then
                                Rom.PrimeiraPesagem = CDbl(txtPrimeiraPesagem.Text)
                                Rom.SegundaPesagem = CDbl(txtSegundaPesagem.Text)
                            End If

                            objNotaFiscal.Romaneio = Rom
                            objNotaFiscal.CodigoRomaneio = -1

                            GridDescontos.DataSource = Rom.DescontosAnalises
                            GridDescontos.DataBind()
                            SessaoSalvaNotaFiscal()
                        End If

                    End If

                    DdlClassificacao.Items.Add(New ListItem(objNotaFiscal.Pedido.Itens(0).Classificacao.Descricao, objNotaFiscal.Pedido.Itens(0).Classificacao.Codigo))
                    DdlClassificacao.SelectedIndex = 0


                    For Each row As GridViewRow In GridDescontos.Rows
                        CType(row.FindControl("txtPercentual"), TextBox).Enabled = True
                        CType(row.FindControl("txtIndice"), TextBox).Enabled = False
                        CType(row.FindControl("txtDesconto"), TextBox).Enabled = False
                    Next
                    Session("ProcurandoClassificacao" & HID.Value) = False
                End If
            End If
        End If
    End Sub

    Private Sub Calcular()

        If String.IsNullOrWhiteSpace(txtPrimeiraPesagem.Text) OrElse CInt(txtPrimeiraPesagem.Text) = 0 Then
            MsgBox(Me.Page, "Pesagem deve ser maior que zero")
        ElseIf txtProduto.Value.Length = 0 Then
            MsgBox(Me.Page, "Produto para Classificação não foi selecionado")
        ElseIf txtEntradaSaida.Value.Length = 0 Then
            MsgBox(Me.Page, "Operação para Classificação não foi selecionado")
        Else
            If GridDescontos.Rows.Count > 0 Then
                Dim libera As Boolean = True

                If libera Then
                    SessaoRecuperaNotaFiscal()

                    Dim CodigoAnalise As Integer
                    Dim Percentual As Decimal

                    txtPesoBruto.Text = Math.Abs(CInt(txtPrimeiraPesagem.Text) - CInt(IIf(Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text), txtSegundaPesagem.Text, "0")))
                    txtLiquido.Text = Math.Abs(CInt(txtPrimeiraPesagem.Text) - CInt(IIf(Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text), txtSegundaPesagem.Text, "0")))

                    objNotaFiscal.Romaneio.PrimeiraPesagem = CInt(txtPrimeiraPesagem.Text)
                    objNotaFiscal.Romaneio.SegundaPesagem = CInt(IIf(Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text), txtSegundaPesagem.Text, "0"))
                    txtDesconto.Text = "0"

                    objNotaFiscal.Romaneio.PesoBruto = txtPesoBruto.Text
                    objNotaFiscal.Romaneio.PesoLiquido = txtLiquido.Text

                    For Each row As GridViewRow In GridDescontos.Rows
                        CodigoAnalise = row.Cells(0).Text()
                        Dim RxD As RomaneioXDesconto = objNotaFiscal.Romaneio.DescontosAnalises.Where(Function(s) s.CodigoAnalise = CodigoAnalise).First
                        If RxD.Analise.Opcao.Length > 0 Then
                            Percentual = CType(row.FindControl("ddlOpcao"), DropDownList).SelectedValue
                        Else
                            Percentual = CType(row.FindControl("txtPercentual"), TextBox).Text
                        End If

                        objNotaFiscal.Romaneio.DescontosAnalises.Where(Function(s) s.CodigoAnalise = CodigoAnalise).First.Percentual = Percentual
                    Next
                    Dim erro As String = ""
                    erro = objNotaFiscal.Romaneio.DescontosAnalises.CalcularDescontos

                    If erro.Length > 0 Then
                        MsgBox(Me.Page, erro)
                        For Each row As GridViewRow In GridDescontos.Rows
                            CType(row.FindControl("txtPercentual"), TextBox).Enabled = True
                            CType(row.FindControl("txtIndice"), TextBox).Text = "0"
                            CType(row.FindControl("txtDesconto"), TextBox).Text = "0"
                        Next
                        Exit Sub
                    Else
                        For Each row As GridViewRow In GridDescontos.Rows
                            CType(row.FindControl("txtPercentual"), TextBox).Enabled = False
                        Next
                        GridDescontos.DataSource = objNotaFiscal.Romaneio.DescontosAnalises
                        GridDescontos.DataBind()
                        SessaoSalvaNotaFiscal()
                    End If

                    txtDesconto.Text = objNotaFiscal.Romaneio.Desconto
                    txtLiquido.Text = objNotaFiscal.Romaneio.PesoLiquido

                    btnOK.Visible = True

                    If objNotaFiscal.IUD = "U" AndAlso Not objNotaFiscal.Romaneio.DescontosAnalises Is Nothing AndAlso objNotaFiscal.Romaneio.DescontosAnalises.Count > 0 Then
                        btnAjustar.Visible = True
                    End If

                    If objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso Not objNotaFiscal.CodigoRomaneio > 0 Then
                        btnCriarRomaneio.Visible = True
                    End If
                End If
            Else
                Dim peso As Integer = txtPrimeiraPesagem.Text

                If Not String.IsNullOrWhiteSpace(txtSegundaPesagem.Text) Then
                    peso = CInt(txtPrimeiraPesagem.Text) - CInt(txtSegundaPesagem.Text)
                End If

                txtPesoBruto.Text = IIf(peso < 0, peso * -1, peso)
                txtDesconto.Text = 0
                txtLiquido.Text = IIf(peso < 0, peso * -1, peso)

                SessaoRecuperaNotaFiscal()
                objNotaFiscal.Romaneio.PesoBruto = txtPesoBruto.Text
                objNotaFiscal.Romaneio.PesoLiquido = txtLiquido.Text

                SessaoSalvaNotaFiscal()

                btnOK.Visible = True

                If objNotaFiscal.IUD = "U" AndAlso Not objNotaFiscal.Romaneio.DescontosAnalises Is Nothing AndAlso objNotaFiscal.Romaneio.DescontosAnalises.Count > 0 Then
                    btnAjustar.Visible = True
                End If

                If objNotaFiscal.IUD = "U" AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico AndAlso Not objNotaFiscal.CodigoRomaneio > 0 Then
                    btnCriarRomaneio.Visible = True
                End If
            End If

            txtSegundaPesagem.Parent.Visible = chkPesagens.Checked
        End If
    End Sub

    Protected Sub GridDescontos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridDescontos.RowDataBound
        Select Case e.Row.RowType
            Case DataControlRowType.DataRow
                Dim RxD As RomaneioXDesconto = CType(e.Row.DataItem, RomaneioXDesconto)
                If RxD.Analise.Opcao.Length > 0 Then
                    CType(e.Row.FindControl("txtPercentual"), TextBox).Visible = False
                    CType(e.Row.FindControl("txtIndice"), TextBox).Visible = False
                    CType(e.Row.FindControl("txtDesconto"), TextBox).Visible = False

                    Dim ddlA As DropDownList = CType(e.Row.FindControl("ddlOpcao"), DropDownList)
                    ddlA.Visible = True
                    ddl.Carregar(ddlA, RxD.Analise.Opcao, ";", "-")

                    ddlA.SelectedValue = CInt(RxD.Percentual)
                    e.Row.Cells(2).ColumnSpan = 3
                End If
        End Select
    End Sub
End Class