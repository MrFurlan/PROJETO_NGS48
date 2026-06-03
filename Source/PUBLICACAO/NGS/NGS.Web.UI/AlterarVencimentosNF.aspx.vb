Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Xml
Imports System.Threading
Imports System.IO

Public Class AlterarVencimentosNF
    Inherits BasePage

    Dim objNotaFiscal As NotaFiscal
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AlterarVencimentosNF", "ACESSAR") Then
                ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsultaNFTitulo" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaNFTitulo" & HID.Value), NotaFiscal))

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                txtUf.Text = objNotaFiscal.Cliente.CodigoEstado
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie

                For Each t In objNotaFiscal.VencimentosNota
                    Dim drTitulo As DataRow = CType(Session("objTitulos" & HID.Value), DataTable).NewRow()

                    drTitulo("Codigo") = t.Codigo
                    drTitulo("Vencimento") = t.Prorrogacao.ToString("dd/MM/yyyy")
                    drTitulo("ValorDoDocumento") = t.ValorDoDocumento

                    CType(Session("objTitulos" & HID.Value), DataTable).Rows.Add(drTitulo)
                Next

                GridConsultaTitulos.DataSource = CType(Session("objTitulos" & HID.Value), DataTable)
                GridConsultaTitulos.DataBind()

                lblTotal.Text = objNotaFiscal.VencimentosNota.Where(Function(s) s.CodigoSituacao = 1).Sum(Function(s) s.ValorDoDocumento).ToString
                lblApurado.Text = CType(Session("objTitulos" & HID.Value), DataTable).Compute("SUM(ValorDoDocumento)", "").ToString
                lblDiferenca.Text = lblTotal.Text - lblApurado.Text

                txtUf.Enabled = False
                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False
                txtDataInicial.Enabled = False
                txtDataFinal.Enabled = False
                ddlEmpresa.Enabled = False
                ddlUnidadeDeNegocio.Enabled = False
                btnCliente.Enabled = False

                SalvaNotaFiscal()
                Session.Remove("objNFConsultaNFTitulo" & HID.Value)

                lnkConsultar.Parent.Visible = False

                If objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Any(Function(s) s.CodigoSituacao = eSituacao.RemessaBancaria) Then
                    MsgBox(Me.Page, "Não é possível alterar os Vencimentos da nota fiscal com o financeiro em Cobrança Bancária.")
                    lnkAtualizar.Parent.Visible = False
                    Exit Sub
                ElseIf objNotaFiscal.VencimentosNota.Where(Function(s) s.IUD <> "D").Any(Function(s) s.CodigoSituacao = eSituacao.EndossoTitulo) Then
                    MsgBox(Me.Page, "Não é possível alterar os Vencimentos da nota fiscal com o financeiro em Endosso.")
                    lnkAtualizar.Parent.Visible = False
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarVencimentosNF", "LEITURA") Then
                If validaConsulta() Then
                    If objNotaFiscal Is Nothing Then
                        objNotaFiscal = New [Lib].Negocio.NotaFiscal
                    End If

                    objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                    objNotaFiscal.Movimento = CDate(txtDataFinal.Text)
                    objNotaFiscal.NossaEmissao = False
                    objNotaFiscal.CodigoSituacao = 1
                    objNotaFiscal.CodigoTipoDeDocumento = 0

                    If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                        objNotaFiscal.CodigoCliente = txtCodigoCliente.Value.Split("-")(0)
                        objNotaFiscal.EnderecoCliente = txtCodigoCliente.Value.Split("-")(1)
                    End If

                    If txtES.Text.Length > 0 Then objNotaFiscal.EntradaSaida = IIf(txtES.Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)

                    If txtNota.Text.Length > 0 Then
                        objNotaFiscal.Codigo = txtNota.Text
                    Else
                        objNotaFiscal.Codigo = 0
                    End If

                    If txtSerie.Text.Length > 0 Then
                        objNotaFiscal.Serie = txtSerie.Text
                    Else
                        objNotaFiscal.Serie = ""
                    End If

                    SalvaNotaFiscal()

                    Session("ssCampo" & HID.Value) = "NotaXItens"

                    txtApurado.Visible = True
                    txtTotal.Visible = True
                    txtDiferenca.Visible = True

                    If ucConsultaPedidosXNotas.BindGridView() > 0 Then
                        Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNFTitulo" & HID.Value.ToString)
                    Else
                        MsgBox(Me.Page, "Nenhum resultado encontrado.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro(s).")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarVencimentosNF", "ALTERAR") Then
                If lblTotal.Text = lblApurado.Text Then
                    Dim Sqls As New ArrayList
                    RecuperaNotaFiscal()

                    For Each t In CType(Session("objTitulos" & HID.Value), DataTable).Rows
                        For Each nt As Titulo In objNotaFiscal.VencimentosNota
                            If t("Codigo") = nt.Codigo Then
                                nt.Carregando = False
                                nt.ValorDoDocumento = t("ValorDoDocumento")
                                nt.Prorrogacao = CDate(t("Vencimento"))
                                nt.IUD = "U"
                                nt.UsuarioAlteracao = Session("ssNomeUsuario")
                                nt.UsuarioAlteracaoData = Date.Now
                                nt.SalvarSql(Sqls)
                                Exit For
                            End If
                        Next
                    Next

                    Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Titulos alterados em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " pelo processo AlterarVencimentosNF "
                    Else
                        obs = "Titulos alterados em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " pelo processo AlterarVencimentosNF."
                    End If

                    Dim Sql As String
                    Sql = " Update NotasFiscais set " & vbCrLf & _
                         "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                         "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                         "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                         "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                         "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                         "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                         "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                         "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                         "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                         "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"

                    Sqls.Add(Sql)

                    If Sqls.Count > 0 Then
                        If Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, "Alterado com Sucesso.", eTitulo.Sucess)
                            limpar()
                        Else
                            Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                            MsgBox(Me.Page, Observacao)
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Valor ajustado está diferente do valor original.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro(s).")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Function validaConsulta() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período para consulta.")
            Return False
        End If
        Return True
    End Function

    Private Sub limpar()
        Session.Remove("objNFConsultaNFTitulo" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("ssCampo" & HID.Value)
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objTitulos" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        Dim dtTitulos As New DataTable("Titulos")

        dtTitulos.Columns.Add("Codigo", Type.GetType("System.String"))
        dtTitulos.Columns.Add("Vencimento", Type.GetType("System.String"))
        dtTitulos.Columns.Add("ValorDoDocumento", Type.GetType("System.Decimal"))
        Session("objTitulos" & HID.Value) = dtTitulos

        txtUf.Enabled = True
        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtDataInicial.Enabled = True
        txtDataFinal.Enabled = True
        ddlEmpresa.Enabled = True
        ddlUnidadeDeNegocio.Enabled = True

        btnCliente.Enabled = True

        lblApurado.Text = String.Empty
        lblTotal.Text = String.Empty
        lblDiferenca.Text = String.Empty
        txtApurado.Visible = False
        txtTotal.Visible = False
        txtDiferenca.Visible = False

        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtUf.Text = String.Empty
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        lnkConsultar.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False

        LiberaEmpresa()

        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)

        GridConsultaTitulos.DataSource = CType(Session("objTitulos" & HID.Value), DataTable)
        GridConsultaTitulos.DataBind()

        Try
            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            SalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub btnAlterarTitulo_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.Visible = False

        Dim linha As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex
        CType(GridConsultaTitulos.Rows(linha).FindControl("btnSalvarTitulo"), Button).Visible = True
        CType(GridConsultaTitulos.Rows(linha).FindControl("btnCancelarTitulo"), Button).Visible = True
        CType(GridConsultaTitulos.Rows(linha).FindControl("txtValorDoDocumento"), TextBox).Enabled = True
        CType(GridConsultaTitulos.Rows(linha).FindControl("txtVencimento"), TextBox).Enabled = True
    End Sub

    Protected Sub btnCancelarTitulo_Click(sender As Object, e As EventArgs)
        GridConsultaTitulos.DataSource = CType(Session("objTitulos" & HID.Value), DataTable)
        GridConsultaTitulos.DataBind()
    End Sub

    Protected Sub btnSalvarTitulo_Click(sender As Object, e As EventArgs)
        Dim btn As Button = CType(sender, Button)
        btn.Visible = False
        RecuperaNotaFiscal()

        Dim linha As Integer = CType(btn.NamingContainer, GridViewRow).RowIndex
        CType(GridConsultaTitulos.Rows(linha).FindControl("btnCancelarTitulo"), Button).Visible = False
        CType(GridConsultaTitulos.Rows(linha).FindControl("btnAlterarTitulo"), Button).Visible = True
        CType(GridConsultaTitulos.Rows(linha).FindControl("txtValorDoDocumento"), TextBox).Enabled = False
        CType(GridConsultaTitulos.Rows(linha).FindControl("txtVencimento"), TextBox).Enabled = False

        For Each t In CType(Session("objTitulos" & HID.Value), DataTable).Rows
            If GridConsultaTitulos.Rows(linha).Cells(0).Text = t("Codigo") Then
                t("ValorDoDocumento") = CDec(CType(GridConsultaTitulos.Rows(linha).FindControl("txtValorDoDocumento"), TextBox).Text)
                t("Vencimento") = Funcoes.ValidaDataUtil(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, CDate(CType(GridConsultaTitulos.Rows(linha).FindControl("txtVencimento"), TextBox).Text)).ToString("dd/MM/yyyy")
            End If
        Next

        lblApurado.Text = CType(Session("objTitulos" & HID.Value), DataTable).Compute("SUM(ValorDoDocumento)", "").ToString
        lblDiferenca.Text = (lblTotal.Text - lblApurado.Text).ToString("N2")

        lnkAtualizar.Parent.Visible = True

        GridConsultaTitulos.DataSource = CType(Session("objTitulos" & HID.Value), DataTable)
        GridConsultaTitulos.DataBind()
    End Sub

End Class