Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Public Class DesvincularRomaneioNotaFiscal
    Inherits BasePage

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("DesvincularRomaneioNotaFiscal", "ACESSAR") Then
                    ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                    limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged() Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CarregarEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("DesvincularRomaneioNotaFiscal", "LEITURA") Then
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

                    If ucConsultaPedidosXNotas.BindGridView() > 0 Then
                        Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsulta" & HID.Value.ToString)
                    Else
                        MsgBox(Me.Page, "Nenhum resultado encontrado.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkRomaneio_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim chk As CheckBox = CType(sender, CheckBox)

            If chk.Checked Then
                lnkAtualizar.Parent.Visible = True
            Else
                lnkAtualizar.Parent.Visible = False
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAtualizar.Click
        RecuperaNotaFiscal()

        If Funcoes.VerificaPermissao("DesvincularRomaneioNotaFiscal", "ALTERAR") Then
            Dim Sqls As New ArrayList
            Dim Sql As String = String.Empty

            Dim obs As String = objNotaFiscal.ObservacoesControleInterno

            If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                obs = obs & ". Removido vinculo do Romaneio em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
            Else
                obs = "Removido vinculo do Romaneio em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
            End If

            Sql = " Update NotasFiscais set " & vbCrLf & _
                      "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                      "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                      "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                      "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                      "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
                      "    and Nota_Id         = " & objNotaFiscal.Codigo
            Sqls.Add(Sql)

            Sql = " Delete from NotasFiscaisXRomaneios " & vbCrLf & _
                      "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "	   and Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                      "    and EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                      "    and Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
                      "    and Romaneio_Id     = " & objNotaFiscal.Romaneio.Codigo
            Sqls.Add(Sql)

            Sql = "UPDATE NotasFiscaisXItens SET" & vbCrLf &
                         "    QuantidadeFisica = 0" & vbCrLf &
                         " WHERE Empresa_id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                         "   AND EndEmpresa_id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                         "   AND Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                         "   AND EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                         "   AND EntradaSaida_Id ='" & Left(objNotaFiscal.EntradaSaida.ToString, 1) & "'" & vbCrLf &
                         "   AND Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                         "   AND Nota_Id         = " & objNotaFiscal.Codigo & ";"
            Sqls.Add(Sql)

            If objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                Sql = " Delete from RomaneiosXDescontos " & vbCrLf & _
                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                          "    and Romaneio_Id     = " & objNotaFiscal.Romaneio.Codigo
                Sqls.Add(Sql)

                Sql = " Delete from Romaneios " & vbCrLf & _
                          "  Where Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                          "    and Romaneio_Id     = " & objNotaFiscal.Romaneio.Codigo
                Sqls.Add(Sql)
            End If

            If Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, "Vínculo do Romaneio removido com sucesso da Nota Fiscal.", eTitulo.Sucess)
                limpar()
            Else
                Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                MsgBox(Me.Page, Observacao)
                Exit Sub
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub CarregarEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
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

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsulta" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsulta" & HID.Value), NotaFiscal))

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                txtUf.Text = objNotaFiscal.Cliente.CodigoEstado
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie

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

                Session.Remove("objNFConsulta" & HID.Value)

                If objNotaFiscal.Romaneio Is Nothing OrElse objNotaFiscal.Romaneio.Codigo = 0 Then
                    MsgBox(Me.Page, "Nota Fiscal não possui romaneio.", eTitulo.Info)
                    Exit Sub
                End If

                lnkConsultar.Parent.Visible = False

                Dim dtRomaneio As New DataTable("Romaneio")
                dtRomaneio.Columns.Add("Produto", Type.GetType("System.String"))
                dtRomaneio.Columns.Add("Nome", Type.GetType("System.String"))
                dtRomaneio.Columns.Add("QuantidadeFisica", Type.GetType("System.Decimal"))
                dtRomaneio.Columns.Add("QuantidadeFiscal", Type.GetType("System.Decimal"))
                dtRomaneio.Columns.Add("Romaneio", Type.GetType("System.String"))
                dtRomaneio.Columns.Add("Origem", Type.GetType("System.String"))

                Dim drItem As DataRow = dtRomaneio.NewRow()

                drItem("Produto") = objNotaFiscal.Itens(0).CodigoProduto
                drItem("Nome") = objNotaFiscal.Itens(0).Produto.Nome
                drItem("QuantidadeFisica") = objNotaFiscal.Itens(0).QuantidadeFisica
                drItem("QuantidadeFiscal") = objNotaFiscal.Itens(0).QuantidadeFiscal
                drItem("Romaneio") = objNotaFiscal.Romaneio.Codigo

                If objNotaFiscal.Romaneio.Processo = "NOTA FISCAL" Then
                    drItem("Origem") = "NOTA FISCAL"
                ElseIf objNotaFiscal.Romaneio.Processo = "Rateio" Then
                    drItem("Origem") = "RATEIO"
                ElseIf objNotaFiscal.Romaneio.Processo = "AGRUPAMENTO" Then
                    drItem("Origem") = "AGRUPAMENTO"
                Else
                    drItem("Origem") = "PESAGEM"
                End If

                dtRomaneio.Rows.Add(drItem)

                grdRomaneioNF.DataSource = dtRomaneio
                grdRomaneioNF.DataBind()

                lnkConsultar.Parent.Visible = False

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub limpar()
        Session.Remove("objNFConsulta" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("ssCampo" & HID.Value)
        Session.Remove("objCliente" & HID.Value)

        txtUf.Enabled = True
        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtDataInicial.Enabled = True
        txtDataFinal.Enabled = True
        ddlEmpresa.Enabled = True
        ddlUnidadeDeNegocio.Enabled = True

        btnCliente.Enabled = True

        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtUf.Text = String.Empty
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        LiberaEmpresa()

        grdRomaneioNF.DataSource = Nothing
        grdRomaneioNF.DataBind()

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)

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

End Class