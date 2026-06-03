Imports NGS.Lib.Negocio

Public Class AlterarMovimentoNotaFiscal
    Inherits BasePage

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AlterarMovimentoNotaFiscal", "ACESSAR") Then
                    ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                    ddl.Carregar(cmbSituacao, CarregarDDL.Tabela.Situacao, "situacao_id in(1,4)")
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
            If validaConsulta() Then
                If objNotaFiscal Is Nothing Then
                    objNotaFiscal = New [Lib].Negocio.NotaFiscal
                End If

                objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                objNotaFiscal.Movimento = CDate(txtDataFinal.Text)
                objNotaFiscal.NossaEmissao = False

                If cmbSituacao.SelectedIndex > 0 Then
                    objNotaFiscal.CodigoSituacao = cmbSituacao.SelectedValue
                Else
                    objNotaFiscal.CodigoSituacao = 1
                    cmbSituacao.SelectedValue = 1
                End If

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
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgMenosHora_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            RecuperaNotaFiscal()

            objNotaFiscal.DataInclusao = objNotaFiscal.DataInclusao.AddHours(-1)

            txtDataInclusao.Text = objNotaFiscal.DataInclusao.ToString("dd/MM/yyyy HH:mm:ss")

            SalvaNotaFiscal()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgMaisHora_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            RecuperaNotaFiscal()

            objNotaFiscal.DataInclusao = objNotaFiscal.DataInclusao.AddHours(1)

            txtDataInclusao.Text = objNotaFiscal.DataInclusao.ToString("dd/MM/yyyy HH:mm:ss")

            SalvaNotaFiscal()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjustarHora_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAjustarHora.Click
        Try
            If Funcoes.VerificaPermissao("AlterarMovimentoNotaFiscal", "ALTERAR") Then

                RecuperaNotaFiscal()

                If objNotaFiscal IsNot Nothing Then
                    If objNotaFiscal.AtualizaHoraNota() Then
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        limpar()
                    Else
                        MsgBox(Me.Page, "Erro alteração da Nota Fiscal.", eTitulo.Sucess)
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


    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarMovimentoNotaFiscal", "ALTERAR") Then

                RecuperaNotaFiscal()

                If validaAlteracao() Then
                    If objNotaFiscal IsNot Nothing Then
                        objNotaFiscal.DataNota = CDate(txtDataEmissao.Text)
                        objNotaFiscal.Movimento = CDate(txtDataES.Text)
                        If objNotaFiscal.AtualizaDatasNota() Then
                            MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                            limpar()
                        Else
                            MsgBox(Me.Page, "Erro alteração da Nota Fiscal.", eTitulo.Sucess)
                        End If
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



    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlterarMovimentoNotaFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsulta" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsulta" & HID.Value), NotaFiscal))

                divDatas.Visible = True

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
                txtUf.Text = objNotaFiscal.Cliente.CodigoEstado
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie
                txtDataEmissao.Text = objNotaFiscal.DataNota.ToString("dd/MM/yyyy")
                txtDataES.Text = objNotaFiscal.Movimento.ToString("dd/MM/yyyy")

                txtDataInclusao.Text = objNotaFiscal.DataInclusao.ToString("dd/MM/yyyy HH:mm:ss")

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

                If objNotaFiscal.NossaEmissao Then

                    Dim Sql As String
                    Sql = " SELECT NFP.Retorno, NFP.MsgRetorno " & vbCrLf & _
                          "  FROM NfePendencias NFP" & vbCrLf & _
                          "	WHERE NFP.Empresa_Id      = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                          "   AND NFP.EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                          "   AND NFP.Cliente_Id      = '" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                          "   AND NFP.EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                          "   AND NFP.EntradaSaida_Id = '" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                          "   AND NFP.Serie_Id        = '" & objNotaFiscal.Serie & "'" & vbCrLf & _
                          "   AND NFP.Nota_Id         = " & objNotaFiscal.Codigo

                    Dim ds As New DataSet

                    ds = Banco.ConsultaDataSet(Sql, "Notas")

                    If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                        If Trim(ds.Tables(0).Rows(0).Item("Retorno")) = "703" Then
                            lnkAjustarHora.Parent.Visible = True
                            lnkAtualizar.Parent.Visible = False
                            lnkConsultar.Parent.Visible = False

                            imgMenosHora.Enabled = True
                            imgMaisHora.Enabled = True

                            txtDataEmissao.Enabled = False
                            txtDataES.Enabled = False
                        Else
                            lnkAjustarHora.Parent.Visible = False
                            lnkAtualizar.Parent.Visible = False
                            lnkConsultar.Parent.Visible = True
                            MsgBox(Me.Page, "Nota Fiscal Nossa Emissão não pode ser alterada.", eTitulo.Info)
                        End If
                    Else
                        lnkAjustarHora.Parent.Visible = False
                        lnkAtualizar.Parent.Visible = False
                        lnkConsultar.Parent.Visible = True
                        MsgBox(Me.Page, "Nota Fiscal Nossa Emissão não pode ser alterada.", eTitulo.Info)
                    End If
                Else
                    txtDataEmissao.Enabled = True
                    txtDataES.Enabled = True

                    lnkAtualizar.Parent.Visible = True
                    lnkAjustarHora.Parent.Visible = False
                    lnkConsultar.Parent.Visible = False
                End If

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
        cmbSituacao.Enabled = True
        ddlUnidadeDeNegocio.Enabled = True
        btnCliente.Enabled = True
        divDatas.Visible = False

        txtCliente.Text = ""
        txtUf.Text = ""
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        txtDataEmissao.Text = ""
        txtDataES.Text = ""
        lnkAtualizar.Parent.Visible = False
        lnkAjustarHora.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        imgMenosHora.Enabled = False
        imgMaisHora.Enabled = False

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)

        LiberaEmpresa()

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

    Private Sub CarregarEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
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

    Private Function validaAlteracao() As Boolean

        If String.IsNullOrWhiteSpace(txtDataEmissao.Text) OrElse String.IsNullOrWhiteSpace(txtDataES.Text) Then
            MsgBox(Me.Page, "Informe a data de emissao e data de entrada/saida para alteração.")
            Return False
        ElseIf Not IsDate(txtDataEmissao.Text) OrElse Not IsDate(txtDataES.Text) Then
            MsgBox(Me.Page, "A data de emissao ou a data de entrada/saida não é uma data válida.")
            Return False
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, txtDataEmissao.Text, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento da Nota Fiscal já Fechado para esta data.")
            Return False
        ElseIf Not Funcoes.VerificaAcesso(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, txtDataEmissao.Text, "CONTABIL") Then
            MsgBox(Me.Page, "Movimento Contábil já Fechado para esta data.")
            Return False
        ElseIf (CDate(txtDataEmissao.Text) > DateTime.Now OrElse CDate(txtDataES.Text) > DateTime.Now) AndAlso Not Funcoes.VerificaPermissao("AlterarMovimentoNotaFiscal", "LIBERAR") Then
            MsgBox(Me.Page, "A data de emissao ou a data de entrada/saida não pode ser maior que a data atual - " & DateTime.Now.ToString("dd/MM/yyyy") & ".")
            Return False
        ElseIf (CDate(txtDataEmissao.Text) > CDate(txtDataES.Text)) AndAlso Not Funcoes.VerificaPermissao("AlterarMovimentoNotaFiscal", "LIBERAR") Then
            MsgBox(Me.Page, "A data de emissao não pode ser maior que a data de entrada/saida.")
            Return False
        ElseIf Not objNotaFiscal.Pedido Is Nothing AndAlso Not objNotaFiscal.Pedido.CodigoMoeda = 1 Then
            MsgBox(Me.Page, "Nota Fiscal em Dólar não pode ser alterada por essa rotina.")
            Return False
        ElseIf Not objNotaFiscal.Romaneio Is Nothing AndAlso _
                objNotaFiscal.Romaneio.Codigo > 0 AndAlso _
                Not objNotaFiscal.Romaneio.Pesagens Is Nothing AndAlso _
                objNotaFiscal.Romaneio.Pesagens.Count > 0 AndAlso _
                CDate(txtDataES.Text) < objNotaFiscal.Romaneio.Movimento Then
            MsgBox(Me.Page, "A data de entrada/saida não pode ser menor que a data da Pesagem.")
            Return False
        Else
            Return True
        End If
    End Function

End Class