Imports NGS.Lib.Negocio
Public Class AlterarFinalidadeNotaFiscal
    Inherits BasePage

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AlterarFinalidadeNotaFiscal", "ACESSAR") Then
                    ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                    ddl.Carregar(ddlFinalidade, CarregarDDL.Tabela.Finalidade)
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
                    objNotaFiscal.Serie = String.Empty
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

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarFinalidadeNotaFiscal", "ALTERAR") Then
                RecuperaNotaFiscal()

                If validaAlteracao() Then
                    If objNotaFiscal IsNot Nothing Then
                        If String.IsNullOrWhiteSpace(objNotaFiscal.ObservacoesControleInterno) Then
                            objNotaFiscal.ObservacoesControleInterno = "Alterado a Finalidade da Nota de " & objNotaFiscal.CodigoFinalidade & " para " & ddlFinalidade.SelectedValue
                        Else
                            objNotaFiscal.ObservacoesControleInterno = objNotaFiscal.ObservacoesControleInterno & " Alterado a Finalidade da Nota de " & objNotaFiscal.CodigoFinalidade & " para " & ddlFinalidade.SelectedValue
                        End If
                        objNotaFiscal.CodigoFinalidade = ddlFinalidade.SelectedValue
                        If objNotaFiscal.AtualizaFinalidadeNota() Then
                            MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                            limpar()
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
            Funcoes.Ajuda(Me.Page, "AlterarFinalidadeNotaFiscal")
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
                ddlFinalidade.SelectedValue = objNotaFiscal.CodigoFinalidade
                ViewState("Finalidade") = objNotaFiscal.CodigoFinalidade

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
                lnkAtualizar.Parent.Visible = True
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
        divDatas.Visible = False

        txtCliente.Text = String.Empty
        txtUf.Text = String.Empty
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        ddlFinalidade.SelectedValue = String.Empty
        ViewState("Finalidade") = String.Empty
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

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
        RecuperaNotaFiscal()

        If ddlFinalidade.SelectedValue.Equals(ViewState("Finalidade").ToString()) Then
            MsgBox(Me.Page, "Nota Fiscal ja consta com esta finalidade.")
            Return False
        ElseIf objNotaFiscal.CodigoFinalidade = 20 AndAlso objNotaFiscal.NotaTrocaOrigem IsNot Nothing AndAlso objNotaFiscal.CodigoTipoDeDocumento = 1 Then
            MsgBox(Me.Page, "Nota Fiscal com vinculo de entrada não pode ser alterado.")
            Return False
        Else
            Return True
        End If
    End Function
End Class