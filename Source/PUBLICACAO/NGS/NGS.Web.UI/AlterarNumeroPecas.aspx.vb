Imports NGS.Lib.Negocio

Public Class AlterarNumeroPecas
    Inherits BasePage

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AlterarNumeroPecas", "ACESSAR") Then
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

        txtCliente.Text = ""
        txtUf.Text = ""
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        LiberaEmpresa()

        GridAlterarNumeroPecas.DataBind()

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

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
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

                Dim dt As New DataTable("Itens")

                dt.Columns.Add("Produto", Type.GetType("System.String"))
                dt.Columns.Add("NomeProduto", Type.GetType("System.String"))
                dt.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
                dt.Columns.Add("Unitario", Type.GetType("System.Decimal"))
                dt.Columns.Add("Valor", Type.GetType("System.Decimal"))
                dt.Columns.Add("NumeroPecas", Type.GetType("System.Decimal"))

                For Each row As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                    Dim dr As DataRow = dt.NewRow()
                    dr("Produto") = row.CodigoProduto
                    dr("NomeProduto") = row.Produto.Nome
                    dr("Quantidade") = row.QuantidadeFiscal.ToString("N4")
                    dr("Unitario") = row.Unitario.ToString("N10")
                    dr("Valor") = row.ValorLiquido.ToString("N4")
                    dr("NumeroPecas") = row.NumeroPecas.ToString()
                    dt.Rows.Add(dr)
                Next

                GridAlterarNumeroPecas.DataSource = dt
                GridAlterarNumeroPecas.DataBind()

                SalvaNotaFiscal()
                Session.Remove("objNFConsulta" & HID.Value)
                lnkAtualizar.Parent.Visible = True
                lnkConsultar.Parent.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
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

    Protected Sub ChkNumeroPecas_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim chk As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(chk.NamingContainer, GridViewRow)
            Dim txt As TextBox = CType(row.FindControl("txtNumeroPecas"), TextBox)
            Dim hdfNum As HiddenField = CType(row.FindControl("hdfNumeroPecas"), HiddenField)

            If chk.Checked Then
                txt.Enabled = True
            Else
                txt.Enabled = False
                txt.Text = hdfNum.Value
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

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarNumeroPecas", "ALTERAR") Then
                Dim sql As String
                Dim sqls As New ArrayList

                RecuperaNotaFiscal()

                For Each row As GridViewRow In GridAlterarNumeroPecas.Rows
                    If CType(row.FindControl("ChkNumeroPecas"), CheckBox).Checked Then

                        For Each item As NotaFiscalXItem In objNotaFiscal.Itens
                            If item.CodigoProduto = Trim(row.Cells(1).Text) Then

                                sql = "Update NotasFiscaisXItens" & vbCrLf &
                                      "   set NumeroPecas = " & CType(row.FindControl("txtNumeroPecas"), TextBox).Text & vbCrLf &
                                      " Where  Empresa_Id     ='" & item.NotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                      "   and EndEmpresa_Id   = " & item.NotaFiscal.EnderecoEmpresa & vbCrLf &
                                      "   and Cliente_Id      ='" & item.NotaFiscal.CodigoCliente & "'" & vbCrLf &
                                      "   and EndCliente_Id   = " & item.NotaFiscal.EnderecoCliente & vbCrLf &
                                      "   and EntradaSaida_Id ='" & item.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                      "   and Nota_Id         = " & item.NotaFiscal.Codigo & vbCrLf &
                                      "   and Serie_Id        ='" & item.NotaFiscal.Serie & "'" & vbCrLf &
                                      "   and Produto_Id      ='" & item.CodigoProduto & "'" & vbCrLf &
                                      "   and CFOP_ID         = " & item.CFOP & vbCrLf &
                                      "   and Sequencia_id    = " & item.Sequencia


                                sqls.Add(sql)

                            End If
                        Next

                    End If

                Next

                If sqls IsNot Nothing AndAlso sqls.Count > 0 Then
                    Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Alterado numero de pecas em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                    Else
                        obs = "Alterado numero de pecas em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                    End If

                    sql = "Update NotasFiscais" & vbCrLf &
                          "   set ObservacoesControleInterno = '" & obs & "'" & vbCrLf &
                          " Where  Empresa_Id     ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                          "   and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                          "   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                          "   and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                          "   and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "   and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                          "   and Serie_Id        ='" & objNotaFiscal.Serie & "'"

                    sqls.Add(sql)

                    If Banco.GravaBanco(sqls) Then
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        limpar()
                    Else
                        MsgBox(Me.Page, "Erro a Altererar Registro.")
                    End If
                Else
                    MsgBox(Me.Page, "Sem Registros para alteração.")
                End If

            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro(s).")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarNumeroPecas", "ALTERAR") Then

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
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro(s).")
                Exit Sub
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlterarMovimentoNotaFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class