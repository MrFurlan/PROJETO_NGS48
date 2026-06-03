Imports System.Data
Imports System.Drawing
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TrocaDeNota
    Inherits BasePage

    Protected objCliente As New Cliente

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TrocaDeNota", "ACESSAR") Then
                CargaEmpresas()
                Limpar()
                txtDataInicial.Text = Now().ToShortDateString()
                txtDataFinal.Text = Now().ToShortDateString()
                txtDataInicialEntrada.Text = Now().ToShortDateString()
                txtDataFinalEntrada.Text = Now().ToShortDateString()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
        ddl.Carregar(ddlEmpresaEntrada, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub Limpar()
        Session.Remove("objNotaParaTroca" & HID.Value.ToString)
        Session.Remove("objListaDeNotas")
        Session.Remove("objTrocaDeNota" & HID.Value.ToString)

        gridTrocaDeNota.DataSource = Nothing
        gridTrocaDeNota.DataBind()

        ddlEmpresaEntrada.ClearSelection()
        ddlEmpresaEntrada.Enabled = False
        cmdConsultaClienteEntrada.Enabled = False
        lblNotaFiscal.Text = ""
        txtClientesEntrada.Text = ""
        txtCodigoClienteEntrada.Value = ""
        txtDataInicialEntrada.Enabled = False
        txtDataFinalEntrada.Enabled = False
        txtNomeDoCliente.Text = ""
        txtboxEntSai.Text = ""
        txtboxSerie.Text = ""
        txtboxNota.Text = ""
        ddlEmpresaEntrada.BackColor = Color.White
        txtClientesEntrada.BackColor = Color.White
        txtDataInicialEntrada.BackColor = Color.White
        txtDataFinalEntrada.BackColor = Color.White
        btnConsultarEntrada.Enabled = False
        btnVincular.Enabled = False
        HID.Value = Guid.NewGuid().ToString
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub ConsultarRegistro()
        Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim strCliente As String = IIf(String.IsNullOrWhiteSpace(hdnEnderecoCliente.Value), "", txtCodigoCliente.Value)
        Dim strEndCliente As String = String.Empty

        If Not String.IsNullOrWhiteSpace(hdnEnderecoCliente.Value) Then
            strEndCliente = hdnEnderecoCliente.Value
        End If

        Dim ListaDeNotas As New [Lib].Negocio.ListNotasFiscais(Empresa(0), Empresa(1), txtDataInicial.Text, txtDataFinal.Text, strCliente, strEndCliente, "S")

        Session("objListaDeNotas") = ListaDeNotas

        If ListaDeNotas.Count > 0 Then
            Dim ds As New DataSet
            Dim tbTroca As New DataTable("Clientes")
            tbTroca.Columns.Add("Nota", Type.GetType("System.String"))
            tbTroca.Columns.Add("Movimento", Type.GetType("System.DateTime"))
            tbTroca.Columns.Add("CodigoPedido", Type.GetType("System.String"))
            tbTroca.Columns.Add("QuantidadeFiscal", Type.GetType("System.String"))
            tbTroca.Columns.Add("OEmpresa", Type.GetType("System.String"))
            tbTroca.Columns.Add("OEndEmpresa", Type.GetType("System.String"))
            tbTroca.Columns.Add("OCliente", Type.GetType("System.String"))
            tbTroca.Columns.Add("OEndCliente", Type.GetType("System.String"))
            tbTroca.Columns.Add("ONomeCliente", Type.GetType("System.String"))
            tbTroca.Columns.Add("OEntradaSaida", Type.GetType("System.String"))
            tbTroca.Columns.Add("OSerie", Type.GetType("System.String"))
            tbTroca.Columns.Add("ONota", Type.GetType("System.String"))
            ds.Tables.Add(tbTroca)

            For Each nf As [Lib].Negocio.NotaFiscal In ListaDeNotas
                Dim drRow As DataRow = ds.Tables(0).NewRow()
                drRow("Nota") = nf.Codigo
                drRow("Movimento") = nf.Movimento
                drRow("CodigoPedido") = nf.CodigoPedido
                drRow("QuantidadeFiscal") = nf.Itens(0).PesoFiscal

                If Not nf.NotasXNotas Is Nothing AndAlso Not nf.NotasXNotas.EmpresaCnpj Is Nothing Then
                    drRow("OEmpresa") = nf.NotasXNotas.OrigemEmpresaCnpj
                    drRow("OEndEmpresa") = nf.NotasXNotas.OrigemEndEmpresa
                    drRow("OCliente") = nf.NotasXNotas.OrigemClienteCnpj
                    drRow("OEndCliente") = nf.NotasXNotas.OrigemEndCliente
                    Dim objCliente As New [Lib].Negocio.Cliente(nf.NotasXNotas.OrigemClienteCnpj, nf.NotasXNotas.OrigemEndCliente)
                    drRow("ONomeCliente") = objCliente.Nome
                    drRow("OEntradaSaida") = nf.NotasXNotas.OrigemEntradaSaida
                    drRow("OSerie") = nf.NotasXNotas.OrigemSerie
                    drRow("ONota") = nf.NotasXNotas.OrigemNota
                End If

                ds.Tables(0).Rows.Add(drRow)
            Next

            gridTrocaDeNota.DataSource = ds
            gridTrocaDeNota.DataBind()

            Dim i As Integer = 0
            While i < gridTrocaDeNota.Rows.Count
                If gridTrocaDeNota.Rows(i).Cells(5).Text = "&nbsp;" Then
                    CType(gridTrocaDeNota.Rows(i).FindControl("imgExcluir"), ImageButton).Visible = False
                    CType(gridTrocaDeNota.Rows(i).FindControl("imgConsultar"), ImageButton).Visible = True
                Else
                    CType(gridTrocaDeNota.Rows(i).FindControl("imgExcluir"), ImageButton).Visible = True
                    CType(gridTrocaDeNota.Rows(i).FindControl("imgConsultar"), ImageButton).Visible = False
                End If

                i += 1
            End While
        End If
    End Sub

    Private Sub ConsultarEntrada()
        If Funcoes.VerificaPermissao("NotaFiscalXItens", "LEITURA") Then
            Dim strJavaScript As String = ""
            Dim Empresa() As String = ddlEmpresaEntrada.SelectedValue.ToString.Split("-")
            Dim Cliente() As String = txtCodigoClienteEntrada.Value.Split("-")

            Session("ssCampo") = "NotaXItens"

            Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal
            objNotaFiscal = CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(HgridRowIndexTroca.Value)
            If objNotaFiscal.Itens.Count = 0 Then
                MsgBox(Me.Page, "A Nota deve conter um Produto para a troca de nota ser realizada")
                Exit Sub
            End If

            If objNotaFiscal.Itens(0).Produto.Agrupar.Equals("S") Then
                MsgBox(Me.Page, "Não é permitido fazer a troca de nota para este produto")
                Exit Sub
            End If
            If objNotaFiscal.Pedido.SubOperacao.PrecoFixo = False And objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) = "E" Then
                MsgBox(Me.Page, "A Devolução para Compra só é permitida em operações de preço Fixo")
                Exit Sub
            End If

            Dim objNotaTroca As New [Lib].Negocio.NotaFiscal
            objNotaTroca.CodigoEmpresa = Empresa(0)
            objNotaTroca.EnderecoEmpresa = Empresa(1)
            objNotaTroca.CodigoCliente = txtCodigoClienteEntrada.Value
            objNotaTroca.EnderecoCliente = hdnEnderecoClienteEntrada.Value
            objNotaTroca.EntradaSaida = objNotaFiscal.EntradaSaida
            objNotaTroca.DataInclusao = txtDataInicialEntrada.Text
            objNotaTroca.DataTermino = txtDataFinalEntrada.Text
            objNotaTroca.Itens = objNotaFiscal.Itens

            Session("objNotaFiscal" & HID.Value.ToString) = objNotaTroca

            ucConsultaNotaTroca.SetarHID(HID.Value)
            ucConsultaNotaTroca.InicializarFormulario(True)
            Popup.ConsultaDeNotaTroca(Me, "objTrocaDeNota" & HID.Value.ToString)
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar Registro")
        End If
    End Sub

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            'Popup.ConsultarClientes(Me, "TROxCLI", txtClientes, txtCodigoCliente, "", True)
            hdnControlePopup.Value = "ClienteNota"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
            Popup.ConsultaDeClientes(Me.Page, "objTrocaDeNota" & HID.Value.ToString)
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparCliente_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            txtClientes.Text = ""
            txtCodigoCliente.Value = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgConsultar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("TrocaDeNota", "LEITURA") Then
                Dim imgNotaFiscal As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgNotaFiscal.NamingContainer, GridViewRow)

                HgridRowIndexTroca.Value = row.RowIndex

                lblNotaFiscal.Text = CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(HgridRowIndexTroca.Value).Codigo
                'sessão para nota pesquisa que será vinculada
                Session("objNotaParaTroca" & HID.Value.ToString) = CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(HgridRowIndexTroca.Value)
                ddlEmpresaEntrada.Enabled = True
                cmdConsultaClienteEntrada.Enabled = True
                txtDataInicialEntrada.Enabled = True
                txtDataFinalEntrada.Enabled = True
                btnConsultarEntrada.Enabled = True
                txtClientesEntrada.Text = ""
                txtCodigoClienteEntrada.Value = ""
                txtNomeDoCliente.Text = ""
                txtboxEntSai.Text = ""
                txtboxSerie.Text = ""
                txtboxNota.Text = ""
                ddlEmpresaEntrada.BackColor = Color.LightBlue
                txtClientesEntrada.BackColor = Color.LightBlue
                txtDataInicialEntrada.BackColor = Color.LightBlue
                txtDataFinalEntrada.BackColor = Color.LightBlue
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("TrocaDeNota", "EXCLUIR") Then
                Dim imgNotaFiscal As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgNotaFiscal.NamingContainer, GridViewRow)

                If CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).TemNotaTroca Then
                    CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).IUD = "D"
                    If CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(row.RowIndex).AtualizarTrocaDeNota Then
                        Limpar()
                        ConsultarRegistro()
                        MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaClienteEntrada_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            'Popup.ConsultarClientes(Me, "TROxCLIE", txtClientesEntrada, txtCodigoClienteEntrada, "", True)
            hdnControlePopup.Value = "ClienteTrocaNota"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
            Popup.ConsultaDeClientes(Me.Page, "objTrocaDeNota" & HID.Value.ToString)
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultarEntrada_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConsultarEntrada.Click
        Try
            If ddlEmpresaEntrada.SelectedIndex = 0 Then
                MsgBox(Me.Page, "A Empresa da Nota Fiscal de Entrada não foi selecionada.")
                ddlEmpresaEntrada.Focus()
            ElseIf String.IsNullOrWhiteSpace(txtCodigoClienteEntrada.Value) Then
                MsgBox(Me.Page, "O Cliente da Nota Fiscal de Entrada não foi selecionado.")
                cmdConsultaClienteEntrada.Focus()
            Else
                ConsultarEntrada()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnVincular_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            'Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal(CType(Session("objListaDeNotas"), [Lib].Negocio.ListNotasFiscais)(HgridRowIndexTroca.Value))
            Dim objNotaFiscal = CType(Session("objNotaParaTroca" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
            objNotaFiscal.IUD = "I"
            objNotaFiscal.NotaTrocaOrigem = CType(Session("objTrocaDeNota" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
            If objNotaFiscal.AtualizarTrocaDeNota Then
                Limpar()
                ConsultarRegistro()
                MsgBox(Me.Page, "Registro incluído com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objTrocaDeNota" & HID.Value.ToString) Is Nothing Then
            objCliente = CType(Session("objTrocaDeNota" & HID.Value.ToString), [Lib].Negocio.Cliente)
            If (hdnControlePopup.Value.Equals("ClienteNota")) Then
                txtClientes.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado & " " & Funcoes.AlinharEsquerda(objCliente.CodigoFormatado, 18, ".") & "-" & CStr(objCliente.CodigoEndereco) & "-" & objCliente.Reduzido
                txtCodigoCliente.Value = objCliente.Codigo
                hdnEnderecoCliente.Value = objCliente.CodigoEndereco
            Else
                txtClientesEntrada.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado & " " & Funcoes.AlinharEsquerda(objCliente.CodigoFormatado, 18, ".") & "-" & CStr(objCliente.CodigoEndereco) & "-" & objCliente.Reduzido
                txtCodigoClienteEntrada.Value = objCliente.Codigo
                hdnEnderecoClienteEntrada.Value = objCliente.CodigoEndereco
            End If
        End If
    End Sub

    Public Sub CarregarNotaTroca()
        Dim objNotaFiscal = CType(Session("objTrocaDeNota" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
        If (objNotaFiscal IsNot Nothing) Then
            txtNomeDoCliente.Text = objNotaFiscal.Cliente.Nome
            txtboxEntSai.Text = objNotaFiscal.EntradaSaida.ToString()
            txtboxSerie.Text = objNotaFiscal.Serie
            txtboxNota.Text = objNotaFiscal.Codigo
            btnVincular.Enabled = True
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao("TrocaDeNota", "RELATORIO") Then
            If ddlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Empresa não foi selecionada")
            Else
                ConsultarRegistro()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
        End If
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
            Funcoes.Ajuda(Me.Page, "TrocaDeNota")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class