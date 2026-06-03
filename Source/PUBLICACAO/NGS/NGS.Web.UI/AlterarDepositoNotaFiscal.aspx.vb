Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AlterarDepositoNotaFiscal
    Inherits BasePage

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AlterarDepositoNotaFiscal", "ACESSAR") Then
                CargaEmpresas()
                Funcoes.VerificaEmpresa(ddlEmpresa)
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objClienteADxCLI" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteADxCLI" & HID.Value), [Lib].Negocio.Cliente))
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteADxCLI" & HID.Value)
        ElseIf Not Session("objClienteADxID" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.Deposito = CType(Session("objClienteADxID" & HID.Value), [Lib].Negocio.Cliente)

            Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Deposito)
            With objNotaFiscal.Deposito
                objNotaFiscal.CodigoDeposito = .Codigo
                objNotaFiscal.EnderecoDeposito = .CodigoEndereco

                txtDepositos.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                txtCodigoDeposito.Value = itemDeposito.Value
            End With

            For Each row As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                row.CodigoDeposito = objNotaFiscal.CodigoDeposito
                row.EnderecoDeposito = objNotaFiscal.EnderecoDeposito
            Next

            SessaoSalvaNotaFiscal()
            Session.Remove("objClienteADxID" & HID.Value)
        ElseIf Not Session("objClienteAODxID" & HID.Value) Is Nothing Then
            SessaoRecuperaNotaFiscal()
            objNotaFiscal.Destino = CType(Session("objClienteAODxID" & HID.Value), [Lib].Negocio.Cliente)

            Dim itemDestino As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Destino)
            With objNotaFiscal.Destino
                objNotaFiscal.CodigoDestino = .Codigo
                objNotaFiscal.EnderecoDestino = .CodigoEndereco

                txtOriDess.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                txtCodigoOriDes.Value = itemDestino.Value
            End With

            SessaoSalvaNotaFiscal()
            Session.Remove("objClienteAODxID" & HID.Value)
        End If

    End Sub

    Public Sub CarregarNotaFiscal()
        If Not Session("objNFConsultaNXI" & HID.Value) Is Nothing Then
            objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaNXI" & HID.Value), NotaFiscal))
            objNotaFiscal.IUD = "U"

            If objNotaFiscal.CodigoDeposito.Trim.Length > 0 Then
                Dim itemDeposito As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Deposito)

                With objNotaFiscal.Deposito
                    txtDepositos.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                    txtCodigoDeposito.Value = itemDeposito.Value
                End With
            End If

            If objNotaFiscal.CodigoDestino.Trim.Length > 0 Then
                Dim itemDestino As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Destino)

                With objNotaFiscal.Destino
                    txtOriDess.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                    txtCodigoOriDes.Value = itemDestino.Value
                End With
            End If

            txtNomeDoCliente.Text = objNotaFiscal.Cliente.Nome
            txtEntSai.Text = objNotaFiscal.EntradaSaida
            txtSerie.Text = objNotaFiscal.Serie
            txtNota.Text = objNotaFiscal.Codigo

            btnDeposito.Enabled = True
            btnOriDes.Enabled = True

            lnkAtualizar.Parent.Visible = True
            lnkConsultar.Parent.Visible = False

            SessaoSalvaNotaFiscal()
            Session.Remove("objNFConsultaNXI" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        Session.Remove("objClienteADxCLI" & HID.Value)
        Session.Remove("objClienteADxID" & HID.Value)
        Session.Remove("objClienteAODxID" & HID.Value)
        Session.Remove("objNFConsultaNXI" & HID.Value)

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)

        objNotaFiscal = New [Lib].Negocio.NotaFiscal
        txtDataInicial.Text = Now().ToString("dd/MM/yyyy")
        txtDataFinal.Text = Now().ToString("dd/MM/yyyy")
        txtClientes.Text = ""
        txtCodigoCliente.Value = ""
        txtNotaFiscal.Text = ""
        txtNomeDoCliente.Text = ""
        txtEntSai.Text = ""
        txtSerie.Text = ""
        txtNota.Text = ""
        txtDepositos.Text = ""
        txtOriDess.Text = ""
        txtCodigoDeposito.Value = ""
        txtCodigoOriDes.Value = ""

        btnDeposito.Enabled = False
        btnOriDes.Enabled = False

        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        LiberaEmpresa()

        SessaoSalvaNotaFiscal()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteADxCLI" & HID.Value.ToString, "txtNome")
    End Sub

    Protected Sub btnDeposito_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteADxID" & HID.Value.ToString, "txtDepositos")
    End Sub

    Protected Sub btnOriDes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteAODxID" & HID.Value.ToString, "txtOriDess")
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarDepositoNotaFiscal", "LEITURA") Then
                If ddlEmpresa.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Empresa não foi selecionada")
                ElseIf txtCodigoCliente.Value.ToString.Length = 0 AndAlso String.IsNullOrWhiteSpace(txtNotaFiscal.Text) Then
                    MsgBox(Me.Page, "Cliente não foi selecionado")
                Else
                    SessaoRecuperaNotaFiscal()
                    Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                    Dim strJavaScript As String = ""
                    objNotaFiscal.CodigoEmpresa = Empresa(0)
                    objNotaFiscal.EnderecoEmpresa = Empresa(1)

                    If Not String.IsNullOrWhiteSpace(txtNotaFiscal.Text) Then
                        objNotaFiscal.Codigo = txtNotaFiscal.Text
                    Else
                        objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                        objNotaFiscal.Movimento = CDate(txtDataFinal.Text)
                        Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                        objNotaFiscal.CodigoCliente = Cliente(0)
                        objNotaFiscal.EnderecoCliente = Cliente(1)
                    End If

                    Session("ssCampo" & HID.Value) = "NotaXItens"
                    Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaNXI" & HID.Value)
                    ucConsultaPedidosXNotas.BindGridView()
                    lnkConsultar.Parent.Visible = False
                    lnkAtualizar.Parent.Visible = True
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarDepositoNotaFiscal", "ALTERAR") Then
                SessaoRecuperaNotaFiscal()
                If objNotaFiscal.AtualizarDeposito Then
                    Limpar()
                    MsgBox(Me.Page, "Depósito alterado com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro")
            End If
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
            Funcoes.Ajuda(Me.Page, "AlterarDepositoNotaFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class