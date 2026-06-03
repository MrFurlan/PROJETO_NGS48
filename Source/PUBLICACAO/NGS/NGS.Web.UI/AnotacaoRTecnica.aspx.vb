Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports CrystalDecisions.CrystalReports.Engine

Public Class AnotacaoRTecnica
    Inherits BasePage

    Private objArt As ART
    Private strJavaScript As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AnotacaoRTecnica", "ACESSAR") Then
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaART" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaART" & HID.Value), Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
            Session.Remove("objEmpresaART" & HID.Value)
        ElseIf Not Session("objClienteART" & HID.Value) Is Nothing Then
            Dim itemAgronomo As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteART" & HID.Value), Cliente))
            txtAgronomo.Text = itemAgronomo.Text
            txtCodigoAgronomo.Value = itemAgronomo.Value
            Session.Remove("objClienteART" & HID.Value)
            txtArtInicial.Focus()
        End If
    End Sub

    Private Sub IniciarART(ByVal Tipo As String)
        If Funcoes.VerificaPermissao("AnotacaoRTecnica", "GRAVAR") Then
            If (Tipo = "I" And Globais.GPermiteGravar = "S") Or (Tipo = "U" And Globais.GPermiteAlterar = "S") Or (Tipo = "D" And Globais.GPermiteExcluir = "S") Then
                If ValidarCampos() Then
                    Dim objArt As New ART
                    Dim Empresa() As String = txtCodigoEmpresa.Value.ToString.Split("-")
                    Dim Agronomo() As String = txtCodigoAgronomo.Value.ToString.Split("-")
                    objArt.CodigoArt = txtCodigo.Text
                    objArt.CodigoEmpresa = Empresa(0)
                    objArt.EndEmpresa = Empresa(1)
                    objArt.CodigoAgronomo = Agronomo(0)
                    objArt.EndAgronomo = Agronomo(1)
                    objArt.ARTInicial = txtArtInicial.Text
                    objArt.ARTFinal = txtArtFinal.Text
                    objArt.ARTValidade = CType(txtValidade.Text, Date)
                    objArt.Status = IIf(radSim.Checked, True, False)
                    objArt.UsuarioInclusao = Session("ssNomeUsuario")
                    objArt.UsuarioInclusaoData = Date.Now
                    objArt.IUD = Tipo
                    If objArt.Salvar Then
                        Select Case Tipo
                            Case "I"
                                MsgBox(Me.Page, "Registro Inserido com Sucesso.", eTitulo.Sucess)
                            Case "U"
                                MsgBox(Me.Page, "Registro Alterado com Sucesso.", eTitulo.Sucess)
                            Case "D"
                                MsgBox(Me.Page, "Registro Deletado com Sucesso.", eTitulo.Sucess)
                        End Select

                        Limpar()
                        AtualizarGrid()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
                    End If
                End If
            Else
                Select Case Tipo
                    Case "I"
                        MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
                    Case "U"
                        MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
                    Case "D"
                        MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
                End Select
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar resgistro.")
        End If

    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtEmpresa.Text = ""
        txtCodigoEmpresa.Value = ""
        txtAgronomo.Text = ""
        txtCodigoAgronomo.Value = ""
        txtArtInicial.Text = ""
        txtArtFinal.Text = ""
        txtValidade.Text = ""

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        Session.Remove("objEmpresaART" & HID.Value)
        Session.Remove("objClienteART" & HID.Value)

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        txtCodigo.Focus()
    End Sub

    Private Sub AtualizarGrid()
        Dim Lista As New ListArt(True)
        gridArt.DataSource = Lista.ToArray
        gridArt.DataBind()
    End Sub

    Private Function ValidarCampos() As Boolean
        If txtCodigo.Text.ToString.Length = 0 OrElse txtCodigo.Text = "0" Then
            MsgBox(Me.Page, "Código ART não foi informado.")
            Return False
        ElseIf txtCodigoEmpresa.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Empresa não foi informada.")
            Return False
        ElseIf txtCodigoAgronomo.Value.ToString.Length = 0 Then
            MsgBox(Me.Page, "Agrônomo não foi informado.")
            Return False
        ElseIf txtArtInicial.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Número ART Inicial não foi informada.")
            Return False
        ElseIf txtArtFinal.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Número ART Final não foi informada.")
            Return False
        ElseIf txtValidade.Text.ToString.Length = 0 Then
            MsgBox(Me.Page, "Válidade da ART não foi informada.")
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "Livre"
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaART" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnAgronomo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAgronomo.Click
        Try
            Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteART" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridArt_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim objArt As New ART(gridArt.SelectedRow.Cells(1).Text())

            txtCodigo.Text = objArt.CodigoArt

            Dim objEmpresa As New Cliente(objArt.CodigoEmpresa, objArt.EndEmpresa)
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value

            Dim objAgronomo As New Cliente(objArt.CodigoAgronomo, objArt.EndAgronomo)
            Dim itemAgronomo As ListItem = Funcoes.FormatarListItemCliente(objAgronomo)
            txtAgronomo.Text = itemAgronomo.Text
            txtCodigoAgronomo.Value = itemAgronomo.Value

            txtArtInicial.Text = objArt.ARTInicial
            txtArtFinal.Text = objArt.ARTFinal
            txtValidade.Text = objArt.ARTValidade.ToString("dd/MM/yyyy")
            txtCodigo.Enabled = False

            If objArt.Status Then
                radSim.Checked = True
            Else
                radNao.Checked = True
            End If

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtArtInicial.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            IniciarART("I")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            IniciarART("U")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            IniciarART("D")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("AnotacaoRTecnica", "RELATORIO") Then
                Dim Lista As New ListArt(True)
                If Lista.Count > 0 Then
                    Dim i As Integer
                    Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
                    Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                    Dim arquivo As String = NomeArquivo
                    Dim linha As String
                    Dim strm As StreamWriter = Nothing

                    If Dir(arquivo).Length > 0 Then Kill(arquivo)

                    linha = "<html>" & vbCrLf & _
                            "<head>" & vbCrLf & _
                            "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & _
                            "<title>NGS - Anotação de Responsabilidade Técnica</title>" & vbCrLf & _
                            "<style type='text/css'>" & vbCrLf & _
                            "H6 {page-break-after:always}" & vbCrLf & _
                            "A.A1 {font-family: Ms serif;font-size: 5pt;line height: 15px}" & vbCrLf & _
                            "A.A2 {font-family: Ms serif;font-size: 6pt;line height: 15px}" & vbCrLf & _
                            "A.A3 {font-family: Ms serif;font-size: 7pt;line height: 15px}" & vbCrLf & _
                            "A.A4 {font-family: Ms serif;font-size: 8pt;line height: 15px}" & vbCrLf & _
                            "A.A5 {font-family: Ms serif;font-size: 9pt;line height: 15px}" & vbCrLf & _
                            "A.A6 {font-family: Ms serif;font-size: 10pt;line height: 15px}" & vbCrLf & _
                            "A.A7 {font-family: Ms serif bold;font-size: 8pt;line height: 15px}" & vbCrLf & _
                            "A.A8 {font-family: Ms serif bold;font-size: 9pt;line height: 15px}" & vbCrLf & _
                            "A.A9 {font-family: Ms serif bold;font-size: 10pt;line height: 15px}" & vbCrLf & _
                            "A.A10 {font-family: Ms serif bold;font-size: 11pt;line height: 15px}" & vbCrLf & _
                            "A.A11 {font-family: Ms serif bold;font-size: 12pt;line height: 15px}" & vbCrLf & _
                            "A.A12 {font-family: Ms serif bold;font-size: 13pt;line height: 15px}" & vbCrLf & _
                            "A.A13 {font-family: Ms serif bold;font-size: 14pt;line height: 15px}" & vbCrLf & _
                            "</style>" & vbCrLf & _
                            "</head>" & vbCrLf & _
                            "<body text=#000000 bgcolor=#FFFFFF>" & vbCrLf

                    'Cabeçalho
                    linha &= "<table border='0' width='100%' height='57'>" & vbCrLf & _
                             "<tr>" & vbCrLf & _
                             "<td width='13%' rowspan='4' HEIGHT='51'><img border='0' src='../images/" & HttpContext.Current.Session("ssImagemEmpresa") & "' width='150' height='70'></td>" & vbCrLf & _
                             "<td colspan='3'><A Class='A10'><b>" & HttpContext.Current.Session("ssNomeEmpresa") & "</b></A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A6'><b>Página: 1</b></A></td>" & vbCrLf & _
                             "</tr>" & vbCrLf & _
                             "<tr>" & vbCrLf & _
                             "<td colspan='3'><A Class='A8'><b>" & HttpContext.Current.Session("ssCidadeEmpresa") & "/" & HttpContext.Current.Session("ssEstadoEmpresa") & "</b></A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A6'><b>Hora: " & Now().ToString("hh:mm:ss") & "</b></A></td>" & vbCrLf & _
                             "</tr>" & vbCrLf & _
                             "<tr>" & vbCrLf & _
                             "<td colspan='3' color='#800000'><A Class='A6'><b>Anotação de Responsabilidade Técnica</b></A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A6'><b>Data: " & Now().ToString("dd-MM-yyyy") & "</b></A></td>" & vbCrLf & _
                             "</tr>" & vbCrLf & _
                             "</table>" & vbCrLf

                    linha &= "<table border='0' width='100%'>" & vbCrLf & _
                             "<tr>" & vbCrLf & _
                             "<td colspan='8'>" & vbCrLf & _
                             "<hr />" & vbCrLf & _
                             "</td>" & vbCrLf & _
                             "</tr>" & vbCrLf & _
                             "<tr>" & vbCrLf & _
                             "<td><A Class='A5'>Código</A></td>" & vbCrLf & _
                             "<td><A Class='A5'>Empresa</A></td>" & vbCrLf & _
                             "<td><A Class='A5'>Agrônomo</A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A5'>ART Inicial</A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A5'>ART Final</A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A5'>Validade</A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A5'>Ativa</A></td>" & vbCrLf & _
                             "<td align='right'><A Class='A5'>Usuário Inclusão</A></td>" & vbCrLf & _
                             "</tr>" & vbCrLf & _
                             "<tr>" & vbCrLf & _
                             "<td colspan='8'>" & vbCrLf & _
                             "<hr />" & vbCrLf & _
                             "</td>" & vbCrLf & _
                             "</tr>" & vbCrLf

                    For i = 0 To Lista.Count - 1
                        linha &= "<tr>" & vbCrLf & _
                                 "<td><A Class='A5'>" & Lista(i).CodigoArt & "</A></td>" & vbCrLf & _
                                 "<td><A Class='A5'>" & Funcoes.FormatarCpfCnpj(Lista(i).CodigoEmpresa) & "</A></td>" & vbCrLf & _
                                 "<td><A Class='A5'>" & Lista(i).Agronomo.Nome & "</A></td>" & vbCrLf & _
                                 "<td align='right'><A Class='A5'>" & Lista(i).ARTInicial & "</A></td>" & vbCrLf & _
                                 "<td align='right'><A Class='A5'>" & Lista(i).ARTFinal & "</A></td>" & vbCrLf & _
                                 "<td align='right'><A Class='A5'>" & Format(Lista(i).ARTValidade, "dd-MM-yyyy") & "</A></td>" & vbCrLf
                        If Lista(i).Status = True Then
                            linha &= "<td align='right'><A Class='A5'>SIM</A></td>" & vbCrLf
                        Else
                            linha &= "<td align='right'><A Class='A5'>NÃO</A></td>" & vbCrLf
                        End If
                        linha &= "<td align='right'><A Class='A5'>" & Trim(Lista(i).UsuarioInclusao) & "</A></td>" & vbCrLf & _
                                 "</tr>" & vbCrLf
                        i += 1
                    Next

                    linha &= "</table>" & vbCrLf & _
                             "</body>" & vbCrLf & _
                             "</html>" & vbCrLf

                    strm = New StreamWriter(arquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
                Else
                    MsgBox(Me.Page, "Nenhum resultado encontrado.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AnotacaoRTecnica")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class