Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading

Public Class ucMonitorDeNotas
    Inherits BaseUserControl

    Dim objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Protected Sub GridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, NfExpress, Operacao, Retorno, MsgRetorno, Recibo, " & vbCrLf &
              "       ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, ISNULL(Protocolo, '') AS Protocolo, RegDPEC " & vbCrLf &
              " FROM NFERealizadas " & vbCrLf &
              " WHERE Empresa_Id      ='" & GridNotas.SelectedRow.Cells(4).Text() & "'" & vbCrLf &
              "   AND EndEmpresa_Id   = " & GridNotas.SelectedRow.Cells(5).Text() & vbCrLf &
              "   AND Cliente_Id      ='" & GridNotas.SelectedRow.Cells(6).Text() & "'" & vbCrLf &
              "   AND EndCliente_Id   = " & GridNotas.SelectedRow.Cells(7).Text() & vbCrLf &
              "   AND EntradaSaida_Id ='" & GridNotas.SelectedRow.Cells(9).Text() & "'" & vbCrLf &
              "   AND Serie_Id        ='" & GridNotas.SelectedRow.Cells(10).Text() & "'" & vbCrLf &
              "   AND Nota_Id         = " & GridNotas.SelectedRow.Cells(11).Text() & vbCrLf &
              "   ORDER BY Data DESC, Empresa_Id DESC, Cliente_Id, Nota_Id DESC"

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "NFERealizadas")

        If ds Is Nothing Then
            HttpContext.Current.Session("ssNFE") = "N"
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            HttpContext.Current.Session("ssNFE") = "N"
        Else
            HttpContext.Current.Session("ssNFE") = "S"
            HttpContext.Current.Session("ssChaveNFE") = ds.Tables(0).Rows(0).Item("ChaveNfe")
            HttpContext.Current.Session("ssProtocolo") = ds.Tables(0).Rows(0).Item("Protocolo")
        End If

        Dim objNFConsulta As New [Lib].Negocio.NotaFiscal()
        objNFConsulta.CodigoEmpresa = Server.HtmlDecode(GridNotas.SelectedRow.Cells(4).Text().Trim())
        objNFConsulta.EnderecoEmpresa = Server.HtmlDecode(GridNotas.SelectedRow.Cells(5).Text().Trim())
        objNFConsulta.CodigoCliente = Server.HtmlDecode(GridNotas.SelectedRow.Cells(6).Text().Trim())
        objNFConsulta.EnderecoCliente = Server.HtmlDecode(GridNotas.SelectedRow.Cells(7).Text().Trim())
        objNFConsulta.EntradaSaida = IIf(Server.HtmlDecode(GridNotas.SelectedRow.Cells(9).Text().Trim()) = "E", 0, 1)
        objNFConsulta.Serie = Server.HtmlDecode(GridNotas.SelectedRow.Cells(10).Text().Trim())
        objNFConsulta.Codigo = Server.HtmlDecode(GridNotas.SelectedRow.Cells(11).Text().Trim())
        objNFConsulta.ChaveNFE = String.Empty
        objNFConsulta.Eletronica = False

        If HttpContext.Current.Session("ssNFE") = "S" Then
            If Not IsDBNull(ds.Tables(0).Rows(0).Item("ChaveNfe")) Then
                objNFConsulta.ChaveNFE = ds.Tables(0).Rows(0).Item("ChaveNfe")
                objNFConsulta.Eletronica = True
            End If
        End If

        Popup.CloseDialog(Me.Page, "divMonitorDeNotas")
        If TypeOf Me.Page Is NotaFiscalXItens Then
            Session("objNFConsultaNXI" & HID.Value) = objNFConsulta
            Session("objNotaFiscal" & HID.Value) = objNFConsulta
            CType(Me.Page, NotaFiscalXItens).CarregarConsulta()
        End If
    End Sub

    Protected Sub GridNotas_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridNotas.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim chkSelecionado As CheckBox = CType(e.Row.FindControl("chkSelecionado"), CheckBox)
            If chkSelecionado IsNot Nothing Then
                chkSelecionado.Visible = ddlSituacao.SelectedValue <> "1"
            End If

            Dim lnkConsultar As LinkButton = CType(e.Row.FindControl("lnkConsultar"), LinkButton)
            Dim str As String() = CType(e.Row.FindControl("imgStatus"), Image).ToolTip.ToString.Split("-")
            Dim hdfNossaEmissao As HiddenField = CType(e.Row.FindControl("hdfNossaEmissao"), HiddenField)

            If Trim(str(0)) = "100" Then
                chkSelecionado.Enabled = True
            Else
                chkSelecionado.Enabled = False
            End If

            If hdfNossaEmissao.Value = "S" Then
                If Not String.IsNullOrWhiteSpace(Trim(str(0))) AndAlso (Trim(str(0)) = "100" Or Trim(str(0)) = "101") Then
                    If Session("ssNomeUsuario") = "FURLAN" Then
                        lnkConsultar.Visible = True
                    Else
                        lnkConsultar.Visible = False
                    End If
                Else
                    lnkConsultar.Visible = True
                End If
            Else
                lnkConsultar.Visible = False
            End If
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            Dim ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)
            If ucConsultaClientes IsNot Nothing Then
                ucConsultaClientes.Limpar()
                ucConsultaClientes.MainUserControl = Me
                Popup.ConsultaDeClientes(Me.Page, "objClienteMonitor" & HID.Value.ToString, "txtCliente")
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgBuscarMovimento_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgBuscarMovimento.Click

        If CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final", eTitulo.Info)
            Exit Sub
        End If

        CargaNotasXItens()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        If TypeOf Me.Page Is NotaFiscalXItens Then
            CType(Me.Page, NotaFiscalXItens).VerContingencia()
        End If
        Popup.CloseDialog(Me.Page, "divMonitorDeNotas")
    End Sub

    Protected Sub ddlSituacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSituacao.SelectedIndexChanged
        CargaNotasXItens()
    End Sub

    Protected Sub lnkEnviarEmail_Click(sender As Object, e As EventArgs) Handles lnkEnviarEmail.Click
        SessaoRecuperaNotaFiscal()

        If String.IsNullOrWhiteSpace(objNotaFiscal.Empresa.EmailNFE) Then
            MsgBox(Me.Page, "Email da Empresa " & objNotaFiscal.Empresa.CodigoFormatado & "-" & objNotaFiscal.Empresa.Nome & " não está cadastrado, favor verificar.")
        Else
            Dim itemSelecionado As Boolean = False

            Dim i As Integer = 0
            While i < GridNotas.Rows.Count
                If CType(GridNotas.Rows(i).FindControl("chkSelecionado"), CheckBox).Checked Then itemSelecionado = True
                i += 1
            End While

            If itemSelecionado Then
                Dim fm As New FilesManager()
                If fm.IsConnect() Then

                    Dim primeira As Boolean = True
                    Dim strMensagem As String = "ANEXO NOTAS FISCAIS: "
                    For Each row As GridViewRow In GridNotas.Rows
                        Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                        If chk IsNot Nothing AndAlso chk.Checked Then
                            If primeira Then
                                strMensagem += row.Cells(11).Text
                                primeira = False
                            Else
                                strMensagem += ", " & row.Cells(11).Text
                            End If
                        End If
                    Next

                    Dim parameters As New Dictionary(Of String, Object)
                    parameters("Assunto") = "XML NOTAS FISCAIS"
                    parameters("Mensagem") = strMensagem
                    parameters("Compactado") = "NAO"

                    Dim ucEmailNFe As ucEmailNFe = CType(Me.Page.FindControlRecursive("ucEmailNFe"), ucEmailNFe)
                    ucEmailNFe.Limpar()
                    ucEmailNFe.EnvioDeEmail(parameters)
                    ucEmailNFe.MainUserControl = Me
                    Dim txtDestinatario As TextBox = CType(ucEmailNFe.FindControlRecursive("txtDestinatario"), TextBox)
                    Popup.ConsultaDeEmailNFe(Me.Page, "objEmailNFe" & HID.Value, txtDestinatario.ClientID, 100)
                Else
                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                End If
            Else
                MsgBox(Me.Page, "Não foi selecionado nenhum item no Grid para Envio!")
            End If
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs)
        Dim rowIndex As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
        ConsultarNFe(rowIndex)
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub BindGridView()
        If Session("ssCampo" & HID.Value) = "NotaXItens" Then

            LimparTela()

            SessaoRecuperaNotaFiscal()

            txtDataInicial.Text = objNotaFiscal.DataNota.ToString("dd/MM/yyyy")
            txtDataFinal.Text = objNotaFiscal.Movimento.ToString("dd/MM/yyyy")

            CarregarSituacoes()

            CargaNotasXItens()
        End If
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub
    Protected Sub chkConsolidarEmpresaMonitor_CheckedChanged(sender As Object, e As EventArgs) Handles chkConsolidarEmpresaMonitor.CheckedChanged
        CargaNotasXItens()
    End Sub

    Protected Sub ddlEmpresaMonitor_CheckedChanged(sender As Object, e As EventArgs)
        CargaNotasXItens()
    End Sub

    Private Sub CargaNotasXItens()

        SessaoRecuperaNotaFiscal()

        ddlEmpresaMonitor.Enabled = False
        ddlEmpresaMonitor.Visible = False
        chkConsolidarEmpresaMonitor.Enabled = False
        chkConsolidarEmpresaMonitor.Visible = False
        divConsolidarEmpresaMonitor.Visible = False

        If ddlSituacao.SelectedValue = 2 Then
            divCliente.Visible = True
            divMovimento.Visible = True
        Else
            divCliente.Visible = False
            divMovimento.Visible = False
        End If

        If ddlEmpresaMonitor.SelectedIndex < 1 Then
            ddl.Carregar(ddlEmpresaMonitor, CarregarDDL.Tabela.EmpresasMonitor, "", True)
        End If

        Dim Sql As String

        If ddlSituacao.SelectedValue = "3" Then
            ddlEmpresaMonitor.Enabled = True
            ddlEmpresaMonitor.Visible = True
            divConsolidarEmpresaMonitor.Visible = True
            chkConsolidarEmpresaMonitor.Enabled = True
            chkConsolidarEmpresaMonitor.Visible = True

            Sql = " SELECT ISNULL(dXML.Emissao, '') AS Movimento, " & vbCrLf &
                  "     ISNULL(dXML.Empresa_Id, '') AS Empresa_Id, " & vbCrLf &
                  "     0 AS EndEmpresa_Id, " & vbCrLf &
                  "     ISNULL(dXML.Cliente_Id, '') AS Cliente, " & vbCrLf &
                  "     0 AS EndCliente, " & vbCrLf &
                  "     ISNULL(ClienteNome, '') AS NomeCliente, " & vbCrLf &
                  "     '' ES, " & vbCrLf &
                  "     ISNULL(dXML.Serie_Id, '') AS Serie, " & vbCrLf &
                  "     ISNULL(dXML.Numero_Id, '') AS Nota, " & vbCrLf &
                  "     0 AS Operacao, " & vbCrLf &
                  "     0 AS SubOperacao, " & vbCrLf &
                  "     ISNULL(dXML.Emissao, '') AS NossaEmissao, " & vbCrLf &
                  "     0 AS Quantidade, " & vbCrLf &
                  "     0 AS Unitario, " & vbCrLf &
                  "     0 AS Valor, " & vbCrLf &
                  "     '' AS Romaneio, " & vbCrLf &
                  "     ISNULL(dXML.MsgImportacaoXML, '') AS Retorno, " & vbCrLf &
                  "     ISNULL(dXML.MsgImportacaoXML, '') AS MsgRetorno " & vbCrLf &
                  " FROM DocumentoXML dXML " & vbCrLf &
                  " WHERE 1=1 " & vbCrLf &
                  " AND dXML.Situacao           = 103 "
            If chkConsolidarEmpresaMonitor.Checked Then
            Else
                Sql &= " AND dXML.Empresa_Id = '" & IIf(ddlEmpresaMonitor.SelectedIndex > 0, ddlEmpresaMonitor.SelectedValue.Split("-")(0), objNotaFiscal.CodigoEmpresa) & "'"
            End If
            Sql &= " ORDER BY Movimento DESC, Empresa_Id DESC, Cliente, Serie DESC, Nota DESC"
        Else

            Sql = " Select DISTINCT  " & vbCrLf &
              "        NF.Movimento, " & vbCrLf &
              "        NF.Empresa_Id, " & vbCrLf &
              "        NF.EndEmpresa_Id, " & vbCrLf &
              "        NF.Cliente_Id As Cliente, " & vbCrLf &
              "        NF.EndCliente_Id As EndCliente, " & vbCrLf &
              "        C.Nome As NomeCliente, " & vbCrLf &
              "        NF.EntradaSaida_Id As ES, " & vbCrLf &
              "        NF.Serie_Id As Serie, " & vbCrLf &
              "        NF.Nota_Id As Nota, " & vbCrLf &
              "        NF.Operacao, " & vbCrLf &
              "        NF.SubOperacao, " & vbCrLf &
              "	       NF.NossaEmissao," & vbCrLf &
              "        SUM(NFxI.QuantidadeFiscal) As Quantidade, " & vbCrLf &
              "        SUM(NFxI.Unitario) As Unitario, " & vbCrLf &
              "        SUM(NFxI.Valor) As Valor, " & vbCrLf &
              "        NFxR.Romaneio_Id As Romaneio, " & vbCrLf &
              "	       ISNULL(NFEr.Retorno,'') AS Retorno, " & vbCrLf &
              "	       ISNULL(NFEr.MsgRetorno,'') AS MsgRetorno " & vbCrLf &
              "   FROM NotasFiscais NF " & vbCrLf &
              "  INNER JOIN NotasFiscaisxItens NFxI " & vbCrLf &
              "     ON NF.Empresa_Id            = NFxI.Empresa_Id " & vbCrLf &
              "    AND NF.EndEmpresa_Id         = NFxI.EndEmpresa_Id " & vbCrLf &
              "    AND NF.Cliente_Id            = NFxI.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id         = NFxI.EndCliente_Id " & vbCrLf &
              "    AND NF.EntradaSaida_Id       = NFxI.EntradaSaida_Id " & vbCrLf &
              "    AND NF.Serie_Id              = NFxI.Serie_Id " & vbCrLf &
              "    AND NF.Nota_Id               = NFxI.Nota_Id " & vbCrLf

            If ddlSituacao.SelectedValue = "4" Then
                Sql &= "INNER JOIN DocumentoXML dXML " & vbCrLf &
                       "     ON NF.Empresa_Id            = dXML.Empresa_Id " & vbCrLf &
                       "     AND NF.Cliente_Id           = dXML.Cliente_Id " & vbCrLf &
                       "     AND NF.Serie_Id             = dXML.Serie_Id " & vbCrLf &
                       "     AND NF.Nota_Id              = dXML.Numero_Id" & vbCrLf
            End If

            Sql &= "LEFT JOIN " & IIf(ddlSituacao.SelectedValue = "1", "NFEPendencias", "NFERealizadas") & " NFEr" & vbCrLf &
              "	    ON NF.Empresa_Id            = NFEr.Empresa_Id" & vbCrLf &
              "     AND NF.EndEmpresa_Id        = NFEr.EndEmpresa_Id" & vbCrLf &
              "     AND NF.Cliente_Id           = NFEr.Cliente_Id" & vbCrLf &
              "     AND NF.EndCliente_Id        = NFEr.EndCliente_Id" & vbCrLf &
              "     AND NF.EntradaSaida_Id      = NFEr.EntradaSaida_Id" & vbCrLf &
              "     AND NF.Serie_Id             = NFEr.Serie_Id" & vbCrLf &
              "     AND NF.Nota_Id              = NFEr.Nota_Id" & vbCrLf &
              "   LEFT OUTER JOIN NotasFiscaisXRomaneios NFxR " & vbCrLf &
              "     ON NFxI.Empresa_Id          = NFxR.Empresa_Id " & vbCrLf &
              "     AND NFxI.EndEmpresa_Id      = NFxR.EndEmpresa_Id " & vbCrLf &
              "     AND NFxI.Cliente_Id         = NFxR.Cliente_Id " & vbCrLf &
              "     AND NFxI.EndCliente_Id      = NFxR.EndCliente_Id " & vbCrLf &
              "     AND NFxI.EntradaSaida_Id    = NFxR.EntradaSaida_Id " & vbCrLf &
              "     AND NFxI.Serie_Id           = NFxR.Serie_Id " & vbCrLf &
              "     AND NFxI.Nota_Id            = NFxR.Nota_Id " & vbCrLf &
              "  INNER JOIN Clientes C " & vbCrLf &
              "     ON C.Cliente_Id             = NF.Cliente_Id " & vbCrLf &
              "     AND C.Endereco_Id           = NF.EndCliente_Id " & vbCrLf &
              "  WHERE 1=1 " & vbCrLf &
              "     AND NF.TipoDeDocumento      = 1" & vbCrLf

            If ddlSituacao.SelectedValue = "1" Then
                Sql &= "    AND NF.Situacao in(4,7) " & vbCrLf
                Sql &= "    AND NF.NossaEmissao = 'S' " & vbCrLf
            ElseIf ddlSituacao.SelectedValue = "4" Then
                ddlEmpresaMonitor.Enabled = True
                ddlEmpresaMonitor.Visible = True
                divConsolidarEmpresaMonitor.Visible = True
                chkConsolidarEmpresaMonitor.Enabled = True
                chkConsolidarEmpresaMonitor.Visible = True

                Sql &= "    AND dXML.Situacao   = 104 " & vbCrLf
                Sql &= "    AND NF.Situacao     = 1 " & vbCrLf
                Sql &= "    AND NF.NossaEmissao = 'N' " & vbCrLf
            Else
                Sql &= "    AND NF.Situacao <> 4 " & vbCrLf
                Sql &= "    AND NF.NossaEmissao = 'S' " & vbCrLf
            End If

            If chkConsolidarEmpresaMonitor.Checked Then
            Else
                If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoEmpresa) Then
                    Sql &= "    AND NF.Empresa_Id       = '" & IIf(ddlEmpresaMonitor.SelectedIndex > 0, ddlEmpresaMonitor.SelectedValue.Split("-")(0), objNotaFiscal.CodigoEmpresa) & "'" & vbCrLf &
                           "    AND NF.EndEmpresa_Id    = 0" & vbCrLf
                End If
            End If

            If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
                Sql &= "    AND NF.Cliente_Id       ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                       "    AND NF.EndCliente_Id    = " & objNotaFiscal.EnderecoCliente & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(objNotaFiscal.Serie) Then Sql &= "    AND NF.Serie_Id  = '" & objNotaFiscal.Serie & "'" & vbCrLf

            If objNotaFiscal.Codigo > 0 Then
                Sql &= "    AND NF.Nota_Id = " & objNotaFiscal.Codigo & vbCrLf
            Else
                If ddlSituacao.SelectedValue = "2" Then
                    Sql &= "    AND (NF.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "')" & vbCrLf

                    If txtCodigoCliente.Value.ToString.Length > 0 Then
                        Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                        Sql &= "    AND NF.Cliente_Id = '" & strCliente(0) & "' And NF.EndCliente_Id = " & strCliente(1) & vbCrLf
                    End If

                End If
            End If

            If objNotaFiscal.CodigoPedido > 0 Then Sql &= "    And NF.Pedido = " & objNotaFiscal.CodigoPedido & vbCrLf

            If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.CodigoSubOperacao > 0 Then
                Sql &= "    AND NF.Operacao         = " & objNotaFiscal.CodigoOperacao & vbCrLf &
                       "    AND NF.SubOperacao      = " & objNotaFiscal.CodigoSubOperacao & vbCrLf
            End If

            Sql &= "  GROUP BY NF.Movimento, " & vbCrLf &
                   "        NF.Empresa_Id, " & vbCrLf &
                   "        NF.EndEmpresa_Id, " & vbCrLf &
                   "        NF.Cliente_Id, " & vbCrLf &
                   "        NF.EndCliente_Id, " & vbCrLf &
                   "        C.Nome, " & vbCrLf &
                   "        NF.EntradaSaida_Id, " & vbCrLf &
                   "        NF.Serie_Id, " & vbCrLf &
                   "        NF.Nota_Id, " & vbCrLf &
                   "        NF.Operacao, " & vbCrLf &
                   "        NF.SubOperacao, " & vbCrLf &
                   "        NF.NossaEmissao, " & vbCrLf &
                   "        NFxR.Romaneio_Id, " & vbCrLf &
                   "        NFEr.Retorno, " & vbCrLf &
                   "        NFEr.MsgRetorno " & vbCrLf &
                   "  ORDER BY NF.Movimento DESC, ES DESC, C.Nome, Serie DESC, Nota DESC" & vbCrLf

        End If

        lnkEnviarEmail.Parent.Visible = Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) AndAlso (ddlSituacao.SelectedValue = "2" Or ddlSituacao.SelectedValue = "4")
        lnkImprimir.Parent.Visible = Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) AndAlso (ddlSituacao.SelectedValue = "2" Or ddlSituacao.SelectedValue = "4")
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Notas")
        GridNotas.DataSource = ds
        GridNotas.DataBind()
    End Sub

    Protected Function GetImagem(ByVal obj As Object) As String
        If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CType(obj, String)) AndAlso CType(obj, String) <> "100" AndAlso CType(obj, String) <> "101" Then
            Return ResolveUrl("~/Images/important.png")
        ElseIf obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CType(obj, String)) AndAlso CType(obj, String) = "100" Then
            Return ResolveUrl("~/Images/ledgreen.png")
        ElseIf obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(CType(obj, String)) AndAlso CType(obj, String) = "101" Then
            Return ResolveUrl("~/Images/ledred.png")
        End If
        Return ResolveUrl("~/Images/lednone.png")
    End Function

    Protected Function GetTooltip(ByVal obj As Object) As String
        If Not String.IsNullOrWhiteSpace(CType(obj, String)) Then
            Dim str As String() = CType(obj, String).Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
            If str IsNot Nothing AndAlso str.Length > 0 Then
                If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(str(0)) Then
                    Return CType(obj, String)
                End If
            End If
        End If
        Return String.Empty
    End Function

    Private Sub CarregarSituacoes()
        ddlSituacao.Items.Clear()
        ddlSituacao.Items.Add(New ListItem() With { _
         .Value = "1", _
         .Text = "NF-e pendências" _
        })
        ddlSituacao.Items.Add(New ListItem() With {
         .Value = "2",
         .Text = "NF-e realizadas"
        })
        ddlSituacao.Items.Add(New ListItem() With {
         .Value = "3",
         .Text = "NF-e imp. MIC pendências"
        })
        ddlSituacao.Items.Add(New ListItem() With {
         .Value = "4",
         .Text = "NF-e imp. MIC realizadas"
        })
        ddlSituacao.SelectedIndex = 0
    End Sub

    Public Overrides Sub Carregar(obj As [Lib].Negocio.IBaseEntity)

        If Session("objClienteMonitor" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteMonitor" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteMonitor" & HID.Value)

            CargaNotasXItens()

        ElseIf Session("objEmailNFe" & HID.Value) IsNot Nothing Then
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                Try
                    Dim Sqls As New ArrayList
                    Dim lstNotas As New List(Of [Lib].Negocio.NotaFiscal)
                    Dim codigoEmpresa As String = String.Empty
                    Dim codigoEnderecoEmpresa As Integer

                    For Each row As GridViewRow In GridNotas.Rows
                        Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                        If chk IsNot Nothing AndAlso chk.Checked Then
                            Dim hdf As HiddenField = CType(row.FindControl("hdfKey"), HiddenField)
                            If hdf IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(hdf.Value) Then
                                Dim keys As String() = hdf.Value.Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
                                Dim nf As New [Lib].Negocio.NotaFiscal()
                                nf.CodigoEmpresa = keys(0)
                                nf.EnderecoEmpresa = keys(1)
                                nf.CodigoCliente = keys(2)
                                nf.EnderecoCliente = keys(3)
                                nf.EntradaSaida = IIf(keys(4) = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                                nf.Codigo = keys(5)
                                nf.Serie = keys(6)
                                nf = New [Lib].Negocio.NotaFiscal(nf)
                                lstNotas.Add(nf)

                                codigoEmpresa = nf.CodigoEmpresa
                                codigoEnderecoEmpresa = nf.EnderecoEmpresa
                            End If
                        End If
                    Next

                    Sqls.Clear()

                    Dim strMsg As String = String.Empty

                    Dim randomNumber As String = FuncoesStrings.GerarNumeroSeguro(9)

                    Dim lst As [Lib].Negocio.ListCliente = CType(Session("objEmailNFe" & HID.Value), [Lib].Negocio.ListCliente)
                    Dim objFil As New [Lib].Negocio.Fil()
                    objFil.IUD = "I"
                    objFil.NomeArquivo = String.Format("email{0:000000000}#{1}.txt", randomNumber, codigoEmpresa)

                    objFil.Texto = getTextoEmail(codigoEmpresa, codigoEnderecoEmpresa, lstNotas, lst, Session("strAssunto" & HID.Value), Session("strMensagem" & HID.Value), Session("strCompactado" & HID.Value))
                    objFil.SalvarSql(Sqls)

                    If Not Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Exit Sub
                    End If

                    Dim resp As [Lib].Negocio.Resp = Nothing
                    Dim fileName As String = String.Format("resp-email{0:000000000}#{1}.txt", randomNumber, codigoEmpresa)

                    While resp Is Nothing
                        resp = GetResp(fileName)
                    End While

                    If resp IsNot Nothing Then
                        Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                        Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                        Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                        Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                        strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                        If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4014" Then
                            strMsg = String.Format("{0} - {1}", strCodigo, strMsg)
                            Exit Sub
                        End If
                    End If

                    For Each row As GridViewRow In GridNotas.Rows
                        Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                        If chk IsNot Nothing AndAlso chk.Checked Then chk.Checked = False
                    Next

                    MsgBox(Me.Page, strMsg)

                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
            End If
        End If
    End Sub

    Private Function getTextoEmail(ByVal nf As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("DESTINATARIO=" & nf.Cliente.EmailNFE & ControlChars.CrLf)
        sb.Append("ASSUNTO=" & "NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & "Envio NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("EMAILEMITENTE=" & nf.Empresa.EmailNFE & ControlChars.CrLf)
        sb.Append("NOMEEMITENTE=" & nf.Empresa.Nome & ControlChars.CrLf)
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & "NAO" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function getTextoEmail(ByVal codEmpresa As String, ByVal codEndEmpresa As Integer, ByVal lstNotas As List(Of [Lib].Negocio.NotaFiscal), ByVal lstClientes As [Lib].Negocio.ListCliente, ByVal strAssunto As String, ByVal strMensagem As String, ByVal strCompactado As Boolean) As String
        Dim sb As New StringBuilder()

        Dim objEmpresa As New [Lib].Negocio.Cliente(codEmpresa, codEndEmpresa)

        For Each nf As [Lib].Negocio.NotaFiscal In lstNotas
            sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        Next

        For Each objCliente As [Lib].Negocio.Cliente In lstClientes.Where(Function(s) Not String.IsNullOrWhiteSpace(s.EmailNFE)).ToList()
            sb.Append("DESTINATARIO=" & objCliente.EmailNFE & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & strAssunto & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & strMensagem & ControlChars.CrLf)
        sb.Append("EMAILEMITENTE=" & objEmpresa.EmailNFE & ControlChars.CrLf)
        sb.Append("NOMEEMITENTE=" & objEmpresa.Nome & ControlChars.CrLf)
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)

        If strCompactado Then
            sb.Append("COMPACTADO=" & "SIM" & ControlChars.CrLf)
        Else
            sb.Append("COMPACTADO=" & "NAO" & ControlChars.CrLf)
        End If

        Return sb.ToString()
    End Function

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

    Private Function getTextoConsultar(ByVal nfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nfe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NRECIBO =" & nfe.ReciboNota & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Sub ConsultarNFe(ByVal rowIndex As Integer)
        Dim Sqls As New ArrayList
        Dim Sql As String = String.Empty
        Dim aux As Boolean = True
        Dim msgNFE As String = String.Empty
        Try
            Dim objNFe As New [Lib].Negocio.NotaFiscal()
            objNFe.CodigoEmpresa = Server.HtmlDecode(GridNotas.Rows(rowIndex).Cells(4).Text().Trim())
            objNFe.EnderecoEmpresa = Server.HtmlDecode(GridNotas.Rows(rowIndex).Cells(5).Text().Trim())
            objNFe.CodigoCliente = Server.HtmlDecode(GridNotas.Rows(rowIndex).Cells(6).Text().Trim())
            objNFe.EnderecoCliente = Server.HtmlDecode(GridNotas.Rows(rowIndex).Cells(7).Text().Trim())
            objNFe.EntradaSaida = IIf(Server.HtmlDecode(GridNotas.Rows(rowIndex).Cells(9).Text().Trim()) = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objNFe.Serie = Server.HtmlDecode(GridNotas.Rows(rowIndex).Cells(10).Text().Trim())
            objNFe.Codigo = Server.HtmlDecode(GridNotas.Rows(rowIndex).Cells(11).Text().Trim())
            objNFe = New [Lib].Negocio.NotaFiscal(objNFe)

            Sql = "SELECT ISNULL(ChaveNFE,'') AS ChaveNFE, Retorno, ISNULL(Protocolo, '') AS Protocolo, DadosAdicionais " & vbCrLf &
                  "  FROM NFEPendencias " & vbCrLf &
                  " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "'" & vbCrLf &
                  "   AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & vbCrLf &
                  "   AND Cliente_Id = '" & objNFe.CodigoCliente & "'" & vbCrLf &
                  "   AND EndCliente_Id = " & objNFe.EnderecoCliente & vbCrLf &
                  "   AND EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                  "   AND Serie_Id = '" & objNFe.Serie & "'" & vbCrLf &
                  "   AND Nota_Id = " & objNFe.Codigo & vbCrLf &
                  "   ORDER BY Data DESC, Empresa_Id DESC, Cliente_Id, Nota_Id DESC"
            Sqls.Add(Sql)

            Dim dsPendencia As New DataSet
            dsPendencia = Banco.ConsultaDataSet(Sql, "Consulta")

            Sqls.Clear()
            Sql = String.Empty

            Dim codRetorno As String = String.Empty

            If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                objNFe.ChaveNFE = dsPendencia.Tables(0).Rows(0).Item("ChaveNfe")

                If Not dsPendencia.Tables(0).Rows(0).Item("Retorno") Is Nothing AndAlso Not IsDBNull(dsPendencia.Tables(0).Rows(0).Item("Retorno")) Then
                    codRetorno = dsPendencia.Tables(0).Rows(0).Item("Retorno")
                End If

                If Not dsPendencia.Tables(0).Rows(0).Item("Protocolo") Is Nothing Then
                    objNFe.ProtocoloNota = dsPendencia.Tables(0).Rows(0).Item("Protocolo")
                End If
            End If

            If codRetorno = "124" OrElse (Not String.IsNullOrWhiteSpace(objNFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objNFe.ProtocoloNota)) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("consulta{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa)
                obj.Texto = getTextoConsultar(objNFe)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Format("resp-consulta{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa)

                Dim tempoLimite As DateTime
                tempoLimite = Now.AddSeconds(30)

                While resp Is Nothing AndAlso Now < tempoLimite
                    resp = GetResp(fileName)
                    System.Threading.Thread.Sleep(3000)
                End While

                If resp IsNot Nothing Then
                    Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                    Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                    Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
                    Dim chave As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CHAVE")).FirstOrDefault()
                    Dim recibo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RECIBO")).FirstOrDefault()
                    Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()
                    Dim lote As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NUMEROLOTE")).FirstOrDefault()

                    Dim strCodigo As String = String.Empty
                    If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                        strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strMsg As String = String.Empty
                    If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                        strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strChave As String = String.Empty
                    If chave IsNot Nothing AndAlso chave.Length > 0 Then
                        strChave = chave.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strRecibo As String = String.Empty
                    If recibo IsNot Nothing AndAlso recibo.Length > 0 Then
                        strRecibo = recibo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strProtocolo As String = String.Empty
                    If protocolo IsNot Nothing AndAlso protocolo.Length > 0 Then
                        strProtocolo = protocolo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strLote As String = String.Empty
                    If lote IsNot Nothing AndAlso lote.Length > 0 Then
                        strLote = lote.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "100" OrElse strCodigo = "101" OrElse strCodigo = "110" OrElse strCodigo = "124" OrElse strCodigo = "302") _
                        And strChave.Length = 44 _
                        And strProtocolo.Length = 15 Then
                        'And strRecibo.Length = 15 _

                        Dim strOperacao = String.Empty
                        Dim strSituacao = String.Empty
                        Sql = String.Empty
                        Sqls.Clear()

                        If strCodigo = "100" OrElse strCodigo = "110" OrElse strCodigo = "124" OrElse strCodigo = "302" Then
                            strOperacao = "Incluir"
                            If strCodigo = "110" Or strCodigo = "302" Then
                                strSituacao = CInt(eSituacao.Denegado)
                            Else
                                strSituacao = CInt(eSituacao.Normal)
                            End If

                            If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                                Sql = " DELETE FROM  NFEPendencias " & vbCrLf & _
                                      "  WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' " & vbCrLf & _
                                      "    AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & vbCrLf & _
                                      "    AND Cliente_Id = '" & objNFe.CodigoCliente & "'" & vbCrLf & _
                                      "    AND EndCliente_Id = " & objNFe.EnderecoCliente & vbCrLf & _
                                      "    AND EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                                      "    AND Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                                Sqls.Add(Sql)

                                If Not dsPendencia.Tables(0).Rows(0).Item("DadosAdicionais").ToString.Contains("CANCELAR") Then
                                    Sql = "INSERT INTO NFERealizadas " & vbCrLf & _
                                          "       (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, " & vbCrLf & _
                                          "        NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) " & vbCrLf & _
                                          "VALUES ('" & objNFe.CodigoEmpresa & "', " & objNFe.EnderecoEmpresa & ", '" & objNFe.CodigoCliente & "'," & vbCrLf & _
                                          "'" & objNFe.EnderecoCliente & "', '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objNFe.Serie & "', '" & objNFe.Codigo & "', " & vbCrLf & _
                                          "'" & objNFe.DataInclusao.ToString("yyyy-MM-dd") & "', '" & objNFe.DataInclusao.ToString("HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', " & vbCrLf & _
                                          "'" & String.Format("nota{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "', 'INCLUIR'," & vbCrLf & _
                                          "'" & strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '1', '" & objNFe.Observacoes & "', '', '" & strProtocolo & "', ''); " & vbCrLf
                                    Sqls.Add(Sql)
                                End If

                                If strCodigo = "100" Or strCodigo = "124" Then
                                    If Not String.IsNullOrWhiteSpace(objNFe.Empresa.EmailNFE) AndAlso Not String.IsNullOrWhiteSpace(objNFe.Cliente.EmailNFE) Then
                                        [Lib].Negocio.DocumentoEletronico.SendMailNFe(objNFe, "", False)
                                    End If

                                    If [Lib].Negocio.DocumentoEletronico.ImprimirNFe(objNFe, objNFe.Empresa.Empresa.ViasNFE, msgNFE, True) Then
                                        Try
                                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNFe.ChaveNFE), eTipoDeDocumento.Nota)
                                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNFe.ChaveNFE) Then
                                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNFe.ChaveNFE))
                                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNFe.ChaveNFE))
                                            End If
                                        Catch ex As Exception
                                            msgNFE = "Não foi possível encontrar o arquivo DANFE da nota fiscal " & objNFe.Codigo & "-" & objNFe.Serie & "!"
                                        End Try
                                    End If
                                End If
                            Else
                                Sql = "DELETE FROM NFEPendencias " & vbCrLf & _
                                      " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND " & vbCrLf & _
                                      "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND " & vbCrLf & _
                                      "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf & _
                                      "       Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                                Sqls.Add(Sql)

                                Sql = "UPDATE NFERealizadas " & vbCrLf & _
                                      "   SET Operacao = '" & strOperacao & "', Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' " & vbCrLf & _
                                      " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND " & vbCrLf & _
                                      "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND " & vbCrLf & _
                                      "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf & _
                                      "       Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                                Sqls.Add(Sql)
                            End If

                            Sql = "UPDATE NotasFiscais " & vbCrLf & _
                                  "   SET Situacao = " & strSituacao & " " & vbCrLf & _
                                  " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = '" & objNFe.EnderecoEmpresa & "' AND " & vbCrLf & _
                                  "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = '" & objNFe.EnderecoCliente & "' AND " & vbCrLf & _
                                  "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND Serie_Id = '" & objNFe.Serie & "' AND " & vbCrLf & _
                                  "       Nota_Id = '" & objNFe.Codigo & "'; " & vbCrLf
                            Sqls.Add(Sql)

                            'Excluir contabilização se a nota for Denegada
                            If strCodigo = "110" OrElse strCodigo = "301" OrElse strCodigo = "302" OrElse strCodigo = "303" Then
                                objNFe.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)

                                If objNFe.VencimentosNota IsNot Nothing AndAlso objNFe.VencimentosNota.Count > 0 Then
                                    For Each tit In objNFe.VencimentosNota
                                        If tit.CodigoSituacao = eSituacao.Normal AndAlso tit.CodigoProvisao = eProvisao.Previsao Then
                                            Sql = "Update ContasAPagar Set Provisao = 3 " & vbCrLf & _
                                                    "Where Registro_id = " & tit.Codigo
                                            Sqls.Add(Sql)

                                            Sql = "Update ContasAReceber Set Provisao = 3 " & vbCrLf & _
                                                    "Where Registro_id = " & tit.Codigo
                                            Sqls.Add(Sql)
                                        End If
                                    Next
                                End If

                                If objNFe.NotasReferenciais IsNot Nothing AndAlso objNFe.NotasReferenciais.Count > 0 Then
                                    For Each item In objNFe.NotasReferenciais
                                        item.IUD = "D"
                                        item.SalvarSql(Sqls)
                                    Next
                                End If

                                Sql = "Delete NotasFiscaisXTransportadores" & vbCrLf & _
                                      " Where Empresa_Id      ='" & objNFe.CodigoEmpresa & "'" & vbCrLf & _
                                      "   and EndEmpresa_Id   = " & objNFe.EnderecoEmpresa & vbCrLf & _
                                      "   and Cliente_Id      ='" & objNFe.CodigoCliente & "'" & vbCrLf & _
                                      "   and EndCliente_Id   = " & objNFe.EnderecoCliente & vbCrLf & _
                                      "   and EntradaSaida_Id ='" & objNFe.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                      "   and Serie_Id        ='" & objNFe.Serie & "'" & vbCrLf & _
                                      "   and Nota_Id         = " & objNFe.Codigo
                                Sqls.Add(Sql)


                                Sql = " Delete NotasFiscaisXRomaneios" & vbCrLf & _
                                      "  Where Empresa_Id      ='" & objNFe.CodigoEmpresa & "'" & vbCrLf & _
                                      "    and EndEmpresa_Id   = " & objNFe.EnderecoEmpresa & vbCrLf & _
                                      "    and Cliente_Id      ='" & objNFe.CodigoCliente & "'" & vbCrLf & _
                                      "    and EndCliente_Id   = " & objNFe.EnderecoCliente & vbCrLf & _
                                      "    and EntradaSaida_Id ='" & objNFe.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                      "    and Serie_Id        ='" & objNFe.Serie & "'" & vbCrLf & _
                                      "    and Nota_Id         = " & objNFe.Codigo & vbCrLf & _
                                      "    and Romaneio_Id     = " & objNFe.CodigoRomaneio
                                Sqls.Add(Sql)

                            End If

                        ElseIf strCodigo = "101" Then
                            CType(Me.Page, NotaFiscalXItens).CancelarNotaFiscal(Sqls, objNFe)
                        Else
                            Sql = "UPDATE NFEPendencias " & vbCrLf & _
                                  "   SET Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' " & vbCrLf & _
                                  " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND " & vbCrLf & _
                                  "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND " & vbCrLf & _
                                  "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf & _
                                  "       Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                            Sqls.Add(Sql)
                        End If

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-consulta{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            Exit Sub
                        ElseIf Not String.IsNullOrEmpty(msgNFE) Then
                            MsgBox(Me.Page, msgNFE)
                            Exit Sub
                        End If
                    End If

                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    BindGridView()
                Else
                    MsgBox(Me.Page, "Sefaz não retornou nenhuma resposta, consulte novamente.")
                End If
            Else
                MsgBox(Me.Page, "Esta nota não têm chave e ou protocolo. Selecione ela e a reenvie por favor.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkImprimir_Click(sender As Object, e As EventArgs) Handles lnkImprimir.Click
        Try
            Dim Sqls As New ArrayList
            Dim Sql As String = ""
            Dim lstNotas As New List(Of [Lib].Negocio.NotaFiscal)

            For Each row As GridViewRow In GridNotas.Rows
                Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                If chk IsNot Nothing AndAlso chk.Checked Then
                    Dim hdf As HiddenField = CType(row.FindControl("hdfKey"), HiddenField)
                    If hdf IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(hdf.Value) Then
                        Dim keys As String() = hdf.Value.Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
                        Dim nf As New [Lib].Negocio.NotaFiscal()
                        nf.CodigoEmpresa = keys(0)
                        nf.EnderecoEmpresa = keys(1)
                        nf.CodigoCliente = keys(2)
                        nf.EnderecoCliente = keys(3)
                        nf.EntradaSaida = IIf(keys(4) = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                        nf.Codigo = keys(5)
                        nf.Serie = keys(6)
                        nf = New [Lib].Negocio.NotaFiscal(nf)
                        lstNotas.Add(nf)
                    End If
                End If
            Next

            For Each nf As [Lib].Negocio.NotaFiscal In lstNotas
                If Not String.IsNullOrWhiteSpace(nf.ChaveNFE) Then
                    Dim obj As New [Lib].Negocio.Fil()
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("danfe{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = [Lib].Negocio.DocumentoEletronico.getPrinterNFE(nf, "I")
                    obj.SalvarSql(Sqls)
                End If
            Next

            If Sqls.Count > 0 Then
                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
            End If

            Sqls.Clear()

            Thread.Sleep(5000)

            For Each nf As [Lib].Negocio.NotaFiscal In lstNotas
                If Not String.IsNullOrWhiteSpace(nf.ChaveNFE) Then
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)
                End If
            Next

            If Sqls.Count > 0 Then
                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
            End If

            MsgBox(Me.Page, "Danfe(s) impressa(s) com Sucesso.", eTitulo.Sucess)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        LimparTela()

        CargaNotasXItens()
    End Sub

    Private Sub LimparTela()
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty

        txtDataInicial.Text = Today.ToString("dd/MM/yyyy")
        txtDataFinal.Text = Today.ToString("dd/MM/yyyy")
    End Sub

End Class