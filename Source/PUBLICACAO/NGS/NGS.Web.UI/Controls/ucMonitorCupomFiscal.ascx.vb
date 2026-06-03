Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading

Public Class ucMonitorCupomFiscal
    Inherits BaseUserControl

    Dim objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Protected Sub GridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim objNFConsulta As New [Lib].Negocio.NotaFiscal()
        objNFConsulta.CodigoEmpresa = Server.HtmlDecode(GridNotas.SelectedRow.Cells(4).Text().Trim())
        objNFConsulta.EnderecoEmpresa = Server.HtmlDecode(GridNotas.SelectedRow.Cells(5).Text().Trim())
        objNFConsulta.CodigoCliente = Server.HtmlDecode(GridNotas.SelectedRow.Cells(6).Text().Trim())
        objNFConsulta.EnderecoCliente = Server.HtmlDecode(GridNotas.SelectedRow.Cells(7).Text().Trim())
        objNFConsulta.EntradaSaida = IIf(Server.HtmlDecode(GridNotas.SelectedRow.Cells(9).Text().Trim()) = "E", 0, 1)
        objNFConsulta.Serie = Server.HtmlDecode(GridNotas.SelectedRow.Cells(10).Text().Trim())
        objNFConsulta.Codigo = Server.HtmlDecode(GridNotas.SelectedRow.Cells(11).Text().Trim())

        Popup.CloseDialog(Me.Page, "divMonitorCupomFiscal")
        CType(Me.Page, CupomFiscal).CarregarNFCe(objNFConsulta)
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

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        If TypeOf Me.Page Is NotaFiscalXItens Then
            CType(Me.Page, NotaFiscalXItens).VerContingencia()
        End If
        Popup.CloseDialog(Me.Page, "divMonitorCupomFiscal")
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
                    Dim ucEmailNFe As ucEmailNFe = CType(Me.Page.FindControlRecursive("ucEmailNFe"), ucEmailNFe)
                    ucEmailNFe.Limpar()
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
        ConsultarNFCe(rowIndex)
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub InicializarUC(ByVal dMovimento As String)
        hdfMovimento.Value = dMovimento
        CarregarSituacoes()
        CargaNotasXItens()
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub CargaNotasXItens()
        SessaoRecuperaNotaFiscal()
        Dim Sql As String
        Sql = " SELECT DISTINCT  " & vbCrLf &
              "        NF.Movimento, " & vbCrLf &
              "        NF.UsuarioInclusaoData, " & vbCrLf &
              "        NF.Empresa_Id, " & vbCrLf &
              "        NF.EndEmpresa_Id, " & vbCrLf &
              "        NF.Cliente_Id AS Cliente, " & vbCrLf &
              "        NF.EndCliente_Id AS EndCliente, " & vbCrLf &
              "        C.Nome as NomeCliente, " & vbCrLf &
              "        NF.EntradaSaida_Id AS ES, " & vbCrLf &
              "        NF.Serie_Id AS Serie, " & vbCrLf &
              "        NF.Nota_Id AS Nota, " & vbCrLf &
              "        NF.Operacao, " & vbCrLf &
              "        NF.SubOperacao, " & vbCrLf &
              "	       NF.NossaEmissao," & vbCrLf &
              "        SUM(NFxI.QuantidadeFiscal) AS Quantidade, " & vbCrLf &
              "        SUM(NFxI.Unitario) AS Unitario, " & vbCrLf &
              "        SUM(NFxI.Valor) AS Valor, " & vbCrLf &
              "	       ISNULL(NFEr.Retorno,'') AS Retorno, " & vbCrLf &
              "	       ISNULL(NFEr.MsgRetorno,'') AS MsgRetorno " & vbCrLf &
              "   FROM NotasFiscais NF " & vbCrLf &
              "  INNER JOIN NotasFiscaisxItens NFxI " & vbCrLf &
              "     ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
              "   LEFT JOIN " & IIf(ddlSituacao.SelectedValue = "1", "NFEPendencias", "NFERealizadas") & " NFEr" & vbCrLf &
              "	    ON NF.Empresa_Id       = NFEr.Empresa_Id" & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFEr.EndEmpresa_Id" & vbCrLf &
              "    AND NF.Cliente_Id      = NFEr.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id   = NFEr.EndCliente_Id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFEr.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NFEr.Serie_Id" & vbCrLf &
              "    AND NF.Nota_Id         = NFEr.Nota_Id" & vbCrLf &
              "  INNER JOIN Clientes C " & vbCrLf &
              "     ON C.Cliente_Id  = NF.Cliente_Id " & vbCrLf &
              "    AND C.Endereco_Id = NF.EndCliente_Id " & vbCrLf &
              "  WHERE NF.TipoDeDocumento in(5,65)" & vbCrLf &
              "    AND NF.Eletronica   = 'S'" & vbCrLf &
              "    AND NF.NossaEmissao = 'S'" & vbCrLf &
              "    AND NF.Movimento    = '" & hdfMovimento.Value & "'" & vbCrLf

        If ddlSituacao.SelectedValue = "1" Then
            Sql &= "    AND NF.Situacao in(4,7) " & vbCrLf
        Else
            Sql &= "    AND NF.Situacao <> 4 " & vbCrLf
        End If

        Sql &= "    AND NF.NossaEmissao = 'S' " & vbCrLf

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoEmpresa) Then
            Sql &= "    AND NF.Empresa_Id       = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                   "    AND NF.EndEmpresa_Id    = 0" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            Sql &= "    AND NF.Cliente_Id       ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                   "    AND NF.EndCliente_Id    = " & objNotaFiscal.EnderecoCliente & vbCrLf
        End If

        'If Not String.IsNullOrWhiteSpace(objNotaFiscal.Serie) Then Sql &= "    AND NF.Serie_Id  = '" & objNotaFiscal.Serie & "'" & vbCrLf

        'If objNotaFiscal.Codigo > 0 Then
        '    Sql &= "    AND NF.Nota_Id = " & objNotaFiscal.Codigo & vbCrLf
        'Else
        '    If ddlSituacao.SelectedValue = "2" Then Sql &= "    AND (NF.Movimento BETWEEN '" & objNotaFiscal.DataNota.ToString("yyyy-MM-dd") & "' AND '" & objNotaFiscal.Movimento.ToString("yyyy-MM-dd") & "')" & vbCrLf
        'End If

        Sql &= "  GROUP BY NF.Movimento, " & vbCrLf &
               "        NF.UsuarioInclusaoData, " & vbCrLf &
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
               "        NFEr.Retorno, " & vbCrLf &
               "        NFEr.MsgRetorno " & vbCrLf &
               "  ORDER BY NF.UsuarioInclusaoData DESC, ES, Serie, Nota" & vbCrLf

        lnkEnviarEmail.Parent.Visible = Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) AndAlso ddlSituacao.SelectedValue = "2"
        lnkImprimir.Parent.Visible = Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) AndAlso ddlSituacao.SelectedValue = "2"
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
        ddlSituacao.Items.Add(New ListItem() With {
         .Value = "1",
         .Text = "NFC-e pendências"
        })
        ddlSituacao.Items.Add(New ListItem() With {
         .Value = "2",
         .Text = "NFC-e realizadas"
        })
        ddlSituacao.SelectedIndex = 0
    End Sub

    Public Overrides Sub Carregar(obj As [Lib].Negocio.IBaseEntity)
        If Session("objEmailNFe" & HID.Value) IsNot Nothing Then
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                Try
                    Dim Sqls As New ArrayList
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

                    Dim strMsg As String = String.Empty

                    For Each nf As [Lib].Negocio.NotaFiscal In lstNotas
                        Sqls.Clear()

                        Dim lst As [Lib].Negocio.ListCliente = CType(Session("objEmailNFe" & HID.Value), [Lib].Negocio.ListCliente)
                        Dim objFil As New [Lib].Negocio.Fil()
                        objFil.IUD = "I"
                        objFil.NomeArquivo = String.Format("emailnfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                        objFil.Texto = getTextoEmail(nf, lst, Session("strAssunto" & HID.Value), Session("strMensagem" & HID.Value))
                        objFil.SalvarSql(Sqls)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If

                        Dim resp As [Lib].Negocio.Resp = Nothing
                        Dim fileName As String = String.Format("resp-emailnfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

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
                    Next

                    MsgBox(Me.Page, strMsg)

                    For Each row As GridViewRow In GridNotas.Rows
                        Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
                        If chk IsNot Nothing AndAlso chk.Checked Then chk.Checked = False
                    Next

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
        sb.Append("CHAVENFCE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("DESTINATARIO=" & nf.Cliente.EmailNFE & ControlChars.CrLf)
        sb.Append("ASSUNTO=" & "NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & "Envio NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("EMAILEMITENTE=" & nf.Empresa.EmailNFE & ControlChars.CrLf)
        sb.Append("NOMEEMITENTE=" & nf.Empresa.Nome & ControlChars.CrLf)
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOQRCODE=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & "NAO" & ControlChars.CrLf)
        sb.Append("IDTOKEN=" & "000001929d25db8c95e991866f1c493c4af50bb741" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function getTextoEmail(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal lstClientes As [Lib].Negocio.ListCliente, ByVal strAssunto As String, ByVal strMensagem As String) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFCE=" & nf.ChaveNFE & ControlChars.CrLf)
        For Each objCliente As [Lib].Negocio.Cliente In lstClientes.Where(Function(s) Not String.IsNullOrWhiteSpace(s.EmailNFE)).ToList()
            sb.Append("DESTINATARIO=" & objCliente.EmailNFE & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & strAssunto & " - NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & strMensagem & " - Envio NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("EMAILEMITENTE=" & nf.Empresa.EmailNFE & ControlChars.CrLf)
        sb.Append("NOMEEMITENTE=" & nf.Empresa.Nome & ControlChars.CrLf)
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOQRCODE=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & "NAO" & ControlChars.CrLf)
        sb.Append("IDTOKEN=" & "000001929d25db8c95e991866f1c493c4af50bb741" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

    'Private Function getTextoConsultar(ByVal nfe As [Lib].Negocio.NotaFiscal) As String
    '    Dim sb As New StringBuilder()
    '    sb.Append("CHAVENFE=" & nfe.ChaveNFE & ControlChars.CrLf)
    '    sb.Append("NRECIBO =" & nfe.ReciboNota & ControlChars.CrLf)
    '    Return sb.ToString()
    'End Function

    Private Function getTextoConsultarNFCe(ByVal nfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFCE=" & nfe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NRECIBO =" & nfe.ReciboNota & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Sub ConsultarNFCe(ByVal rowIndex As Integer)
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

            Sql = "SELECT ISNULL(ChaveNFE,'') AS ChaveNFE, ISNULL(Protocolo, '') AS Protocolo, DadosAdicionais " & vbCrLf &
                  "  FROM NFEPendencias " & vbCrLf &
                  " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "'" & vbCrLf &
                  "   AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & vbCrLf &
                  "   AND Cliente_Id = '" & objNFe.CodigoCliente & "'" & vbCrLf &
                  "   AND EndCliente_Id = " & objNFe.EnderecoCliente & vbCrLf &
                  "   AND EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                  "   AND Serie_Id = '" & objNFe.Serie & "'" & vbCrLf &
                  "   AND Nota_Id = " & objNFe.Codigo & vbCrLf
            Sqls.Add(Sql)

            Dim dsPendencia As New DataSet
            dsPendencia = Banco.ConsultaDataSet(Sql, "Consulta")

            Sqls.Clear()
            Sql = String.Empty

            If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                objNFe.ChaveNFE = dsPendencia.Tables(0).Rows(0).Item("ChaveNfe")
                objNFe.ProtocoloNota = dsPendencia.Tables(0).Rows(0).Item("Protocolo")
            End If

            If Not String.IsNullOrWhiteSpace(objNFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objNFe.ProtocoloNota) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("consultanfce{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa)
                obj.Texto = getTextoConsultarNFCe(objNFe)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Format("resp-consultanfce{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa)

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

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "100" OrElse strCodigo = "101" OrElse strCodigo = "110" OrElse strCodigo = "302") _
                        And strChave.Length = 44 _
                        And strProtocolo.Length = 15 Then
                        'And strRecibo.Length = 15 _

                        Dim strOperacao = String.Empty
                        Dim strSituacao = String.Empty
                        Sql = String.Empty
                        Sqls.Clear()

                        If strCodigo = "100" OrElse strCodigo = "110" OrElse strCodigo = "302" Then
                            strOperacao = "Incluir"
                            If strCodigo = "110" Or strCodigo = "302" Then
                                strSituacao = CInt(eSituacao.Denegado)
                            Else
                                strSituacao = CInt(eSituacao.Normal)
                            End If

                            If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                                Sql = " DELETE FROM  NFEPendencias " & vbCrLf &
                                      "  WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' " & vbCrLf &
                                      "    AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & vbCrLf &
                                      "    AND Cliente_Id = '" & objNFe.CodigoCliente & "'" & vbCrLf &
                                      "    AND EndCliente_Id = " & objNFe.EnderecoCliente & vbCrLf &
                                      "    AND EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                                      "    AND Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                                Sqls.Add(Sql)

                                If Not dsPendencia.Tables(0).Rows(0).Item("DadosAdicionais").ToString.Contains("CANCELAR") Then
                                    Sql = "INSERT INTO NFERealizadas " & vbCrLf &
                                          "       (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, " & vbCrLf &
                                          "        NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) " & vbCrLf &
                                          "VALUES ('" & objNFe.CodigoEmpresa & "', " & objNFe.EnderecoEmpresa & ", '" & objNFe.CodigoCliente & "'," & vbCrLf &
                                          "'" & objNFe.EnderecoCliente & "', '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objNFe.Serie & "', '" & objNFe.Codigo & "', " & vbCrLf &
                                          "'" & objNFe.DataInclusao.ToString("yyyy-MM-dd") & "', '" & objNFe.DataInclusao.ToString("HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', " & vbCrLf &
                                          "'" & String.Format("nota{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "', 'INCLUIR'," & vbCrLf &
                                          "'" & strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '1', '" & objNFe.Observacoes & "', '', '" & strProtocolo & "', ''); " & vbCrLf
                                    Sqls.Add(Sql)
                                End If

                                If strCodigo = "100" Then
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
                                            msgNFE = "Não foi possível encontrar o arquivo DANFCE da nota fiscal " & objNFe.Codigo & "-" & objNFe.Serie & "!"
                                        End Try
                                    End If
                                End If
                            Else
                                Sql = "DELETE FROM NFEPendencias " & vbCrLf &
                                      " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND " & vbCrLf &
                                      "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND " & vbCrLf &
                                      "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf &
                                      "       Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                                Sqls.Add(Sql)

                                Sql = "UPDATE NFERealizadas " & vbCrLf &
                                      "   SET Operacao = '" & strOperacao & "', Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' " & vbCrLf &
                                      " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND " & vbCrLf &
                                      "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND " & vbCrLf &
                                      "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf &
                                      "       Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                                Sqls.Add(Sql)
                            End If

                            Sql = "UPDATE NotasFiscais " & vbCrLf &
                                  "   SET Situacao = " & strSituacao & " " & vbCrLf &
                                  " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = '" & objNFe.EnderecoEmpresa & "' AND " & vbCrLf &
                                  "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = '" & objNFe.EnderecoCliente & "' AND " & vbCrLf &
                                  "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND Serie_Id = '" & objNFe.Serie & "' AND " & vbCrLf &
                                  "       Nota_Id = '" & objNFe.Codigo & "'; " & vbCrLf
                            Sqls.Add(Sql)

                            'Excluir contabilização se a nota for Denegada
                            If strCodigo = "110" OrElse strCodigo = "301" OrElse strCodigo = "302" OrElse strCodigo = "303" Then
                                objNFe.Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)

                                If objNFe.VencimentosNota IsNot Nothing AndAlso objNFe.VencimentosNota.Count > 0 Then
                                    For Each tit In objNFe.VencimentosNota
                                        If tit.CodigoSituacao = eSituacao.Normal AndAlso tit.CodigoProvisao = eProvisao.Previsao Then
                                            Sql = "Update ContasAPagar Set Provisao = 3 " & vbCrLf &
                                                    "Where Registro_id = " & tit.Codigo
                                            Sqls.Add(Sql)

                                            Sql = "Update ContasAReceber Set Provisao = 3 " & vbCrLf &
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

                                Sql = "Delete NotasFiscaisXTransportadores" & vbCrLf &
                                      " Where Empresa_Id      ='" & objNFe.CodigoEmpresa & "'" & vbCrLf &
                                      "   and EndEmpresa_Id   = " & objNFe.EnderecoEmpresa & vbCrLf &
                                      "   and Cliente_Id      ='" & objNFe.CodigoCliente & "'" & vbCrLf &
                                      "   and EndCliente_Id   = " & objNFe.EnderecoCliente & vbCrLf &
                                      "   and EntradaSaida_Id ='" & objNFe.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                      "   and Serie_Id        ='" & objNFe.Serie & "'" & vbCrLf &
                                      "   and Nota_Id         = " & objNFe.Codigo
                                Sqls.Add(Sql)


                                Sql = " Delete NotasFiscaisXRomaneios" & vbCrLf &
                                      "  Where Empresa_Id      ='" & objNFe.CodigoEmpresa & "'" & vbCrLf &
                                      "    and EndEmpresa_Id   = " & objNFe.EnderecoEmpresa & vbCrLf &
                                      "    and Cliente_Id      ='" & objNFe.CodigoCliente & "'" & vbCrLf &
                                      "    and EndCliente_Id   = " & objNFe.EnderecoCliente & vbCrLf &
                                      "    and EntradaSaida_Id ='" & objNFe.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                      "    and Serie_Id        ='" & objNFe.Serie & "'" & vbCrLf &
                                      "    and Nota_Id         = " & objNFe.Codigo & vbCrLf &
                                      "    and Romaneio_Id     = " & objNFe.CodigoRomaneio
                                Sqls.Add(Sql)

                            End If

                        ElseIf strCodigo = "101" Then
                            CType(Me.Page, NotaFiscalXItens).CancelarNotaFiscal(Sqls, objNFe)
                        Else
                            Sql = "UPDATE NFEPendencias " & vbCrLf &
                                  "   SET Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' " & vbCrLf &
                                  " WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND " & vbCrLf &
                                  "       Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND " & vbCrLf &
                                  "       EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf &
                                  "       Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; " & vbCrLf
                            Sqls.Add(Sql)
                        End If

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-consultanfce{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailnfce{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfenfce{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
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

                    InicializarUC(hdfMovimento.Value)
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
                Dim msgErro = String.Empty
                If Not String.IsNullOrWhiteSpace(nf.ChaveNFE) Then
                    [Lib].Negocio.DocumentoEletronico.ImprimirNFCe(nf, 1, msgErro)
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
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfenfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)
                End If
            Next

            If Sqls.Count > 0 Then
                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
            End If

            MsgBox(Me.Page, "Cupom Fiscal impresso com sucesso.", eTitulo.Sucess)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class