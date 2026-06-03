Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaCTeXNotas
    Inherits BaseUserControl

    Dim objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Protected Sub grd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

    Protected Sub grd_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles grd.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim strES = e.Row.Cells(6).Text.Trim()
            Dim lnkConsultar As LinkButton = CType(e.Row.FindControl("lnkConsultar"), LinkButton)
            If lnkConsultar IsNot Nothing Then
                lnkConsultar.Visible = strES <> "E"
            End If
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs)
        Dim rowIndex As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
        ConsultarCTe(rowIndex)
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaCTeXNotas")
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        grd.DataSource = New List(Of Object)
        grd.DataBind()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, NfExpress, " & vbCrLf & _
                                "       Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, ISNULL(Protocolo, '') AS Protocolo, RegDPEC " & vbCrLf & _
                                "  FROM NFERealizadas " & vbCrLf & _
                                " WHERE (Empresa_Id = '" & grd.SelectedRow.Cells(2).Text() & "') " & vbCrLf & _
                                "   AND (EndEmpresa_Id = " & grd.SelectedRow.Cells(3).Text() & ") " & vbCrLf & _
                                "   AND (Cliente_Id = '" & grd.SelectedRow.Cells(4).Text() & "') " & vbCrLf & _
                                "   AND (EndCliente_Id = " & grd.SelectedRow.Cells(5).Text() & ") " & vbCrLf & _
                                "   AND (EntradaSaida_Id = '" & grd.SelectedRow.Cells(7).Text() & "') " & vbCrLf & _
                                "   AND (Serie_Id = '" & Server.HtmlDecode(Trim(grd.SelectedRow.Cells(8).Text())) & "') " & vbCrLf & _
                                "   AND (Nota_Id = " & Server.HtmlDecode(grd.SelectedRow.Cells(9).Text()) & ")" & vbCrLf

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "NFERealizadas")

            Dim obj As New [Lib].Negocio.NotaFiscal()
            obj.CodigoEmpresa = Server.HtmlDecode(grd.SelectedRow.Cells(2).Text().Trim())
            obj.EnderecoEmpresa = Server.HtmlDecode(grd.SelectedRow.Cells(3).Text().Trim())
            obj.CodigoCliente = Server.HtmlDecode(grd.SelectedRow.Cells(4).Text().Trim())
            obj.EnderecoCliente = Server.HtmlDecode(grd.SelectedRow.Cells(5).Text().Trim())
            obj.EntradaSaida = IIf(Server.HtmlDecode(grd.SelectedRow.Cells(7).Text().Trim()) = "E", 0, 1)
            obj.Serie = Server.HtmlDecode(grd.SelectedRow.Cells(8).Text().Trim())
            obj.Codigo = Server.HtmlDecode(grd.SelectedRow.Cells(9).Text().Trim())
            obj.Eletronica = False
            obj.ChaveNFE = ""

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 Then
                If ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 AndAlso Not IsDBNull(ds.Tables(0).Rows(0).Item("ChaveNfe")) Then
                    obj.ChaveNFE = ds.Tables(0).Rows(0).Item("ChaveNfe")
                    obj.Eletronica = True
                End If
            End If

            Popup.CloseDialog(Me.Page, "divConsultaCTeXNotas")
            If TypeOf Me.Page Is ConhecimentoDeTransporte Then
                Session("NfConsultaCTRC" & HID.Value) = obj
                CType(Me.Page, ConhecimentoDeTransporte).CarregarNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BindGridView()
        If Session("ssCampo" & HID.Value) = "ConhecimentoDeTransporte" Then
            CargaNotaFiscalCTe()
        End If
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub CargaNotaFiscalCTe()
        SessaoRecuperaNotaFiscal()
        Dim Empresa As String = IIf(String.IsNullOrWhiteSpace(objNotaFiscal.CodigoEmpresa), String.Empty, Funcoes.FormatarCpfCnpj(objNotaFiscal.CodigoEmpresa))
        Dim Cliente As String() = (objNotaFiscal.CodigoCliente & "-" & objNotaFiscal.EnderecoCliente).Split("-")
        Dim DataInicial As String = objNotaFiscal.DataNota
        Dim DataFinal As String = objNotaFiscal.Movimento
        Dim Pedido As String = objNotaFiscal.CodigoPedido
        Dim Operacao As String() = (objNotaFiscal.CodigoOperacao & "-" & objNotaFiscal.CodigoSubOperacao).Split("-")
        Dim EntSai As String = IIf(objNotaFiscal.EntradaSaida = [Lib].Negocio.eEntradaSaida.Entrada, "E", "S")
        Dim Serie As String = objNotaFiscal.Serie
        Dim Nota As String = objNotaFiscal.Codigo
        Dim Sql As String = ""

        Sql = " SELECT DISTINCT NF.Movimento," & vbCrLf &
              "		   NF.Empresa_id," & vbCrLf &
              "		   NF.EndEmpresa_id," & vbCrLf &
              "		   NF.Cliente_Id AS Cliente," & vbCrLf &
              "		   NF.EndCliente_Id AS EndCliente," & vbCrLf &
              "		   Cli.Nome AS ClienteNome," & vbCrLf &
              "		   NF.EntradaSaida_Id AS ES," & vbCrLf &
              "		   NF.Serie_Id AS Serie," & vbCrLf &
              "		   NF.Nota_Id AS Nota," & vbCrLf &
              "		   NF.Operacao," & vbCrLf &
              "		   NF.SubOperacao," & vbCrLf &
              "		   NFxI.QuantidadeFiscal AS Quantidade," & vbCrLf &
              "		   NFxI.Unitario," & vbCrLf &
              "		   NFxI.Valor," & vbCrLf &
              "		   ISNULL(NFEr.Retorno,'') AS Retorno" & vbCrLf &
              "   FROM NotasFiscais NF" & vbCrLf &
              "  INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
              "	    ON NF.Empresa_Id       = NFxI.Empresa_Id" & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
              "   LEFT JOIN NFERealizadas NFEr" & vbCrLf &
              "	    ON NF.Empresa_Id      = NFEr.Empresa_Id" & vbCrLf &
              "    AND NF.EndEmpresa_Id   = NFEr.EndEmpresa_Id" & vbCrLf &
              "    AND NF.Cliente_Id      = NFEr.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id   = NFEr.EndCliente_Id" & vbCrLf &
              "    AND NF.EntradaSaida_Id = NFEr.EntradaSaida_Id" & vbCrLf &
              "    AND NF.Serie_Id        = NFEr.Serie_Id" & vbCrLf &
              "    AND NF.Nota_Id         = NFEr.Nota_Id" & vbCrLf &
              "   JOIN Clientes Cli" & vbCrLf &
              "     ON NF.Cliente_id    = Cli.Cliente_Id" & vbCrLf &
              "    AND NF.EndCliente_Id = Cli.Endereco_Id" & vbCrLf &
              "  WHERE NF.Situacao IN (1,2,4) " & vbCrLf

        If objNotaFiscal.CodigoTipoDeDocumento > 0 Then
            Sql &= "   AND NF.TipoDeDocumento in (" & getTipoDeDocumento() & ")" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Empresa) AndAlso Empresa <> "0" Then
            Sql &= "    AND NF.Empresa_Id       = '" & Empresa.Replace(".", "").Replace("/", "").Replace("-", "") & "' " & vbCrLf & _
                   "    AND NF.EndEmpresa_Id    = 0" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Nota) AndAlso Nota <> "0" Then
            Sql &= "    AND NF.Nota_Id          = '" & Nota & "'" & vbCrLf
        Else
            Sql &= "    AND NF.Movimento BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Serie) Then Sql &= "    AND NF.Serie_Id          = '" & Serie & "'" & vbCrLf

        Sql &= "  ORDER BY NF.Movimento desc, NF.Nota_Id, NF.Serie_Id, NF.EntradaSaida_Id " & vbCrLf
        grd.DataSource = Banco.ConsultaDataSet(Sql, "Notas")
        grd.DataBind()
        grd.SelectedIndex = -1
    End Sub

    Private Function getTipoDeDocumento() As String
        Return CInt(eTipoDeDocumento.CTRC).ToString() & ", " & _
                CInt(eTipoDeDocumento.ComplementoDeFrete).ToString & ", " & _
                CInt(eTipoDeDocumento.Estadia).ToString & ", " & _
                CInt(eTipoDeDocumento.CT_E).ToString & ", " & _
                CInt(eTipoDeDocumento.Anulacao).ToString & ", " & _
                CInt(eTipoDeDocumento.RPA).ToString
    End Function


    Private Function getTextoConsultar(ByVal CTe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVECTE=" & CTe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NRECIBO =" & CTe.ReciboNota & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Sub ConsultarCTe(ByVal rowIndex As Integer)
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True
        Try
            Dim objConhecimento As New [Lib].Negocio.NotaFiscal()
            objConhecimento.CodigoEmpresa = grd.Rows(rowIndex).Cells(2).Text.Trim()
            objConhecimento.EnderecoEmpresa = grd.Rows(rowIndex).Cells(3).Text.Trim()
            objConhecimento.CodigoCliente = grd.Rows(rowIndex).Cells(4).Text.Trim()
            objConhecimento.EnderecoCliente = grd.Rows(rowIndex).Cells(5).Text.Trim()
            objConhecimento.EntradaSaida = IIf(grd.Rows(rowIndex).Cells(7).Text.Trim() = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objConhecimento.Serie = grd.Rows(rowIndex).Cells(8).Text.Trim()
            objConhecimento.Codigo = grd.Rows(rowIndex).Cells(9).Text.Trim()
            objConhecimento = New [Lib].Negocio.NotaFiscal(objConhecimento)

            If Not String.IsNullOrWhiteSpace(objConhecimento.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.ProtocoloNota) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("consultacte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                obj.Texto = getTextoConsultar(objConhecimento)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Format("resp-consultacte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)

                While resp Is Nothing
                    resp = GetResp(fileName)
                End While

                If resp IsNot Nothing Then
                    Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                    Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                    Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
                    Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "100" OrElse strCodigo = "101" OrElse strCodigo = "102") Then
                        Dim strOperacao = String.Empty
                        Dim strSituacao = String.Empty

                        If strCodigo = "100" Then
                            strSituacao = CInt(eSituacao.Normal)
                            strOperacao = "Incluir"
                        ElseIf strCodigo = "101" Then
                            strSituacao = CInt(eSituacao.Cancelado)
                            strOperacao = "Cancelado"
                        ElseIf strCodigo = "102" Then
                            strSituacao = CInt(eSituacao.Inutilizada)
                            strOperacao = "Inutilizar"
                        End If

                        Sqls.Clear()
                        Dim Sql As String = "UPDATE NFERealizadas " & vbCrLf & _
                                            "   SET Operacao = '" & strOperacao & "', Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' " & vbCrLf & _
                                            " WHERE Empresa_Id      = '" & objConhecimento.CodigoEmpresa & "' " & vbCrLf & _
                                            "   AND EndEmpresa_Id   = " & objConhecimento.EnderecoEmpresa & vbCrLf & _
                                            "   AND Cliente_Id      = '" & objConhecimento.CodigoCliente & "'" & vbCrLf & _
                                            "   AND EndCliente_Id   = " & objConhecimento.EnderecoCliente & vbCrLf & _
                                            "   AND EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                                            "   AND Serie_Id        = '" & objConhecimento.Serie & "'" & vbCrLf & _
                                            "   AND Nota_Id         = " & objConhecimento.Codigo & "; " & vbCrLf
                        Sqls.Add(Sql)

                        Sql = "UPDATE NotasFiscais " & vbCrLf & _
                              "   SET Situacao = " & strSituacao & vbCrLf & _
                              " WHERE Empresa_Id      = '" & objConhecimento.CodigoEmpresa & "'" & vbCrLf & _
                              "   AND EndEmpresa_Id   = '" & objConhecimento.EnderecoEmpresa & "'" & vbCrLf & _
                              "   AND Cliente_Id      = '" & objConhecimento.CodigoCliente & "'" & vbCrLf & _
                              "   AND EndCliente_Id   = '" & objConhecimento.EnderecoCliente & "'" & vbCrLf & _
                              "   AND EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                              "   AND Serie_Id        = '" & objConhecimento.Serie & "'" & vbCrLf & _
                              "   AND Nota_Id         = '" & objConhecimento.Codigo & "'; " & vbCrLf
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            Exit Sub
                        End If
                    End If

                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    BindGridView()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

End Class