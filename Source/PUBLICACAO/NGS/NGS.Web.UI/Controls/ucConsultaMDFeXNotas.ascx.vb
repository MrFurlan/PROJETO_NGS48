Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaMDFeXNotas
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

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs)
        Dim rowIndex As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
        ConsultarMDFe(rowIndex)
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaMDFeXNotas")
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim sql As String = "SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, NfExpress, "
            sql &= "Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, ISNULL(Protocolo, '') AS Protocolo, RegDPEC "
            sql &= "FROM NFERealizadas "
            sql &= "WHERE (Empresa_Id = '" & grd.SelectedRow.Cells(3).Text() & "') "
            sql &= "AND (EndEmpresa_Id = " & grd.SelectedRow.Cells(4).Text() & ") "
            sql &= "AND (Cliente_Id = '" & grd.SelectedRow.Cells(5).Text() & "') "
            sql &= "AND (EndCliente_Id = " & grd.SelectedRow.Cells(6).Text() & ") "
            sql &= "AND (EntradaSaida_Id = '" & grd.SelectedRow.Cells(7).Text() & "') "
            sql &= "AND (Serie_Id = '" & grd.SelectedRow.Cells(8).Text() & "') "
            sql &= "AND (Nota_Id = " & grd.SelectedRow.Cells(9).Text() & ")"

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "NFERealizadas")

            Dim obj As New [Lib].Negocio.NotaFiscal()
            obj.CodigoEmpresa = Server.HtmlDecode(grd.SelectedRow.Cells(3).Text().Trim())
            obj.EnderecoEmpresa = Server.HtmlDecode(grd.SelectedRow.Cells(4).Text().Trim())
            obj.CodigoCliente = Server.HtmlDecode(grd.SelectedRow.Cells(5).Text().Trim())
            obj.EnderecoCliente = Server.HtmlDecode(grd.SelectedRow.Cells(6).Text().Trim())
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

            Popup.CloseDialog(Me.Page, "divConsultaMDFeXNotas")
            If TypeOf Me.Page Is MDFe Then
                Session("NfConsultaMDFe" & HID.Value) = obj
                CType(Me.Page, MDFe).CarregarNotaFiscal()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BindGridView()
        If Session("ssCampo" & HID.Value) = "MDFe" Then
            CargaNotaFiscalMDFe()
        End If
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub CargaNotaFiscalMDFe()
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

        Sql = " SELECT DISTINCT NF.Movimento," & vbCrLf & _
              "		NF.Empresa_id," & vbCrLf & _
              "		NF.EndEmpresa_id," & vbCrLf & _
              "		NF.Cliente_Id AS Cliente," & vbCrLf & _
              "		NF.EndCliente_Id AS EndCliente," & vbCrLf & _
              "		NF.EntradaSaida_Id AS ES," & vbCrLf & _
              "		NF.Serie_Id AS Serie," & vbCrLf & _
              "		NF.Nota_Id AS Nota," & vbCrLf & _
              "		NF.Operacao," & vbCrLf & _
              "		NF.SubOperacao," & vbCrLf & _
              "		NFxI.QuantidadeFiscal AS Quantidade," & vbCrLf & _
              "		NFxI.Unitario," & vbCrLf & _
              "		NFxI.Valor," & vbCrLf & _
              "		NFxR.Romaneio_Id AS Romaneio," & vbCrLf & _
              "		isnull(NFEr.Retorno,'') AS Retorno" & vbCrLf & _
              " FROM NotasFiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf & _
              "	  ON NF.Empresa_Id       = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " LEFT JOIN NFERealizadas NFEr" & vbCrLf & _
              "	  ON NF.Empresa_Id       = NFEr.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFEr.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFEr.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFEr.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFEr.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFEr.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFEr.Nota_Id" & vbCrLf & _
              "  LEFT JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf & _
              "	  ON NFxI.Empresa_Id       = NFxR.Empresa_Id" & vbCrLf & _
              "   AND NFxI.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf & _
              "   AND NFxI.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf & _
              "   AND NFxI.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf & _
              "   AND NFxI.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf & _
              "   AND NFxI.Serie_Id        = NFxR.Serie_Id" & vbCrLf & _
              "   AND NFxI.Nota_Id         = NFxR.Nota_Id" & vbCrLf & _
              " WHERE 1=1 AND NF.Situacao in (1,4) AND NF.Movimento BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "'" & vbCrLf

        If objNotaFiscal.CodigoTipoDeDocumento > 0 Then
            Sql &= "   AND NF.TipoDeDocumento in (" & CInt(eTipoDeDocumento.ManifestoEletronico).ToString() & ")"
        End If

        If Not String.IsNullOrWhiteSpace(Empresa) AndAlso Empresa <> "0" Then
            Sql &= "   AND NF.Empresa_Id       = '" & Empresa.Replace(".", "").Replace("/", "").Replace("-", "") & "' " & vbCrLf & _
                   "   AND NF.EndEmpresa_Id    = 0" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Nota) AndAlso Nota <> "0" Then Sql &= "   AND NF.Nota_Id          = '" & Nota & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(Serie) Then Sql &= "   AND NF.Serie_Id          = '" & Serie & "'" & vbCrLf

        Sql &= " ORDER BY NF.Movimento, ES, Serie, Nota" & vbCrLf
        grd.DataSource = Banco.ConsultaDataSet(Sql, "Notas")
        grd.DataBind()
    End Sub

    Protected Function GetImagem(ByVal obj As Object) As String
        If obj IsNot Nothing AndAlso CType(obj, String) = "135" Then
            Return ResolveUrl("~/Images/ledgreen.png")
        End If
        Return ResolveUrl("~/Images/ledred.png")
    End Function

    Protected Function GetTooltip(ByVal obj As Object) As String
        If obj IsNot Nothing AndAlso CType(obj, String) = "135" Then
            Return "Encerrado"
        End If
        Return "Não Encerrado"
    End Function

    Private Function getTextoConsultar(ByVal mdfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVEMDFE=" & mdfe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NRECIBO =" & mdfe.ReciboNota & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Sub ConsultarMDFe(ByVal rowIndex As Integer)
        Dim Sqls As New ArrayList
        Dim Sql As String = String.Empty
        Dim aux As Boolean = True
        Try
            Dim objMDFe As New [Lib].Negocio.NotaFiscal()
            objMDFe.CodigoEmpresa = grd.Rows(rowIndex).Cells(3).Text.Trim()
            objMDFe.EnderecoEmpresa = grd.Rows(rowIndex).Cells(4).Text.Trim()
            objMDFe.CodigoCliente = grd.Rows(rowIndex).Cells(5).Text.Trim()
            objMDFe.EnderecoCliente = grd.Rows(rowIndex).Cells(6).Text.Trim()
            objMDFe.EntradaSaida = IIf(grd.Rows(rowIndex).Cells(7).Text.Trim() = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objMDFe.Serie = grd.Rows(rowIndex).Cells(8).Text.Trim()
            objMDFe.Codigo = grd.Rows(rowIndex).Cells(9).Text.Trim()
            objMDFe = New [Lib].Negocio.NotaFiscal(objMDFe)

            Sql = "Select * from  NFEPendencias "
            Sql &= "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & " AND "
            Sql &= "  Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente & " AND "
            Sql &= "  EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND "
            Sql &= "  Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo & "; "
            Sqls.Add(Sql)

            Dim dsPendencia As New DataSet
            dsPendencia = Banco.ConsultaDataSet(Sql, "Consulta")

            Sqls.Clear()
            Sql = String.Empty

            If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                objMDFe.ChaveNFE = dsPendencia.Tables(0).Rows(0).Item("ChaveNfe")
                objMDFe.ProtocoloNota = dsPendencia.Tables(0).Rows(0).Item("Protocolo")
            End If

            If Not String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.ProtocoloNota) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("consultamdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)
                obj.Texto = getTextoConsultar(objMDFe)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    Exit Sub
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Format("resp-consultamdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)

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

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "100" OrElse strCodigo = "101" OrElse strCodigo = "132") Then
                        Dim strOperacao = String.Empty
                        Dim strSituacao = String.Empty
                        Sqls.Clear()
                        Sql = String.Empty

                        If strCodigo = "100" Then
                            strSituacao = CInt(eSituacao.Normal)
                            strOperacao = "Incluir"
                        ElseIf strCodigo = "101" Then
                            strSituacao = CInt(eSituacao.Cancelado)
                            strOperacao = "Cancelado"
                        ElseIf strCodigo = "102" Then
                            strSituacao = CInt(eSituacao.Inutilizada)
                            strOperacao = "Inutilizar"
                        ElseIf strCodigo = "132" Then
                            strSituacao = CInt(eSituacao.Normal)
                            strCodigo = "135"
                            strOperacao = "INCLUIR"
                            strMsg = "Evento registrado e vinculado ao MDF-e"
                        End If

                        objMDFe.ChaveNFE = strChave
                        objMDFe.ProtocoloNota = strProtocolo
                        objMDFe.ReciboNota = strRecibo

                        If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                            Sql = "Delete from  NFEPendencias "
                            Sql &= "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & " AND "
                            Sql &= "  Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente & " AND "
                            Sql &= "  EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND "
                            Sql &= "  Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo & "; "
                            Sqls.Add(Sql)

                            Sql = "INSERT INTO NFERealizadas "
                            Sql &= "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, "
                            Sql &= "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) "
                            Sql &= "VALUES ('" & objMDFe.CodigoEmpresa & "', " & objMDFe.EnderecoEmpresa & ", '" & objMDFe.CodigoCliente & "', '"
                            Sql &= objMDFe.EnderecoCliente & "', '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objMDFe.Serie & "', '" & objMDFe.Codigo & "', '"
                            Sql &= objMDFe.DataInclusao.ToSqlDate() & "', '" & Format(objMDFe.DataInclusao, "HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', '"
                            Sql &= String.Format("mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "', 'INCLUIR', '"
                            Sql &= strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '1', '" & objMDFe.Observacoes & "', '', '" & strProtocolo & "', ''); "
                            Sqls.Add(Sql)
                        Else
                            Sql = "UPDATE NFERealizadas SET Operacao = '" & strOperacao & "', Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' "
                            Sql &= "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & " AND "
                            Sql &= "      Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente & " AND "
                            Sql &= "      EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND "
                            Sql &= "      Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo & "; "
                            Sqls.Add(Sql)
                        End If

                        Sql = "UPDATE NotasFiscais SET Situacao = " & strSituacao & " "
                        Sql &= "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = '" & objMDFe.EnderecoEmpresa & "' AND "
                        Sql &= "      Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = '" & objMDFe.EnderecoCliente & "' AND "
                        Sql &= "      EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND Serie_Id = '" & objMDFe.Serie & "' AND "
                        Sql &= "      Nota_Id = '" & objMDFe.Codigo & "'; "
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                            Exit Sub
                        End If

                        If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                            If CType(Me.Page, MDFe).SendMailMDFe(objMDFe) Then
                                CType(Me.Page, MDFe).ImprimirMDFe(objMDFe)
                            End If
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