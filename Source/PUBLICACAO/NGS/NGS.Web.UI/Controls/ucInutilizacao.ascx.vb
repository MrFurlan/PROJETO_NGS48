Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucInutilizacao
    Inherits BaseUserControl

    Dim Sql As String

    Private Property objEmpresa As [Lib].Negocio.Cliente
        Get
            Return CType(Session("_objEmpresa" & HID.Value), [Lib].Negocio.Cliente)
        End Get
        Set(ByVal value As [Lib].Negocio.Cliente)
            Session("_objEmpresa" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Private Sub SelecionarEmpresa(ByVal Empresa As String, ByVal EndEmpresa As String)
        objEmpresa = New [Lib].Negocio.Cliente(Empresa, EndEmpresa)
        With objEmpresa
            txtEmpresa.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
        End With
    End Sub

    Public Sub CarregarInutilizacao(ByVal nf As [Lib].Negocio.NotaFiscal)
        Limpar()
        hdfTipoDeDocumento.Value = nf.CodigoTipoDeDocumento
        SelecionarEmpresa(nf.CodigoEmpresa, nf.EnderecoEmpresa)

        If nf IsNot Nothing AndAlso nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Nota) Then
            If Left(nf.Empresa.Codigo, 8) = "24450490" Then
                txtSerie.Text = "10"
            Else
                txtSerie.Text = "1"
            End If
            txtJustificativa.Text = "ERRO NA SEQUÊNCIA CRONOLÓGICA DA NOTA FiSCAL"
        ElseIf nf IsNot Nothing AndAlso (nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse nf.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E)) Then
            txtSerie.Text = "0"
            txtSerie.Enabled = False
            txtJustificativa.Text = "ERRO NA SEQUÊNCIA CRONOLÓGICA DO CONHECIMENTO DE TRANSPORTE"
        Else
            txtSerie.Text = ""
            txtSerie.Enabled = True
            txtJustificativa.Text = "ERRO NA SEQUÊNCIA CRONOLÓGICA"
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaNUNXI" & HID.Value) Is Nothing Then
            SelecionarEmpresa(CType(Session("objEmpresaNUNXI" & HID.Value), [Lib].Negocio.Cliente).Codigo, CType(Session("objEmpresaNUNXI" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco)
            Session.Remove("objEmpresaNUNXI" & HID.Value)
        ElseIf Not Session("objLiberaInutilizacao" & HID.Value) Is Nothing Then
            hdLiberar.Value = CType(Session("objLiberaInutilizacao" & HID.Value), Boolean)
        End If
    End Sub

    Private Function GerarInutilizacao() As Boolean
        Dim Sqls As New ArrayList
        Dim i As Integer = CInt(txtNotaInicial.Text)

        Dim verPendencias As Boolean = True

        While i <= CInt(txtNotaFinal.Text)
            Dim dsNotaFiscal As New DataSet
            Dim dsPendencias As New DataSet
            Dim dsRealizadas As New DataSet

            Sql = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_id " & vbCrLf & _
                  "  FROM NotasFiscais " & vbCrLf & _
                  " WHERE Empresa_id = '" & objEmpresa.Codigo & "'" & vbCrLf & _
                  "   AND EndEmpresa_id = " & objEmpresa.CodigoEndereco & vbCrLf & _
                  "   AND Nota_id = " & i & vbCrLf & _
                  "   AND Serie_id= '" & txtSerie.Text & "'" & vbCrLf & _
                  "   AND " & IIf(rdSaida.Checked, "EntradaSaida_Id = 'S' ", "EntradaSaida_Id = 'E' ")

            dsNotaFiscal = Banco.ConsultaDataSet(Sql, "NotaFiscal")

            If dsNotaFiscal.Tables(0).Rows.Count > 0 Then

                If Not CInt(txtNotaInicial.Text) = CInt(txtNotaFinal.Text) Then
                    MsgBox(Me.Page, "Só pode ser excluido/inutilizado uma Nota fiscal por vez!")
                    Return False
                End If

                Dim objNFe As New [Lib].Negocio.NotaFiscal()
                objNFe.CodigoEmpresa = dsNotaFiscal.Tables(0).Rows(0)("Empresa_Id")
                objNFe.EnderecoEmpresa = dsNotaFiscal.Tables(0).Rows(0)("EndEmpresa_Id")
                objNFe.CodigoCliente = dsNotaFiscal.Tables(0).Rows(0)("Cliente_Id")
                objNFe.EnderecoCliente = dsNotaFiscal.Tables(0).Rows(0)("EndCliente_Id")
                objNFe.EntradaSaida = IIf(dsNotaFiscal.Tables(0).Rows(0)("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                objNFe.Serie = dsNotaFiscal.Tables(0).Rows(0)("Serie_Id")
                objNFe.Codigo = dsNotaFiscal.Tables(0).Rows(0)("Nota_id")
                objNFe = New [Lib].Negocio.NotaFiscal(objNFe)

                If objNFe.CodigoSituacao = CInt(eSituacao.Bloqueado) Then

                    Sql = "SELECT Retorno, ChaveNfe " & vbCrLf & _
                             "  FROM NfePendencias " & vbCrLf & _
                             " WHERE Empresa_id = '" & objNFe.CodigoEmpresa & "'" & vbCrLf & _
                             "   AND EndEmpresa_id = " & objNFe.EnderecoEmpresa & vbCrLf & _
                             "   AND Cliente_Id = '" & objNFe.CodigoCliente & "'" & vbCrLf & _
                             "   AND EndCliente_Id = " & objNFe.EnderecoCliente & vbCrLf & _
                             "   AND Nota_id = " & objNFe.Codigo & vbCrLf & _
                             "   AND Serie_id = '" & objNFe.Serie & "'" & vbCrLf & _
                             "   AND " & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "EntradaSaida_Id = 'E' ", "EntradaSaida_Id = 'S' ")

                    dsPendencias = Banco.ConsultaDataSet(Sql, "Pendencias")

                    If dsPendencias.Tables(0).Rows.Count > 0 Then
                        '233 - IE do destinatario nao cadastrada. Orientações: Verifique se a Inscrição Estadual do Destinatário/Remetente informada está cadastrada na SEFAZ
                        '302 - Rejeição Irregularidade fiscal do destinatário
                        '303 - Rejeição: Destinatário não habilitado a operar na UF
                        '305 - Rejeição: Destinatário bloqueado na UF
                        '306 - Rejeição: IE do destinatário não está ativa na UF
                        '307 - Rejeição: Emitente bloqueado pela UF de destino, em operação com consumidor final
                        '508 - Rejeição: CST incompatível na operação com Não Contribuinte [nItem:999]
                        '795 - Rejeicao: Total do ICMS desonerado difere do somatorio dos itens
                        '805 - Rejeição: A SEFAZ do destinatario nao permite Contribuinte Isento de Inscricao Estadual. 
                        '817 - Unidade Tributavel incompativel com o NCM informado na operacao com Comercio Exterior. [nItem:1]
                        If Session("ssNomeUsuario") = "FURLAN" OrElse
                            hdLiberar.Value = True OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 233 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 302 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 303 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 305 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 306 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 307 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 508 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 795 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 805 OrElse
                            CInt(dsPendencias.Tables(0).Rows(0)("Retorno")) = 817 Then

                            If RemoverEInutilizar(objNFe, dsPendencias.Tables(0).Rows(0)("ChaveNfe")) Then
                                If exclusaoLiberada(objNFe) Then
                                    verPendencias = False
                                Else
                                    MsgBox(Me.Page, "Nota fiscal " & i & " está em pendências no monitoramento de nota fiscal, portanto não pode ser inutilizada!")
                                    Return False
                                End If
                            Else
                                MsgBox(Me.Page, "Nota fiscal " & i & " não teve nenhum retorno da Sefaz. Aguarde um pouco e reenvie a solicitação!")
                                Return False
                            End If
                        Else
                            MsgBox(Me.Page, "Nota fiscal " & i & " está em pendências com RETORNO " & dsPendencias.Tables(0).Rows(0)("Retorno").ToString & ", portanto não pode ser inutilizada!")
                            lnkLiberar.Parent.Visible = True
                            Return False
                        End If
                    Else
                        MsgBox(Me.Page, "Nota fiscal " & i & " não está em pendências no monitoramento de nota fiscal, portanto não pode ser inutilizada!")
                        Return False
                    End If
                Else
                    MsgBox(Me.Page, "Nota fiscal " & i & " já existe lançada, portanto não pode ser inutilizada!")
                    Return False
                End If
            End If

            If verPendencias Then
                Sql = "SELECT Nota_id " & vbCrLf & _
                      "  FROM NfePendencias " & vbCrLf & _
                      " WHERE Empresa_id = '" & objEmpresa.Codigo & "'" & vbCrLf & _
                      "   AND EndEmpresa_id = " & objEmpresa.CodigoEndereco & vbCrLf & _
                      "   AND Nota_id = " & i & vbCrLf & _
                      "   AND Serie_id = '" & txtSerie.Text & "'" & vbCrLf & _
                      "   AND " & IIf(rdSaida.Checked, "EntradaSaida_Id = 'S' ", "EntradaSaida_Id = 'E' ")

                dsPendencias = Banco.ConsultaDataSet(Sql, "Pendencias")

                If dsPendencias.Tables(0).Rows.Count > 0 Then
                    MsgBox(Me.Page, "Nota fiscal " & i & " está em pendências no monitoramento de nota fiscal, portanto não pode ser inutilizada!")
                    Return False
                End If

                Sql = "SELECT Nota_id " & vbCrLf & _
                      "  FROM NfeRealizadas " & vbCrLf & _
                      " WHERE Empresa_id = '" & objEmpresa.Codigo & "'" & vbCrLf & _
                      "   AND EndEmpresa_id = " & objEmpresa.CodigoEndereco & vbCrLf & _
                      "   AND Nota_id = " & i & vbCrLf & _
                      "   AND Serie_id = '" & txtSerie.Text & "'" & vbCrLf & _
                      "   AND " & IIf(rdSaida.Checked, "EntradaSaida_Id = 'S' ", "EntradaSaida_Id = 'E' ")

                dsRealizadas = Banco.ConsultaDataSet(Sql, "Realizadas")

                If dsRealizadas.Tables(0).Rows.Count > 0 Then
                    MsgBox(Me.Page, "Nota fiscal " & i & " está em realizadas no monitoramento de nota fiscal, portanto não pode ser inutilizada!")
                    Return False
                End If
            End If

            i += 1
        End While

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        If CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.Nota Then
            obj.NomeArquivo = String.Format("inutiliza{0}-{1}#{2}.txt", txtNotaInicial.Text.Trim(), txtNotaFinal.Text.Trim(), objEmpresa.Codigo)
        ElseIf CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.CTRC OrElse CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.CT_E Then
            obj.NomeArquivo = String.Format("inutilizacte{0}-{1}#{2}.txt", txtNotaInicial.Text.Trim(), txtNotaFinal.Text.Trim(), objEmpresa.Codigo)
        End If
        obj.Texto = getTextoInutilizacao()

        obj.SalvarSql(Sqls)

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Exit Function
        End If

        'AGUARDANDO RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Empty
        If CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.Nota Then
            fileName = String.Format("resp-inutiliza{0}-{1}#{2}.txt", txtNotaInicial.Text.Trim(), txtNotaFinal.Text.Trim(), objEmpresa.Codigo)
        ElseIf CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.CTRC OrElse CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.CT_E Then
            fileName = String.Format("resp-inutilizacte{0}-{1}#{2}.txt", txtNotaInicial.Text.Trim(), txtNotaFinal.Text.Trim(), objEmpresa.Codigo)
        End If

        While resp Is Nothing
            resp = GetResp(fileName)
        End While

        If resp IsNot Nothing Then
            Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()

            Dim strCodigo As String = String.Empty
            If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            End If

            Dim strMsg As String = String.Empty
            If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            End If

            Dim strProtocolo As String = String.Empty
            If protocolo IsNot Nothing AndAlso protocolo.Length > 0 Then
                strProtocolo = protocolo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            End If

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "102" Or (Session("ssNomeUsuario") = "FURLAN" And strCodigo = "563")) Then
                Sqls.Clear()
                Dim index As Integer = CInt(txtNotaInicial.Text)
                While index <= CInt(txtNotaFinal.Text)

                    'Insere Notas, Itens e Encargos somente como registro demonstrativo.
                    Sql = Inutilizar(index)

                    'GRAVAR NAS REALIZADAS QUANDO A INUTILIZAÇÃO DER CERTO
                    Sql &= "INSERT INTO NFERealizadas " & vbCrLf &
                          "       (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora," & vbCrLf &
                          "        Usuario, TpEmis, ChaveNfe, ObservacoesFiscais, DadosAdicionais, Protocolo) " & vbCrLf &
                          "VALUES ('" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", '" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", " & vbCrLf &
                          "        " & IIf(rdSaida.Checked, "'S'", "'E'") & ",'" & txtSerie.Text.Trim() & "', " & index & ", " & vbCrLf &
                          "        '" & DateTime.Now.ToString("yyyy-MM-dd") & "', '" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "', " & vbCrLf &
                          "        '" & HttpContext.Current.Session("ssNomeUsuario") & "',  1, '', '" & Funcoes.EliminarCaracteresEspeciais(txtJustificativa.Text.Trim()) & "', " & vbCrLf &
                          "        'INUTILIZAR', '" & strProtocolo & "'); " & vbCrLf
                    Sqls.Add(Sql)
                    index += 1
                End While

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Function
                End If
            End If
            MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
        End If
        Return True
    End Function

    Private Function RemoverEInutilizar(ByVal objNFe As [Lib].Negocio.NotaFiscal, ByVal ChaveNFE As String) As Boolean
        Dim SqlsR As New ArrayList
        Dim strMsg As String = String.Empty

        Dim liberado As Boolean = False

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("consulta{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa)
        obj.Texto = getTextoConsultar(objNFe, ChaveNFE)
        obj.SalvarSql(SqlsR)

        If Not Banco.GravaBanco(SqlsR) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Return liberado
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

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            '217 - Rejeição: NF-e não consta na base de dados da Sefaz
            '233 - IE do destinatario nao cadastrada. Orientações: Verifique se a Inscrição Estadual do Destinatário/Remetente informada está cadastrada na SEFAZ
            '305 - Rejeição: Destinatário bloqueado na UF
            '306 - Rejeição: IE do destinatário não está ativa na UF
            '508 - Rejeição: CST incompatível na operação com Não Contribuinte [nItem:999]
            '795	Rejeicao: Total do ICMS desonerado difere do somatorio dos itens
            '805 - Rejeição: A SEFAZ do destinatario nao permite Contribuinte Isento de Inscricao Estadual. 
            '817 - Unidade Tributavel incompativel com o NCM informado na operacao com Comercio Exterior. [nItem:1]
            If Session("ssNomeUsuario") = "FURLAN" OrElse hdLiberar.Value = True OrElse (Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "217" OrElse strCodigo = "233" OrElse strCodigo = "302" OrElse strCodigo = "303" OrElse strCodigo = "305" OrElse strCodigo = "306" OrElse strCodigo = "307" OrElse strCodigo = "508" OrElse strCodigo = "795" OrElse strCodigo = "805" OrElse strCodigo = "817")) Then
                liberado = True
            Else
                strMsg = String.Format("{0} - {1}", strCodigo, strMsg)
                MsgBox(Me.Page, strMsg)
            End If
        End If

        Return liberado
    End Function

    Private Function getTextoConsultar(ByVal nfe As [Lib].Negocio.NotaFiscal, ByVal ChaveNfe As String) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & ChaveNfe & ControlChars.CrLf)
        sb.Append("NRECIBO =" & nfe.ReciboNota & ControlChars.CrLf)
        Return sb.ToString()
    End Function


    Private Function exclusaoLiberada(ByVal objNFe As [Lib].Negocio.NotaFiscal) As Boolean
        Dim SqlsE As New ArrayList

        objNFe.IUD = "D"
        objNFe.SalvarSql(SqlsE)

        If Banco.GravaBanco(SqlsE) Then
            Return True
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Return False
        End If

    End Function


    Private Function Inutilizar(ByVal nf As String) As String
        'GRAVAR REGISTRO NA TABELA NOTASFISCAIS COM SITUAÇÃO 9
        Sql = "INSERT INTO NotasFiscais " & vbCrLf & _
              "            (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
              "             Formulario, Pedido, Procuracao, Operacao, SubOperacao, Finalidade, Movimento, DataDaNota, " & vbCrLf & _
              "             NossaEmissao, SerieNotadoProdutor, NumeroNotadoProdutor, PedidoOrigem, EntradaSaidaOrigem, " & vbCrLf & _
              "             SerieOrigem, NotaOrigem, Deposito, EndDeposito, Destino, EndDestino, Transbordo, EndTransbordo, " & vbCrLf & _
              "             Observacoes, Eletronica, UsuarioInclusao, UsuarioInclusaoData, Situacao, TipoDeDocumento) " & vbCrLf & _
              "VALUES ('" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", '" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", " & vbCrLf & _
              "        '" & IIf(rdSaida.Checked, "S", "E") & "', '" & txtSerie.Text & "', " & nf & ", 0, NULL, 0, " & IIf(rdSaida.Checked, "21", "1") & ", 1, 1, '" & Now().ToString("yyyy-MM-dd") & "', '" & Now().ToString("yyyy-MM-dd") & "', " & vbCrLf & _
              "        'S', '', 0, 0, '', '', 0, '" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", '" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", " & vbCrLf & _
              "        '', 0, '" & Funcoes.EliminarCaracteresEspeciais(RTrim(txtJustificativa.Text)) & "', 'S', '" & HttpContext.Current.Session("ssNomeUsuario") & "', '" & Now().ToString("yyyy-MM-dd") & "', 9,1); " & vbCrLf

        'GRAVAR REGISTRO NA TABELA NOTASFISCAISXITENS
        Sql &= "INSERT INTO NotasFiscaisXItens " & vbCrLf & _
               "            (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
               "             Produto_Id, CFOP_Id, PesoFiscal, QuantidadeFisica, QuantidadeFiscal, Unitario, Valor, PesoQuantidade, Sequencia_Id, Operacao, SubOperacao, OperacaoXEstado) " & vbCrLf & _
               "VALUES ('" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", '" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", " & vbCrLf & _
               "        '" & IIf(rdSaida.Checked, "S", "E") & "','" & txtSerie.Text & "', " & nf & ", " & vbCrLf

        If Left(objEmpresa.Codigo, 8) = "03189063" Then
            Sql &= "'10101001', " & vbCrLf
        Else
            Sql &= "'101010001', " & vbCrLf
        End If

        Sql &= "" & IIf(rdSaida.Checked, "5101", "1101") & "," & "0, 0, 0, 0, 0, '', 0," & IIf(rdSaida.Checked, "21", "1") & ",1" & vbCrLf

        If objEmpresa.Codigo = "05272759000147" Then
            If rdSaida.Checked Then
                Sql &= ",4);"
            Else
                Sql &= ",3);"
            End If
        ElseIf Left(objEmpresa.Codigo, 8) = "04854422" Then
            If rdSaida.Checked Then
                Sql &= ",3);"
            Else
                Sql &= ",2);"
            End If
        ElseIf Left(objEmpresa.Codigo, 8) = "04440724" Then
            If rdSaida.Checked Then
                Sql &= ",97);"
            Else
                Sql &= ",96);"
            End If
        ElseIf Left(objEmpresa.Codigo, 8) = "03189063" Then
            If rdSaida.Checked Then
                Sql &= ",92);"
            Else
                Sql &= ",26);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "05366261" Then
            If rdSaida.Checked Then
                Sql &= ",17);"
            Else
                Sql &= ",18);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "40938762" Then
            If rdSaida.Checked Then
                Sql &= ",17);"
            Else
                Sql &= ",18);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "44005444" Then
            If rdSaida.Checked Then
                Sql &= ",1278);"
            Else
                Sql &= ",2467);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "48984539" Then
            If rdSaida.Checked Then
                Sql &= ",1278);"
            Else
                Sql &= ",2467);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "44979506" Then
            If rdSaida.Checked Then
                Sql &= ",297);"
            Else
                Sql &= ",353);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "24450490" Then
            If rdSaida.Checked Then
                Sql &= ",126518);"
            Else
                Sql &= ",126520);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "62747840" Then
            If rdSaida.Checked Then
                Sql &= ",14);"
            Else
                Sql &= ",15);"
            End If
        ElseIf Left(objEmpresa.Codigo, 8) = "62780383" Then
            If rdSaida.Checked Then
                Sql &= ",810);"
            Else
                Sql &= ",113);"
            End If

        ElseIf Left(objEmpresa.Codigo, 8) = "63358210" Then
            If rdSaida.Checked Then
                Sql &= ",1648);"
            Else
                Sql &= ",951);"
            End If

        End If

        'GRAVAR REGISTRO NA TABELA NOTASFISCAISXENCARGOS
        Sql &= "INSERT INTO NotasFiscaisXEncargos " & vbCrLf & _
               "           (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
               "            Produto_Id, CFOP_Id, Encargo_Id, Base, Percentual, Valor, Sequencia_Id) " & vbCrLf & _
               "VALUES ('" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", '" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", " & vbCrLf & _
               "        '" & IIf(rdSaida.Checked, "S", "E") & "','" & txtSerie.Text & "', " & nf & ", " & vbCrLf
        If Left(objEmpresa.Codigo, 8) = "03189063" Then
            Sql &= "'10101001', " & vbCrLf
        Else
            Sql &= "'101010001', " & vbCrLf
        End If
        Sql &= "" & IIf(rdSaida.Checked, "5101", "1101") & ",'PRODUTO', 0, 0, 0, 0); " & vbCrLf


        Sql &= "INSERT INTO NotasFiscaisXEncargos " & vbCrLf & _
               "            (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
               "             Produto_Id, CFOP_Id, Encargo_Id, Base, Percentual, Valor, Sequencia_Id)" & vbCrLf & _
               "VALUES ('" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", '" & objEmpresa.Codigo & "', " & objEmpresa.CodigoEndereco & ", " & vbCrLf & _
               "        '" & IIf(rdSaida.Checked, "S", "E") & "','" & txtSerie.Text & "', " & nf & ", " & vbCrLf
        If Left(objEmpresa.Codigo, 8) = "03189063" Then
            Sql &= "'10101001', " & vbCrLf
        Else
            Sql &= "'101010001', " & vbCrLf
        End If
        Sql &= "" & IIf(rdSaida.Checked, "5101", "1101") & ",'LIQUIDO', 0, 0, 0, 0); " & vbCrLf

        Return Sql
    End Function

    Private Function ValidarCampos() As Boolean
        If txtEmpresa.Text.Length = 0 Then
            MsgBox(Me.Page, "Empresa não foi selecionada!")
            txtEmpresa.Focus()
            Return False
        ElseIf txtNotaInicial.Text.Length = 0 Then
            MsgBox(Me.Page, "Sequência da Nota Fiscal inicial não foi informada!")
            txtNotaInicial.Focus()
            Return False
        ElseIf txtNotaFinal.Text.Length = 0 Then
            MsgBox(Me.Page, "Sequência da Nota Fiscal final não doi informada!")
            txtNotaFinal.Focus()
            Return False
        ElseIf CInt(txtNotaInicial.Text) > CInt(txtNotaFinal.Text) Then
            MsgBox(Me.Page, "Sequência da Nota Fiscal final não pode ser menor que inicial!")
            txtNotaInicial.Focus()
            Return False
        ElseIf txtSerie.Text.Length = 0 Then
            MsgBox(Me.Page, "Série da Nota Fiscal não foi informada!")
            txtSerie.Focus()
            Return False
        ElseIf txtJustificativa.Text.Length = 0 Then
            MsgBox(Me.Page, "Justificativa de ter pelo menos 15 caracteres!")
            txtJustificativa.Focus()
            Return False
        ElseIf txtJustificativa.Text.Length < 15 Then
            MsgBox(Me.Page, "Justificativa de ter pelo menos 15 caracteres!")
            txtJustificativa.Focus()
            Return False
        ElseIf txtJustificativa.Text.Length > 255 Then
            MsgBox(Me.Page, "Justificativa não pode ter mais que 255 caracteres!")
            txtJustificativa.Focus()
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub Limpar()
        lnkLiberar.Parent.Visible = False
        hdLiberar.Value = False
        txtEmpresa.Text = ""
        txtNotaInicial.Text = ""
        txtNotaFinal.Text = ""
        txtSerie.Text = ""
        txtJustificativa.Text = ""
        hdfTipoDeDocumento.Value = ""
        Session.Remove("objLiberaInutilizacao" & HID.Value)
        SelecionarEmpresa(Session("ssEmpresa"), Session("ssEndEmpresa"))
    End Sub

    Private Function getTextoInutilizacao() As String
        Dim sb As New StringBuilder()
        If CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.Nota Then
            sb.Append("CODIGOUF=" & objEmpresa.Municipio.EstadoIbge & ControlChars.CrLf)
            sb.Append("ANO =" & Mid(Now.Year.ToString, 3, 2) & ControlChars.CrLf)
            sb.Append("CNPJ =" & objEmpresa.Codigo & ControlChars.CrLf)
            sb.Append("MODELO = 55" & ControlChars.CrLf)
            sb.Append("SERIE =" & txtSerie.Text.Trim() & ControlChars.CrLf)
            sb.Append("NFEINI =" & txtNotaInicial.Text.Trim() & ControlChars.CrLf)
            sb.Append("NFEFIM =" & txtNotaFinal.Text.Trim() & ControlChars.CrLf)
            sb.Append("JUSTIFICATIVA =" & Funcoes.EliminarCaracteresEspeciais(txtJustificativa.Text.ToUpper().Trim()) & ControlChars.CrLf)
        ElseIf CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.CTRC OrElse CType(CInt(hdfTipoDeDocumento.Value), eTipoDeDocumento) = eTipoDeDocumento.CT_E Then
            sb.Append("CODIGOUF=" & objEmpresa.Municipio.EstadoIbge & ControlChars.CrLf)
            sb.Append("ANO =" & Mid(Now.Year.ToString, 3, 2) & ControlChars.CrLf)
            sb.Append("CNPJ =" & objEmpresa.Codigo & ControlChars.CrLf)
            sb.Append("MODELO = 57" & ControlChars.CrLf)
            sb.Append("SERIE =" & txtSerie.Text.Trim() & ControlChars.CrLf)
            sb.Append("CTEINI =" & txtNotaInicial.Text.Trim() & ControlChars.CrLf)
            sb.Append("CTEFIM =" & txtNotaFinal.Text.Trim() & ControlChars.CrLf)
            sb.Append("JUSTIFICATIVA =" & Funcoes.EliminarCaracteresEspeciais(txtJustificativa.Text.ToUpper().Trim()) & ControlChars.CrLf)
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

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo") = "Livre"
        If TypeOf Me.Page Is NotaFiscalXItens Then
            Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(CType(Me.Page, NotaFiscalXItens), "ucConsultaEmpresas"), UserControl)
            CType(uc, ucConsultaEmpresas).Limpar()
            Me.MainUserControl = Me
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaNUNXI" & HID.Value)
        ElseIf TypeOf Me.Page Is ConhecimentoDeTransporte Then
            Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(CType(Me.Page, ConhecimentoDeTransporte), "ucConsultaEmpresas"), UserControl)
            CType(uc, ucConsultaEmpresas).Limpar()
            Me.MainUserControl = Me
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaNUNXI" & HID.Value)
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("INUTILIZACAOFISCAL", "GRAVAR") Then
            If ValidarCampos() Then
                If GerarInutilizacao() Then
                    Popup.CloseDialog(Me.Page, "divInutilizacao")
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para inutilizar sequência de nota fiscal!")
        End If
    End Sub

    Protected Sub lnkLiberar_Click(sender As Object, e As EventArgs) Handles lnkLiberar.Click
        Dim ucLiberacao = CType(Me.Page.FindControlRecursive("ucLiberacao"), ucLiberacao)
        If ucLiberacao IsNot Nothing Then
            hdLiberar.Value = False
            ucLiberacao.Limpar()
            ucLiberacao.MainUserControl = Me
        End If

        Popup.ConsultaLiberacao(Me.Page, "objLiberaInutilizacao" & HID.Value)
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divInutilizacao")
    End Sub

End Class