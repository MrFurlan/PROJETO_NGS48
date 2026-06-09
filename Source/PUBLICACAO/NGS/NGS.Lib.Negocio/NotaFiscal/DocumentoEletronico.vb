Imports System.Net.Configuration
Imports System.Threading
Imports System.Xml
Imports System.Text
Imports System.Data
Imports System.Web
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Web.UI.WebControls
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Globalization

<Serializable()>
Public Class DocumentoEletronico
    Implements IBaseEntity

#Region "NF-e"
    Public Shared Function getTextoConsultar(ByVal nfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nfe.ChaveNFE & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function getTextoConsultarCTe(ByVal nfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVECTE=" & nfe.ChaveNFE & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

    Public Shared Function VerificarModoOperacaoNFe(ByRef Codigo As String, ByRef Empresa As String, ByRef msgErro As String, ByRef codRetorno As String) As Boolean
        Dim aux As Boolean = True

        Try
            Dim Sqls As New ArrayList
            Dim Sql As String = String.Empty

            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("status{0:000000000}#{1}.txt", Codigo, Empresa)
            obj.Texto = String.Empty
            obj.SalvarSql(Sqls)

            Dim Banco As New AcessaBanco
            If Not Banco.GravaBanco(Sqls) Then
                msgErro = HttpContext.Current.Session("ssMessage")
                Return False
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-status{0:000000000}#{1}.txt", Codigo, Empresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(30)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If resp IsNot Nothing Then
                Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                Sqls.Clear()

                If String.IsNullOrWhiteSpace(strCodigo) Then
                    msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE"
                    aux = False
                ElseIf strCodigo = "107" Then
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", Codigo, Empresa) & "';"
                    Sqls.Add(Sql)

                    codRetorno = "107"
                    msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - MODO NORMAL"
                ElseIf strCodigo = "4036" Then
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", Codigo, Empresa) & "';"
                    Sqls.Add(Sql)

                    codRetorno = "4036"
                    msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - MODO CONTINGÊNCIA"
                Else
                    msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                    aux = False
                End If

                If Sqls.Count > 0 Then Banco.GravaBanco(Sqls)
            Else
                msgErro = "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação. "
                aux = False
            End If
        Catch ex As Exception
            msgErro = ex.Message
            aux = False
        End Try

        Return aux
    End Function

    Public Shared Function ConsultaStatusNFE(ByVal nfe As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, ByRef codRetorno As String, ByRef msgRetorno As String, ByRef codRecibo As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("consulta{0:000000000}#{1}.txt", nfe.Codigo, nfe.CodigoEmpresa)
        obj.Texto = getTextoConsultar(nfe)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-consulta{0:000000000}#{1}.txt", nfe.Codigo, nfe.CodigoEmpresa)

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
            Dim strRecibo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RECIBO")).FirstOrDefault()
            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) Then
                codRetorno = strCodigo
                msgRetorno = strMsg
                codRecibo = strRecibo
                Return True
            Else
                msgErro = "999 - Consulte novamente." & strMsg
                Return False
            End If
        Else
            msgErro = "999 - Sefaz não retornou nenhuma resposta, consulte novamente."
            Return False
        End If
    End Function

    Public Shared Function ConsultaNFEFornecedor(ByVal nfe As [Lib].Negocio.NotaFiscal) As String
        Dim Sqls As New ArrayList
        Dim Resposta As String = String.Empty
        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("consulta{0:000000000}#{1}.txt", CInt(Mid(nfe.ChaveNFE, 26, 9)), nfe.CodigoEmpresa)
        obj.Texto = getTextoConsultar(nfe)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            Resposta = "999;" & HttpContext.Current.Session("ssMessage") & ";"
            Return Resposta
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-consulta{0:000000000}#{1}.txt", CInt(Mid(nfe.ChaveNFE, 26, 9)), nfe.CodigoEmpresa)

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
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If strCodigo = "217" Then
                Return "999;Verifique se a chave de NF-e informada está correta.;" & strMsg & ";"
            End If

            Dim strData As String = String.Empty

            If lstMsg.Count > 3 AndAlso Not String.IsNullOrWhiteSpace(lstMsg(3)) Then
                strData = lstMsg(3).Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            End If

            If Not String.IsNullOrWhiteSpace(strCodigo) Then
                Return String.Format("{0};{1};{2}", strCodigo, strMsg, strData)
            Else
                Return "999;Consulte novamente.;" & strMsg & ";"
            End If
        Else
            Return "999;Sefaz não retornou nenhuma resposta, consulte novamente.;"
        End If

        Return Resposta
    End Function

    Public Shared Function getTextoNFe4G(ByVal nfe As [Lib].Negocio.NotaFiscal, ByVal HomologAlfasig As Boolean, ByVal ModoDeEmissao As Integer) As String
        Dim sb As New StringBuilder()
        Dim sqls As New ArrayList

        sb.Append("[IDE]" & ControlChars.CrLf)
        'Código da UF do Emitente Documento Fiscal
        sb.Append("CUF	        = " & nfe.Empresa.Municipio.EstadoIbge & ControlChars.CrLf)
        'Descrição da Natureza da Operação
        If (nfe.CFOP.Codigo = 5949 OrElse nfe.CFOP.Codigo = 6949) AndAlso nfe.CodigoFinalidade = 31 Then
            sb.Append("NATOP        = NOTA FISCAL DE ESTORNO" & ControlChars.CrLf)
        ElseIf nfe.SubOperacao.FinalidadeDaNota = 3 Then
            'Foi comentdo devido a uma exigência da NUTRI - 28/02/2025
            'sb.Append("NATOP        = NOTA FISCAL DE ESTORNO" & ControlChars.CrLf)
            sb.Append("NATOP        = TRANSFERENCIA DE SALDO" & ControlChars.CrLf)
        Else
            sb.Append("NATOP        = " & Funcoes.EliminarCaracteresEspeciais(Mid(nfe.CFOP.Descricao, 1, 60)) & ControlChars.CrLf)
        End If

        'Indicador da Forma de Pagamento
        'If (nfe.VencimentosNota IsNot Nothing AndAlso nfe.VencimentosNota.Count > 0) OrElse (nfe.CodigoPedido > 0 AndAlso nfe.Pedido.Vencimentos.Count > 0) Then
        '    sb.Append("INDPAG      = 1" & ControlChars.CrLf)
        'Else
        '    sb.Append("INDPAG      = 2" & ControlChars.CrLf)
        'End If

        'Código do Modelo do Documento Fiscal
        sb.Append("MOD         = 55" & ControlChars.CrLf)
        'Série do Documento Fiscal
        sb.Append("SERIE       = " & nfe.Serie & ControlChars.CrLf)
        'Número do Documento Fiscal
        sb.Append("NNF         = " & nfe.Codigo & ControlChars.CrLf)
        'Data/Hora de emissão do Documento Fiscal
        sb.Append("DHEMI       = " & nfe.Movimento.ToString("yyyy-MM-dd") & "T" & nfe.DataInclusao.ToString("HH:mm:ss") & ControlChars.CrLf)
        'Data/Hora de Saída ou da Entrada da Mercadoria/Produto
        sb.Append("DHSAIENT    = " & nfe.Movimento.ToString("yyyy-MM-dd") & "T" & nfe.DataInclusao.ToString("HH:mm:ss") & ControlChars.CrLf)
        'Tipo de Operação - ENTRADA = 0 - SAÍDA = 1
        If nfe.EntradaSaida = eEntradaSaida.Entrada Then
            sb.Append("TPNF        = 0" & ControlChars.CrLf)
        Else
            sb.Append("TPNF        = 1" & ControlChars.CrLf)
        End If

        'Identificador de Local de destino da operação (1 - Interno, 2 - Interestadual e 3 - Exterior)
        If nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado Then
            If nfe.CodigoEmpresa = nfe.CodigoCliente AndAlso nfe.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS AndAlso Not nfe.Empresa.CodigoEstado = nfe.Destino.CodigoEstado Then
                sb.Append("IDDEST      = 2" & ControlChars.CrLf)
            Else
                sb.Append("IDDEST      = 1" & ControlChars.CrLf)
            End If
        ElseIf Not nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado AndAlso nfe.Cliente.CodigoEstado = "EX" Then
            sb.Append("IDDEST      = 3" & ControlChars.CrLf)
        ElseIf Not nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado Then
            sb.Append("IDDEST      = 2" & ControlChars.CrLf)
        End If

        'MODE DE IMPRESSÃO  1 - RETRATO   2 - PAISAGEM
        sb.Append("TPIMP       = 1" & ControlChars.CrLf)

        'CÓDIGO DO MUNICÍPIO DO IBGE de Ocorrência do Fato Gerador
        sb.Append("CMUNFG      = " & nfe.Empresa.Municipio.EstadoIbge & nfe.Empresa.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)

        'Finalidade de emissão da NF-E: 1 - NFE-NORMAL    2 – NF-E COMPLEMENTAR   3 – NF-E DE AJUSTE - 4 Devolução
        Dim ModoAjuste As Boolean = False
        For Each item In nfe.Itens
            If item.QuantidadeFiscal = 0 AndAlso item.Unitario = 0 Then ModoAjuste = True
        Next

        Dim finNFE As Integer = 0

        If nfe.SubOperacao.FinalidadeDaNota = 2 Then 'FINALIDADE DA NOTA NA SUBOPERAÇÃO FOR COMPLEMENTO - 13/08/2024
            sb.Append("FINNFE      = 2" & ControlChars.CrLf)
        ElseIf nfe.SubOperacao.FinalidadeDaNota = 3 Then 'FINALIDADE DA NOTA NA SUBOPERAÇÃO FOR AJUSTE - 13/08/2024
            sb.Append("FINNFE      = 3" & ControlChars.CrLf)
        ElseIf nfe.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
            sb.Append("FINNFE      = 1" & ControlChars.CrLf)
            finNFE = 1
        ElseIf nfe.CFOP.Codigo = 5601 Then 'TRANSFERENCIA DE CREDITO DE ICMS - FURLAN - 28/11/2023
            sb.Append("FINNFE      = 1" & ControlChars.CrLf)
            finNFE = 1
        ElseIf (nfe.CFOP.Codigo = 5949 OrElse nfe.CFOP.Codigo = 6949) AndAlso nfe.CodigoFinalidade = 31 Then
            sb.Append("FINNFE      = 3" & ControlChars.CrLf)
        ElseIf nfe.CodigoFinalidade = 28 Then
            sb.Append("FINNFE      = 2" & ControlChars.CrLf)
        ElseIf nfe.CFOP.Descricao.ToString.Contains("DEV.") OrElse
                nfe.CFOP.Descricao.ToString.Contains("DEVOLUCAO") OrElse
                nfe.CFOP.Descricao.ToString.Contains("DEVOLUÇÃO") Then
            sb.Append("FINNFE      = 4" & ControlChars.CrLf)
        ElseIf (nfe.ModoComplemento AndAlso nfe.NotasReferenciais.Count > 0) OrElse
                (ModoAjuste And nfe.VencimentosNota.Count > 0 And nfe.SubOperacao.Financeiro) Then 'SE POSSUI VENCIMENTOS MODO DE AJUSTE DEVE SER COMPLEMENTO- FURLAN - 24/01/2022
            sb.Append("FINNFE      = 2" & ControlChars.CrLf)
        ElseIf ModoAjuste Then
            sb.Append("FINNFE      = 3" & ControlChars.CrLf)
        Else
            sb.Append("FINNFE      = 1" & ControlChars.CrLf)
            finNFE = 1
        End If

        'Indica Operação com Consumidor Final "0 - Normal    1 - Consumidor final"
        If Not nfe.Cliente.CodigoEstado = "EX" Then
            If nfe.Cliente.InscricaoEstadual.ToString.Contains("ISENTO") Or nfe.Cliente.InscricaoEstadual.Length = 0 Then
                sb.Append("INDFINAL    = 1" & ControlChars.CrLf)
            Else
                sb.Append("INDFINAL    = 0" & ControlChars.CrLf)
            End If
        Else
            sb.Append("INDFINAL    = 0" & ControlChars.CrLf)
        End If

        If ModoAjuste Then
            sb.Append("INDPRES     = 0" & ControlChars.CrLf)
        Else
            sb.Append("INDPRES     = 9" & ControlChars.CrLf)
        End If

        'Obrigatório à partir de 01/09/2021 - Furlan 
        If finNFE = 1 Then 'Em conversa com a Alfasig só colocar se for FINNFE = 1 - Furlan - 07/06/2022
            sb.Append("INDINTERMED=0" & ControlChars.CrLf)
        End If

        sb.Append("   PROCEMI=0" & ControlChars.CrLf)

        'ENTRA EM VIFOR À PARTIR DE 02/07/2018 - FURLAN
        sb.Append("   VERPROC=nfe4g" & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)

        'Grupo com as informações das NF/NF-e /NF de produtor/Cupom Fiscal referenciadas.
        'Esta informação será utilizada nas hipóteses previstas na legislação. (Ex.: Devolução de Mercadorias, Substituição de NF cancelada, Complementação de NF, etc.).
        Dim NotasDevolucao = IIf(nfe.SubOperacao.Descricao = "COMPLEMENTACAO DE ICMS" Or nfe.SubOperacao.Descricao = "COMPLEMENTO DE ICMS-ST" Or nfe.SubOperacao.Descricao = "COMPLEMENTACAO DE IPI", (From NF In nfe.Itens.SelectMany(Function(s) s.NotasDevolucao.Where(Function(i) i.CodigoCFOP > 0))
                                                                                                                                                                                                       Select NF.Nota.Codigo, NF.Nota.CodigoTipoDeDocumento, NF.Nota.Empresa.Municipio.EstadoIbge, NF.Nota.Movimento, NF.Nota.CodigoEmpresa, NF.Nota.Serie, NF.Nota.ChaveNFE).Distinct(), (From NF In nfe.Itens.SelectMany(Function(s) s.NotasDevolucao.Where(Function(i) i.ValorDevolucao > 0 OrElse i.QuantidadeDevolucao > 0))
                                                                                                                                                                                                                                                                                                                                                                                           Select NF.Nota.Codigo, NF.Nota.CodigoTipoDeDocumento, NF.Nota.Empresa.Municipio.EstadoIbge, NF.Nota.Movimento, NF.Nota.CodigoEmpresa, NF.Nota.Serie, NF.Nota.ChaveNFE).Distinct())
        For Each t In NotasDevolucao
            sb.Append("[NFREF]" & ControlChars.CrLf)
            If t.ChaveNFE.Length > 0 Then

                If t.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC OrElse
                    t.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E OrElse
                    t.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E_OUT OrElse
                    t.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E_TOM OrElse
                    t.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC_SEM_NF Then

                    sb.Append("   REFCTE    =" & t.ChaveNFE & ControlChars.CrLf)
                Else
                    sb.Append("   REFNFE    =" & t.ChaveNFE & ControlChars.CrLf)
                End If
            Else
                sb.Append("[REFNF]" & ControlChars.CrLf)
                sb.Append("   CUF    =" & t.EstadoIbge & ControlChars.CrLf)
                Dim Movimento As DateTime = DateTime.Parse(t.Movimento)
                sb.Append("   AAMM   =" & Movimento.ToString("yyMM") & ControlChars.CrLf)
                sb.Append("   CNPJ   =" & t.CodigoEmpresa & ControlChars.CrLf)
                sb.Append("   MOD    = 01" & ControlChars.CrLf)
                sb.Append("   SERIE  =" & t.Serie & ControlChars.CrLf)
                sb.Append("   NNF    =" & t.Codigo & ControlChars.CrLf)
            End If

            sb.Append(ControlChars.CrLf)
        Next

        If nfe.ObrigaNFProdutor AndAlso Not nfe.SubOperacao.Classe = eClassesOperacoes.REAJUSTES AndAlso Not nfe.NotaTrocaOrigem Is Nothing Then
            sb.Append("[NFREF]" & ControlChars.CrLf)

            sb.Append("[REFNFP]" & ControlChars.CrLf)
            sb.Append("   CUF    =" & nfe.Cliente.Municipio.EstadoIbge & ControlChars.CrLf)
            sb.Append("   AAMM   =" & nfe.Movimento.ToString("yyMM") & ControlChars.CrLf)

            If nfe.Cliente.Codigo.Length = 11 Then
                sb.Append("   CPF    =" & nfe.Cliente.Codigo & ControlChars.CrLf)
            Else
                sb.Append("   CNPJ   =" & nfe.Cliente.Codigo & ControlChars.CrLf)
            End If

            sb.Append("   IE     = " & nfe.Cliente.InscricaoEstadual & ControlChars.CrLf)

            sb.Append("   MOD    = 04" & ControlChars.CrLf)

            If nfe.NotaTrocaOrigem.Serie.Contains("D") OrElse nfe.NotaTrocaOrigem.Serie.Contains("F") Then
                sb.Append("   SERIE  = 0" & ControlChars.CrLf)
            Else
                sb.Append("   SERIE  =" & nfe.NotaTrocaOrigem.Serie & ControlChars.CrLf)
            End If

            sb.Append("   NNF    =" & nfe.NotaTrocaOrigem.Codigo & ControlChars.CrLf)

            sb.Append(ControlChars.CrLf)
        End If

        'sb.Append("[NFREF]" & ControlChars.CrLf)
        'sb.Append("   REFNFE    =41210303189063000126550020000004281426269580" & ControlChars.CrLf)


        'Notas referenciais de exportação.
        'nfe.NotasReferenciais
        'Dim listreferenciais As New ListNotaFiscalReferencial(nfe, eTipoReferencial.NFC)
        'ALTEREI PARA PEGAR A LISTA - FURLAN - 22/07/2024
        If nfe.NotasReferenciais IsNot Nothing AndAlso nfe.NotasReferenciais.Count > 0 Then

            sb.Append("[NFREF]" & ControlChars.CrLf)

            For Each nfref In nfe.NotasReferenciais

                Dim objNFe As New [Lib].Negocio.NotaFiscal()
                objNFe.CodigoEmpresa = nfref.EmpresaReferencial_Id
                objNFe.EnderecoEmpresa = nfref.EndEmpresaReferencial_Id
                objNFe.CodigoCliente = nfref.ClienteReferencial_Id
                objNFe.EnderecoCliente = nfref.EndClienteReferencial_Id
                objNFe.EntradaSaida = IIf(nfref.EntradaSaida_Id = eEntradaSaida.Entrada, eEntradaSaida.Entrada, eEntradaSaida.Saida)
                objNFe.Serie = nfref.Serie_Id
                objNFe.Codigo = nfref.Nota_Id
                objNFe = New [Lib].Negocio.NotaFiscal(objNFe)

                If objNFe.ChaveNFE.Length > 0 Then
                    If nfref.ParentOrigem.NotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC OrElse
                        nfref.ParentOrigem.NotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E OrElse
                        nfref.ParentOrigem.NotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E_OUT OrElse
                        nfref.ParentOrigem.NotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E_TOM OrElse
                        nfref.ParentOrigem.NotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC_SEM_NF Then

                        sb.Append("   REFCTE    =" & objNFe.ChaveNFE & ControlChars.CrLf)
                    Else
                        sb.Append("   REFNFE    =" & objNFe.ChaveNFE & ControlChars.CrLf)
                    End If
                Else
                    sb.Append("[REFNF]" & ControlChars.CrLf)
                    sb.Append("   CUF    =" & nfref.ParentOrigem.NotaFiscal.Empresa.Municipio.EstadoIbge & ControlChars.CrLf)
                    sb.Append("   AAMM   =" & nfref.ParentOrigem.NotaFiscal.Movimento.ToString("yyMM") & ControlChars.CrLf)
                    sb.Append("   CNPJ   =" & nfref.ParentOrigem.NotaFiscal.CodigoEmpresa & ControlChars.CrLf)
                    sb.Append("   MOD    = 01" & ControlChars.CrLf)
                    sb.Append("   SERIE  =" & nfref.ParentOrigem.NotaFiscal.Serie & ControlChars.CrLf)
                    sb.Append("   NNF    =" & nfref.ParentOrigem.NotaFiscal.Codigo & ControlChars.CrLf)
                End If
            Next

            sb.Append(ControlChars.CrLf)
        End If
        '**********************************

        'DANFE - impressão
        'EMAIL - Envio de Email.

        'Emitente
        sb.Append("[EMIT]" & ControlChars.CrLf)
        sb.Append("   CNPJ   = " & nfe.CodigoEmpresa & ControlChars.CrLf)
        sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(nfe.Empresa.Nome) & ControlChars.CrLf)
        sb.Append("   IE     = " & nfe.Empresa.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
        'INCLUIDO PARA BAXI DE APUCARANA PARA DESOBRIGAR EMISSÃO DA GUIA DE ICMS POR NOTA - FURLAN - 29/07/2024
        'ACRESCENDO REGRAS PARA SP E RJ - FURLAN - 29/07/2024
        If nfe.CodigoEmpresa = "40938762000239" AndAlso nfe.Cliente.CodigoEstado = "PR" AndAlso (nfe.CFOP.Codigo = 5401 OrElse nfe.CFOP.Codigo = 5403 OrElse nfe.CFOP.Codigo = 5910) Then
            sb.Append("   IEST   = " & nfe.Empresa.SubstitutoTributario.Where(Function(s) s.Estado_Id = nfe.Cliente.CodigoEstado).FirstOrDefault.IESubstitutoTributario & ControlChars.CrLf)
        ElseIf nfe.CodigoEmpresa = "40938762000239" AndAlso nfe.Cliente.CodigoEstado = "SP" AndAlso (nfe.CFOP.Codigo = 6401 OrElse nfe.CFOP.Codigo = 6403 OrElse nfe.CFOP.Codigo = 6910) Then
            sb.Append("   IEST   = " & nfe.Empresa.SubstitutoTributario.Where(Function(s) s.Estado_Id = nfe.Cliente.CodigoEstado).FirstOrDefault.IESubstitutoTributario & ControlChars.CrLf)
        ElseIf nfe.CodigoEmpresa = "40938762000239" AndAlso nfe.Cliente.CodigoEstado = "RJ" AndAlso (nfe.CFOP.Codigo = 6401 OrElse nfe.CFOP.Codigo = 6403 OrElse nfe.CFOP.Codigo = 6910) Then
            sb.Append("   IEST   = " & nfe.Empresa.SubstitutoTributario.Where(Function(s) s.Estado_Id = nfe.Cliente.CodigoEstado).FirstOrDefault.IESubstitutoTributario & ControlChars.CrLf)
        End If
        sb.Append("   CRT    = " & nfe.Empresa.Empresa.Crt & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)

        sb.Append("[ENDEREMIT]" & ControlChars.CrLf)
        sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(nfe.Empresa.Endereco) & ControlChars.CrLf)
        sb.Append("   NRO	  = " & nfe.Empresa.Numero & ControlChars.CrLf)
        sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(nfe.Empresa.Complemento) & ControlChars.CrLf)
        sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(nfe.Empresa.Bairro) & ControlChars.CrLf)
        sb.Append("   CMUN    = " & nfe.Empresa.Municipio.EstadoIbge & nfe.Empresa.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)
        sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(nfe.Empresa.Cidade) & ControlChars.CrLf)
        sb.Append("   UF	  = " & nfe.Empresa.CodigoEstado & ControlChars.CrLf)
        sb.Append("   CEP	  = " & nfe.Empresa.CEP.Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        sb.Append("   CPAIS   = " & nfe.Empresa.CodigoPais & ControlChars.CrLf)
        sb.Append("   XPAIS   = " & nfe.Empresa.Pais.Descricao & ControlChars.CrLf)
        Dim Fone As String = nfe.Empresa.Telefone.Replace("(", "").Replace(")", "").Replace(" ", "").Replace("-", "").Replace("/", "")
        If Fone.Length > 0 Then sb.Append("   FONE   =" & Mid(Fone, 1, 10) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        'Destinatario
        sb.Append("[DEST]" & ControlChars.CrLf)
        If Not HomologAlfasig Then
            If nfe.Cliente.CodigoEstado = "EX" Then
                sb.Append("   IDESTRANGEIRO   = " & ControlChars.CrLf)
                sb.Append("   INDIEDEST  =9" & ControlChars.CrLf)
            ElseIf nfe.CodigoCliente.Length > 11 Then
                sb.Append("   CNPJ   =" & nfe.CodigoCliente & ControlChars.CrLf)
                If nfe.Cliente.InscricaoEstadual.Length > 0 AndAlso Not nfe.Cliente.InscricaoEstadual.ToString.Contains("ISENTO") Then
                    sb.Append("   IE	  =" & nfe.Cliente.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                    sb.Append("   INDIEDEST  =1" & ControlChars.CrLf)
                Else
                    sb.Append("   INDIEDEST  =9" & ControlChars.CrLf)
                End If
            Else
                sb.Append("   CPF    =" & nfe.CodigoCliente & ControlChars.CrLf)
                If nfe.Cliente.InscricaoEstadual.Length > 0 Then
                    If nfe.Cliente.InscricaoEstadual.ToString.Contains("ISENTO") Then
                        sb.Append("   INDIEDEST  =9" & ControlChars.CrLf)
                    Else
                        sb.Append("   IE	  =" & nfe.Cliente.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                        sb.Append("   INDIEDEST  =1" & ControlChars.CrLf)
                    End If
                Else
                    sb.Append("   INDIEDEST  =2" & ControlChars.CrLf)
                End If
            End If

            sb.Append("   XNOME   =" & Funcoes.EliminarCaracteresEspeciais(nfe.Cliente.Nome.ToString.ToUpper) & ControlChars.CrLf)
        Else
            'If nfe.Cliente.CodigoEstado = "GO" OrElse nfe.Cliente.CodigoEstado = "MS" Then

            If nfe.Cliente.CodigoEstado = "EX" Then
                sb.Append("   IDESTRANGEIRO   = " & ControlChars.CrLf)
                sb.Append("   INDIEDEST  =9" & ControlChars.CrLf)
            ElseIf nfe.CodigoCliente.Length > 11 Then
                sb.Append("   CNPJ   =" & nfe.CodigoCliente & ControlChars.CrLf)
                If nfe.Cliente.InscricaoEstadual.Length > 0 AndAlso Not nfe.Cliente.InscricaoEstadual.ToString.Contains("ISENTO") Then
                    sb.Append("   IE	  = " & nfe.Cliente.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                    sb.Append("   INDIEDEST  =1" & ControlChars.CrLf)
                Else
                    sb.Append("   INDIEDEST  =9" & ControlChars.CrLf)
                End If
            Else
                sb.Append("   CPF    =" & nfe.CodigoCliente & ControlChars.CrLf)
                If nfe.Cliente.InscricaoEstadual.Length > 0 Then
                    If nfe.Cliente.InscricaoEstadual.ToString.Contains("ISENTO") Then
                        sb.Append("   INDIEDEST  =9" & ControlChars.CrLf)
                    Else
                        sb.Append("   IE	  =" & nfe.Cliente.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                        sb.Append("   INDIEDEST  =1" & ControlChars.CrLf)
                    End If
                Else
                    sb.Append("   INDIEDEST  =2" & ControlChars.CrLf)
                End If
            End If

            sb.Append("   XNOME  = NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
            'Else
            '    sb.Append("   CNPJ   = 99999999000191" & ControlChars.CrLf)
            '    sb.Append("   IE     =" & ControlChars.CrLf)
            '    sb.Append("   INDIEDEST  =2" & ControlChars.CrLf)
            '    sb.Append("   XNOME  = NF-E EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
            'End If
        End If

        If nfe.Cliente.EmailNFE IsNot Nothing AndAlso nfe.Cliente.EmailNFE.Length > 0 Then
            If nfe.Cliente.EmailNFE.Contains(";") Then
                Dim emailCli() As String = nfe.Cliente.EmailNFE.Split(";")
                sb.Append("  EMAIL   =" & emailCli(0) & ControlChars.CrLf)
            Else
                sb.Append("  EMAIL   =" & nfe.Cliente.EmailNFE & ControlChars.CrLf)
            End If
        End If

        sb.Append(ControlChars.CrLf)

        'Endereço do Destinatário
        sb.Append("[ENDERDEST]" & ControlChars.CrLf)
        sb.Append("   XLGR   =" & Funcoes.EliminarCaracteresEspeciais(nfe.Cliente.Endereco.ToString.ToUpper) & ControlChars.CrLf)
        sb.Append("   NRO	  =" & nfe.Cliente.Numero & ControlChars.CrLf)

        If nfe.Cliente.Complemento.Length = 0 Then
            sb.Append("   XCPL   =." & ControlChars.CrLf)
        Else
            sb.Append("   XCPL   =" & Funcoes.EliminarCaracteresEspeciais(nfe.Cliente.Complemento.ToString.ToUpper) & ControlChars.CrLf)
        End If

        If nfe.Cliente.CodigoEstado = "EX" Then
            sb.Append("   CMUN   =9999999" & ControlChars.CrLf)
            sb.Append("   XBAIRRO=" & Funcoes.EliminarCaracteresEspeciais(nfe.Cliente.Cidade.ToString.ToUpper) & ControlChars.CrLf)
            sb.Append("   XMUN   = " & Funcoes.EliminarCaracteresEspeciais(nfe.Cliente.Cidade.ToString.ToUpper) & ControlChars.CrLf)
        Else
            If nfe.Cliente.Bairro.Length = 0 Then
                sb.Append("   XBAIRRO=." & ControlChars.CrLf)
            Else
                sb.Append("   XBAIRRO=" & Funcoes.EliminarCaracteresEspeciais(nfe.Cliente.Bairro.ToString.ToUpper) & ControlChars.CrLf)
            End If
            sb.Append("   CMUN   = " & nfe.Cliente.Municipio.EstadoIbge & nfe.Cliente.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)
            sb.Append("   XMUN   =" & Funcoes.EliminarCaracteresEspeciais(nfe.Cliente.Cidade.ToString.ToUpper) & ControlChars.CrLf)
            sb.Append("   CEP	  =" & nfe.Cliente.CEP.Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        End If

        sb.Append("   UF	  =" & nfe.Cliente.CodigoEstado & ControlChars.CrLf)
        sb.Append("   CPAIS  =" & nfe.Cliente.CodigoPais & ControlChars.CrLf)
        sb.Append("   XPAIS  =" & nfe.Cliente.Pais.Descricao & ControlChars.CrLf)
        Fone = nfe.Cliente.Telefone.Replace("(", "").Replace(")", "").Replace(" ", "").Replace("-", "").Replace("/", "")
        If Fone.Length > 0 Then sb.Append("   FONE   =" & Mid(Fone, 1, 10) & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)


        'Entrega - Solicitação Verde em 13/08/2025 - Furlan
        If nfe.CodigoEntrega.Length > 0 Then
            sb.Append("[ENTREGA]" & ControlChars.CrLf)

            If nfe.CodigoEntrega.Length > 11 Then
                sb.Append("   CNPJ   = " & nfe.CodigoEntrega & ControlChars.CrLf)
            Else
                sb.Append("   CPF   = " & nfe.CodigoEntrega & ControlChars.CrLf)
            End If

            sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(nfe.Entrega.Nome) & ControlChars.CrLf)
            sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(nfe.Entrega.Endereco) & ControlChars.CrLf)
            sb.Append("   NRO	  = " & nfe.Entrega.Numero & ControlChars.CrLf)
            sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(nfe.Entrega.Complemento) & ControlChars.CrLf)
            sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(nfe.Entrega.Bairro) & ControlChars.CrLf)
            sb.Append("   CMUN    = " & nfe.Entrega.Municipio.EstadoIbge & nfe.Entrega.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)
            sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(nfe.Entrega.Cidade) & ControlChars.CrLf)
            sb.Append("   UF	  = " & nfe.Entrega.CodigoEstado & ControlChars.CrLf)
            sb.Append("   CEP	  = " & nfe.Entrega.CEP.Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            sb.Append("   CPAIS   = " & nfe.Entrega.CodigoPais & ControlChars.CrLf)
            sb.Append("   XPAIS   = " & nfe.Entrega.Pais.Descricao & ControlChars.CrLf)
            Fone = nfe.Entrega.Telefone.Replace("(", "").Replace(")", "").Replace(" ", "").Replace("-", "").Replace("/", "")
            If Fone.Length > 0 Then sb.Append("   FONE   =" & Mid(Fone, 1, 10) & ControlChars.CrLf)

            If nfe.Entrega.EmailNFE IsNot Nothing AndAlso nfe.Entrega.EmailNFE.Length > 0 Then
                If nfe.Entrega.EmailNFE.Contains(";") Then
                    Dim emailCli() As String = nfe.Entrega.EmailNFE.Split(";")
                    sb.Append("  EMAIL   =" & emailCli(0) & ControlChars.CrLf)
                Else
                    sb.Append("  EMAIL   =" & nfe.Entrega.EmailNFE & ControlChars.CrLf)
                End If
            End If

            If nfe.Entrega.InscricaoEstadual.Length > 0 Then
                sb.Append("   IE     = " & nfe.Entrega.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
            End If

            sb.Append(ControlChars.CrLf)
        End If

        If nfe.Empresa.CodigoEstado.Equals("BA") Then
            sb.Append("[autXML]" & ControlChars.CrLf)
            sb.Append("   CNPJ   =13937073000156" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        End If

        If Left(nfe.CodigoEmpresa, 8) = "05366261" OrElse
            Left(nfe.CodigoEmpresa, 8) = "38198213" OrElse
            Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse
            Left(nfe.CodigoEmpresa, 8) = "49673784" OrElse
            Left(nfe.CodigoEmpresa, 8) = "53267147" Then
            sb.Append("[autXML]" & ControlChars.CrLf)
            sb.Append("   CNPJ   =41459558000117" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[autXML]" & ControlChars.CrLf)
            sb.Append("   CNPJ   =50161164000179" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[autXML]" & ControlChars.CrLf)
            sb.Append("   CNPJ   =49359541000108" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        End If

        'Dim j As Integer = 1

        Dim Valor As String = ""
        Dim ValorBruto As Decimal = 0
        Dim ValorLiquido As Decimal = 0

        Dim SituacaoTributaria As Integer = 0
        Dim SituacaoTributariaIPI As Integer = 0
        Dim SituacaoTributariaPISCOFINS As Integer = 0
        Dim SituacaoTributariaIBSCBS As Integer = 0
        Dim ClassificacaoIBSCBS As Integer
        Dim ReducaoIBS_Perc As Decimal = 0
        Dim ReducaoCBS_Perc As Decimal = 0

        Dim BaseIcms As String = "0.00"
        Dim TotalBaseIcms As Decimal = 0
        Dim AliIcms As String = ""
        Dim ValorIcms As String = ""
        Dim ValIcms As Decimal = 0
        Dim TotalValIcms As Decimal = 0

        Dim BaseIcmsST As String = "0.00"
        Dim TotalBaseIcmsST As Decimal = 0
        Dim AliIcmsST As String = ""
        Dim ValorIcmsST As String = ""
        Dim TotalValIcmsST As Decimal = 0

        Dim valorFPobresa As Decimal = 0
        Dim TotalValorFPobresa As Decimal = 0

        Dim ValorIPI As String = "0.00"
        Dim ValIPI As Decimal = 0
        Dim TotalValIPI As Decimal = 0
        Dim BaseIPI As String = ""
        Dim AliIPI As String = ""

        Dim BasePis As String = ""
        Dim AliPis As String = ""
        Dim ValorPis As String = ""
        Dim ValPis As Decimal = 0
        Dim TotalValPis As Decimal = 0

        Dim BaseCofins As String = ""
        Dim AliCofins As String = ""
        Dim ValorCofins As String = ""
        Dim ValCofins As Decimal = 0
        Dim TotalValCofins As Decimal = 0

        Dim BaseIBS As String = ""
        Dim AliIBS As String = ""
        Dim ValorIBS As String = ""
        Dim AliIBSOri As Decimal = 0
        Dim ValIBS As Decimal = 0
        Dim TotalBaseIBS As Decimal = 0
        Dim TotalValIBS As Decimal = 0

        Dim BaseCBS As String = ""
        Dim AliCBS As String = ""
        Dim ValorCBS As String = ""
        Dim AliCBSOri As Decimal = 0
        Dim ValCBS As Decimal = 0
        Dim TotalBaseCBS As Decimal = 0
        Dim TotalValCBS As Decimal = 0

        Dim bTemIBSCBS As Boolean = False

        Dim ValOutros As Decimal = 0
        Dim ValFrete As Decimal = 0
        Dim ValSeguro As Decimal = 0
        Dim ValDesconto As Decimal = 0

        Dim TotalValOutros As Decimal = 0
        Dim TotalValFrete As Decimal = 0
        Dim TotalValSeguro As Decimal = 0
        Dim TotalValDesconto As Decimal = 0
        Dim BaseAduaneira As Decimal = 0
        Dim ValAduaneira As Decimal = 0

        Dim ICMSDIFERENCIAL As Decimal = 0
        Dim TOTALICMSDIFERENCIAL As Decimal = 0

        Dim IcmsDESONERADO51 As Boolean = False
        Dim ValIcmsDESONERADO As Decimal = 0
        Dim TotalValIcmsDESONERADO As Decimal = 0

        For Each item In nfe.Itens

            'DET ateração para o PROD
            sb.Append("[DET]" & ControlChars.CrLf)

            Dim msgimposto As String = Funcoes.EliminarCaracteresEspeciais(item.Encargos.MensagemImpostos)
            Dim msgprd As String = Funcoes.EliminarCaracteresEspeciais(item.ObservacoesDoProduto)

            Dim desPRD As String = String.Empty

            If nfe.Empresa.Empresa.UsarRegistroMinAgr AndAlso item.Produto.RegistroMinisterioAgricultura.Length > 0 Then
                desPRD = Funcoes.EliminarCaracteresEspeciais(item.Produto.Descricao)

                If msgprd.Length > 0 Then
                    desPRD = desPRD & ". " & msgprd
                End If

                If msgimposto.Length > 0 Then
                    desPRD = desPRD & ". " & msgimposto
                End If

                sb.Append("   INFADPROD=" & desPRD & ControlChars.CrLf)

            ElseIf nfe.Empresa.Empresa.UsarDescricaoProduto AndAlso Not item.Produto.Nome = item.Produto.Descricao Then
                desPRD = Funcoes.EliminarCaracteresEspeciais(item.Produto.Descricao)

                If msgprd.Length > 0 Then
                    desPRD = desPRD & ". " & msgprd
                End If

                If msgimposto.Length > 0 Then
                    desPRD = desPRD & ". " & msgimposto
                End If

                sb.Append("   INFADPROD=" & desPRD & ControlChars.CrLf)
            Else
                If Not item.ObservacoesDoProduto Is Nothing AndAlso item.ObservacoesDoProduto.Length > 0 Then
                    desPRD = Funcoes.EliminarCaracteresEspeciais(item.ObservacoesDoProduto)
                End If

                If Not item.Encargos.MensagemImpostos Is Nothing AndAlso item.Encargos.MensagemImpostos.Length > 0 Then
                    If desPRD.Length > 0 Then
                        desPRD = desPRD & ". " & Funcoes.EliminarCaracteresEspeciais(item.Encargos.MensagemImpostos)
                    Else
                        desPRD = Funcoes.EliminarCaracteresEspeciais(item.Encargos.MensagemImpostos)
                    End If
                End If

                sb.Append("   INFADPROD=" & desPRD & ControlChars.CrLf)
            End If

            'If Not item.Produto.ProdutoAgrupador Is Nothing AndAlso item.Produto.ProdutoAgrupador.Count > 0 Then
            '    sb.Append("   INFADPROD=" & item.Produto.Descricao & "|" & msgprd & IIf(msgprd.Length > 0 And msgimposto.Length > 0, " / ", "") & msgimposto & ControlChars.CrLf)
            'Else
            '    sb.Append("   INFADPROD=" & msgprd & IIf(msgprd.Length > 0 And msgimposto.Length > 0, " / ", "") & msgimposto & ControlChars.CrLf)
            'End If
            sb.Append(ControlChars.CrLf)

            sb.Append("[PROD]" & ControlChars.CrLf)

            If Not item.Produto.ProdutoAgrupador Is Nothing AndAlso item.Produto.ProdutoAgrupador.Count > 0 Then
                sb.Append("   CPROD   =" & item.Produto.ProdutoAgrupador(0).CodigoProdutoAgrupado & ControlChars.CrLf)
                sb.Append("   XPROD   =" & Funcoes.EliminarCaracteresEspeciais(item.Produto.ProdutoAgrupador(0).NomeProduto) & " " & ControlChars.CrLf)
            Else
                sb.Append("   CPROD   =" & item.CodigoProduto & ControlChars.CrLf)

                'Dim desPRD As String = String.Empty

                'If nfe.Empresa.Empresa.UsarRegistroMinAgr AndAlso item.Produto.RegistroMinisterioAgricultura.Length > 0 Then
                '    desPRD = Funcoes.EliminarCaracteresEspeciais(item.Produto.Nome) & "-" & Funcoes.EliminarCaracteresEspeciais(item.Produto.Descricao)
                'ElseIf nfe.Empresa.Empresa.UsarDescricaoProduto AndAlso Not item.Produto.Nome = item.Produto.Descricao Then
                '    desPRD = Funcoes.EliminarCaracteresEspeciais(item.Produto.Nome) & "-" & Funcoes.EliminarCaracteresEspeciais(item.Produto.Descricao)
                'Else
                '    desPRD = Funcoes.EliminarCaracteresEspeciais(item.Produto.Nome)
                'End If

                sb.Append("   XPROD   =" & Funcoes.EliminarCaracteresEspeciais(item.Produto.Nome) & " " & ControlChars.CrLf)

            End If

            Dim varCean As String = "SEM GTIN"
            Dim varCeanTrib As String = "SEM GTIN"

            ' Criar uma lista das variáveis não vazias
            Dim listaGTIN As List(Of String) = New List(Of String) From {item.Produto.Gtin8, item.Produto.Gtin12, item.Produto.Gtin13, item.Produto.Gtin14}
            listaGTIN = listaGTIN.Where(Function(gtin) Not String.IsNullOrEmpty(gtin)).ToList()

            If listaGTIN.Where(Function(x) x.Length >= 8).Count > 0 Then
                ' Identificar o maior e o menor GTIN com base no comprimento
                varCean = listaGTIN.OrderByDescending(Function(gtin) gtin.Length).First()
                varCeanTrib = listaGTIN.OrderBy(Function(gtin) gtin.Length).First()
            End If

            'Se tiver apenas um GTIN, informamos apenas o CEANTRIB
            'Se tiver mais de um, o GTIN principal ou seja o código mais abrangente, será o CEAN e o codigo menor entre os 2 será o CEANTRIB
            'Conforme documentação Alfasig o CEANTRIB é informado mais abaixo
            sb.Append(String.Format("   CEAN={0}", varCean) & ControlChars.CrLf)

            sb.Append("   NCM	  =" & item.Produto.NCM.Replace(".", "") & ControlChars.CrLf)

            sb.Append("   CFOP    =" & item.CFOP & ControlChars.CrLf)

            Dim undPedido As String = String.Empty

            For Each itemPed In nfe.Pedido.Itens
                If itemPed.CodigoProduto = item.CodigoProduto Then
                    undPedido = itemPed.CodigoUnidadeComercializacao
                    Exit For
                End If
            Next

            If nfe.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                sb.Append("   UCOM    =KGS" & ControlChars.CrLf)
            Else
                'sb.Append("   UCOM    =" & item.Produto.Unidade & ControlChars.CrLf)

                sb.Append("   UCOM    =" & undPedido & ControlChars.CrLf)
            End If

            If Not item.Produto.Unidade = "TON" AndAlso undPedido = "TON" Then
                Valor = (item.QuantidadeFiscal / 1000).ToString("N4")
            Else
                Valor = item.QuantidadeFiscal.ToString("N4")
            End If

            'versão 1.22 da NT 2013/005 
            'Recebemos um email dizendo que essa validação estava sendo feita incorretamente pelo sefaz nas Finalidadess nfe's (Campo FinNfe) <> 1
            'If CDec(Valor) = 0 Then Valor = "1"
            sb.Append("   QCOM    =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

            If nfe.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                If item.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = item.Produto.Unidade).FirstOrDefault.FatorConversao = 1000 Then
                    Valor = Math.Round((item.ValorTotal / item.QuantidadeFiscal), 10, MidpointRounding.AwayFromZero)
                Else
                    Valor = item.Unitario.ToString("N10")
                End If
            Else
                If Not item.Produto.Unidade = "TON" AndAlso undPedido = "TON" Then
                    Valor = CDec((item.Unitario * 1000)).ToString("N10")

                Else
                    Valor = CDec(item.Unitario).ToString("N10")
                End If
            End If

            'versão 1.22 da NT 2013/005 
            'Recebemos um email dizendo que essa validação estava sendo feita incorretamente pelo sefaz nas Finalidadess nfe's (Campo FinNfe) <> 1
            'If CDec(Valor) = 0 Then Valor = CDec(item.ValorTotal).ToString("N10")
            sb.Append("   VUNCOM  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            '-----
            Valor = item.ValorTotal.ToString("N2")
            sb.Append("   VPROD   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)


            sb.Append(String.Format("   CEANTRIB={0}", varCeanTrib) & ControlChars.CrLf)

            'If (nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.Cliente.CodigoEstado = "EX") OrElse (item.CFOP = 1501 OrElse item.CFOP = 2501) Then
            'estava vendo talvez não precise
            'End If

            'SOLICITAÇÃO VIA E-MAIL EM 14/01/2021 - JONATHAN
            'If (Left(nfe.Empresa.Codigo, 8) = "05366261" Or Left(nfe.Empresa.Codigo, 8) = "38198213" Or Left(nfe.Empresa.Codigo, 8) = "40938762") Then
            If nfe.Pedido.XPedNFe.Length > 0 Then
                sb.Append("   xPed    =" & nfe.Pedido.XPedNFe & ControlChars.CrLf)

                If nfe.Pedido.ItemXPedNFe.Length > 0 Then
                    sb.Append("   nItemPed=" & nfe.Pedido.ItemXPedNFe & ControlChars.CrLf)
                Else
                    sb.Append("   nItemPed=" & item.Sequencia & ControlChars.CrLf)
                End If

                sb.Append(ControlChars.CrLf)
            End If
            'End If

            '-----

            'VERIFICA ENCARGOS DO ITEM DA NOTA

            BaseIcms = "0.00"
            ValorIcms = "0.00"
            AliIcms = "0.00"
            ValIcms = 0

            ValOutros = 0.0
            ValFrete = 0.0
            ValSeguro = 0.0
            ValDesconto = 0.0
            BaseAduaneira = 0.0
            ValAduaneira = 0.0

            BasePis = "0.00"
            AliPis = "0.00"
            ValorPis = "0.00"
            BaseCofins = "0.00"
            AliCofins = "0.00"
            ValorCofins = "0.00"

            ' >>> Reseta bases/valores IBS/CBS por item
            BaseCBS = "0.00"
            AliCBS = "0.00"
            ValorCBS = "0.00"
            ValCBS = 0

            BaseIBS = "0.00"
            AliIBS = "0.00"
            ValorIBS = "0.00"
            ValIBS = 0
            ' <<< IBS/CBS

            'SUBSTITUIÇÃO TRIBUTÁRIA - FURLAN - 20/10/2023
            BaseIcmsST = "0.00"
            AliIcmsST = "0.00"
            ValorIcmsST = "0.00"
            '---------------------------------------------
            BaseIPI = "0.00"
            AliIPI = "0.00"
            ValorIPI = 0.0
            ValIcmsDESONERADO = 0.0
            ValIPI = 0.0

            valorFPobresa = 0

            ICMSDIFERENCIAL = 0
            Dim BaseIcmsDIFERENCIAL As String = "0.00"
            Dim AliIcmsDIFERENCIAL As String = "0.00"

            Dim Operacao As Integer = 0

            Dim objEncargoXTaxa As New EncargoXTaxa()

            For Each enc In item.Encargos
                Operacao = enc.CodigoOperacao

                SituacaoTributariaIBSCBS = enc.SituacaoTributariaIBSCBS
                ClassificacaoIBSCBS = enc.ClassificacaoIBSCBS

                If enc.Codigo = "LIQUIDO" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo = "PRODUTO" Then
                    ValorBruto += enc.Valor
                    SituacaoTributaria = enc.SituacaoTributaria
                    SituacaoTributariaIPI = enc.SituacaoTributariaIPI
                    SituacaoTributariaPISCOFINS = enc.SituacaoTributariaPISCOFINS
                ElseIf (enc.Codigo = "ICMS" OrElse enc.Codigo = "ICMS A REC.") Then
                    BaseIcms = enc.Base
                    TotalBaseIcms += enc.Base
                    If enc.PercentualExibicao > 0 Then
                        AliIcms = enc.PercentualExibicao
                    Else
                        AliIcms = enc.Percentual
                    End If
                    ValIcms = enc.Valor
                    TotalValIcms += enc.Valor

                    If SituacaoTributaria = 20 OrElse SituacaoTributaria = 120 OrElse SituacaoTributaria = 620 Then
                        ValIcmsDESONERADO = Math.Round(((ValorBruto * enc.Percentual) / 100) - ValIcms, 2, MidpointRounding.AwayFromZero)
                    End If

                ElseIf enc.Codigo = "FUNDO FECP" Then
                    valorFPobresa = enc.Valor
                    TotalValorFPobresa += enc.Valor
                ElseIf enc.Codigo = "ICMS-ST" Then
                    BaseIcmsST = enc.Base
                    TotalBaseIcmsST += enc.Base

                    If enc.PercentualExibicao > 0 Then
                        AliIcmsST = enc.PercentualExibicao
                    Else
                        AliIcmsST = enc.Percentual
                    End If
                    ValorIcmsST = enc.Valor
                    TotalValIcmsST += enc.Valor
                ElseIf enc.Codigo = "ICMS DIFERENCIAL" Then
                    BaseIcmsDIFERENCIAL = enc.Base
                    AliIcmsDIFERENCIAL = enc.PercentualExibicao
                    ICMSDIFERENCIAL = enc.Valor
                    TOTALICMSDIFERENCIAL += enc.Valor

                ElseIf enc.Codigo.Contains("DESONERADO") Then
                    BaseIcms = enc.Base
                    TotalBaseIcms += enc.Base
                    If enc.PercentualExibicao > 0 Then
                        AliIcms = enc.PercentualExibicao
                    Else
                        AliIcms = enc.Percentual
                    End If
                    ValIcms = enc.Valor
                    ValIcmsDESONERADO = enc.Valor

                    If enc.SituacaoTributaria = 51 Then IcmsDESONERADO51 = True

                ElseIf enc.Codigo.Trim = "IPI" OrElse enc.Codigo.Trim = "IPI A RECUP." Then
                    BaseIPI = enc.Base
                    AliIPI = enc.Percentual
                    ValorIPI = enc.Valor
                    If enc.OperacaoEncargo.Sinal = "+" Then
                        ValIPI = enc.Valor
                        TotalValIPI += enc.Valor
                    End If
                ElseIf enc.Codigo = "PIS" Then
                    BasePis = enc.Base
                    AliPis = enc.Percentual
                    ValorPis = enc.Valor
                    ValPis = enc.Valor
                    TotalValPis += enc.Valor
                ElseIf enc.Codigo = "COFINS" Then
                    BaseCofins = enc.Base
                    AliCofins = enc.Percentual
                    ValorCofins = enc.Valor
                    ValCofins = enc.Valor
                    TotalValCofins += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("FACS") AndAlso enc.OperacaoEncargo.Sinal = "-" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("FETHAB") AndAlso enc.OperacaoEncargo.Sinal = "-" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("IAGRO") AndAlso enc.OperacaoEncargo.Sinal = "-" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("FUNRURAL") AndAlso enc.OperacaoEncargo.Sinal = "-" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("SENAR") AndAlso enc.OperacaoEncargo.Sinal = "-" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("FUNDERSUL") AndAlso enc.OperacaoEncargo.Sinal = "-" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("FUNDEMS") AndAlso enc.OperacaoEncargo.Sinal = "-" Then
                    ValorLiquido += enc.Valor
                ElseIf enc.Codigo.Contains("FRETES") Then
                    ValFrete = enc.Valor
                    TotalValFrete += enc.Valor
                ElseIf enc.Codigo.Contains("SEGURO") Then
                    ValSeguro = enc.Valor
                    TotalValSeguro += enc.Valor
                ElseIf enc.Codigo.ToString.Contains("AFIXAR") Then
                    'NÃO FAZ NADA
                ElseIf enc.Codigo.Contains("DESCONTOS") Then
                    ValDesconto = enc.Valor
                    TotalValDesconto += enc.Valor
                ElseIf enc.Codigo.Contains("ADUANEIRAS") OrElse enc.Codigo.Contains("DESP. ACESSORIA") OrElse enc.Codigo.Contains("CUSTO PIS") OrElse enc.Codigo.Contains("CUSTO COFINS") Then
                    BaseAduaneira = enc.Base
                    ValAduaneira = enc.Valor
                    ValOutros = enc.Valor
                    TotalValOutros += enc.Valor
                ElseIf enc.Codigo = "IBS" Then
                    BaseIBS = enc.Base.ToString("N2").Replace(".", "").Replace(",", ".")

                    AliIBSOri = item.OperacaoEstado.Encargos.Where(Function(s) s.CodigoEncargo = enc.Codigo).FirstOrDefault.Aliquota
                    AliIBS = enc.Percentual.ToString("N4").Replace(".", "").Replace(",", ".")

                    ValorIBS = enc.Valor.ToString("N2").Replace(".", "").Replace(",", ".")
                    ValIBS = enc.Valor
                    TotalBaseIBS += enc.Base
                    TotalValIBS += enc.Valor



                    'If enc.ReducaoIBS_Perc > 0 Then
                    '    TotalValIBS -= (enc.Valor * enc.ReducaoIBS_Perc / 100D)
                    'End If

                    ReducaoIBS_Perc = enc.ReducaoIBS_Perc
                    ReducaoCBS_Perc = enc.ReducaoCBS_Perc
                    bTemIBSCBS = True
                ElseIf enc.Codigo = "CBS" Then
                    BaseCBS = enc.Base.ToString("N2").Replace(".", "").Replace(",", ".")

                    AliCBSOri = item.OperacaoEstado.Encargos.Where(Function(s) s.CodigoEncargo = enc.Codigo).FirstOrDefault.Aliquota
                    AliCBS = enc.Percentual.ToString("N4").Replace(".", "").Replace(",", ".")

                    ValorCBS = enc.Valor.ToString("N2").Replace(".", "").Replace(",", ".")
                    ValCBS = enc.Valor
                    TotalBaseCBS += enc.Base
                    TotalValCBS += enc.Valor

                    'If enc.ReducaoCBS_Perc > 0 Then
                    '    TotalValCBS -= (enc.Valor * enc.ReducaoCBS_Perc / 100D)
                    'End If

                    ReducaoIBS_Perc = enc.ReducaoIBS_Perc
                    ReducaoCBS_Perc = enc.ReducaoCBS_Perc
                    bTemIBSCBS = True
                Else
                    If Not enc.Codigo = "CUSTO ICMS" Then
                        If enc.OperacaoEncargo.Sinal = "+" Then
                            ValOutros = ValOutros + enc.Valor
                            TotalValOutros += enc.Valor
                        ElseIf enc.OperacaoEncargo.Sinal = "-" Then
                            ValDesconto = ValDesconto + enc.Valor
                            TotalValDesconto += enc.Valor
                        End If
                    End If
                End If

            Next

            If nfe.EntradaSaida = eEntradaSaida.Saida AndAlso (item.CFOP = 5502 OrElse item.CFOP = 6502) AndAlso item.Produto.CodigoGrupo = "10101" Then
                undPedido = "TON"
            End If

            If nfe.EntradaSaida = eEntradaSaida.Saida AndAlso (nfe.Cliente.CodigoEstado = "EX" OrElse item.CFOP = 5502 OrElse item.CFOP = 6502) AndAlso undPedido = "TON" Then
                If Left(nfe.CodigoEmpresa, 8) = "24450490" AndAlso
                    item.CFOP = 7102 AndAlso
                    (item.Produto.CodigoGrupo = "20101" OrElse
                     item.Produto.CodigoGrupo = "20102" OrElse
                     item.Produto.CodigoGrupo = "20103" OrElse
                     item.Produto.CodigoGrupo = "20104" OrElse
                     item.Produto.CodigoGrupo = "20105" OrElse
                     item.Produto.CodigoGrupo = "20106") Then
                    undPedido = "KG"
                    sb.Append("   UTRIB   =" & undPedido & ControlChars.CrLf)
                ElseIf Left(nfe.CodigoEmpresa, 8) = "05366261" AndAlso item.Produto.CodigoGrupo = "10102" Then
                    undPedido = "KG"
                    sb.Append("   UTRIB   =" & undPedido & ControlChars.CrLf)
                Else
                    sb.Append("   UTRIB   =" & undPedido & ControlChars.CrLf)
                End If

                If undPedido = "TON" Then
                    Valor = (item.QuantidadeFiscal / 1000).ToString("N4")
                Else
                    Valor = item.QuantidadeFiscal.ToString("N4")
                End If

                sb.Append("   QTRIB   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                If undPedido = "TON" Then
                    Valor = CDec((item.Unitario * 1000)).ToString("N10")

                Else
                    Valor = CDec(item.Unitario).ToString("N10")
                End If

                sb.Append("   VUNTRIB =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                If Left(nfe.CodigoEmpresa, 8) = "24450490" AndAlso
                   item.CFOP = 7102 AndAlso
                   (item.Produto.CodigoGrupo = "20101" OrElse
                    item.Produto.CodigoGrupo = "20102" OrElse
                    item.Produto.CodigoGrupo = "20103" OrElse
                    item.Produto.CodigoGrupo = "20104" OrElse
                    item.Produto.CodigoGrupo = "20105" OrElse
                    item.Produto.CodigoGrupo = "20106") Then
                    For Each itemPed In nfe.Pedido.Itens
                        If itemPed.CodigoProduto = item.CodigoProduto Then
                            undPedido = itemPed.CodigoUnidadeComercializacao
                            Exit For
                        End If
                    Next
                End If
            Else
                If nfe.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    sb.Append("   UTRIB   =KGS" & ControlChars.CrLf)
                Else
                    sb.Append("   UTRIB   =" & undPedido & ControlChars.CrLf)
                End If

                If Not item.Produto.Unidade = "TON" AndAlso undPedido = "TON" Then
                    Valor = (item.QuantidadeFiscal / 1000).ToString("N4")
                Else
                    Valor = item.QuantidadeFiscal.ToString("N4")
                End If

                'versão 1.22 da NT 2013/005 
                'Recebemos um email dizendo que essa validação estava sendo feita incorretamente pelo sefaz nas Finalidadess nfe's (Campo FinNfe) <> 1
                'If CDec(Valor) = 0 Then Valor = "1"
                sb.Append("   QTRIB   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                '-----
                If nfe.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    If item.Produto.UnidadesDeComercializacao.Where(Function(s) s.CodigoUnidade = item.Produto.Unidade).FirstOrDefault.FatorConversao = 1000 Then
                        Valor = Math.Round((item.ValorTotal / item.QuantidadeFiscal), 10, MidpointRounding.AwayFromZero)
                    Else
                        Valor = item.Unitario.ToString("N10")
                    End If
                Else
                    If Not item.Produto.Unidade = "TON" AndAlso undPedido = "TON" Then
                        Valor = CDec((item.Unitario * 1000)).ToString("N10")

                    Else
                        Valor = CDec(item.Unitario).ToString("N10")
                    End If
                End If

                'versão 1.22 da NT 2013/005
                'Recebemos um email dizendo que essa validação estava sendo feita incorretamente pelo sefaz nas Finalidadess nfe's (Campo FinNfe) <> 1
                'If Valor = 0 Then Valor = CDec(item.ValorTotal).ToString("N10")
                sb.Append("   VUNTRIB =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

            End If

            '-------

            If ValFrete > 0 Then
                Valor = CDec(ValFrete).ToString("N2")
                sb.Append("   VFRETE  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            If ValSeguro > 0 Then
                Valor = CDec(ValSeguro).ToString("N2")
                sb.Append("   VSEG    =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            If ValDesconto > 0 Then
                Valor = CDec(ValDesconto).ToString("N2")
                sb.Append("   VDESC   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            If ValOutros > 0 Then
                Valor = CDec(ValOutros).ToString("N2")
                sb.Append("  VOUTRO   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            sb.Append("   INDTOT  =1" & ControlChars.CrLf)

            'If nfe.Empresa.CodigoEstado = "PR" AndAlso item.OperacaoEstado.CodigoBeneficio.Length > 0 Then
            '    sb.Append("   CBENEF  =" & item.OperacaoEstado.CodigoBeneficio & ControlChars.CrLf)
            'End If
            If item.OperacaoEstado.CodigoBeneficio.Length > 0 Then
                sb.Append("   CBENEF  =" & item.OperacaoEstado.CodigoBeneficio & ControlChars.CrLf)
            End If

            sb.Append(ControlChars.CrLf)


            'SOLICITAÇÃO VIA WHATS EM 14/04/2022 - JEFFERSON QUIMICA
            If item.Lotes IsNot Nothing AndAlso item.Lotes.Count > 0 Then

                If Left(nfe.Empresa.Codigo, 8) = "40938762" OrElse Left(nfe.Empresa.Codigo, 8) = "49673784" Then
                    'CASO VOLTE A MOSTRAR MEXER NO ucNFObsProduto EM Carregar informações do Lote
                    'SE FOR BAXI(FOODS E DISTRIBUIDORA) NÃO INFORMAR LOTE ATÉ SEGUNDA ORDEM, SOLICITAÇÃO DOUGLAS - FURLAN - 21/08/2024
                Else
                    For Each lt In item.Lotes
                        sb.Append("[RASTRO]" & ControlChars.CrLf)
                        sb.Append("   NLOTE   =" & lt.Lote & ControlChars.CrLf)
                        Valor = lt.Quantidade.ToString("N3")
                        sb.Append("   QLOTE   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                        sb.Append("    DFAB   =" & lt.Fabricado.ToString("yyyy-MM-dd") & ControlChars.CrLf)
                        sb.Append("    DVAL   =" & lt.Validade.ToString("yyyy-MM-dd") & ControlChars.CrLf)

                        sb.Append(ControlChars.CrLf)
                    Next
                End If

            End If

            'Acrescentei para saber se for Importação Origem da Mercadoria deve ser Estrangeira(1) - furlan - 23/07/2014
            Dim temImportacao As Boolean = False
            'Declaração de Importação
            If Not nfe.DadosDaImportacao Is Nothing AndAlso nfe.DadosDaImportacao.NumeroDeclaracaoImportacao.Length > 0 Then
                temImportacao = True

                sb.Append("[DI]" & ControlChars.CrLf)
                sb.Append("        NDI =" & nfe.DadosDaImportacao.NumeroDeclaracaoImportacao & ControlChars.CrLf)
                sb.Append("        DDI =" & nfe.DadosDaImportacao.DataDeclaracaoImportacao.ToString("yyyy-MM-dd") & ControlChars.CrLf)
                sb.Append(" XLOCDESEMB =" & nfe.DadosDaImportacao.LocalDesembarqueImportacao & ControlChars.CrLf)
                sb.Append("   UFDESEMB =" & nfe.DadosDaImportacao.EstadoDesembarqueImportacao & ControlChars.CrLf)
                sb.Append("    DDESEMB =" & nfe.DadosDaImportacao.DataDesembarqueImportacao.ToString("yyyy-MM-dd") & ControlChars.CrLf)
                'Campos Novos. 3G
                sb.Append("TPVIATRANSP =" & nfe.DadosDaImportacao.ViaDeTransporte & ControlChars.CrLf) 'OBRIGATORIO NA 4G - FURAN 28/03/2018
                sb.Append("TPINTERMEDIO=" & nfe.DadosDaImportacao.TipoDeImportacao & ControlChars.CrLf)
                If nfe.DadosDaImportacao.ViaDeTransporte = 1 Then
                    sb.Append("VAFRMM =" & nfe.DadosDaImportacao.ValorVAFRMM.ToString.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                End If
                sb.Append("CEXPORTADOR =" & nfe.CodigoCliente & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)

                sb.Append("[ADI]" & ControlChars.CrLf)
                sb.Append("   NADICAO =" & item.Sequencia & ControlChars.CrLf)
                sb.Append("  NSEQADIC =1" & ControlChars.CrLf)
                sb.Append("CFABRICANTE=" & nfe.DadosDaImportacao.CodigoFabricante & ControlChars.CrLf)
                If Not String.IsNullOrWhiteSpace(nfe.DadosDaImportacao.NumAtoConcessorio) Then
                    sb.Append("NDRAW=" & nfe.DadosDaImportacao.NumAtoConcessorio & ControlChars.CrLf)
                End If

                sb.Append(ControlChars.CrLf)
            End If

            'ERRO: Número do processo de drawback não informado na exportação
            'Obrigatória informação do número do processo de drawback para CFOP: 
            ' 7127: Venda de produção do estabelecimento sob o regime de drawback 
            ' 7211: Devolução de compras p/ industrialização sob o regime de drawback
            If nfe.CFOP.Codigo = 7127 OrElse nfe.CFOP.Codigo = 7211 Then
                sb.Append("[DETEXPORT]" & ControlChars.CrLf)
                sb.Append("   NDRAW = " & nfe.DadosDaExportacao.NumAtoConcessorio & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)
            End If

            'Impostos
            sb.Append("[IMPOSTO]" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
            'sb.Append("VTOTTRIB=" & ControlChars.CrLf) 'Total de impostos federais, estaduais e municipais.

            If SituacaoTributaria = 0 OrElse SituacaoTributaria = 100 OrElse SituacaoTributaria = 600 Then
                sb.Append("[ICMS00]" & ControlChars.CrLf)

                If SituacaoTributaria.ToString.Length = 3 OrElse SituacaoTributaria.ToString.Length = 6 Then
                    sb.Append("   ORIG   =" & Left(SituacaoTributaria, 1) & ControlChars.CrLf)
                ElseIf temImportacao Then
                    sb.Append("   ORIG   =1" & ControlChars.CrLf)
                Else
                    sb.Append("   ORIG   =0" & ControlChars.CrLf)
                End If

                If SituacaoTributaria = 0 Then
                    sb.Append("   CST    =0" & SituacaoTributaria & ControlChars.CrLf)
                ElseIf SituacaoTributaria.ToString.Length = 3 Then
                    sb.Append("   CST    =" & Mid(SituacaoTributaria, 2, 2) & ControlChars.CrLf)
                Else
                    sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)
                End If

                sb.Append("   MODBC  =3" & ControlChars.CrLf)
                sb.Append("   VBC    =" & BaseIcms.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliIcms).ToString("N2")
                sb.Append("   PICMS  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(ValIcms).ToString("N2")
                sb.Append("   VICMS  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                'ENTRA EM VIFOR À PARTIR DE 02/04/2018 - FURLAN
                'sb.Append("   PFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("   VFCP   =0.00" & ControlChars.CrLf)
            End If

            If SituacaoTributaria = 10 OrElse SituacaoTributaria = 110 Then
                sb.Append("[ICMS10]" & ControlChars.CrLf)
                If temImportacao Then
                    sb.Append("   ORIG   =1" & ControlChars.CrLf)
                Else
                    sb.Append("   ORIG   =0" & ControlChars.CrLf)
                End If
                sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)
                sb.Append("   MODBC  =3" & ControlChars.CrLf)
                sb.Append("   VBC    =" & BaseIcms.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliIcms).ToString("N2")
                sb.Append("   PICMS  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(ValIcms).ToString("N2")
                sb.Append("   VICMS  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                'Aliquota ICMS-ST EXEMPLOS
                '<pMVAST>71.9400</pMVAST>
                '<pMVAST>74.9400</pMVAST>
                '----------------------------------------------------------------------------------------------------------------------------------------
                'Rejeicao: Informada modalidade de determinacao da BC da ST diferente de MVA e informado o campo pMVAST [nItem: 1]. 
                'Orientações: Verifique se o campo MODBCST é diferente de “4” Margem Valor Agregado, pois então não deverá ser preenchido o campo PMVAST.
                'sb.Append("   MODBCST=0" & ControlChars.CrLf)
                '----------------------------------------------------------------------------------------------------------------------------------------
                sb.Append("   MODBCST=4" & ControlChars.CrLf)

                If nfe.Cliente.CodigoEstado = "MG" Then
                    objEncargoXTaxa.SelecionarVigente(nfe.Cliente.CodigoEstado, "ICMS ST MG", nfe.Movimento, "")
                Else
                    objEncargoXTaxa.SelecionarVigente(nfe.Cliente.CodigoEstado, "ICMS-ST", nfe.Movimento, "")
                End If

                If Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse Left(nfe.CodigoEmpresa, 8) = "49673784" Then

                    Valor = objEncargoXTaxa.SimplesNacional.ToString("N2")
                    sb.Append("  pMVAST  =" & Valor.Replace(",", ".") & ControlChars.CrLf)

                ElseIf nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado Then
                    sb.Append("  pMVAST  =53.45" & ControlChars.CrLf)
                ElseIf CDec(AliIcms) = 4 Then
                    sb.Append("  pMVAST  =77.49" & ControlChars.CrLf)
                ElseIf CDec(AliIcms) = 7 Then
                    sb.Append("  pMVAST  =71.94" & ControlChars.CrLf)
                ElseIf CDec(AliIcms) = 12 Then
                    sb.Append("  pMVAST  =62.70" & ControlChars.CrLf)
                End If

                sb.Append("   vBCST  =" & BaseIcmsST.Replace(",", ".") & ControlChars.CrLf)
                'Alterado conforme solicitação Leandro - FURLAN - 22/02/2024
                'Valor = CDec(AliIcmsST).ToString("N2")

                If Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse Left(nfe.CodigoEmpresa, 8) = "49673784" Then

                    Dim vlrST As Decimal = Math.Round((CDec(BaseIcmsST) * objEncargoXTaxa.Percentual / 100), 2, MidpointRounding.AwayFromZero)

                    If vlrST > 0 Then
                        Valor = Math.Round(((vlrST - ValIcms) / CDec(BaseIcmsST)) * 100, 2, MidpointRounding.AwayFromZero).ToString("N2")
                    End If

                    If nfe.Cliente.CodigoEstado = "RJ" OrElse
                        nfe.Cliente.CodigoEstado = "CE" OrElse
                        nfe.Cliente.CodigoEstado = "DF" OrElse
                        nfe.Cliente.CodigoEstado = "PB" OrElse
                        nfe.Cliente.CodigoEstado = "RR" OrElse
                        nfe.Cliente.CodigoEstado = "TO" Then
                        sb.Append(" pICMSST  = 20.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "SP" OrElse
                        nfe.Cliente.CodigoEstado = "AP" OrElse
                        nfe.Cliente.CodigoEstado = "MG" Then
                        sb.Append(" pICMSST  = 18.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PA" OrElse
                        nfe.Cliente.CodigoEstado = "SE" Then
                        sb.Append(" pICMSST  = 19.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "ES" OrElse
                        nfe.Cliente.CodigoEstado = "MS" OrElse
                        nfe.Cliente.CodigoEstado = "MT" OrElse
                        nfe.Cliente.CodigoEstado = "RS" Then
                        sb.Append(" pICMSST  = 17.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "MA" Then
                        sb.Append(" pICMSST  = 23.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PE" Then
                        sb.Append(" pICMSST  = 20.50" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PI" Then
                        sb.Append(" pICMSST  = 21.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PR" Then
                        sb.Append(" pICMSST  = 19.50" & ControlChars.CrLf)

                    Else
                        sb.Append(" pICMSST  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                    End If
                Else
                    Valor = CDec(AliIcms).ToString("N2")
                    sb.Append(" pICMSST  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                End If

                sb.Append(" vICMSST  =" & ValorIcmsST.Replace(",", ".") & ControlChars.CrLf)

                If nfe.Cliente.CodigoEstado = "AL" OrElse
                    nfe.Cliente.CodigoEstado = "CE" OrElse
                    nfe.Cliente.CodigoEstado = "MA" OrElse
                    nfe.Cliente.CodigoEstado = "PB" OrElse
                    nfe.Cliente.CodigoEstado = "RJ" OrElse
                    nfe.Cliente.CodigoEstado = "SE" Then

                    sb.Append(" VBCFCPST   =" & BaseIcmsST.Replace(",", ".") & ControlChars.CrLf)

                    sb.Append(" PFCPST   =2.00" & ControlChars.CrLf)

                    Valor = CDec(valorFPobresa).ToString("N2")
                    sb.Append(" VFCPST  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                End If
            End If

            If SituacaoTributaria = 20 OrElse SituacaoTributaria = 120 OrElse SituacaoTributaria = 620 Then
                sb.Append("[ICMS20]" & ControlChars.CrLf)
                If SituacaoTributaria = 620 Then
                    sb.Append("   ORIG   =6" & ControlChars.CrLf)
                ElseIf temImportacao Then
                    sb.Append("   ORIG   =1" & ControlChars.CrLf)
                Else
                    sb.Append("   ORIG   =0" & ControlChars.CrLf)
                End If

                If SituacaoTributaria = 120 OrElse SituacaoTributaria = 620 Then
                    sb.Append("   CST    =" & Mid(SituacaoTributaria, 2, 2) & ControlChars.CrLf)
                Else
                    sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)
                End If

                sb.Append("   MODBC  =3" & ControlChars.CrLf)

                Dim basePercentual As Decimal = 100
                basePercentual = item.OperacaoEstado.Encargos.Where(Function(s) (s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.")).First.AliquotaBase

                If basePercentual < 100 Then
                    Valor = Math.Round(100 - basePercentual, 2, MidpointRounding.AwayFromZero)
                Else
                    Valor = (100 - Math.Round(((12 * 100) / 18), 2, MidpointRounding.AwayFromZero))
                End If

                sb.Append("  PREDBC  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VBC    =" & BaseIcms.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliIcms).ToString("N2")
                sb.Append("   PICMS  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(ValIcms).ToString("N2")
                sb.Append("   VICMS  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                'ENTRA EM VIFOR À PARTIR DE 02/04/2018 - FURLAN
                'sb.Append(" VBCFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("   PFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("   VFCP   =0.00" & ControlChars.CrLf)

                ValIcmsDESONERADO = Math.Round(((item.ValorTotal * CDec(AliIcms)) / 100), 2, MidpointRounding.AwayFromZero)

                ValIcmsDESONERADO = ValIcmsDESONERADO - Math.Round(ValIcms, 2)
                Valor = ValIcmsDESONERADO

                sb.Append("VICMSDESON=" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append("MOTDESICMS=3" & ControlChars.CrLf)

            End If

            If SituacaoTributaria = 30 Then
                sb.Append("[ICMS30]" & ControlChars.CrLf)
                sb.Append("   ORIG   =0" & ControlChars.CrLf)
                sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)

                sb.Append("   MODBCST=4" & ControlChars.CrLf)

                objEncargoXTaxa.SelecionarVigente(nfe.Cliente.CodigoEstado, "ICMS-ST", nfe.Movimento, "")

                If Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse Left(nfe.CodigoEmpresa, 8) = "49673784" Then

                    Valor = objEncargoXTaxa.SimplesNacional.ToString("N2")
                    sb.Append("  pMVAST  =" & Valor.Replace(",", ".") & ControlChars.CrLf)

                ElseIf nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado Then
                    sb.Append("  pMVAST  =53.45" & ControlChars.CrLf)
                ElseIf CDec(AliIcms) = 4 Then
                    sb.Append("  pMVAST  =77.49" & ControlChars.CrLf)
                ElseIf CDec(AliIcms) = 7 Then
                    sb.Append("  pMVAST  =71.94" & ControlChars.CrLf)
                ElseIf CDec(AliIcms) = 12 Then
                    sb.Append("  pMVAST  =62.70" & ControlChars.CrLf)
                End If

                sb.Append("   vBCST  =" & BaseIcmsST.Replace(",", ".") & ControlChars.CrLf)
                'Alterado conforme solicitação Leandro - FURLAN - 22/02/2024
                'Valor = CDec(AliIcmsST).ToString("N2")

                If Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse Left(nfe.CodigoEmpresa, 8) = "49673784" Then
                    Valor = CDec(objEncargoXTaxa.Percentual).ToString("N2")

                    If nfe.Cliente.CodigoEstado = "RJ" OrElse
                        nfe.Cliente.CodigoEstado = "CE" OrElse
                        nfe.Cliente.CodigoEstado = "DF" OrElse
                        nfe.Cliente.CodigoEstado = "PB" OrElse
                        nfe.Cliente.CodigoEstado = "RR" OrElse
                        nfe.Cliente.CodigoEstado = "TO" Then
                        sb.Append(" pICMSST  = 20.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "SP" OrElse
                        nfe.Cliente.CodigoEstado = "AP" OrElse
                        nfe.Cliente.CodigoEstado = "MG" Then
                        sb.Append(" pICMSST  = 18.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PA" OrElse
                        nfe.Cliente.CodigoEstado = "SE" Then
                        sb.Append(" pICMSST  = 19.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "ES" OrElse
                        nfe.Cliente.CodigoEstado = "MS" OrElse
                        nfe.Cliente.CodigoEstado = "MT" OrElse
                        nfe.Cliente.CodigoEstado = "RS" Then
                        sb.Append(" pICMSST  = 17.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "MA" Then
                        sb.Append(" pICMSST  = 23.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PE" Then
                        sb.Append(" pICMSST  = 20.50" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PI" Then
                        sb.Append(" pICMSST  = 21.00" & ControlChars.CrLf)

                    ElseIf nfe.Cliente.CodigoEstado = "PR" Then
                        sb.Append(" pICMSST  = 19.50" & ControlChars.CrLf)

                    Else
                        sb.Append(" pICMSST  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                    End If
                Else
                    Valor = CDec(AliIcms).ToString("N2")
                    sb.Append(" pICMSST  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                End If

                sb.Append(" vICMSST  =" & ValorIcmsST.Replace(",", ".") & ControlChars.CrLf)

                sb.Append("motDesICMS=7" & ControlChars.CrLf)

                Valor = CDec(ValIcmsDESONERADO).ToString("N2")
                sb.Append("VICMSDESON=" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

            End If

            If SituacaoTributaria = 40 OrElse SituacaoTributaria = 140 OrElse
                SituacaoTributaria = 41 OrElse SituacaoTributaria = 141 OrElse SituacaoTributaria = 341 OrElse
                SituacaoTributaria = 50 OrElse SituacaoTributaria = 150 OrElse SituacaoTributaria = 640 OrElse SituacaoTributaria = 641 OrElse SituacaoTributaria = 650 Then
                sb.Append("[ICMS40]" & ControlChars.CrLf)

                If SituacaoTributaria = 341 Then
                    sb.Append("   ORIG   =3" & ControlChars.CrLf)
                ElseIf SituacaoTributaria = 640 Then
                    sb.Append("   ORIG   =6" & ControlChars.CrLf)
                ElseIf SituacaoTributaria = 641 Then
                    sb.Append("   ORIG   =6" & ControlChars.CrLf)
                ElseIf SituacaoTributaria = 650 Then
                    sb.Append("   ORIG   =6" & ControlChars.CrLf)
                ElseIf temImportacao OrElse (SituacaoTributaria.ToString.Length = 3 AndAlso Mid(SituacaoTributaria, 1, 1) = 1) Then
                    sb.Append("   ORIG   =1" & ControlChars.CrLf)
                Else
                    sb.Append("   ORIG   =0" & ControlChars.CrLf)
                End If

                If SituacaoTributaria = 640 OrElse SituacaoTributaria = 641 OrElse SituacaoTributaria = 650 OrElse SituacaoTributaria = 341 OrElse (SituacaoTributaria.ToString.Length = 3 AndAlso Mid(SituacaoTributaria, 1, 1) = 1) Then
                    sb.Append("   CST    =" & Mid(SituacaoTributaria, 2, 2) & ControlChars.CrLf)
                Else
                    sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)
                End If

                If ValIcmsDESONERADO > 0 Then
                    Valor = CDec(ValIcmsDESONERADO).ToString("N2")
                    sb.Append("VICMSDESON=" & Valor.Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("MOTDESICMS=9" & SituacaoTributaria & ControlChars.CrLf)
                End If
            End If

            If SituacaoTributaria = 51 OrElse SituacaoTributaria = 151 OrElse SituacaoTributaria = 651 Then
                sb.Append("[ICMS51]" & ControlChars.CrLf)
                If SituacaoTributaria = 651 Then
                    sb.Append("   ORIG   =6" & ControlChars.CrLf)
                ElseIf temImportacao Then
                    sb.Append("   ORIG   =1" & ControlChars.CrLf)
                Else
                    sb.Append("   ORIG   =0" & ControlChars.CrLf)
                End If

                If SituacaoTributaria = 151 OrElse SituacaoTributaria = 651 Then
                    sb.Append("   CST    =" & Mid(SituacaoTributaria, 2, 2) & ControlChars.CrLf)
                Else
                    sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)
                End If

                If nfe.Empresa.CodigoEstado = "PR" Then
                    sb.Append("  MODBC   =3" & ControlChars.CrLf)
                    sb.Append(" PREDBC   =0.00" & ControlChars.CrLf)

                    Dim valorVBC As Decimal = 0
                    Dim VICMSOP As Decimal = 0

                    If Left(nfe.CodigoEmpresa, 8) = "24450490" Then
                        valorVBC = CDec(BaseIcms)
                    ElseIf ValDesconto > 0 Then
                        valorVBC = item.ValorTotal - ValDesconto
                    Else
                        'valorVBC = item.ValorTotal
                        valorVBC = CDec(BaseIcms)
                    End If

                    Valor = valorVBC.ToString("N2")
                    sb.Append("   VBC    =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                    If CDec(AliIcms) = 0 Then

                        If nfe.Movimento.Year > 2022 Then

                            'Nutri, Baxi, RTGrãos
                            If Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse Left(nfe.CodigoEmpresa, 8) = "49673784" OrElse Left(nfe.CodigoEmpresa, 8) = "05366261" OrElse Left(nfe.CodigoEmpresa, 8) = "44979506" OrElse Left(nfe.CodigoEmpresa, 8) = "24450490" OrElse Left(nfe.CodigoEmpresa, 8) = "62747840" OrElse Left(nfe.CodigoEmpresa, 8) = "62780383" OrElse Left(nfe.CodigoEmpresa, 8) = "63358210" Then
                                sb.Append("   PICMS  =19.50" & ControlChars.CrLf)
                                VICMSOP = Math.Round((valorVBC * 19.5) / 100, 2, MidpointRounding.AwayFromZero)
                            Else
                                sb.Append("   PICMS  =19.00" & ControlChars.CrLf)
                                VICMSOP = Math.Round((valorVBC * 19) / 100, 2, MidpointRounding.AwayFromZero)
                            End If

                            Valor = VICMSOP.ToString("N2")
                            sb.Append(" VICMSOP  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                            sb.Append("    PDIF  =100.00" & ControlChars.CrLf)
                            sb.Append("VICMSDIF  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                        Else
                            sb.Append("   PICMS  =18.00" & ControlChars.CrLf)
                            VICMSOP = Math.Round((valorVBC * 18) / 100, 2, MidpointRounding.AwayFromZero)
                            Valor = VICMSOP.ToString("N2")
                            sb.Append(" VICMSOP  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                            sb.Append("    PDIF  =100.00" & ControlChars.CrLf)
                            sb.Append("VICMSDIF  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                        End If

                        sb.Append("   VICMS  =0.00" & ControlChars.CrLf)

                    Else
                        Valor = CDec(AliIcms).ToString("N2")
                        sb.Append("   PICMS  =" & Valor.Replace(",", ".") & ControlChars.CrLf)

                        VICMSOP = Math.Round((valorVBC * CDec(AliIcms)) / 100, 2, MidpointRounding.AwayFromZero)
                        Valor = VICMSOP.ToString("N2")
                        sb.Append(" VICMSOP  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                        Dim basePercentual As Decimal = 100
                        basePercentual = item.OperacaoEstado.Encargos.Where(Function(s) (s.CodigoEncargo = "ICMS" OrElse s.CodigoEncargo = "ICMS A REC.")).First.AliquotaBase

                        If basePercentual < 100 Then
                            Valor = Math.Round(100 - basePercentual, 2, MidpointRounding.AwayFromZero)
                        Else
                            Valor = 100
                        End If
                        sb.Append("    PDIF  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                        Dim vdif As Decimal = (valorVBC * CDec(AliIcms)) / 100
                        vdif = Math.Round(vdif - ValIcms, 2, MidpointRounding.AwayFromZero)

                        Valor = vdif.ToString("N2")
                        sb.Append("VICMSDIF  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                        Valor = CDec(ValIcms).ToString("N2")
                        sb.Append("   VICMS  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                    End If

                ElseIf nfe.Empresa.CodigoEstado = "RS" AndAlso nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.Itens(0).CodigoProduto = "101110001" Then
                    sb.Append("  MODBC   =3" & ControlChars.CrLf)
                    sb.Append(" PREDBC   =0.00" & ControlChars.CrLf)

                    Dim valorVBC As Decimal = 0
                    Dim VICMSOP As Decimal = 0

                    If ValDesconto > 0 Then
                        valorVBC = item.ValorTotal - ValDesconto
                    Else
                        valorVBC = item.ValorTotal
                    End If

                    Valor = valorVBC.ToString("N2")
                    sb.Append("   VBC    =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("   PICMS  =17.00" & ControlChars.CrLf)

                    VICMSOP = (valorVBC * 17) / 100
                    Valor = VICMSOP.ToString("N2")
                    sb.Append(" VICMSOP  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                    sb.Append("    PDIF  =100.00" & ControlChars.CrLf)
                    sb.Append("VICMSDIF  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("   VICMS  =0.00" & ControlChars.CrLf)

                ElseIf ValIcmsDESONERADO > 0 Then
                    sb.Append("   VBC    =" & BaseIcms.Replace(",", ".") & ControlChars.CrLf)
                    Valor = CDec(AliIcms).ToString("N2")
                    sb.Append("   PICMS  =" & Valor.Replace(",", ".") & ControlChars.CrLf)

                    Valor = CDec(ValIcms).ToString("N2")
                    sb.Append(" VICMSOP  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("    PDIF  =100.00" & ControlChars.CrLf)
                    sb.Append("VICMSDIF  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                End If

                'sb.Append("   MODBC  =3" & ControlChars.CrLf)
                'Valor = (100 - Math.Round(((12 * 100) / 18), 2, MidpointRounding.AwayFromZero))
                'If BaseIcms = "0.00" Then
                '    sb.Append("  PREDBC  =0.00" & ControlChars.CrLf)
                'Else
                '    sb.Append("  PREDBC  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                'End If

                'BaseIcms = Math.Round((BaseIcms - ((BaseIcms * Valor) / 100)), 2)

                'If BaseIcms = "0" Or BaseIcms = "0.00" Then
                '    sb.Append("   VBC    =257506.90" & ControlChars.CrLf)
                'Else
                '    sb.Append("   VBC    =" & BaseIcms.Replace(",", ".") & ControlChars.CrLf)
                'End If
                'Valor = CDec(AliIcms).ToString("N2")

                'sb.Append("   PICMS  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                'sb.Append("   VICMS  =" & ValorIcms.Replace(",", ".") & ControlChars.CrLf)

                'ENTRA EM VIFOR À PARTIR DE 02/04/2018 - FURLAN
                'sb.Append(" VBCFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("   PFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("   VFCP   =0.00" & ControlChars.CrLf)
            End If

            If SituacaoTributaria = 60 OrElse SituacaoTributaria = 160 Then
                sb.Append("[ICMS60]" & ControlChars.CrLf)
                If temImportacao Then
                    sb.Append("   ORIG   =1" & ControlChars.CrLf)
                Else
                    sb.Append("   ORIG   =0" & ControlChars.CrLf)
                End If
                sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)

                If Left(nfe.CodigoEmpresa, 8) = "49673784" Then
                    If nfe.CodigoCliente.Length > 11 AndAlso (nfe.Cliente.InscricaoEstadual.ToString.Contains("ISENTO") Or nfe.Cliente.InscricaoEstadual.Length = 0) Then
                        'NÃO INFORMA A TAG
                    Else
                        sb.Append("    PST   =0.00" & ControlChars.CrLf)
                        sb.Append("vBCSTRet =0.00" & ControlChars.CrLf)
                        sb.Append("vICMSSTRet =0.00" & ControlChars.CrLf)
                    End If
                End If

                'ENTRA EM VIFOR À PARTIR DE 02/04/2018 - FURLAN
                'sb.Append("VBCFCPSTRET=0.00" & ControlChars.CrLf)
                'sb.Append("PFCPSTRET =0.00" & ControlChars.CrLf)
                'sb.Append("VFCPSTRET =0.00" & ControlChars.CrLf)
            End If

            If SituacaoTributaria = 90 OrElse SituacaoTributaria = 190 OrElse SituacaoTributaria = 690 Then
                sb.Append("[ICMS90]" & ControlChars.CrLf)

                If SituacaoTributaria = 690 Then
                    sb.Append("   ORIG   =6" & ControlChars.CrLf)
                ElseIf temImportacao Then
                    sb.Append("   ORIG   =1" & ControlChars.CrLf)
                Else
                    sb.Append("   ORIG   =0" & ControlChars.CrLf)
                End If

                If SituacaoTributaria = 190 OrElse SituacaoTributaria = 690 Then
                    sb.Append("   CST    =" & Mid(SituacaoTributaria, 2, 2) & ControlChars.CrLf)
                Else
                    sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)
                End If

                If Not BaseIcms = "0.00" Then
                    sb.Append("   MODBC  =0" & ControlChars.CrLf)
                    sb.Append("   VBC    =" & BaseIcms.Replace(",", ".") & ControlChars.CrLf)
                    Valor = (100 - Math.Round(((12 * 100) / 18), 2, MidpointRounding.AwayFromZero))
                    sb.Append("  PREDBC  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                    Valor = CDec(AliIcms).ToString("N2")
                    sb.Append("   PICMS  =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                    Valor = CDec(ValIcms).ToString("N2")
                    sb.Append("   VICMS  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                End If

                'ENTRA EM VIFOR À PARTIR DE 02/04/2018 - FURLAN
                'sb.Append(" VBCFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("   PFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("   VFCP   =0.00" & ControlChars.CrLf)
                'sb.Append("VBCFCPST  =0.00" & ControlChars.CrLf)
                'sb.Append(" PFCPST   =0.00" & ControlChars.CrLf)
                'sb.Append(" VFCPST   =0.00" & ControlChars.CrLf)
            End If

            If ValIcmsDESONERADO > 0 Then
                TotalValIcmsDESONERADO += ValIcmsDESONERADO
            End If

            sb.Append(ControlChars.CrLf)

            If ICMSDIFERENCIAL > 0 Then
                sb.Append("[ICMSUFDest]" & ControlChars.CrLf)
                sb.Append(" vBCUFDest=" & BaseIcmsDIFERENCIAL.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("vBCFCPUFDest=0.00" & ControlChars.CrLf)
                sb.Append("pFCPUFDest=0.00" & ControlChars.CrLf)
                Valor = CDec(AliIcmsDIFERENCIAL).ToString("N2")
                sb.Append("pICMSUFDest=" & Valor.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliIcms).ToString("N2")
                sb.Append("pICMSInter=" & Valor.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("pICMSInterPart=100.00" & ControlChars.CrLf)
                sb.Append("vFCPUFDest=0.00" & ControlChars.CrLf)
                Valor = ICMSDIFERENCIAL.ToString("N2")
                sb.Append("vICMSUFDest=" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append("vICMSUFRemet=0.00" & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)
            End If

            'Imposto de Importação
            If Not nfe.DadosDaImportacao Is Nothing AndAlso nfe.DadosDaImportacao.NumeroDeclaracaoImportacao.Length > 0 Then
                sb.Append("[II]" & ControlChars.CrLf)
                Valor = CDec(BaseAduaneira).ToString("N2")
                sb.Append("   VBC    =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(ValAduaneira).ToString("N2")
                sb.Append("VDESPADU  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VII    =0.00" & ControlChars.CrLf)
                sb.Append("   VIOF	 =0.00" & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)
            End If

            'IPI
            sb.Append("[IPI]" & ControlChars.CrLf)

            If nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.Cliente.CodigoEstado = "EX" AndAlso undPedido = "TON" AndAlso SituacaoTributariaIPI = "54" Then
                sb.Append("   CENQ   =002" & ControlChars.CrLf)
            Else
                sb.Append("   CENQ   =999" & ControlChars.CrLf)
            End If

            sb.Append(ControlChars.CrLf)

            'IPI
            If ValIPI > 0 Then
                sb.Append("[IPITRIB]" & ControlChars.CrLf)
                'If nfe.EntradaSaida = eEntradaSaida.Entrada Then
                '    sb.Append("   CST	  =00" & ControlChars.CrLf)
                'Else
                '    sb.Append("   CST	  =50" & ControlChars.CrLf)
                'End If
                sb.Append("   CST	  =" & SituacaoTributariaIPI.ToString("00") & ControlChars.CrLf)
                sb.Append("   VBC    =" & BaseIPI.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliIPI).ToString("N2")
                sb.Append("   PIPI   =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VIPI   =" & ValorIPI.Replace(",", ".") & ControlChars.CrLf)
            Else
                sb.Append("[IPINT]" & ControlChars.CrLf)
                'If nfe.EntradaSaida = eEntradaSaida.Entrada Then
                '    sb.Append("   CST	  =03" & ControlChars.CrLf)
                'Else
                '    sb.Append("   CST	  =53" & ControlChars.CrLf)
                'End If
                sb.Append("   CST	  =" & SituacaoTributariaIPI.ToString("00") & ControlChars.CrLf)
            End If
            sb.Append(ControlChars.CrLf)

            'PIS
            If SituacaoTributariaPISCOFINS = 1 Or SituacaoTributariaPISCOFINS = 2 Then
                sb.Append("[PISALIQ]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
                sb.Append("   VBC    =" & BasePis.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliPis).ToString("N2")
                sb.Append("   PPIS   =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VPIS   =" & ValorPis.Replace(",", ".") & ControlChars.CrLf)

                sb.Append(ControlChars.CrLf)

                sb.Append("[COFINSALIQ]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
                sb.Append("   VBC    =" & BaseCofins.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliCofins).ToString("N2")
                sb.Append("   PCOFINS   =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VCOFINS   =" & ValorCofins.Replace(",", ".") & ControlChars.CrLf)
            ElseIf SituacaoTributariaPISCOFINS = 3 Then
                sb.Append("[PISQTDE]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
                Valor = item.QuantidadeFiscal.ToString("N4")
                sb.Append(" QBCPROD  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                Valor = Math.Round((CDec(BasePis) - CDec(ValorPis)), 2, MidpointRounding.AwayFromZero).ToString("N4")
                sb.Append("VALIQPROD =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VPIS   =" & ValorPis.Replace(",", ".") & ControlChars.CrLf)

                sb.Append(ControlChars.CrLf)

                sb.Append("[COFINSQTDE]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
                Valor = item.QuantidadeFiscal.ToString("N4")
                sb.Append(" QBCPROD  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                Valor = Math.Round((CDec(BaseCofins) - CDec(AliCofins)), 2, MidpointRounding.AwayFromZero).ToString("N4")
                sb.Append("VALIQPROD =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VCOFINS   =" & ValorCofins.Replace(",", ".") & ControlChars.CrLf)
            ElseIf SituacaoTributariaPISCOFINS = 4 Or SituacaoTributariaPISCOFINS = 6 Or SituacaoTributariaPISCOFINS = 7 Or SituacaoTributariaPISCOFINS = 8 Or SituacaoTributariaPISCOFINS = 9 Then
                sb.Append("[PISNT]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)

                sb.Append(ControlChars.CrLf)

                sb.Append("[COFINSNT]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
            Else
                sb.Append("[PISOUTR]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
                sb.Append("   VBC    =" & BasePis.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliPis).ToString("N2")
                sb.Append("   PPIS   =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VPIS   =" & ValorPis.Replace(",", ".") & ControlChars.CrLf)

                sb.Append(ControlChars.CrLf)

                sb.Append("[COFINSOUTR]" & ControlChars.CrLf)
                sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
                sb.Append("   VBC    =" & BaseCofins.Replace(",", ".") & ControlChars.CrLf)
                Valor = CDec(AliCofins).ToString("N2")
                sb.Append("   PCOFINS   =" & Valor.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VCOFINS   =" & ValorCofins.Replace(",", ".") & ControlChars.CrLf)
            End If

            'IBS-CBS
            sb.Append(ControlChars.CrLf)

            If nfe.SubOperacao.Classe = eClassesOperacoes.TRANSFERENCIAS AndAlso nfe.EntradaSaida = eEntradaSaida.Saida AndAlso SituacaoTributariaIBSCBS = 410 AndAlso ClassificacaoIBSCBS = 410002 Then
                sb.Append("[IBSCBS]" & ControlChars.CrLf)
                sb.Append("   CST                   = " & SituacaoTributariaIBSCBS.ToString("000") & ControlChars.CrLf)
                sb.Append("   CCLASSTRIB                = " & ClassificacaoIBSCBS.ToString("000000") & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)
            ElseIf bTemIBSCBS AndAlso Not nfe.CFOP.Codigo = 5605 Then

                sb.Append("[IBSCBS]" & ControlChars.CrLf)
                sb.Append("   CST                   = " & SituacaoTributariaIBSCBS.ToString("000") & ControlChars.CrLf)
                sb.Append("   CCLASSTRIB                = " & ClassificacaoIBSCBS.ToString("000000") & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)

                sb.Append("[GIBSCBS]" & ControlChars.CrLf)
                sb.Append("   VBC                       = " & BaseIBS.Replace(",", ".") & ControlChars.CrLf)

                Dim dValorIBS As Decimal = 0D
                If Not String.IsNullOrWhiteSpace(dValorIBS) Then
                    dValorIBS = Decimal.Parse(ValorIBS, CultureInfo.InvariantCulture)
                End If

                If (SituacaoTributariaIBSCBS = 515 And ClassificacaoIBSCBS = 515001) OrElse (SituacaoTributariaIBSCBS = 510 And ClassificacaoIBSCBS = 510001) Then
                    sb.Append("   VIBS                    = 0.00" & ControlChars.CrLf)
                Else
                    sb.Append("   VIBS                    = " & dValorIBS.ToString("N2").Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                End If

                sb.Append(ControlChars.CrLf)

                'If ReducaoIBS_Perc > 0 Then
                '    dValorIBS -= (dValorIBS * ReducaoIBS_Perc / 100D)
                'End If

                sb.Append("[GIBSUF]" & ControlChars.CrLf)
                If ReducaoIBS_Perc > 0 Then
                    sb.Append("   PIBSUF                    = " & AliIBSOri.ToString("N4").Replace(",", ".") & ControlChars.CrLf)
                Else
                    sb.Append("   PIBSUF                    = " & AliIBS.Replace(",", ".") & ControlChars.CrLf)
                End If

                If (SituacaoTributariaIBSCBS = 515 And ClassificacaoIBSCBS = 515001) OrElse (SituacaoTributariaIBSCBS = 510 And ClassificacaoIBSCBS = 510001) Then
                    sb.Append("   VIBSUF                    = 0.00" & ControlChars.CrLf)

                    sb.Append("[GDIF]" & ControlChars.CrLf)

                    sb.Append("   PDIF                      = 100.00" & ControlChars.CrLf)
                    sb.Append("   VDIF                      = " & dValorIBS.ToString("N2").Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                Else
                    sb.Append("   VIBSUF                    = " & dValorIBS.ToString("N2").Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                End If

                sb.Append(ControlChars.CrLf)

                If ReducaoIBS_Perc > 0 Then

                    Dim dAliIBS As Decimal = 0D

                    If Not String.IsNullOrWhiteSpace(dAliIBS) Then
                        dAliIBS = Decimal.Parse(AliIBS, CultureInfo.InvariantCulture)
                    End If

                    Dim pAliqEfet As Decimal
                    'pAliqEfet = alíquota_nominal × (1 − pRedAliq / 100)
                    'pAliqEfet = dAliIBS * (1D - (ReducaoIBS_Perc / 100D))
                    pAliqEfet = dAliIBS

                    sb.Append("[GRED]" & ControlChars.CrLf)
                    sb.Append("   PREDALIQ                  = " & ReducaoIBS_Perc.ToString("N2").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("   PALIQEFET                 = " & pAliqEfet.ToString("N4").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append(ControlChars.CrLf)

                End If

                sb.Append(ControlChars.CrLf)
                sb.Append("[GDEVTRIB]" & ControlChars.CrLf)
                sb.Append("   VDEVTRIB                  = 0 " & ControlChars.CrLf)  'Valor do tributo devolvido
                sb.Append(ControlChars.CrLf)

                sb.Append("[GIBSMUN]" & ControlChars.CrLf)
                sb.Append("   PIBSMUN                   = 0" & ControlChars.CrLf)
                sb.Append("   VIBSMUN                   = 0" & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)

                If (SituacaoTributariaIBSCBS = 515 And ClassificacaoIBSCBS = 515001) OrElse (SituacaoTributariaIBSCBS = 510 And ClassificacaoIBSCBS = 510001) Then
                    sb.Append("[GDIF_MUN]" & ControlChars.CrLf)

                    sb.Append("   PDIF                      = 100.00" & ControlChars.CrLf)
                    sb.Append("   VDIF                      = 0.00" & ControlChars.CrLf)
                End If

                sb.Append(ControlChars.CrLf)

                If ReducaoIBS_Perc > 0 Then

                    sb.Append("[GRED_MUN]" & ControlChars.CrLf)
                    sb.Append("   PREDALIQ                  = " & ReducaoIBS_Perc.ToString("N2").Replace(",", ".") & ControlChars.CrLf)
                    'pAliqEfet = alíquota_nominal × (1 − pRedAliq / 100)
                    sb.Append("   PALIQEFET                 = 0 " & ControlChars.CrLf)
                    sb.Append(ControlChars.CrLf)
                End If

                sb.Append("[GDEVTRIB_MUN]" & ControlChars.CrLf)
                sb.Append("   VDEVTRIB                  = 0 " & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)

                sb.Append("[GCBS]" & ControlChars.CrLf)
                If ReducaoIBS_Perc > 0 Then
                    sb.Append("   PCBS                      = " & AliCBSOri.ToString("N4").Replace(",", ".") & ControlChars.CrLf)      'Alíquota da CBS (em percentual)
                Else
                    sb.Append("   PCBS                      = " & AliCBS.Replace(",", ".") & ControlChars.CrLf)      'Alíquota da CBS (em percentual)
                End If

                Dim dValorCBS As Decimal = 0D
                If Not String.IsNullOrWhiteSpace(ValorCBS) Then
                    dValorCBS = Decimal.Parse(ValorCBS, CultureInfo.InvariantCulture)
                End If

                'If ReducaoCBS_Perc > 0D Then
                '    dValorCBS -= (dValorCBS * ReducaoCBS_Perc / 100D)
                'End If

                If (SituacaoTributariaIBSCBS = 515 And ClassificacaoIBSCBS = 515001) OrElse (SituacaoTributariaIBSCBS = 510 And ClassificacaoIBSCBS = 510001) Then
                    sb.Append("   VCBS                    = 0.00" & ControlChars.CrLf)

                    sb.Append(ControlChars.CrLf)

                    sb.Append("[GDIF_CBS]" & ControlChars.CrLf)

                    sb.Append("   PDIF                      = 100.00" & ControlChars.CrLf)
                    sb.Append("   VDIF                      = " & dValorCBS.ToString("N2").Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                Else
                    sb.Append("   VCBS                    = " & dValorCBS.ToString("N2").Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                End If

                sb.Append(ControlChars.CrLf)

                If ReducaoCBS_Perc > 0 Then

                    Dim dAliCBS As Decimal = 0D

                    If Not String.IsNullOrWhiteSpace(dAliCBS) Then
                        dAliCBS = Decimal.Parse(AliCBS, CultureInfo.InvariantCulture)
                    End If

                    Dim pAliqEfet As Decimal
                    'pAliqEfet = alíquota_nominal × (1 − pRedAliq / 100)
                    'pAliqEfet = dAliCBS * (1D - (ReducaoIBS_Perc / 100D))
                    pAliqEfet = dAliCBS

                    sb.Append("[GRED_CBS]" & ControlChars.CrLf)
                    sb.Append("   PREDALIQ                  = " & ReducaoCBS_Perc.ToString("N2").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("   PALIQEFET                 = " & pAliqEfet.ToString("N4").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append(ControlChars.CrLf)
                End If

                sb.Append("[GDEVTRIB_CBS]" & ControlChars.CrLf)
                sb.Append("   VDEVTRIB                  = 0 " & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)

            End If

            sb.Append(ControlChars.CrLf)

            'j += 1
        Next

        sb.Append(ControlChars.CrLf)

        sb.Append("--" & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)

        sb.Append("[ICMSTOT]" & ControlChars.CrLf)

        If SituacaoTributaria = 51 OrElse SituacaoTributaria = 151 OrElse SituacaoTributaria = 251 OrElse SituacaoTributaria = 351 OrElse SituacaoTributaria = 651 Then
            If nfe.Empresa.CodigoEstado = "PR" Then
                If Left(nfe.CodigoEmpresa, 8) = "24450490" Then
                    Valor = CDec(TotalBaseIcms).ToString("N2")
                    sb.Append("   VBC	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                ElseIf ValDesconto > 0 Then
                    Valor = CDec((ValorBruto - ValDesconto)).ToString("N2")
                    sb.Append("   VBC	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                Else
                    Valor = CDec(TotalBaseIcms).ToString("N2")
                    sb.Append("   VBC	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                End If
            ElseIf CDec(TotalValIcmsDESONERADO) > 0 Then
                Valor = CDec(TotalBaseIcms).ToString("N2")
                sb.Append("   VBC	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            Else
                Valor = (100 - Math.Round(((12 * 100) / 18), 2, MidpointRounding.AwayFromZero))
                TotalBaseIcms = (TotalBaseIcms - ((TotalBaseIcms * Valor) / 100))
                Valor = CDec(TotalBaseIcms).ToString("N2")
                sb.Append("   VBC    =" & Valor.Replace(",", ".") & ControlChars.CrLf)
            End If
        Else

            Dim teveEncrgoIcms As Boolean = False

            For Each item In nfe.Itens
                For Each enc In item.Encargos
                    If (enc.Codigo = "ICMS" OrElse enc.Codigo = "ICMS A REC.") Then
                        teveEncrgoIcms = True
                    End If
                Next
            Next

            If (Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse Left(nfe.CodigoEmpresa, 8) = "49673784") AndAlso nfe.Cliente.CodigoEstado = "AM" AndAlso Not teveEncrgoIcms Then
                Dim icmsZerado As String = "0.00"
                Valor = CDec(icmsZerado).ToString("N2")
                sb.Append("   VBC	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            Else
                Valor = CDec(TotalBaseIcms).ToString("N2")
                sb.Append("   VBC	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If
        End If

        Valor = CDec(TotalValIcms).ToString("N2")
        sb.Append("   VICMS  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

        If IcmsDESONERADO51 Then
            sb.Append("VICMSDESON=0.00" & ControlChars.CrLf)
        Else
            Valor = CDec(TotalValIcmsDESONERADO).ToString("N2")
            sb.Append("VICMSDESON=" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        End If

        If TOTALICMSDIFERENCIAL > 0 Then
            Valor = TOTALICMSDIFERENCIAL.ToString("N2")
            sb.Append("vICMSUFDest=" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            'Else
            '    sb.Append("vICMSUFDest=0.00" & ControlChars.CrLf)
        End If

        'SUBSTITUIÇÃO TRIBUTÁRIA - FURLAN - 20/10/2023
        If TotalValIcmsST > 0 Then
            Valor = CDec(TotalBaseIcmsST).ToString("N2")
            sb.Append("   VBCST	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            Valor = CDec(TotalValIcmsST).ToString("N2")
            sb.Append("   VST     =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

            If nfe.Cliente.CodigoEstado = "AL" OrElse
                  nfe.Cliente.CodigoEstado = "CE" OrElse
                  nfe.Cliente.CodigoEstado = "MA" OrElse
                  nfe.Cliente.CodigoEstado = "PB" OrElse
                  nfe.Cliente.CodigoEstado = "RJ" OrElse
                  nfe.Cliente.CodigoEstado = "SE" Then

                Valor = CDec(TotalValorFPobresa).ToString("N2")
                sb.Append(" VFCPST  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            Else
                sb.Append(" VFCPST =0.00" & ControlChars.CrLf)
            End If
        Else
            sb.Append("   VBCST  =0.00" & ControlChars.CrLf)
            sb.Append("   VST	 =0.00" & ControlChars.CrLf)
            sb.Append("   VFCPST =0.00" & ControlChars.CrLf)
        End If

        Valor = CDec(ValorBruto).ToString("N2")
        sb.Append("   VPROD  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(TotalValFrete).ToString("N2")
        sb.Append("   VFRETE =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(TotalValSeguro).ToString("N2")
        sb.Append("   VSEG   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(TotalValDesconto).ToString("N2")
        sb.Append("   VDESC  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append("   VII	  =0.00" & ControlChars.CrLf)
        Valor = CDec(TotalValIPI).ToString("N2")
        sb.Append("   VIPI   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(TotalValPis).ToString("N2")
        sb.Append("   VPIS   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(TotalValCofins).ToString("N2")
        sb.Append("   VCOFINS=" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(TotalValOutros).ToString("N2")
        sb.Append("   VOUTRO =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)


        'ENTRA EM VIFOR À PARTIR DE 02/07/2018 - FURLAN
        sb.Append("   VFCP   =0.00" & ControlChars.CrLf)

        sb.Append(" VFCPSTRET=0.00" & ControlChars.CrLf)

        sb.Append("VFCPUFDEST=0.00" & ControlChars.CrLf)

        sb.Append("VICMSUFREMET=0.00" & ControlChars.CrLf)

        sb.Append(" VIPIDEVOL=0.00" & ControlChars.CrLf)

        Valor = CDec(ValorLiquido).ToString("N2")
        sb.Append("   VNF	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

        'IBS-CBS
        sb.Append(ControlChars.CrLf)

        If bTemIBSCBS Then

            'Dim dBaseIBS As Decimal = 0D
            'Dim dBaseCBS As Decimal = 0D

            'If Not String.IsNullOrWhiteSpace(BaseIBS) Then
            '    dBaseIBS = Decimal.Parse(TotalBaseIBS, CultureInfo.InvariantCulture)
            'End If

            'If Not String.IsNullOrWhiteSpace(BaseCBS) Then
            '    dBaseCBS = Decimal.Parse(TotalBaseCBS, CultureInfo.InvariantCulture)
            'End If

            'Pegamos o MAX, pq encargo nova, apesar dos 2 usar a mesma BASE
            'ainda não tem conhecimento se poderá haver situações onde um dos impostos
            'possa vim com base zero
            'Dim Resultado As Decimal = Math.Max(dBaseIBS, dBaseCBS)
            'Dim Resultado As Decimal = TotalBaseCBS + TotalBaseIBS

            'Valor total da BC do IBS e da CBS
            sb.Append("[IBSCBSTOT]" & ControlChars.CrLf)

            sb.Append("   VBCIBSCBS                     = " & TotalBaseCBS.ToString("0.00", CultureInfo.InvariantCulture) & ControlChars.CrLf)

            sb.Append("[GIBS]" & ControlChars.CrLf)

            If (SituacaoTributariaIBSCBS = 515 And ClassificacaoIBSCBS = 515001) OrElse (SituacaoTributariaIBSCBS = 510 And ClassificacaoIBSCBS = 510001) Then
                sb.Append("   VIBS                          = 0.00" & ControlChars.CrLf)
            Else
                sb.Append("   VIBS                          = " & TotalValIBS.ToString("0.00", CultureInfo.InvariantCulture) & ControlChars.CrLf)
            End If

            sb.Append("   VCREDPRES                     = 0" & ControlChars.CrLf)   'Valor total do crédito presumido
            sb.Append("   VCREDPRESCONDSUS              = 0" & ControlChars.CrLf)   'Valor total do crédito presumido em condição suspensiva
            sb.Append(ControlChars.CrLf)

            sb.Append("[GIBSUF]" & ControlChars.CrLf)

            If (SituacaoTributariaIBSCBS = 515 And ClassificacaoIBSCBS = 515001) OrElse (SituacaoTributariaIBSCBS = 510 And ClassificacaoIBSCBS = 510001) Then
                sb.Append("   VDIF                          = " & TotalValIBS.ToString("0.00", CultureInfo.InvariantCulture) & ControlChars.CrLf)   'Valor total do diferimento
                sb.Append("   VDEVTRIB                      = 0" & ControlChars.CrLf)   'Valor total de devolução de tributos
                sb.Append("   VIBSUF                        = 0" & ControlChars.CrLf)
            Else
                sb.Append("   VDIF                          = 0" & ControlChars.CrLf)   'Valor total do diferimento
                sb.Append("   VDEVTRIB                      = 0" & ControlChars.CrLf)   'Valor total de devolução de tributos
                sb.Append("   VIBSUF                        = " & TotalValIBS.ToString("0.00", CultureInfo.InvariantCulture) & ControlChars.CrLf)
            End If

            sb.Append(ControlChars.CrLf)

            sb.Append("[GIBSMUN]" & ControlChars.CrLf)
            sb.Append("   VDIF                          = 0" & ControlChars.CrLf)
            sb.Append("   VDEVTRIB                      = 0" & ControlChars.CrLf)
            sb.Append("   VIBSMUN                       = 0" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[GCBS]" & ControlChars.CrLf)

            If (SituacaoTributariaIBSCBS = 515 And ClassificacaoIBSCBS = 515001) OrElse (SituacaoTributariaIBSCBS = 510 And ClassificacaoIBSCBS = 510001) Then
                sb.Append("   VDIF                          = " & TotalValCBS.ToString("0.00", CultureInfo.InvariantCulture) & ControlChars.CrLf)   'Valor total do diferimento
                sb.Append("   VDEVTRIB                      = 0" & ControlChars.CrLf)   'Valor total de devolução de tributos
                sb.Append("   VCBS                          = 0" & ControlChars.CrLf)    'Valor total da CBS
            Else
                sb.Append("   VDIF                          = 0" & ControlChars.CrLf)   'Valor total do diferimento
                sb.Append("   VDEVTRIB                      = 0" & ControlChars.CrLf)   'Valor total de devolução de tributos
                sb.Append("   VCBS                          = " & TotalValCBS.ToString("0.00", CultureInfo.InvariantCulture) & ControlChars.CrLf)    'Valor total da CBS
            End If

            sb.Append("   VCREDPRES                     = 0" & ControlChars.CrLf)   'Valor total do crédito presumido
            sb.Append("   VCREDPRESCONDSUS              = 0" & ControlChars.CrLf)   'Valor total do crédito presumido em condição suspensiva
            sb.Append(ControlChars.CrLf)

            sb.Append("[GMONO]" & ControlChars.CrLf)
            sb.Append("   VIBSMONO                      = 0" & ControlChars.CrLf)   'Total do IBS monofásico
            sb.Append("   VCBSMONO                      = 0" & ControlChars.CrLf)   'Total da CBS monofásica
            sb.Append("   VIBSMONORETEN                 = 0" & ControlChars.CrLf)   'Total do IBS monofásico sujeito a retenção
            sb.Append("   VCBSMONORETEN                 = 0" & ControlChars.CrLf)   'Total da CBS monofásica sujeita a retenção
            sb.Append("   VIBSMONORET                   = 0" & ControlChars.CrLf)   'Total do IBS monofásico retido anteriormente
            sb.Append("   VCBSMONORET                   = 0" & ControlChars.CrLf)   'Total da CBS monofásica retira anteriormente

        End If

        sb.Append(ControlChars.CrLf)

        sb.Append("[TRANSP]" & ControlChars.CrLf)

        sb.Append("   MODFRETE = " & nfe.TipoFreteSefaz & ControlChars.CrLf)

        If (nfe.EntradaSaida = eEntradaSaida.Entrada AndAlso nfe.CIFFOB = eTiposFrete.FOB) OrElse (nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF) Then
            If Not nfe.PlacaDetalhes Is Nothing AndAlso nfe.PlacaDetalhes.ViaDeTransporte = 2 Then

                sb.Append("   VAGAO  =" & nfe.PlacaDetalhes.Placa01.Replace("-", "") & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)

            End If
        End If

        sb.Append(ControlChars.CrLf)

        'Transportador.
        If Not nfe.Transportador Is Nothing AndAlso nfe.Transportador.Codigo.Length > 0 Then
            sb.Append("[TRANSPORTA]" & ControlChars.CrLf)
            If nfe.Transportador.CodigoEstado = "EX" Then
            ElseIf nfe.Transportador.Codigo.Length > 11 Then
                sb.Append("   CNPJ	  =" & nfe.Transportador.Codigo & ControlChars.CrLf)
            Else
                sb.Append("   CPF	  =" & nfe.Transportador.Codigo & ControlChars.CrLf)
            End If
            sb.Append("   XNOME  =" & Funcoes.EliminarCaracteresEspeciais(nfe.Transportador.Nome) & ControlChars.CrLf)

            If nfe.Transportador.Numero = 0 Then
                sb.Append("   XENDER =" & Funcoes.EliminarCaracteresEspeciais(nfe.Transportador.Endereco) & ControlChars.CrLf)
            Else
                sb.Append("   XENDER =" & Funcoes.EliminarCaracteresEspeciais(nfe.Transportador.Endereco) & ", " & nfe.Transportador.Numero & ControlChars.CrLf)
            End If
            sb.Append("   XMUN   =" & Funcoes.EliminarCaracteresEspeciais(nfe.Transportador.Cidade) & ControlChars.CrLf)
            sb.Append("   UF	  =" & nfe.Transportador.CodigoEstado & ControlChars.CrLf)
            If nfe.Transportador.InscricaoEstadual.Length > 0 Then sb.Append("   IE	  =" & nfe.Transportador.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        End If

        If nfe.Transportador Is Nothing Then
        ElseIf nfe.Transportador.Codigo.Length = 0 Then
        ElseIf nfe.PlacaTransportador.Length = 0 Then
        ElseIf nfe.PlacaDetalhes Is Nothing Then
        ElseIf nfe.PlacaDetalhes.Placa01.Length = 0 Then
        ElseIf nfe.PlacaDetalhes.Placa01 = "." Then
        ElseIf nfe.PlacaDetalhes.Placa01 = "," Then
        ElseIf nfe.PlacaDetalhes.Placa01 = "-" Then
        ElseIf nfe.PlacaDetalhes.Placa01 = "BONIFICA" Then
        ElseIf nfe.PlacaDetalhes.Placa01 = "BONIFICACAO" Then
        ElseIf nfe.PlacaDetalhes.Placa01 = "DEVOLUCAO" Then
            'ElseIf Left(nfe.CodigoEmpresa, 8) = "24450490" AndAlso nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF AndAlso nfe.Empresa.CodigoEstado = "MG" AndAlso nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado Then
            'LIBERADO RTGRÃOS PARA APARECER PLACA, ACOMPANHAR RETORNO DA SEFAZ - FURLAN 02/12/2024
            'elseif (nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado OrElse nfe.Cliente.CodigoEstado = "EX")
        Else

            'LIBERA INFORMAÇÃO/IMPRESSÃO DA PLACA NO XML/DANFE
            Dim liberaPlaca As Boolean = False

            'FRETE POT CONTA DA EMPRESA
            If ((nfe.EntradaSaida = eEntradaSaida.Entrada AndAlso nfe.CIFFOB = eTiposFrete.FOB) OrElse
                (nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF)) AndAlso (nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado OrElse nfe.Cliente.CodigoEstado = "EX") Then
                liberaPlaca = True

                If Left(nfe.CodigoEmpresa, 8) = "24450490" AndAlso (nfe.Empresa.CodigoEstado = "MG" AndAlso nfe.Cliente.CodigoEstado = "MG") Then
                    liberaPlaca = False
                End If
            End If

            'FRETE POR CONTA DO DESTINATÁRIO/REMETENTE
            If ((nfe.EntradaSaida = eEntradaSaida.Entrada AndAlso nfe.CIFFOB = eTiposFrete.CIF) OrElse
                (nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.FOB)) AndAlso (nfe.Empresa.CodigoEstado = nfe.Cliente.CodigoEstado) Then
                liberaPlaca = True
            End If

            'FRETE POR CONTA DO DESTINATÁRIO/REMETENTE - MINAS GERAIS NÃO ACEITA
            If ((nfe.EntradaSaida = eEntradaSaida.Entrada AndAlso nfe.CIFFOB = eTiposFrete.CIF) OrElse
                (nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.FOB)) AndAlso (nfe.Empresa.CodigoEstado = "MG" AndAlso nfe.Cliente.CodigoEstado = "MG") Then
                liberaPlaca = False
            ElseIf ((nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF)) AndAlso (nfe.Empresa.CodigoEstado = "PR" AndAlso nfe.Cliente.CodigoEstado = "SC") Then
                liberaPlaca = False
            ElseIf ((nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF)) AndAlso (nfe.Empresa.CodigoEstado = "PR" AndAlso nfe.Cliente.CodigoEstado = "SP") Then
                liberaPlaca = False
            ElseIf ((nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF)) AndAlso (nfe.Empresa.CodigoEstado = "PR" AndAlso nfe.Cliente.CodigoEstado = "RS") Then
                liberaPlaca = False
            ElseIf ((nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF)) AndAlso (nfe.Empresa.CodigoEstado = "PR" AndAlso nfe.Cliente.CodigoEstado = "GO") Then
                liberaPlaca = False
            End If

            If ((nfe.EntradaSaida = eEntradaSaida.Entrada AndAlso nfe.CIFFOB = eTiposFrete.FOB) OrElse
                 (nfe.EntradaSaida = eEntradaSaida.Saida AndAlso nfe.CIFFOB = eTiposFrete.CIF)) AndAlso (nfe.Empresa.CodigoEstado = "PR" AndAlso nfe.Cliente.CodigoEstado = "MT") Then
                liberaPlaca = False
            End If


            If liberaPlaca Then

                If Not nfe.PlacaDetalhes Is Nothing AndAlso nfe.PlacaDetalhes.ViaDeTransporte = 1 Then

                    If nfe.PlacaDetalhes.Placa04.Length > 0 Then
                        sb.Append("[VEICTRANSP]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa01.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca01 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)

                        sb.Append("[REBOQUE]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa02.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca02 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)

                        sb.Append("[REBOQUE]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa03.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca03 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)

                        sb.Append("[REBOQUE]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa04.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca04 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)
                    ElseIf nfe.PlacaDetalhes.Placa03.Length > 0 Then
                        sb.Append("[VEICTRANSP]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa01.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca01 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)

                        sb.Append("[Reboque]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa02.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca02 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)

                        sb.Append("[REBOQUE]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa03.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca03 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)
                    ElseIf nfe.PlacaDetalhes.Placa02.Length > 0 Then
                        sb.Append("[VEICTRANSP]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa01.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca01 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)

                        sb.Append("[REBOQUE]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa02.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca02 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)
                    ElseIf nfe.PlacaDetalhes.Placa01.Length > 0 Then
                        sb.Append("[VEICTRANSP]" & ControlChars.CrLf)
                        sb.Append("   PLACA  =" & nfe.PlacaDetalhes.Placa01.Replace("-", "") & ControlChars.CrLf)
                        sb.Append("   UF	  =" & nfe.PlacaDetalhes.EstadoPlaca01 & ControlChars.CrLf)
                        sb.Append(ControlChars.CrLf)
                    End If

                End If

            End If
        End If

        'Informações de Volume
        sb.Append("[VOL]" & ControlChars.CrLf)

        If nfe.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
            sb.Append("   QVOL  =" & CInt(nfe.NotaDeTroca.Quantidade).ToString & ControlChars.CrLf)
            sb.Append("   ESP	  =" & nfe.NotaDeTroca.Especie.ToString & ControlChars.CrLf)
            If Not nfe.Marca Is Nothing AndAlso nfe.Marca.Length > 0 Then sb.Append("   MARCA  =" & nfe.NotaDeTroca.Marca & ControlChars.CrLf)
            If nfe.Numero.Length > 0 Then sb.Append("   NVOL   =" & nfe.NotaDeTroca.Numero & ControlChars.CrLf)

            Valor = nfe.NotaDeTroca.PesoLiquido.ToString("N3")
            sb.Append("   PESOL  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            Valor = nfe.NotaDeTroca.PesoBruto.ToString("N3")
            sb.Append("   PESOB  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Else
            If nfe.Quantidade = 0 Then
                If nfe.Especie = "SACAS DE 50 KGS" Then
                    sb.Append("   QVOL  =" & CInt(nfe.Itens(0).QuantidadeFiscal / 50).ToString("00000") & ControlChars.CrLf)
                End If
                sb.Append("   ESP	  =" & nfe.Especie & ControlChars.CrLf)
                Valor = nfe.PesoLiquido.ToString("N3")
                sb.Append("   PESOL  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                Valor = nfe.PesoBruto.ToString("N3")
                sb.Append("   PESOB  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            Else
                sb.Append("   QVOL  =" & CInt(nfe.Quantidade).ToString & ControlChars.CrLf)
                sb.Append("   ESP	  =" & nfe.Especie.ToString & ControlChars.CrLf)
                If nfe.Marca.Length > 0 Then sb.Append("   MARCA  =" & nfe.Marca & ControlChars.CrLf)
                If nfe.Numero.Length > 0 Then sb.Append("   NVOL   =" & nfe.Numero & ControlChars.CrLf)

                Valor = nfe.PesoLiquido.ToString("N3")
                sb.Append("   PESOL  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                Valor = nfe.PesoBruto.ToString("N3")
                sb.Append("   PESOB  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If
        End If

        sb.Append(ControlChars.CrLf)

        'Informações adicionais a NF
        sb.Append("[INFADIC]" & ControlChars.CrLf)
        If nfe.ObservacoesDeEmbarque.Length > 0 Then
            sb.Append("   INFCPL    = " & nfe.ObservacoesDeEmbarque & "|" & Funcoes.EliminarCaracteresEspeciaisNF(nfe.Observacoes))
        Else
            sb.Append("   INFCPL    = " & Funcoes.EliminarCaracteresEspeciaisNF(nfe.Observacoes))
        End If

        sb.Append(ControlChars.CrLf)

        If nfe.CodigoLocalEmbarque.Length > 0 AndAlso
            nfe.EntradaSaida = eEntradaSaida.Saida Then

            If nfe.Cliente.CodigoEstado.Equals("EX") Then
                If Not [Enum].Parse(GetType(eClassesOperacoes), nfe.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.IMPORTACOES Then
                    Dim strLocal As String = ""
                    Dim strTraco As String = ""

                    If nfe.LocalEmbarque.Endereco.Length > 0 Then
                        strLocal &= strTraco & nfe.LocalEmbarque.Endereco
                        strTraco = " - "
                    End If

                    If nfe.LocalEmbarque.Cidade.Length > 0 Then
                        strLocal &= strTraco & nfe.LocalEmbarque.Cidade.ToString
                        strTraco = " - "
                    End If

                    sb.Append("[EXPORTA]" & ControlChars.CrLf)
                    'sb.Append("   UFEMBARQ  =" & nfe.LocalEmbarque.CodigoEstado & ControlChars.CrLf)
                    'sb.Append("   XLOCEMBARQ=" & Mid(strLocal, 1, 60) & ControlChars.CrLf)
                    sb.Append("   UFSAIDAPAIS =" & nfe.LocalEmbarque.CodigoEstado & ControlChars.CrLf)
                    sb.Append("   XLOCEXPORTA =" & Mid(strLocal, 1, 60) & ControlChars.CrLf)
                    sb.Append(ControlChars.CrLf)
                End If
            ElseIf nfe.Empresa.CodigoEstado.Equals("MS") Then
                sb.Append("[RETIRADA]" & ControlChars.CrLf)

                If nfe.CodigoLocalEmbarque.Length > 11 Then
                    sb.Append("   CNPJ   = " & nfe.CodigoLocalEmbarque & ControlChars.CrLf)
                Else
                    sb.Append("    CPF   = " & nfe.CodigoLocalEmbarque & ControlChars.CrLf)
                End If
                sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(nfe.LocalEmbarque.Nome) & ControlChars.CrLf)
                sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(nfe.LocalEmbarque.Endereco) & ControlChars.CrLf)
                sb.Append("   NRO	  = " & nfe.LocalEmbarque.Numero & ControlChars.CrLf)
                sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(nfe.LocalEmbarque.Bairro) & ControlChars.CrLf)
                sb.Append("   CMUN    = " & nfe.LocalEmbarque.Municipio.EstadoIbge & nfe.LocalEmbarque.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)
                sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(nfe.LocalEmbarque.Cidade) & ControlChars.CrLf)
                sb.Append("   UF	  = " & nfe.LocalEmbarque.CodigoEstado & ControlChars.CrLf)
                sb.Append("   CEP	  = " & nfe.LocalEmbarque.CEP.Replace(".", "").Replace("-", "") & ControlChars.CrLf)
                sb.Append("   CPAIS   = " & nfe.LocalEmbarque.CodigoPais & ControlChars.CrLf)
                sb.Append("   XPAIS   = " & nfe.LocalEmbarque.Pais.Descricao & ControlChars.CrLf)
                sb.Append("   IE	  = " & nfe.LocalEmbarque.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)
            End If
        End If

        'Informações da Fatura de cobrança
        Dim Vencimentos As Decimal = 0
        Dim ValorParcela As Decimal = 0

        If (Mid(nfe.Empresa.Nome, 1, 16) = "CURTUME PANORAMA" OrElse
            Mid(nfe.Empresa.Nome, 1, 15) = "QUIMICA CENTRAL" OrElse
            nfe.CodigoEmpresa = "15204808000168" OrElse
            nfe.CodigoEmpresa = "15204808000249" OrElse
            Left(nfe.CodigoEmpresa, 8) = "05366261" OrElse
            Left(nfe.CodigoEmpresa, 8) = "38198213" OrElse
            Left(nfe.CodigoEmpresa, 8) = "40938762" OrElse
            Left(nfe.CodigoEmpresa, 8) = "49673784" OrElse
            Left(nfe.CodigoEmpresa, 8) = "44005444" OrElse
            Left(nfe.CodigoEmpresa, 8) = "44979506" OrElse
            Left(nfe.CodigoEmpresa, 8) = "24450490" OrElse
            Left(nfe.CodigoEmpresa, 8) = "48984539" OrElse
            Left(nfe.CodigoEmpresa, 8) = "53267147" OrElse
            Left(nfe.CodigoEmpresa, 8) = "62747840" OrElse
            Left(nfe.CodigoEmpresa, 8) = "62780383" OrElse
            Left(nfe.CodigoEmpresa, 8) = "63358210" OrElse
            nfe.CodigoEmpresa = "04854422000266") AndAlso
            nfe.EntradaSaida = eEntradaSaida.Saida AndAlso
            Not nfe.SubOperacao.FinalidadeDaNota = 3 AndAlso
            Not nfe.SubOperacao.Devolucao Then
            If nfe.VencimentosNota.Count > 0 And nfe.SubOperacao.Financeiro Then
                Vencimentos = 0
                ValorParcela = 0
                For Each tit In nfe.VencimentosNota
                    Vencimentos = Vencimentos + tit.ValorDoDocumento
                Next

                sb.Append("[FAT]" & ControlChars.CrLf)
                sb.Append("   NFAT   =" & nfe.Codigo & ControlChars.CrLf)

                Valor = CDec(Vencimentos).ToString("N2")
                sb.Append("   VORIG  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VDESC  =0.00" & ControlChars.CrLf)
                sb.Append("   VLIQ   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

                sb.Append(ControlChars.CrLf)

                Dim i As Integer = 0
                For Each tit In nfe.VencimentosNota
                    sb.Append("[DUP]" & ControlChars.CrLf)
                    sb.Append("   NDUP   =" & (i + 1).ToString("000") & ControlChars.CrLf)
                    sb.Append("   DVENC  =" & tit.Prorrogacao.ToString("yyyy-MM-dd") & ControlChars.CrLf)
                    Valor = tit.ValorDoDocumento.ToString("N2")
                    sb.Append("   VDUP   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                    sb.Append(ControlChars.CrLf)
                    i += 1
                Next
            End If

            sb.Append(ControlChars.CrLf)
        End If

        sb.Append(ControlChars.CrLf)

        'ENTRA EM VIFOR À PARTIR DE 02/04/2018 - FURLAN
        sb.Append("[PAG]" & ControlChars.CrLf)
        'sb.Append(" VTROCO   =0.00" & ControlChars.CrLf)

        If ((nfe.CFOP.Codigo = 5949 OrElse nfe.CFOP.Codigo = 6949) AndAlso nfe.CodigoFinalidade = 31) OrElse (nfe.SubOperacao.FinalidadeDaNota = 3) Then
            sb.Append("[DETPAG]" & ControlChars.CrLf)
            sb.Append("   TPAG   =90" & ControlChars.CrLf)
            sb.Append("   VPAG   =0.00" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        ElseIf nfe.VencimentosNota IsNot Nothing AndAlso nfe.VencimentosNota.Count > 0 And nfe.SubOperacao.Financeiro And Not nfe.SubOperacao.Devolucao Then
            Dim i As Integer = 0
            For Each tit In nfe.VencimentosNota
                sb.Append("[DETPAG]" & ControlChars.CrLf)
                sb.Append("   INDPAG =1" & ControlChars.CrLf)
                sb.Append("   TPAG   =99" & ControlChars.CrLf)

                'ENTRA EM VIFOR À PARTIR DE 01/09/2021 - FURLAN
                If nfe.EntradaSaida = eEntradaSaida.Saida Then
                    sb.Append("   XPAG   =RECEBIBEMTO POR MEIO BANCARIO" & ControlChars.CrLf)
                Else
                    sb.Append("   XPAG   =PAGAMENTO POR MEIO BANCARIO" & ControlChars.CrLf)
                End If

                Valor = tit.ValorLiquido.ToString("N2")
                sb.Append("   VPAG   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)
                i += 1
            Next
        Else
            sb.Append("[DETPAG]" & ControlChars.CrLf)
            sb.Append("   TPAG   =90" & ControlChars.CrLf)
            sb.Append("   VPAG   =0.00" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        End If

        Return sb.ToString()
    End Function

    'PROCESSO CUPOM FISCAL - FURLAN 10/05/2023
    Public Shared Function getTextoNFCe4G(ByVal nfCe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        Dim sqls As New ArrayList

        sb.Append("[IDE]" & ControlChars.CrLf)
        'Código da UF do Emitente Documento Fiscal
        sb.Append("CUF	        = " & nfCe.Empresa.Municipio.EstadoIbge & ControlChars.CrLf)
        'Descrição da Natureza da Operação
        sb.Append("NATOP        = " & Funcoes.EliminarCaracteresEspeciais(Mid(nfCe.CFOP.Descricao, 1, 60)) & ControlChars.CrLf)
        'Código do Modelo do Documento Fiscal
        sb.Append("MOD         = 65" & ControlChars.CrLf)
        'Série do Documento Fiscal
        sb.Append("SERIE       = " & nfCe.Serie & ControlChars.CrLf)
        'Número do Documento Fiscal
        sb.Append("NNF         = " & nfCe.Codigo & ControlChars.CrLf)
        'Data/Hora de emissão do Documento Fiscal
        sb.Append("DHEMI       = " & nfCe.DataInclusao.ToString("yyyy-MM-dd") & "T" & nfCe.DataInclusao.ToString("HH:mm:ss") & ControlChars.CrLf)
        'Tipo de Operação - ENTRADA = 0 - SAÍDA = 1
        sb.Append("TPNF        = 1" & ControlChars.CrLf)
        'Identificador de Local de destino da operação (1 - Interno, 2 - Interestadual e 3 - Exterior)
        sb.Append("IDDEST      = 1" & ControlChars.CrLf)
        'MODE DE IMPRESSÃO  1 - RETRATO   2 - PAISAGEM
        sb.Append("TPIMP       = 1" & ControlChars.CrLf)
        'CÓDIGO DO MUNICÍPIO DO IBGE de Ocorrência do Fato Gerador
        sb.Append("CMUNFG      = " & nfCe.Empresa.Municipio.EstadoIbge & nfCe.Empresa.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)
        'Tipo da impressão do DANFE NFC-e    4 – DANFE NFC-e   5 – DANFE NFC-e em mensagem eletrônica
        sb.Append("TPIMP       = 4" & ControlChars.CrLf)
        'Finalidade de emissão da NFC-E: 1 - nfCe-NORMAL    2 – NF-E COMPLEMENTAR   3 – NF-E DE AJUSTE - 4 Devolução
        sb.Append("FINNFE      = 1" & ControlChars.CrLf)
        'Indica Operação com Consumidor Final "0 - Normal    1 - Consumidor final"
        sb.Append("INDFINAL    = 1" & ControlChars.CrLf)
        'Indicador de presença do comprador no estabelecimento comercial no momento da operação     1 – Operação Presencial   4 – NFC-e em operação com entrega a domicílio
        sb.Append("INDPRES     = 1" & ControlChars.CrLf)

        sb.Append("   PROCEMI=0" & ControlChars.CrLf)
        sb.Append("   VERPROC=nfce4g" & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)

        'Emitente
        sb.Append("[EMIT]" & ControlChars.CrLf)
        sb.Append("   CNPJ   = " & nfCe.CodigoEmpresa & ControlChars.CrLf)
        sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(nfCe.Empresa.Nome) & ControlChars.CrLf)
        sb.Append("   IE     = " & nfCe.Empresa.InscricaoEstadual.Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
        sb.Append("   CRT    = " & nfCe.Empresa.Empresa.Crt & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[ENDEREMIT]" & ControlChars.CrLf)
        sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(nfCe.Empresa.Endereco) & ControlChars.CrLf)
        sb.Append("   NRO	  = " & nfCe.Empresa.Numero & ControlChars.CrLf)
        sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(nfCe.Empresa.Complemento) & ControlChars.CrLf)
        sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(nfCe.Empresa.Bairro) & ControlChars.CrLf)
        sb.Append("   CMUN    = " & nfCe.Empresa.Municipio.EstadoIbge & nfCe.Empresa.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)
        sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(nfCe.Empresa.Cidade) & ControlChars.CrLf)
        sb.Append("   UF	  = " & nfCe.Empresa.CodigoEstado & ControlChars.CrLf)
        sb.Append("   CEP	  = " & nfCe.Empresa.CEP.Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        sb.Append("   CPAIS   = " & nfCe.Empresa.CodigoPais & ControlChars.CrLf)
        sb.Append("   XPAIS   = " & nfCe.Empresa.Pais.Descricao & ControlChars.CrLf)
        Dim Fone As String = nfCe.Empresa.Telefone.Replace("(", "").Replace(")", "").Replace(" ", "").Replace("-", "").Replace("/", "")
        If Fone.Length > 0 Then sb.Append("   FONE   =" & Mid(Fone, 1, 10) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        If Not nfCe.CodigoEmpresa = nfCe.CodigoCliente Then
            'Destinatario
            sb.Append("[DEST]" & ControlChars.CrLf)

            sb.Append("    CPF    =" & nfCe.CodigoCliente & ControlChars.CrLf)
            sb.Append("INDIEDEST  =9" & ControlChars.CrLf)
            sb.Append("   XNOME   =" & Funcoes.EliminarCaracteresEspeciais(nfCe.Cliente.Nome.ToString.ToUpper) & ControlChars.CrLf)

            If nfCe.Cliente.EmailNFE IsNot Nothing AndAlso nfCe.Cliente.EmailNFE.Length > 0 Then
                If nfCe.Cliente.EmailNFE.Contains(";") Then
                    Dim emailCli() As String = nfCe.Cliente.EmailNFE.Split(";")
                    sb.Append("   EMAIL   =" & emailCli(0) & ControlChars.CrLf)
                Else
                    sb.Append("   EMAIL   =" & nfCe.Cliente.EmailNFE & ControlChars.CrLf)
                End If
            End If

            sb.Append(ControlChars.CrLf)

            If String.IsNullOrWhiteSpace(nfCe.Cliente.Endereco) Then
                'Endereço do Destinatário
                sb.Append("[ENDERDEST]" & ControlChars.CrLf)
                sb.Append("   XLGR   =" & Funcoes.EliminarCaracteresEspeciais(nfCe.Cliente.Endereco.ToString.ToUpper) & ControlChars.CrLf)
                sb.Append("   NRO	  =" & nfCe.Cliente.Numero & ControlChars.CrLf)

                If nfCe.Cliente.Complemento.Length = 0 Then
                    sb.Append("   XCPL   =." & ControlChars.CrLf)
                Else
                    sb.Append("   XCPL   =" & Funcoes.EliminarCaracteresEspeciais(nfCe.Cliente.Complemento.ToString.ToUpper) & ControlChars.CrLf)
                End If

                If nfCe.Cliente.Bairro.Length = 0 Then
                    sb.Append("   XBAIRRO=." & ControlChars.CrLf)
                Else
                    sb.Append("   XBAIRRO=" & Funcoes.EliminarCaracteresEspeciais(nfCe.Cliente.Bairro.ToString.ToUpper) & ControlChars.CrLf)
                End If

                sb.Append("   CMUN   = " & nfCe.Cliente.Municipio.EstadoIbge & nfCe.Cliente.Municipio.CodigoIbge.ToString("00000") & ControlChars.CrLf)
                sb.Append("   XMUN   =" & Funcoes.EliminarCaracteresEspeciais(nfCe.Cliente.Cidade.ToString.ToUpper) & ControlChars.CrLf)
                sb.Append("   CEP	  =" & nfCe.Cliente.CEP.Replace(".", "").Replace("-", "") & ControlChars.CrLf)

                sb.Append("   UF	  =" & nfCe.Cliente.CodigoEstado & ControlChars.CrLf)
                sb.Append("   CPAIS  =" & nfCe.Cliente.CodigoPais & ControlChars.CrLf)
                sb.Append("   XPAIS  =" & nfCe.Cliente.Pais.Descricao & ControlChars.CrLf)
                Fone = nfCe.Cliente.Telefone.Replace("(", "").Replace(")", "").Replace(" ", "").Replace("-", "").Replace("/", "")
                If Fone.Length > 0 Then sb.Append("   FONE   =" & Mid(Fone, 1, 10) & ControlChars.CrLf)
                sb.Append(ControlChars.CrLf)
            End If
        End If

        Dim BaseProduto As Decimal = 0

        Dim ValorBruto As Decimal = 0
        Dim ValorLiquido As Decimal = 0
        Dim SituacaoTributaria As Integer = 0


        Dim BaseIcms As String = "0.00"
        Dim ValIcms As Decimal = 0
        Dim AliIcms As String = ""
        Dim ValorIcms As String = ""

        Dim BasePis As String = ""
        Dim AliPis As String = ""
        Dim ValorPis As String = ""
        Dim ValPis As Decimal = 0
        Dim BaseCofins As String = ""
        Dim AliCofins As String = ""
        Dim ValorCofins As String = ""
        Dim ValCofins As Decimal = 0

        Dim ValOutros As Decimal = 0
        Dim ValDesconto As Decimal = 0
        Dim Valor As String = ""
        Dim ValFrete As Decimal = 0
        Dim ValSeguro As Decimal = 0

        Dim ValIPI As Decimal = 0
        Dim BaseIPI As String = ""
        Dim AliIPI As String = ""
        Dim ValorIPI As String = "0.00"

        For Each item In nfCe.Itens
            For Each enc In item.Encargos
                If enc.Valor > 0 Then
                    If enc.Codigo = "LIQUIDO" Then
                        ValorLiquido = ValorLiquido + enc.Valor
                    ElseIf enc.Codigo = "PRODUTO" Then
                        ValorBruto = ValorBruto + enc.Valor
                        SituacaoTributaria = enc.SituacaoTributaria
                    ElseIf enc.Codigo = "ICMS" Then
                        BaseIcms = BaseIcms + enc.Base
                        ValIcms = ValIcms + enc.Valor
                    ElseIf enc.Codigo = "PIS" Then
                        ValPis = ValPis + enc.Valor
                    ElseIf enc.Codigo = "COFINS" Then
                        ValCofins = ValCofins + enc.Valor
                    Else
                        If enc.OperacaoEncargo.Sinal = "+" Then
                            ValOutros = ValOutros + enc.Valor
                        ElseIf enc.OperacaoEncargo.Sinal = "-" Then
                            ValDesconto = ValDesconto + enc.Valor
                        End If
                    End If
                End If
            Next
        Next

        sb.Append("[ICMSTOT]" & ControlChars.CrLf)

        Valor = CDec(BaseIcms).ToString("N2")
        sb.Append("   VBC	  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValIcms).ToString("N2")
        sb.Append("   VICMS  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append("VICMSDESON=0.00" & ControlChars.CrLf)
        sb.Append("   VFCP	 =0.00" & ControlChars.CrLf)
        sb.Append("   VBCST  =0.00" & ControlChars.CrLf)
        sb.Append("   VST	 =0.00" & ControlChars.CrLf)
        sb.Append("   VFCPST =0.00" & ControlChars.CrLf)
        sb.Append("VFCPSTRET =0.00" & ControlChars.CrLf)
        Valor = CDec(ValorBruto).ToString("N2")
        sb.Append("   VPROD  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValFrete).ToString("N2")
        sb.Append("   VFRETE =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValSeguro).ToString("N2")
        sb.Append("   VSEG   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValDesconto).ToString("N2")
        sb.Append("   VDESC  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append("   VII	  =0.00" & ControlChars.CrLf)
        Valor = CDec(ValIPI).ToString("N2")
        sb.Append("   VIPI   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValPis).ToString("N2")
        sb.Append("   VPIS   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValCofins).ToString("N2")
        sb.Append("   VCOFINS=" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValOutros).ToString("N2")
        sb.Append("   VOUTRO =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        Valor = CDec(ValorLiquido).ToString("N2")
        sb.Append("   VNF	 =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)

        sb.Append("[TRANSP]" & ControlChars.CrLf)
        sb.Append("   MODFRETE = 9" & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)

        sb.Append("[PAG]" & ControlChars.CrLf)

        sb.Append("[DETPAG]" & ControlChars.CrLf)
        'Ver Tabela dos meios de pagamento
        sb.Append("   TPAG   =01" & ControlChars.CrLf)
        sb.Append("   VPAG   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

        sb.Append(ControlChars.CrLf)

        sb.Append("[OBSCONT]" & ControlChars.CrLf)
        sb.Append("   XCAMPO = IDTOKEN" & ControlChars.CrLf)
        sb.Append("   XTEXTO = 000001a8fdcf15f3e02f86e774f8b3addd8bfcabea" & ControlChars.CrLf)  'Produção
        'sb.Append("   XTEXTO = 0000021223b4f0870a79248aad9ce091bfc69c69fe" & ControlChars.CrLf) 'Homologação

        sb.Append(ControlChars.CrLf)

        Dim j As Integer = 1

        For Each item In nfCe.Itens

            sb.Append("[DET]" & ControlChars.CrLf)

            sb.Append(ControlChars.CrLf)

            sb.Append("[PROD]" & ControlChars.CrLf)

            If Not item.Produto.ProdutoAgrupador Is Nothing AndAlso item.Produto.ProdutoAgrupador.Count > 0 Then
                sb.Append("   CPROD   =" & item.Produto.ProdutoAgrupador(0).CodigoProdutoAgrupado & ControlChars.CrLf)
                sb.Append("   XPROD   = " & Funcoes.EliminarCaracteresEspeciais(item.Produto.ProdutoAgrupador(0).NomeProduto) & " " & ControlChars.CrLf)
                'sb.Append("   XPROD   = " & "NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & " " & ControlChars.CrLf) 'Homologação
            Else
                sb.Append("   CPROD   =" & item.CodigoProduto & ControlChars.CrLf)
                sb.Append("   XPROD   =" & Funcoes.EliminarCaracteresEspeciais(item.Produto.Nome) & " " & ControlChars.CrLf)
                'sb.Append("   XPROD   = " & "NOTA FISCAL EMITIDA EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & " " & ControlChars.CrLf) 'Homologação
            End If

            sb.Append("   CEAN    =SEM GTIN" & ControlChars.CrLf)
            sb.Append("   CFOP    =" & item.CFOP & ControlChars.CrLf)

            Dim undPedido As String = String.Empty

            For Each itemPed In nfCe.Pedido.Itens
                If itemPed.CodigoProduto = item.CodigoProduto Then
                    undPedido = itemPed.CodigoUnidadeComercializacao
                    Exit For
                End If
            Next

            sb.Append("   UCOM    =" & undPedido & ControlChars.CrLf)
            Valor = item.QuantidadeFiscal.ToString("N4")
            sb.Append("   QCOM    =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            sb.Append("   NCM	  =" & item.Produto.NCM.Replace(".", "") & ControlChars.CrLf)

            Valor = CDec(item.Unitario).ToString("N10")
            sb.Append("   VUNCOM  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            Valor = item.ValorTotal.ToString("N2")
            sb.Append("   VPROD   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)


            'VERIFICA ENCARGOS DO ITEM DA NOTA
            BaseProduto = 0.0
            BasePis = "0.00"
            AliPis = "0.00"
            ValorPis = "0.00"
            BaseCofins = "0.00"
            AliCofins = "0.00"
            ValorCofins = "0.00"
            BaseIcms = "0.00"
            ValorIcms = "0.00"
            AliIcms = "0.00"
            BaseIPI = "0.00"
            AliIPI = "0.00"
            ValorIPI = 0.0
            ValFrete = 0.0
            ValSeguro = 0.0
            ValDesconto = 0.0
            ValOutros = 0.0

            Dim Operacao As Integer = 0
            Dim SituacaoTributariaPISCOFINS As Integer = 0

            For Each enc In item.Encargos
                Operacao = enc.CodigoOperacao
                If enc.Codigo = "ICMS" Then
                    BaseIcms = enc.Base
                    If enc.PercentualExibicao > 0 Then
                        AliIcms = enc.PercentualExibicao
                    Else
                        AliIcms = enc.Percentual
                    End If
                    ValorIcms = enc.Valor
                ElseIf enc.Codigo.Contains("PRODUTO") Then
                    BaseProduto = enc.Valor
                    SituacaoTributaria = enc.SituacaoTributaria
                    SituacaoTributariaPISCOFINS = enc.SituacaoTributariaPISCOFINS
                End If

                If enc.Codigo.Trim = "IPI" Then
                    BaseIPI = enc.Base
                    AliIPI = enc.Percentual
                    ValorIPI = enc.Valor
                End If

                If enc.Codigo.Contains("FRETES") Then
                    ValFrete = enc.Valor
                End If

                If enc.Codigo.Contains("SEGURO") Then
                    ValSeguro += enc.Valor
                End If

                If enc.Codigo.Contains("DESCONTOS") Then
                    ValDesconto = enc.Valor
                End If

                If enc.Codigo.Contains("ADUANEIRAS") OrElse enc.Codigo.Contains("DESP. ACESSORIA") Then
                    ValOutros += enc.Valor
                End If

                If enc.Codigo = "PIS" Then
                    BasePis = enc.Base
                    AliPis = enc.Percentual
                    ValorPis = enc.Valor
                End If

                If enc.Codigo = "COFINS" Then
                    BaseCofins = enc.Base
                    AliCofins = enc.Percentual
                    ValorCofins = enc.Valor
                End If
            Next

            sb.Append("   UTRIB   =" & undPedido & ControlChars.CrLf)
            Valor = item.QuantidadeFiscal.ToString("N4")
            sb.Append("   QTRIB   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            sb.Append("   CEANTRIB=SEM GTIN" & ControlChars.CrLf)

            Valor = CDec(item.Unitario).ToString("N10")
            sb.Append("   VUNTRIB =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)

            If ValFrete > 0 Then
                Valor = CDec(ValFrete).ToString("N2")
                sb.Append("   VFRETE  =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            If ValSeguro > 0 Then
                Valor = CDec(ValSeguro).ToString("N2")
                sb.Append("   VSEG    =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            If ValDesconto > 0 Then
                Valor = CDec(ValDesconto).ToString("N2")
                sb.Append("   VDESC   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            If ValOutros > 0 Then
                Valor = CDec(ValOutros).ToString("N2")
                sb.Append("  VOUTRO   =" & Valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            End If

            sb.Append("   INDTOT  =1" & ControlChars.CrLf)

            sb.Append(ControlChars.CrLf)

            sb.Append("[IMPOSTO]" & ControlChars.CrLf)

            sb.Append(ControlChars.CrLf)

            If SituacaoTributaria = 40 OrElse SituacaoTributaria = 140 OrElse
                SituacaoTributaria = 41 OrElse SituacaoTributaria = 141 OrElse SituacaoTributaria = 341 OrElse
                SituacaoTributaria = 50 OrElse SituacaoTributaria = 150 Then
                sb.Append("[ICMS40]" & ControlChars.CrLf)
                sb.Append("   ORIG   =0" & ControlChars.CrLf)
                sb.Append("   CST    =" & SituacaoTributaria & ControlChars.CrLf)
            End If

            sb.Append(ControlChars.CrLf)

            sb.Append("[PISALIQ]" & ControlChars.CrLf)
            sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
            sb.Append("   VBC    =" & BasePis.Replace(",", ".") & ControlChars.CrLf)
            Valor = CDec(AliPis).ToString("N2")
            sb.Append("   PPIS   =" & Valor.Replace(",", ".") & ControlChars.CrLf)
            sb.Append("   VPIS   =" & ValorPis.Replace(",", ".") & ControlChars.CrLf)

            sb.Append(ControlChars.CrLf)

            sb.Append("[COFINSALIQ]" & ControlChars.CrLf)
            sb.Append("   CST	  =" & SituacaoTributariaPISCOFINS.ToString("00") & ControlChars.CrLf)
            sb.Append("   VBC    =" & BaseCofins.Replace(",", ".") & ControlChars.CrLf)
            Valor = CDec(AliCofins).ToString("N2")
            sb.Append("   PCOFINS   =" & Valor.Replace(",", ".") & ControlChars.CrLf)
            sb.Append("   VCOFINS   =" & ValorCofins.Replace(",", ".") & ControlChars.CrLf)
        Next

        Return sb.ToString()
    End Function


    Public Shared Function getTextoCancelar(ByVal nf As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("NPROTOCOLO =" & nf.ProtocoloNota & ControlChars.CrLf)
        sb.Append("JUSTIFICATIVA=" & Funcoes.EliminarCaracteresEspeciaisNF(nf.ObservacaoCancelamento) & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function getTextoCancelarNFCe(ByVal nf As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFCE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("NPROTOCOLO =" & nf.ProtocoloNota & ControlChars.CrLf)
        sb.Append("JUSTIFICATIVA=" & Funcoes.EliminarCaracteresEspeciaisNF(nf.ObservacaoCancelamento) & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function CancelarNFe(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("cancel{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoCancelar(nf)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-cancel{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

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
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "101" Then
                Return True
            ElseIf strCodigo = "151" Then
                Return True
            ElseIf strCodigo = "420" Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Entre no Monitor de Notas, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
                Return False
            ElseIf strCodigo = "690" Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Não pode proceder com o Cancelamento."
                Return False
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Public Shared Function CancelarNFCe(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("cancelnfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoCancelarNFCe(nf)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-cancelnfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

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
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "101" Then
                Return True
            ElseIf strCodigo = "151" Then
                Return True
            ElseIf strCodigo = "420" Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Entre no Monitor de Notas, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
                Return False
            ElseIf strCodigo = "690" Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Não pode proceder com o Cancelamento."
                Return False
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Public Shared Function ImprimirNFeDanfe(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer, ByRef msgErro As String, Optional ByVal IsFirst As Boolean = False) As Boolean
        Dim Sqls As New ArrayList
        Dim pdf As New [Lib].Negocio.Fil()

        pdf.IUD = "I"
        pdf.NomeArquivo = String.Format("danfe{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        pdf.Texto = getTextoApenasDANFE(nf)
        pdf.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        Dim obj As [Lib].Negocio.Resp = Nothing

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While obj Is Nothing AndAlso Now < tempoLimite
            obj = [Lib].Negocio.DocumentoEletronico.GetResp(String.Format("resp-{0}", pdf.NomeArquivo))
            If obj Is Nothing Then
                System.Threading.Thread.Sleep(3000)
            End If
        End While

        If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(obj.Texto) Then
            Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4012" Then
                msgErro = strMsg
                Return False
            Else
                Return True
            End If
        Else
            msgErro = "Resultado=4009  Mensagem=Erro na impressão do DANFE erro: (Tempo de espera excedido, verifique a impressora de PDF!)"
            Return False
        End If
    End Function

    Public Shared Function ImprimirCTeDacte(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer, ByRef msgErro As String, Optional ByVal IsFirst As Boolean = False) As Boolean
        Dim Sqls As New ArrayList
        Dim pdf As New [Lib].Negocio.Fil()

        pdf.IUD = "I"
        pdf.NomeArquivo = String.Format("dacte{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        pdf.Texto = getTextoDACTE(nf)
        pdf.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        Dim obj As [Lib].Negocio.Resp = Nothing

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While obj Is Nothing AndAlso Now < tempoLimite
            obj = [Lib].Negocio.DocumentoEletronico.GetResp(String.Format("resp-{0}", pdf.NomeArquivo))
            If obj Is Nothing Then
                System.Threading.Thread.Sleep(3000)
            End If
        End While

        If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(obj.Texto) Then
            Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4012" Then
                msgErro = strMsg
                Return False
            Else
                Return True
            End If
        Else
            msgErro = "Resultado=4009  Mensagem=Erro na impressão do DACTE erro: (Tempo de espera excedido, verifique a impressora de PDF!)"
            Return False
        End If
    End Function

    Public Shared Function ImprimirNFeImpressora(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer, ByRef msgErro As String, Optional ByVal IsFirst As Boolean = False) As Boolean
        Dim Sqls As New ArrayList
        Dim pdf As New [Lib].Negocio.Fil()

        pdf.IUD = "I"
        pdf.NomeArquivo = String.Format("danfe{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        pdf.Texto = getTextoApenasImpressora(nf, Vias)
        pdf.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        Dim obj As [Lib].Negocio.Resp = Nothing

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While obj Is Nothing AndAlso Now < tempoLimite
            obj = [Lib].Negocio.DocumentoEletronico.GetResp(String.Format("resp-{0}", pdf.NomeArquivo))
            System.Threading.Thread.Sleep(3000)
        End While

        If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(obj.Texto) Then
            Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4012" Then
                msgErro = strMsg
                Return False
            Else
                Return True
            End If
        Else
            msgErro = "Resultado=4009  Mensagem=Erro na impressão do DANFE erro: (Tempo de espera excedido, verifique a impressora de PDF!)"
            Return False
        End If
    End Function

    Public Shared Function ImprimirNFe(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer, ByRef msgErro As String, Optional ByVal IsFirst As Boolean = False) As Boolean
        Dim Sqls As New ArrayList
        Dim pdf As New [Lib].Negocio.Fil()

        pdf.IUD = "I"
        pdf.NomeArquivo = String.Format("danfe{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        pdf.Texto = getTextoDANFE(nf, Vias)
        pdf.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        Dim obj As [Lib].Negocio.Resp = Nothing

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While obj Is Nothing AndAlso Now < tempoLimite
            obj = [Lib].Negocio.DocumentoEletronico.GetResp(String.Format("resp-{0}", pdf.NomeArquivo))
            System.Threading.Thread.Sleep(3000)
        End While

        If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(obj.Texto) Then
            Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4012" Then
                msgErro = strMsg
                Return False
            Else
                Return True
            End If
        Else
            msgErro = "Resultado=4009  Mensagem=Erro na impressão do DANFE erro: (Tempo de espera excedido, verifique a impressora de PDF!)"
            Return False
        End If
    End Function

    Public Shared Function ImprimirNFCe(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer, ByRef msgErro As String, Optional ByVal IsFirst As Boolean = False) As Boolean
        Dim Sqls As New ArrayList
        Dim pdf As New [Lib].Negocio.Fil()

        pdf.IUD = "I"
        pdf.NomeArquivo = String.Format("danfenfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        pdf.Texto = getTextoDANFCe(nf, Vias)
        pdf.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        Dim obj As [Lib].Negocio.Resp = Nothing

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While obj Is Nothing AndAlso Now < tempoLimite
            obj = [Lib].Negocio.DocumentoEletronico.GetResp(String.Format("resp-{0}", pdf.NomeArquivo))
            System.Threading.Thread.Sleep(3000)
        End While

        If obj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(obj.Texto) Then
            Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4012" Then
                msgErro = strMsg
                Return False
            Else
                Return True
            End If
        Else
            msgErro = "Resultado=4009  Mensagem=Erro na impressão do DANFE erro: (Tempo de espera excedido, verifique a impressora de PDF!)"
            Return False
        End If
    End Function

    Public Shared Function getTextoDANFE(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("IMPRESSORA=pdf_nfe" & ControlChars.CrLf)

        If UsuarioServidor.ImprimirDanfe AndAlso Vias > 0 Then

            If nf.CodigoEmpresa = "04854422000266" Then
                For j = 0 To Vias - 1
                    sb.Append("IMPRESSORA=\\10.1.1.71\HP LaserJet P1005" & ControlChars.CrLf)
                Next
            ElseIf Mid(nf.CodigoEmpresa = "04854422", 1, 8) Then
                For j = 0 To Vias - 1
                    sb.Append("IMPRESSORA=\\10.1.1.46\HP LaserJet M1120 MFP" & ControlChars.CrLf)
                Next
            Else
                For j = 0 To Vias - 1
                    sb.Append("IMPRESSORA=Laser" & ControlChars.CrLf)
                Next
            End If
        End If
        Return sb.ToString()
    End Function

    Public Shared Function getTextoDANFCe(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFCE=" & nf.ChaveNFE & ControlChars.CrLf)
        If UsuarioServidor.ImprimirDanfe AndAlso Vias > 0 Then
            sb.Append("IMPRESSORA=IMP-LOJA-FISCAL" & ControlChars.CrLf)
        End If
        sb.Append("TIPOIMPR=10" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function getTextoApenasDANFE(ByVal nf As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("IMPRESSORA=pdf_nfe" & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Private Shared Function getTextoDACTE(ByVal cte As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVECTE=" & cte.ChaveNFE & ControlChars.CrLf)
        sb.Append("IMPRESSORA=pdf_cte" & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Public Shared Function getTextoApenasImpressora(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Vias As Integer) As String
        Dim sb As New StringBuilder()
        If UsuarioServidor.ImprimirDanfe AndAlso Vias > 0 Then
            If nf.CodigoEmpresa = "04854422000266" Then
                For j = 0 To Vias - 1
                    sb.Append("IMPRESSORA=\\10.1.1.71\HP LaserJet P1005" & ControlChars.CrLf)
                Next
            ElseIf Mid(nf.CodigoEmpresa = "04854422", 1, 8) Then
                For j = 0 To Vias - 1
                    sb.Append("IMPRESSORA=\\10.1.1.46\HP LaserJet M1120 MFP" & ControlChars.CrLf)
                Next
            Else
                For j = 0 To Vias - 1
                    sb.Append("IMPRESSORA=Laser" & ControlChars.CrLf)
                Next
            End If
        End If

        Return sb.ToString()
    End Function

    Public Shared Function SendMailNFe(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, Optional ByVal compactado As Boolean = False) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("email{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoEmail(nf, compactado)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        Return True
    End Function

    Public Shared Function SendMailNFce(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, Optional ByVal compactado As Boolean = False) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("emailnfce{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoEmail(nf, compactado)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        Return True
    End Function

    Public Shared Function getTextoEmail(ByVal nf As [Lib].Negocio.NotaFiscal, Optional ByVal compactado As Boolean = False) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        Dim strCliente As String() = nf.Cliente.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
        For Each cliMail As String In strCliente
            sb.Append("DESTINATARIO=" & cliMail & ControlChars.CrLf)
        Next

        'ENVIAR NOTA FISCAL PARA TRANSPORTADOR EMPRESA RTGRÃOS - FURLAN - 02/07/2024
        If Left(nf.Empresa.Codigo, 8) = "24450490" AndAlso nf.EntradaSaida = eEntradaSaida.Saida AndAlso Not nf.Transportador Is Nothing AndAlso nf.Transportador.EmailNFE.Length > 0 Then
            Dim strTransportador As String() = nf.Transportador.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each cliMailTransportador As String In strTransportador
                sb.Append("DESTINATARIO=" & cliMailTransportador & ControlChars.CrLf)
            Next
        End If

        sb.Append("ASSUNTO=" & "NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & "Envio NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        If Not String.IsNullOrWhiteSpace(nf.Empresa.EmailNFE) Then
            Dim strEmpresa As String = nf.Empresa.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)(0).Trim()
            sb.Append("EMAILEMITENTE=" & strEmpresa & ControlChars.CrLf)
            sb.Append("NOMEEMITENTE=" & nf.Empresa.Nome & ControlChars.CrLf)
        End If
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & IIf(compactado, "SIM", "NAO") & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function CCeVolumes(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, ByVal Volumes As Boolean, ByVal Especie As Boolean, ByVal Marca As Boolean, ByVal Numeracao As Boolean, ByVal PesoBruto As Boolean, ByVal PesoLiquido As Boolean) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoVolumes(nf, Volumes, Especie, Marca, Numeracao, PesoBruto, PesoLiquido)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        Dim lstMsg As String()
        Dim resultado As String
        Dim mensagem As String
        Dim strCodigo As String
        Dim strMsg As String

        If resp IsNot Nothing Then
            lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "135" Then 'Evento registrado e vinculado a NF-e 
                Sqls.Clear()

                ''Arquivo
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "A")
                obj.SalvarSql(Sqls)
                'Impressora
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "I")
                obj.SalvarSql(Sqls)
                'Email
                If Not String.IsNullOrWhiteSpace(nf.Cliente.EmailNFE) Then
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("emailcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = getTextoEmailCCe(nf, False)
                    obj.SalvarSql(Sqls)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgErro = "Entre em contato com o Suporte, Erro na solicitação da impressão: " & HttpContext.Current.Session("ssMessage")
                End If

                Return True
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Public Shared Function getTextoVolumes(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Volumes As Boolean, ByVal Especie As Boolean, ByVal Marca As Boolean, ByVal Numeracao As Boolean, ByVal PesoBruto As Boolean, ByVal PesoLiquido As Boolean) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)

        If Volumes Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = qvol " & ControlChars.CrLf)
            sb.Append("nomepaicampo = vol " & ControlChars.CrLf)
            sb.Append("valornovo = " & nf.Quantidade & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If Especie Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = esp " & ControlChars.CrLf)
            sb.Append("nomepaicampo = vol " & ControlChars.CrLf)
            sb.Append("valornovo = " & nf.Especie & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If Marca Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = marca " & ControlChars.CrLf)
            sb.Append("nomepaicampo = vol " & ControlChars.CrLf)
            sb.Append("valornovo = " & nf.Marca & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If Numeracao Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = nvol " & ControlChars.CrLf)
            sb.Append("nomepaicampo = vol " & ControlChars.CrLf)
            sb.Append("valornovo = " & nf.Numero & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If PesoLiquido Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = pesol " & ControlChars.CrLf)
            sb.Append("nomepaicampo = vol " & ControlChars.CrLf)
            sb.Append("valornovo = " & nf.PesoLiquido & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If PesoBruto Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = pesob " & ControlChars.CrLf)
            sb.Append("nomepaicampo = vol " & ControlChars.CrLf)
            sb.Append("valornovo = " & nf.PesoBruto & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        Return sb.ToString()
    End Function

    Public Shared Function CCeProduto(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, ByRef gridProduto As System.Web.UI.WebControls.GridView) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoProduto(nf, gridProduto)


        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        Dim lstMsg As String()
        Dim resultado As String
        Dim mensagem As String
        Dim strCodigo As String
        Dim strMsg As String

        If resp IsNot Nothing Then
            lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "135" Then 'Evento registrado e vinculado a NF-e 
                Sqls.Clear()

                'Arquivo
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "A")
                obj.SalvarSql(Sqls)

                'Impressora
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "I")
                obj.SalvarSql(Sqls)

                'Email
                If Not String.IsNullOrWhiteSpace(nf.Cliente.EmailNFE) Then
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("emailcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = getTextoEmailCCe(nf, False)
                    obj.SalvarSql(Sqls)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgErro = "Entre em contato com o Suporte, Erro na solicitação da impressão: " & HttpContext.Current.Session("ssMessage")
                End If

                Return True
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Private Shared Function getTextoProduto(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef gridProduto As System.Web.UI.WebControls.GridView) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)

        Dim prd As String = String.Empty
        Dim nomeprd As String = String.Empty
        Dim unidade As String = String.Empty
        Dim ncm As String = String.Empty
        Dim obsProd As String = String.Empty
        Dim index As Integer = 0

        For Each row In gridProduto.Rows
            If CType(row.FindControl("chkProduto"), CheckBox).Checked Then
                index = CType(CType(row.FindControl("chkProduto"), CheckBox).NamingContainer, GridViewRow).RowIndex
                prd = CType(gridProduto.Rows(index).FindControl("txtProduto"), TextBox).Text
                nomeprd = CType(gridProduto.Rows(index).FindControl("txtNomeProduto"), TextBox).Text
                unidade = CType(gridProduto.Rows(index).FindControl("txtUnidade"), TextBox).Text

                sb.Append("[campo]" & ControlChars.CrLf)
                sb.Append("nomecampo = cprod " & ControlChars.CrLf)
                sb.Append("nomepaicampo = prod " & ControlChars.CrLf)
                sb.Append("valornovo = " & prd & ControlChars.CrLf)
                sb.Append("ordem = " & gridProduto.Rows(index).Cells(0).Text & ControlChars.CrLf)

                sb.Append("[campo]" & ControlChars.CrLf)
                sb.Append("nomecampo = xprod " & ControlChars.CrLf)
                sb.Append("nomepaicampo = prod " & ControlChars.CrLf)
                sb.Append("valornovo = " & nomeprd & ControlChars.CrLf)
                sb.Append("ordem = " & gridProduto.Rows(index).Cells(0).Text & ControlChars.CrLf)

                sb.Append("[campo]" & ControlChars.CrLf)
                sb.Append("nomecampo = ucom " & ControlChars.CrLf)
                sb.Append("nomepaicampo = prod " & ControlChars.CrLf)
                sb.Append("valornovo = " & unidade & ControlChars.CrLf)
                sb.Append("ordem = " & gridProduto.Rows(index).Cells(0).Text & ControlChars.CrLf)
            End If

            If CType(row.FindControl("chkNCM"), CheckBox).Checked Then
                index = CType(CType(row.FindControl("chkNCM"), CheckBox).NamingContainer, GridViewRow).RowIndex
                ncm = CType(gridProduto.Rows(index).FindControl("txtNcm"), TextBox).Text

                sb.Append("[campo]" & ControlChars.CrLf)
                sb.Append("nomecampo = ncm " & ControlChars.CrLf)
                sb.Append("nomepaicampo = prod " & ControlChars.CrLf)
                sb.Append("valornovo = " & ncm & ControlChars.CrLf)
                sb.Append("ordem = " & gridProduto.Rows(index).Cells(0).Text & ControlChars.CrLf)
            End If

            If CType(row.FindControl("chkObservacaoProd"), CheckBox).Checked Then
                index = CType(CType(row.FindControl("chkObservacaoProd"), CheckBox).NamingContainer, GridViewRow).RowIndex
                Dim obs As String = Funcoes.EliminarCaracteresEspeciaisNF(CType(gridProduto.Rows(index).FindControl("txtObservacaoProd"), TextBox).Text)
                obsProd = obs.ToUpper

                sb.Append("[campo]" & ControlChars.CrLf)
                sb.Append("nomecampo = infadprod " & ControlChars.CrLf)
                sb.Append("nomepaicampo = det " & ControlChars.CrLf)
                sb.Append("valornovo = " & obsProd.Trim & ControlChars.CrLf)
                sb.Append("ordem = " & gridProduto.Rows(index).Cells(0).Text & ControlChars.CrLf)
            End If
        Next
        Return sb.ToString()
    End Function

    Public Shared Function CCeObservacao(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, ByVal Observacao As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoObservacao(nf, Observacao)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(90)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        Dim lstMsg As String()
        Dim resultado As String
        Dim mensagem As String
        Dim strCodigo As String
        Dim strMsg As String

        If resp IsNot Nothing Then
            lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "135" Then 'Evento registrado e vinculado a NF-e 
                Sqls.Clear()

                ''Arquivo
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "A")
                obj.SalvarSql(Sqls)
                'Impressora
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "I")
                obj.SalvarSql(Sqls)
                'Email
                If Not String.IsNullOrWhiteSpace(nf.Cliente.EmailNFE) Then
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("emailcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = getTextoEmailCCe(nf, False)
                    obj.SalvarSql(Sqls)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgErro = "Entre em contato com o Suporte, Erro na solicitação da impressão: " & HttpContext.Current.Session("ssMessage")
                End If

                Return True
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Public Shared Function getTextoObservacao(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Observacao As String) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("[campo]" & ControlChars.CrLf)
        sb.Append("nomecampo = INFCPL " & ControlChars.CrLf)
        sb.Append("nomepaicampo = infAdic " & ControlChars.CrLf)
        sb.Append("valornovo = " & Observacao & ControlChars.CrLf)
        sb.Append("ordem = 1" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function CCeCIFFOB(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoCIFFOB(nf, nf.TipoFreteSefaz)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        Dim lstMsg As String()
        Dim resultado As String
        Dim mensagem As String
        Dim strCodigo As String
        Dim strMsg As String

        If resp IsNot Nothing Then
            lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "135" Then 'Evento registrado e vinculado a NF-e 
                Sqls.Clear()
                'Arquivo
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "A")
                obj.SalvarSql(Sqls)
                'Impressora
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "I")
                obj.SalvarSql(Sqls)
                'Email
                If Not String.IsNullOrWhiteSpace(nf.Cliente.EmailNFE) Then
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("emailcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = getTextoEmailCCe(nf, False)
                    obj.SalvarSql(Sqls)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgErro = "Entre em contato com o Suporte, Erro na solicitação da impressão: " & HttpContext.Current.Session("ssMessage")
                End If

                Return True
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Public Shared Function getTextoCIFFOB(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal tipo As Integer) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("[campo]" & ControlChars.CrLf)
        sb.Append("nomecampo = MODFRETE " & ControlChars.CrLf)
        sb.Append("nomepaicampo = TRANSP " & ControlChars.CrLf)
        sb.Append("valornovo = " & tipo & ControlChars.CrLf)
        sb.Append("ordem = 1" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function CCeTransportador(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, ByVal Transportador As Boolean, ByVal Codigo As String,
                                            ByVal Veiculo As Boolean, ByVal VeiculoPlaca As String, ByVal VeiculoEstado As Boolean, ByVal VeiculoUF As String,
                                            ByVal Reboque1 As Boolean, ByVal Reboque1Placa As String, ByVal Reboque1Estado As Boolean, ByVal Reboque1UF As String,
                                            ByVal Reboque2 As Boolean, ByVal Reboque2Placa As String, ByVal Reboque2Estado As Boolean, ByVal Reboque2UF As String,
                                            ByVal Reboque3 As Boolean, ByVal Reboque3Placa As String, ByVal Reboque3Estado As Boolean, ByVal Reboque3UF As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoTransportador(nf, Transportador, Codigo, Veiculo, VeiculoPlaca, VeiculoEstado, VeiculoUF,
                                          Reboque1, Reboque1Placa, Reboque1Estado, Reboque1UF,
                                          Reboque2, Reboque2Placa, Reboque2Estado, Reboque2UF,
                                          Reboque3, Reboque3Placa, Reboque3Estado, Reboque3UF)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(90)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        Dim lstMsg As String()
        Dim resultado As String
        Dim mensagem As String
        Dim strCodigo As String
        Dim strMsg As String

        If resp IsNot Nothing Then
            lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "135" Then 'Evento registrado e vinculado a NF-e 
                Sqls.Clear()
                'Arquivo
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "A")
                obj.SalvarSql(Sqls)
                'Impressora
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "I")
                obj.SalvarSql(Sqls)
                'Email
                If Not String.IsNullOrWhiteSpace(nf.Cliente.EmailNFE) Then
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("emailcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = getTextoEmailCCe(nf, False)
                    obj.SalvarSql(Sqls)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgErro = "Entre em contato com o Suporte, Erro na solicitação da impressão: " & HttpContext.Current.Session("ssMessage")
                End If

                Return True
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Public Shared Function getTextoTransportador(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Transportador As Boolean, ByVal Codigo As String,
                                                 ByVal Veiculo As Boolean, ByVal VeiculoPlaca As String, ByVal VeiculoEstado As Boolean, ByVal VeiculoUF As String,
                                                 ByVal Reboque1 As Boolean, ByVal Reboque1Placa As String, ByVal Reboque1Estado As Boolean, ByVal Reboque1UF As String,
                                                 ByVal Reboque2 As Boolean, ByVal Reboque2Placa As String, ByVal Reboque2Estado As Boolean, ByVal Reboque2UF As String,
                                                 ByVal Reboque3 As Boolean, ByVal Reboque3Placa As String, ByVal Reboque3Estado As Boolean, ByVal Reboque3UF As String) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)

        If Transportador Then
            Dim transp() As String = Codigo.ToString.Split("-")
            Dim objTransp As New [Lib].Negocio.Cliente(transp(0), transp(1))

            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = CNPJ " & ControlChars.CrLf)
            sb.Append("nomepaicampo = TRANSPORTA " & ControlChars.CrLf)
            sb.Append("valornovo = " & objTransp.Codigo & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)

            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = XNOME " & ControlChars.CrLf)
            sb.Append("nomepaicampo = TRANSPORTA " & ControlChars.CrLf)
            sb.Append("valornovo = " & objTransp.Nome & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)

            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = XENDER " & ControlChars.CrLf)
            sb.Append("nomepaicampo = TRANSPORTA " & ControlChars.CrLf)
            If objTransp.Numero > 0 Then
                sb.Append("valornovo = " & Trim(objTransp.Endereco.ToUpper) & ", " & objTransp.Numero & ControlChars.CrLf)
            Else
                sb.Append("valornovo = " & Trim(objTransp.Endereco.ToUpper) & ControlChars.CrLf)
            End If
            sb.Append("ordem = 1" & ControlChars.CrLf)

            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = XMUN " & ControlChars.CrLf)
            sb.Append("nomepaicampo = TRANSPORTA " & ControlChars.CrLf)
            sb.Append("valornovo = " & Trim(objTransp.Cidade.ToUpper) & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)

            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = UF " & ControlChars.CrLf)
            sb.Append("nomepaicampo = TRANSPORTA " & ControlChars.CrLf)
            sb.Append("valornovo = " & objTransp.CodigoEstado.ToUpper & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If Veiculo Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = PLACA " & ControlChars.CrLf)
            sb.Append("nomepaicampo = VEICTRANSP " & ControlChars.CrLf)
            sb.Append("valornovo = " & VeiculoPlaca & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If VeiculoEstado Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = UF " & ControlChars.CrLf)
            sb.Append("nomepaicampo = VEICTRANSP " & ControlChars.CrLf)
            sb.Append("valornovo = " & VeiculoUF.Substring(0, 2) & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If Reboque1 Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = PLACA " & ControlChars.CrLf)
            sb.Append("nomepaicampo = REBOQUE " & ControlChars.CrLf)
            sb.Append("valornovo = " & Reboque1Placa & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If Reboque1Estado Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = UF " & ControlChars.CrLf)
            sb.Append("nomepaicampo = REBOQUE " & ControlChars.CrLf)
            sb.Append("valornovo = " & Reboque1UF.Substring(0, 2) & ControlChars.CrLf)
            sb.Append("ordem = 1" & ControlChars.CrLf)
        End If

        If Reboque2 Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = PLACA " & ControlChars.CrLf)
            sb.Append("nomepaicampo = REBOQUE " & ControlChars.CrLf)
            sb.Append("valornovo = " & Reboque2Placa & ControlChars.CrLf)
            sb.Append("ordem = 2" & ControlChars.CrLf)
        End If

        If Reboque2Estado Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = UF " & ControlChars.CrLf)
            sb.Append("nomepaicampo = REBOQUE " & ControlChars.CrLf)
            sb.Append("valornovo = " & Reboque2UF.Substring(0, 2) & ControlChars.CrLf)
            sb.Append("ordem = 2" & ControlChars.CrLf)
        End If

        If Reboque3 Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = PLACA " & ControlChars.CrLf)
            sb.Append("nomepaicampo = REBOQUE " & ControlChars.CrLf)
            sb.Append("valornovo = " & Reboque3Placa & ControlChars.CrLf)
            sb.Append("ordem = 3" & ControlChars.CrLf)
        End If

        If Reboque3Estado Then
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = UF " & ControlChars.CrLf)
            sb.Append("nomepaicampo = REBOQUE " & ControlChars.CrLf)
            sb.Append("valornovo = " & Reboque3UF.Substring(0, 2) & ControlChars.CrLf)
            sb.Append("ordem = 3" & ControlChars.CrLf)
        End If

        Return sb.ToString()
    End Function

    Public Shared Function CCeCFOP(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoCFOP(nf)


        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        Dim lstMsg As String()
        Dim resultado As String
        Dim mensagem As String
        Dim strCodigo As String
        Dim strMsg As String

        If resp IsNot Nothing Then
            lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "135" Then 'Evento registrado e vinculado a NF-e 
                Sqls.Clear()

                'Arquivo
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "A")
                obj.SalvarSql(Sqls)

                'Impressora
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "I")
                obj.SalvarSql(Sqls)

                'Email
                If Not String.IsNullOrWhiteSpace(nf.Cliente.EmailNFE) Then
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("emailcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = getTextoEmailCCe(nf, False)
                    obj.SalvarSql(Sqls)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgErro = "Entre em contato com o Suporte, Erro na solicitação da impressão: " & HttpContext.Current.Session("ssMessage")
                End If

                Return True
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Private Shared Function getTextoCFOP(ByVal nf As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)

        Dim ncm As String = ""
        Dim obsProd As String = ""
        Dim index As Integer = 0

        sb.Append("[campo]" & ControlChars.CrLf)
        sb.Append("nomecampo = natop " & ControlChars.CrLf)
        sb.Append("nomepaicampo = ide " & ControlChars.CrLf)
        sb.Append("valornovo = " & nf.CFOP.Descricao & ControlChars.CrLf)
        sb.Append("ordem = 1" & ControlChars.CrLf)

        For Each item In nf.Itens
            sb.Append("[campo]" & ControlChars.CrLf)
            sb.Append("nomecampo = cfop " & ControlChars.CrLf)
            sb.Append("nomepaicampo = prod " & ControlChars.CrLf)
            sb.Append("valornovo = " & nf.CFOP.Codigo & ControlChars.CrLf)
            sb.Append("ordem = " & item.Sequencia & ControlChars.CrLf)
        Next
        Return sb.ToString()
    End Function

    Public Shared Function CCeLocalEmbarque(ByVal nf As [Lib].Negocio.NotaFiscal, ByRef msgErro As String, ByVal Endereco As String) As Boolean
        Dim Sqls As New ArrayList

        Dim obj As New [Lib].Negocio.Fil()
        obj.IUD = "I"
        obj.NomeArquivo = String.Format("carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
        obj.Texto = getTextoLocalEmbarque(nf, Endereco)
        obj.SalvarSql(Sqls)

        Dim Banco As New AcessaBanco
        If Not Banco.GravaBanco(Sqls) Then
            msgErro = HttpContext.Current.Session("ssMessage")
            Return False
        End If

        'AGUARDAR RESPOSTA SEFAZ
        Dim resp As [Lib].Negocio.Resp = Nothing
        Dim fileName As String = String.Format("resp-carta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While resp Is Nothing AndAlso Now < tempoLimite
            resp = GetResp(fileName)
            System.Threading.Thread.Sleep(3000)
        End While

        Dim lstMsg As String()
        Dim resultado As String
        Dim mensagem As String
        Dim strCodigo As String
        Dim strMsg As String

        If resp IsNot Nothing Then
            lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
            strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If String.IsNullOrWhiteSpace(strCodigo) Then
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                Return False
            ElseIf strCodigo = "135" Then 'Evento registrado e vinculado a NF-e 
                Sqls.Clear()

                ''Arquivo
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "A")
                obj.SalvarSql(Sqls)
                'Impressora
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("impcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                obj.Texto = getPrinterNFE(nf, "I")
                obj.SalvarSql(Sqls)
                'Email
                If Not String.IsNullOrWhiteSpace(nf.Cliente.EmailNFE) Then
                    obj.IUD = "I"
                    obj.NomeArquivo = String.Format("emailcarta{0:000000000}#{1}.txt", nf.Codigo, nf.CodigoEmpresa)
                    obj.Texto = getTextoEmailCCe(nf, False)
                    obj.SalvarSql(Sqls)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    msgErro = "Entre em contato com o Suporte, Erro na solicitação da impressão: " & HttpContext.Current.Session("ssMessage")
                End If

                Return True
            Else
                msgErro = String.Format("{0} - {1}", strCodigo, strMsg)
                Return False
            End If
        Else
            msgErro = "Sefaz não retornou nenhuma resposta, Consulte ou Reenvie a Nota para proceder com o Cancelamento."
            Return False
        End If
    End Function

    Public Shared Function getTextoLocalEmbarque(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal Endereco As String) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("[campo]" & ControlChars.CrLf)
        sb.Append("nomecampo = UFSAIDAPAIS " & ControlChars.CrLf)
        sb.Append("nomepaicampo = EXPORTA " & ControlChars.CrLf)
        sb.Append("valornovo = " & nf.LocalEmbarque.CodigoEstado & ControlChars.CrLf)
        sb.Append("ordem = 1" & ControlChars.CrLf)
        sb.Append("[campo]" & ControlChars.CrLf)
        sb.Append("nomecampo = XLOCEXPORTA " & ControlChars.CrLf)
        sb.Append("nomepaicampo = EXPORTA " & ControlChars.CrLf)
        sb.Append("valornovo = " & Endereco & ControlChars.CrLf)
        sb.Append("ordem = 1" & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Public Shared Function getPrinterNFE(ByVal nf As [Lib].Negocio.NotaFiscal, ByVal impressao As String) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        sb.Append("impressora=pdf_nfe" & ControlChars.CrLf)

        If UsuarioServidor.ImprimirDanfe Then
            If impressao = "I" Then
                sb.Append("impressora=Laser" & ControlChars.CrLf)
            End If
        End If

        Return sb.ToString()
    End Function

    Public Shared Function getTextoEmailCCe(ByVal nf As [Lib].Negocio.NotaFiscal, Optional ByVal compactado As Boolean = False) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nf.ChaveNFE & ControlChars.CrLf)
        Dim strCliente As String() = nf.Cliente.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
        For Each cliMail As String In strCliente
            sb.Append("DESTINATARIO=" & cliMail & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & "Carta de Correcao NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & "Envio Carta de Correcao NF-e Nr. " & nf.Codigo & ControlChars.CrLf)
        If Not String.IsNullOrWhiteSpace(nf.Empresa.EmailNFE) Then
            Dim strEmpresa As String = nf.Empresa.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)(0).Trim()
            sb.Append("EMAILEMITENTE=" & strEmpresa & ControlChars.CrLf)
            sb.Append("NOMEEMITENTE=" & nf.Empresa.Nome & ControlChars.CrLf)
        End If
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & IIf(compactado, "SIM", "NAO") & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Public Shared Function getTextoConsultaCadastro(ByVal CpfCnpj As String, ByVal Uf As String) As String
        Dim sb As New StringBuilder()

        Dim Tipo As String = IIf(CpfCnpj.Length = 14, "CNPJ", "CPF")

        sb.Append(Tipo & "=" & CpfCnpj & ControlChars.CrLf)
        sb.Append("UF=" & Uf & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Public Shared Function ConsultaCadastro(ByVal CpfCnpj As String, ByVal Uf As String, ByRef MsgDeErro As String) As String()
        Dim MsgRetorno() As String = {"", ""}

        Try


            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"

            Dim Empresa As Cliente = New Cliente(HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa"))

            obj.NomeArquivo = String.Format("cadastro1#{0}.txt", Empresa.Codigo)
            obj.Texto = [Lib].Negocio.DocumentoEletronico.getTextoConsultaCadastro(CpfCnpj, Uf)

            If String.IsNullOrWhiteSpace(obj.Texto) Then
                MsgDeErro = "Não foi possível construir o arquivo texto para emissão da nota fiscal!"
                '    Return False
            End If

            obj.SalvarSql(Sqls)


            Dim Banco As New AcessaBanco
            If Not Banco.GravaBanco(Sqls) Then
                MsgDeErro = HttpContext.Current.Session("ssMessage")
                'Return False
            End If

            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-cadastro1#{0:000000000}.txt", Empresa.Codigo)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(30)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                'System.Threading.Thread.Sleep(3000)
            End While

            If resp IsNot Nothing Then
                MsgRetorno = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
            Else
                MsgDeErro = "Falha no retorno da Sefaz, tente o reenvio da consulta de cadastro novamente."
            End If
        Catch ex As Exception
            'aux = False
            Throw New Exception(ex.Message)
        End Try
        Return MsgRetorno
    End Function

    Public Shared Function ManifestoNFe(ByVal pNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal pTipoManifesto As String, ByRef msgResultado As String) As Boolean

        Dim Sqls As New ArrayList
        Dim Sql As String = String.Empty

        Dim result As Boolean = False

        Try
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"

            obj.NomeArquivo = String.Format("evento{0:000000000}#{1}.txt", CInt(Mid(pNotaFiscal.ChaveNFE, 26, 9)), pNotaFiscal.CodigoEmpresa)
            obj.Texto = DocumentoEletronico.getTextoConsultar(pNotaFiscal)
            obj.Texto &= "  tipo=" & pTipoManifesto & ControlChars.CrLf
            If pTipoManifesto = "210240" Then obj.Texto &= "justificativa=" & pNotaFiscal.Observacoes
            obj.SalvarSql(Sqls)

            Dim Banco As New AcessaBanco
            If Not Banco.GravaBanco(Sqls) Then
                msgResultado = HttpContext.Current.Session("ssMessage")
                Return False
            End If

            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-evento{0:000000000}#{1}.txt", CInt(Mid(pNotaFiscal.ChaveNFE, 26, 9)), pNotaFiscal.CodigoEmpresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(30)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            Dim lstMsg As String()
            Dim resultado As String = String.Empty
            Dim mensagem As String = String.Empty
            Dim strCodigo As String = String.Empty
            Dim strMsg As String = String.Empty
            Dim Protocolo As String = String.Empty
            Dim DataEvento As String = String.Empty
            Dim HoraEvento As String = String.Empty

            If resp IsNot Nothing Then
                lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Protocolo = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()
                DataEvento = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("DATAEVENTO")).FirstOrDefault()
                HoraEvento = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("HORAEVENTO")).FirstOrDefault()

                strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                If String.IsNullOrWhiteSpace(strCodigo) Then
                    msgResultado = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                    result = False
                ElseIf strCodigo = "135" Then
                    msgResultado &= String.Format("{0} - {1}", strCodigo, strMsg) & " \n"
                    result = True
                ElseIf strCodigo = "573" Then
                    msgResultado &= String.Format("{0} - {1}", strCodigo, " Manifesto NFe já realizado!") & "\n"
                    result = True
                ElseIf strCodigo = "596" Then
                    msgResultado &= "Prazo para manisfesto da NFe expirado. \n"
                    result = False
                ElseIf strCodigo = "4036" Then
                    msgResultado = "Guardian em Modo de Contingência não pode ser usado para Emissão em Modo Normal."
                    result = False
                Else
                    msgResultado &= String.Format("{0} - {1}", strCodigo, strMsg) & " \n"
                    result = False
                End If

                Sqls.Clear()

                Sqls.Add("DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-evento{0:000000000}#{1}.txt", CInt(Mid(pNotaFiscal.ChaveNFE, 26, 9)), pNotaFiscal.CodigoEmpresa) & "';")

                If strCodigo = "135" OrElse strCodigo = "573" OrElse strCodigo = "596" Then
                    Sql = "SELECT ChaveNFe FROM ManifestoNFE" & vbCrLf &
                                                "Where ChaveNFe = '" & pNotaFiscal.ChaveNFE & "'" & vbCrLf &
                                                "and CodigoEvento = " & pTipoManifesto

                    Dim ds As New DataSet
                    ds = Banco.ConsultaDataSet(Sql, "ManifestoNFE")

                    If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                        Sql = "Insert into ManifestoNFE (ChaveNFe, CodigoEvento, Justificativa, " & vbCrLf &
                              "CodigoResultado, MensagemRetorno, Protocolo, " & vbCrLf &
                              "DataEvento, HoraEvento, UsuarioInclusao, UsuarioInclusaoDate)" & vbCrLf &
                              "values ('" & pNotaFiscal.ChaveNFE & "', " & pTipoManifesto & ", '" & pNotaFiscal.Observacoes & "'," & vbCrLf &
                              strCodigo & ",'" & msgResultado & "','" & Protocolo.Split("=")(1) & "'," & vbCrLf &
                              "'" & DataEvento.Split("=")(1) & "','" & HoraEvento.Split("=")(1) & "'," & vbCrLf &
                              "'" & pNotaFiscal.UsuarioInclusao & "','" & pNotaFiscal.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "')"
                    Else
                        Sql = "Update ManifestoNFE set Justificativa = '" & pNotaFiscal.Observacoes & "'," & vbCrLf &
                              "CodigoResultado = " & strCodigo & "," & vbCrLf &
                              "MensagemRetorno = '" & msgResultado & "'," & vbCrLf
                        If Protocolo.Split("=")(1) IsNot Nothing AndAlso Protocolo.Split("=")(1).Length > 0 Then
                            Sql &= "Protocolo = '" & Protocolo.Split("=")(1) & "'," & vbCrLf
                        End If
                        Sql &= "DataEvento = '" & DataEvento.Split("=")(1) & "'," & vbCrLf &
                              "HoraEvento = '" & HoraEvento.Split("=")(1) & "'," & vbCrLf &
                              "UsuarioInclusao = '" & pNotaFiscal.UsuarioInclusao & "'," & vbCrLf &
                              "UsuarioInclusaoDate = '" & pNotaFiscal.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                              "Where ChaveNFe = '" & pNotaFiscal.ChaveNFE & "'" & vbCrLf &
                              "and CodigoEvento = " & pTipoManifesto
                    End If

                    Sqls.Add(Sql)

                    If pTipoManifesto = "210240" Then
                        Sql = "Update DocumentoXML set Situacao = 2, RecusaNFE = 1" & vbCrLf &
                              "Where chave_id = '" & pNotaFiscal.ChaveNFE & "'"
                    Else
                        Sql = "Update DocumentoXML set Situacao = 1, RecusaNFE = 0" & vbCrLf &
                              "Where chave_id = '" & pNotaFiscal.ChaveNFE & "'"
                    End If

                    Sqls.Add(Sql)

                End If

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        result = True
                    Else
                        msgResultado = HttpContext.Current.Session("ssMessage")
                        result = False
                    End If
                End If
            Else
                msgResultado = "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação."
                Return False
            End If
        Catch ex As Exception
            msgResultado = String.Format("{0} - {1}", "", "Não foi possível manifestar a Nota Fiscal, verifique a Disponibilidade da Sefaz.")
            Return False
        End Try

        Return result
    End Function

    Public Shared Function ManifestoCTe(ByVal pNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal pTipoManifesto As String, ByRef msgResultado As String) As Boolean
        Dim Sqls As New ArrayList
        Dim Sql As String = String.Empty

        Dim result As Boolean = False

        Try
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"

            obj.NomeArquivo = String.Format("eventocte{0:000000000}#{1}.txt", pNotaFiscal.Codigo, pNotaFiscal.CodigoEmpresa)
            obj.Texto = DocumentoEletronico.getTextoConsultarCTe(pNotaFiscal)
            obj.Texto &= "  tpevento=" & pTipoManifesto & ControlChars.CrLf
            obj.Texto &= "justificativa=" & pNotaFiscal.Observacoes
            obj.SalvarSql(Sqls)

            Dim Banco As New AcessaBanco
            If Not Banco.GravaBanco(Sqls) Then
                msgResultado = HttpContext.Current.Session("ssMessage")
                Return False
            End If

            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-eventocte{0:000000000}#{1}.txt", pNotaFiscal.Codigo, pNotaFiscal.CodigoEmpresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(30)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            Dim lstMsg As String()
            Dim resultado As String = String.Empty
            Dim mensagem As String = String.Empty
            Dim strCodigo As String = String.Empty
            Dim strMsg As String = String.Empty
            Dim Protocolo As String = String.Empty

            If resp IsNot Nothing Then
                lstMsg = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                resultado = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                mensagem = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Protocolo = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()

                strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                If String.IsNullOrWhiteSpace(strCodigo) Then
                    msgResultado = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                    result = False
                ElseIf strCodigo = "135" Then
                    msgResultado &= String.Format("{0} - {1}", strCodigo, strMsg) & " \n"
                    result = True
                ElseIf strCodigo = "573" Then
                    msgResultado &= String.Format("{0} - {1}", strCodigo, " Manifesto CTe já realizado!") & "\n"
                    result = True
                ElseIf strCodigo = "596" Then
                    msgResultado &= "Prazo para manisfesto da CTe expirado. \n"
                    result = False
                ElseIf strCodigo = "4036" Then
                    msgResultado = "Guardian em Modo de Contingência não pode ser usado para Emissão em Modo Normal."
                    result = False
                Else
                    msgResultado &= String.Format("{0} - {1}", strCodigo, strMsg) & " \n"
                    result = False
                End If

                Sqls.Clear()

                Sqls.Add("DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-eventocte{0:000000000}#{1}.txt", pNotaFiscal.Codigo, pNotaFiscal.CodigoEmpresa) & "';")

                If strCodigo = "135" OrElse strCodigo = "573" OrElse strCodigo = "596" Then
                    Sql = "SELECT ChaveNFe FROM ManifestoNFE" & vbCrLf &
                                                "Where ChaveNFe = '" & pNotaFiscal.ChaveNFE & "'" & vbCrLf &
                                                "and CodigoEvento = " & pTipoManifesto

                    Dim ds As New DataSet
                    ds = Banco.ConsultaDataSet(Sql, "ManifestoNFE")

                    If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                        Sql = "Insert into ManifestoNFE (ChaveNFe, CodigoEvento, Justificativa, " & vbCrLf &
                              "CodigoResultado, MensagemRetorno, Protocolo, " & vbCrLf &
                              "DataEvento, HoraEvento, UsuarioInclusao, UsuarioInclusaoDate)" & vbCrLf &
                              "values ('" & pNotaFiscal.ChaveNFE & "', " & pTipoManifesto & ", '" & pNotaFiscal.Observacoes & "'," & vbCrLf &
                              strCodigo & ",'" & msgResultado & "','" & Protocolo.Split("=")(1) & "'," & vbCrLf &
                              "'" & Now.ToString("yyyy-MM-dd") & "','" & Now.ToString("HH:mm:ss") & "'," & vbCrLf &
                              "'" & pNotaFiscal.UsuarioInclusao & "','" & pNotaFiscal.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "')"
                    Else
                        Sql = "Update ManifestoNFE set Justificativa = '" & pNotaFiscal.Observacoes & "'," & vbCrLf &
                              "CodigoResultado = " & strCodigo & "," & vbCrLf &
                              "MensagemRetorno = '" & msgResultado & "'," & vbCrLf
                        If Protocolo.Split("=")(1) IsNot Nothing AndAlso Protocolo.Split("=")(1).Length > 0 Then
                            Sql &= "Protocolo = '" & Protocolo.Split("=")(1) & "'," & vbCrLf
                        End If
                        Sql &= "DataEvento = '" & Now.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                              "HoraEvento = '" & Now.ToString("HH:mm:ss") & "'," & vbCrLf &
                              "UsuarioInclusao = '" & pNotaFiscal.UsuarioInclusao & "'," & vbCrLf &
                              "UsuarioInclusaoDate = '" & pNotaFiscal.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                              "Where ChaveNFe = '" & pNotaFiscal.ChaveNFE & "'" & vbCrLf &
                              "and CodigoEvento = " & pTipoManifesto
                    End If

                    Sqls.Add(Sql)

                    If pTipoManifesto = "610110" Then
                        Sql = "Update DocumentoXML set Situacao = 2, RecusaNFE = 1" & vbCrLf &
                              "Where chave_id = '" & pNotaFiscal.ChaveNFE & "'"
                    Else
                        Sql = "Update DocumentoXML set Situacao = 1, RecusaNFE = 0" & vbCrLf &
                              "Where chave_id = '" & pNotaFiscal.ChaveNFE & "'"
                    End If

                    Sqls.Add(Sql)

                End If

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        result = True
                    Else
                        msgResultado = HttpContext.Current.Session("ssMessage")
                        result = False
                    End If
                End If
            Else
                msgResultado = "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação."
                Return False
            End If

        Catch ex As Exception
            msgResultado = String.Format("{0} - {1}", "", "Não foi possível manifestar a Nota Fiscal, verifique a Disponibilidade da Sefaz.")
            Return False
        End Try

        Return result

    End Function

    Public Shared Function DownloadXml(ByVal pNotaFiscal As [Lib].Negocio.NotaFiscal, ByRef msgResultado As String) As String
        Dim result As Boolean = False
        'Empresa  do usuário logado
        Dim Empresa As New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
        Dim Sqls As New ArrayList
        Try
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("downloadnfe{0:000000000}#{1}.txt", pNotaFiscal.Codigo, pNotaFiscal.CodigoEmpresa)
            obj.Texto = getTextoConsultar(pNotaFiscal)
            obj.Texto &= "salvarem=" & pNotaFiscal.Empresa.Empresa.PathDownloadNFe
            obj.SalvarSql(Sqls)
            Dim Banco As New AcessaBanco
            If Not Banco.GravaBanco(Sqls) Then
                msgResultado = HttpContext.Current.Session("ssMessage")
                Return False
            End If

            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-downloadnfe{0:000000000}#{1}.txt", pNotaFiscal.Codigo, pNotaFiscal.CodigoEmpresa)

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(30)

            While resp Is Nothing AndAlso Now < tempoLimite
                resp = [Lib].Negocio.DocumentoEletronico.GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If resp IsNot Nothing Then
                Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                If String.IsNullOrWhiteSpace(strCodigo) Then
                    msgResultado = String.Format("{0} - {1}", strCodigo, strMsg) & " - Erro na Consulta do Guardian, verifique o Servidor de NFE."
                    result = False
                ElseIf strCodigo = "140" Then
                    msgResultado &= String.Format("{0} - {1}", strCodigo, strMsg) & " \n"
                    result = True
                Else
                    msgResultado &= String.Format("{0} - {1}", strCodigo, strMsg) & " \n"
                    result = False
                End If

                Sqls.Clear()

                Sqls.Add("DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-downloadnfe{0:000000000}#{1}.txt", pNotaFiscal.Codigo, pNotaFiscal.CodigoEmpresa) & "';")

                If Sqls.Count > 0 Then Banco.GravaBanco(Sqls)
            Else
                msgResultado = "Verifique o Guardian no Servidor, não foi recebida nenhuma resposta do Modo de Operação."
            End If
        Catch ex As Exception
            msgResultado = ex.ToString
            Return False
        End Try
        Return result
    End Function

    Private Shared Function GetCnpjOuCpfEmitente(row As DataRow) As String
        If row.Table.Columns.Contains("CNPJ") AndAlso Not IsDBNull(row("CNPJ")) Then
            Return row("CNPJ").ToString().Trim()
        ElseIf row.Table.Columns.Contains("CPF") AndAlso Not IsDBNull(row("CPF")) Then
            Return row("CPF").ToString().Trim()
        End If

        Return String.Empty
    End Function

    'Importação de xml de entrada
    Public Shared Sub PreencherNFeComXML(ByRef objNotaFiscal As NotaFiscal, ByVal DsXml As DataSet, ByVal emissaoNossaImportacaoXML As Boolean, ByVal usarProdutoXML As Boolean, ByVal notaDeTerceiro As Boolean, ByVal pedido As String, ByVal bImportarProdutoUnico As Boolean, ByVal bAGruparNCM As Boolean)

        If Not objNotaFiscal Is Nothing AndAlso DsXml.Tables.Count > 0 Then
            Try

                '    If Not DsXml.Tables("dest").Rows(0)("CNPJ").ToString().Equals(objNotaFiscal.CodigoEmpresa) Then
                '        MsgBox(Me.Page, "Empresa Selecionada não pode ser diferente do destinatário da NFe! Destinatário da NFe: " & DsXml.Tables("dest").Rows(0)("CNPJ").ToString())
                '        Exit Function
                '    End If

                If emissaoNossaImportacaoXML Then
                    Dim emitRow = DsXml.Tables("emit").Rows(0)
                    objNotaFiscal.CodigoEmpresa = GetCnpjOuCpfEmitente(emitRow)
                    objNotaFiscal.EnderecoEmpresa = 0
                    objNotaFiscal.Empresa = Nothing
                Else
                    Dim destRow = DsXml.Tables("dest").Rows(0)
                    objNotaFiscal.CodigoEmpresa = GetCnpjOuCpfEmitente(destRow)
                    objNotaFiscal.EnderecoEmpresa = 0
                    objNotaFiscal.Empresa = Nothing
                End If

                If notaDeTerceiro Then

                    Dim ehCTe As Boolean = False

                    If DsXml.Tables.Contains("infProt") AndAlso DsXml.Tables("infProt").Rows.Count > 0 Then
                        Dim tblProt As DataTable = DsXml.Tables("infProt")

                        If tblProt.Columns.Contains("chCTe") Then
                            ehCTe = True
                        ElseIf tblProt.Columns.Contains("chNFe") Then
                            ehCTe = False
                        End If
                    End If

                    If ehCTe Then
                        Throw New Exception("Não é possível importar um CTe como nota de terceiro!")
                    End If

                    'Se nota terceiro, precisamos ver se a empresa da nota de terceiro existe
                    'Se não, precisamos incluir
                    Dim ClienteEmpresa As New ListCliente("", Nothing, objNotaFiscal.CodigoEmpresa)

                    If ClienteEmpresa.Count() = 0 Then

                        Dim verCliente As String = Funcoes.EliminarCaracteresEspeciais(objNotaFiscal.CodigoEmpresa)

                        IncluirDadosNovoClienteDestinatario(objNotaFiscal, DsXml, ClienteEmpresa, verCliente)

                    End If

                End If

                Dim tentativas As Integer = 0

                Do
                    tentativas += 1

                    Dim temEmitente As Boolean = False

                    'Verificar o cadastro do cliente no maximo 2 vezes

                    If emissaoNossaImportacaoXML Then

                        If DsXml.Tables("dest").Columns.Contains("CNPJ") Then
                            objNotaFiscal.CodigoCliente = DsXml.Tables("dest").Rows(0)("CNPJ").ToString()
                            temEmitente = True
                        ElseIf DsXml.Tables("dest").Columns.Contains("CPF") Then
                            objNotaFiscal.CodigoCliente = DsXml.Tables("dest").Rows(0)("CPF").ToString()
                            temEmitente = True
                        End If

                        If temEmitente Then

                            Dim EmitenteCidadeTemp As String = DsXml.Tables("enderDest").Rows(0)("xMun").ToString()
                            Dim EmitenteUFTemp As String = DsXml.Tables("enderDest").Rows(0)("UF").ToString()
                            Dim EmitenteEnderecoTemp As String = DsXml.Tables("enderDest").Rows(0)("xLgr").ToString()

                            Dim EmitenteEnd As Integer = 0
                            Dim temEndEmitente As Boolean = False
                            Dim ClienteTemp As New ListCliente("", Nothing, objNotaFiscal.CodigoCliente)

                            If ClienteTemp.Count() = 0 Then

                                Dim verCliente As String = Funcoes.EliminarCaracteresEspeciais(objNotaFiscal.CodigoCliente)

                                IncluirDadosNovoClienteDestinatario(objNotaFiscal, DsXml, ClienteTemp, verCliente)

                            Else

                                For Each c In ClienteTemp
                                    'c.CodigoEndereco = objNotaFiscal.EnderecoCliente AndAlso AndAlso c.Endereco.Contains(EmitenteEnderecoTemp) - Removido - Furlan - 11/01/2024

                                    If DsXml.Tables("dest").Columns.Contains("IE") AndAlso NormalizarIE(c.InscricaoEstadual) = NormalizarIE(DsXml.Tables("dest").Rows(0)("IE").ToString()) Then
                                        EmitenteEnd = c.CodigoEndereco
                                        temEndEmitente = True
                                        Exit For
                                    ElseIf Not DsXml.Tables("dest").Columns.Contains("IE") AndAlso c.CodigoEstado = EmitenteUFTemp AndAlso c.Cidade.ToUpper() = EmitenteCidadeTemp.ToUpper() Then
                                        EmitenteEnd = c.CodigoEndereco
                                        temEndEmitente = True
                                        Exit For
                                    ElseIf ClienteTemp.Count() = 1 AndAlso c.InscricaoEstadual.Length = 0 Then
                                        c.InscricaoEstadual = DsXml.Tables("dest").Rows(0)("IE").ToString()
                                        EmitenteEnd = c.CodigoEndereco
                                        temEndEmitente = True
                                        c.IUD = "U"
                                        If Not c.Salvar Then
                                            Throw New Exception("Não foi possível atualizar dados do cliente!")
                                        End If

                                    End If
                                Next

                                If temEndEmitente Then

                                    objNotaFiscal.EnderecoCliente = EmitenteEnd

                                Else

                                    'Incluir cliente novo
                                    Dim verCliente As String = Funcoes.EliminarCaracteresEspeciais(objNotaFiscal.CodigoCliente)
                                    IncluirDadosNovoClienteDestinatario(objNotaFiscal, DsXml, ClienteTemp, verCliente)

                                End If

                            End If

                        End If

                    Else

                        If DsXml.Tables("emit").Columns.Contains("CNPJ") Then
                            objNotaFiscal.CodigoCliente = DsXml.Tables("emit").Rows(0)("CNPJ").ToString()
                            temEmitente = True
                        ElseIf DsXml.Tables("emit").Columns.Contains("CPF") Then
                            objNotaFiscal.CodigoCliente = DsXml.Tables("emit").Rows(0)("CPF").ToString()
                            temEmitente = True
                        End If


                        If temEmitente Then
                            Dim EmitenteCidadeTemp As String = DsXml.Tables("Enderemit").Rows(0)("xMun").ToString()
                            Dim EmitenteUFTemp As String = DsXml.Tables("Enderemit").Rows(0)("UF").ToString()
                            Dim EmitenteEnderecoTemp As String = DsXml.Tables("Enderemit").Rows(0)("xLgr").ToString()

                            Dim EmitenteEnd As Integer = 0
                            Dim temEndEmitente As Boolean = False
                            Dim ClienteTemp As New ListCliente("", Nothing, objNotaFiscal.CodigoCliente)

                            For Each c In ClienteTemp
                                'c.CodigoEndereco = objNotaFiscal.EnderecoCliente AndAlso AndAlso c.Endereco.Contains(EmitenteEnderecoTemp) - Removido - Furlan - 11/01/2024

                                If DsXml.Tables("emit").Columns.Contains("IE") AndAlso NormalizarIE(c.InscricaoEstadual) = NormalizarIE(DsXml.Tables("emit").Rows(0)("IE").ToString()) Then
                                    EmitenteEnd = c.CodigoEndereco
                                    temEndEmitente = True
                                    Exit For
                                ElseIf Not DsXml.Tables("emit").Columns.Contains("IE") AndAlso c.CodigoEstado = EmitenteUFTemp AndAlso c.Cidade.ToUpper() = EmitenteCidadeTemp.ToUpper() Then
                                    EmitenteEnd = c.CodigoEndereco
                                    temEndEmitente = True
                                    Exit For
                                ElseIf ClienteTemp.Count() = 1 AndAlso c.InscricaoEstadual.Length = 0 Then
                                    c.InscricaoEstadual = DsXml.Tables("emit").Rows(0)("IE").ToString()
                                    EmitenteEnd = c.CodigoEndereco
                                    temEndEmitente = True
                                    c.IUD = "U"
                                    If Not c.Salvar Then
                                        Throw New Exception("Não foi possível atualizar dados do cliente!")
                                    End If
                                End If
                            Next

                            If temEndEmitente Then

                                objNotaFiscal.EnderecoCliente = EmitenteEnd

                            Else

                                'Cadastrar cliente com a inscrição do XML
                                'Incluir cliente novo
                                Dim verCliente As String = Funcoes.EliminarCaracteresEspeciais(objNotaFiscal.CodigoCliente)
                                IncluirDadosNovoClienteEmitente(objNotaFiscal, DsXml, ClienteTemp, verCliente)

                            End If

                        End If

                    End If

                Loop While tentativas < 2

                '**************
                If DsXml.Tables.Contains("infProt") AndAlso DsXml.Tables("infProt").Rows.Count > 0 Then
                    'Chave NFe
                    objNotaFiscal.ChaveNFE = DsXml.Tables("infProt").Rows(0)("chNFe").ToString()

                Else

                    If DsXml.Tables.Contains("infNFE") Then
                        Dim infNFE As DataTable = DsXml.Tables("infNFE")

                        ' Procura a coluna chNFe na tabela protNFe
                        If infNFE.Columns.Contains("Id") Then
                            ' Itera sobre as linhas da tabela para obter o valor de chNFe
                            For Each row As DataRow In infNFE.Rows
                                objNotaFiscal.ChaveNFE = row("Id").ToString().ToUpper().Replace("NFE", "")
                                Exit For
                            Next
                        End If

                    Else

                        Throw New Exception("Não foi possível encontrar a tag chNFe!")

                    End If

                End If

                Dim chaveNFEOrigem As String = ""

                If DsXml.Tables.Contains("NFref") Then
                    Dim nfRef As DataTable = DsXml.Tables("NFref")

                    If nfRef.Columns.Contains("refNFe") Then
                        For Each row As DataRow In nfRef.Rows

                            chaveNFEOrigem = row("refNFe").ToString().Trim()

                            Dim nf As New [Lib].Negocio.NotaFiscal()
                            If Not (objNotaFiscal.NotaOrigemImportacaoXML IsNot Nothing AndAlso objNotaFiscal.NotaOrigemImportacaoXML.Count > 0) Then
                                objNotaFiscal.NotaOrigemImportacaoXML = New List(Of [Lib].Negocio.NotaFiscal)
                            End If

                            nf.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                            nf.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                            nf.CarregarNotaComChaveXML(chaveNFEOrigem)

                            If nf.Codigo > 0 Then

                                If objNotaFiscal.NotaOrigemImportacaoXML.Count = 0 Or objNotaFiscal.NotaOrigemImportacaoXML.Where(Function(x) x.Empresa.Codigo = nf.Empresa.Codigo And x.Cliente.Codigo = nf.Cliente.Codigo And x.Codigo = nf.Codigo).Count = 0 Then
                                    objNotaFiscal.NotaOrigemImportacaoXML.Add(nf)
                                End If

                            End If

                        Next

                    End If
                End If

                'Codigo NFe
                objNotaFiscal.Codigo = DsXml.Tables("ide").Rows(0)("nNF").ToString()
                'Série
                objNotaFiscal.Serie = DsXml.Tables("ide").Rows(0)("serie").ToString()

                'SegCodBarras
                If DsXml.Tables("ide").Rows(0)("tpEmis").ToString.Contains("2,4,5") Then
                    objNotaFiscal.TemSegCodBarra = True
                Else
                    objNotaFiscal.TemSegCodBarra = False
                End If

                'Emissão
                If DsXml.Tables("ide").Columns.Contains("dEmi") Then
                    objNotaFiscal.DataNota = DsXml.Tables("ide").Rows(0)("dEmi").ToString()
                    objNotaFiscal.DataInclusao = objNotaFiscal.DataNota
                    objNotaFiscal.Movimento = objNotaFiscal.DataNota
                    objNotaFiscal.DataTermino = objNotaFiscal.DataNota
                Else
                    Dim strData As String = DsXml.Tables("ide").Rows(0)("dhEmi").ToString()
                    objNotaFiscal.DataNota = strData.Remove(strData.Length - 6)
                    objNotaFiscal.DataInclusao = objNotaFiscal.DataNota
                    objNotaFiscal.Movimento = objNotaFiscal.DataNota
                    objNotaFiscal.DataTermino = objNotaFiscal.DataNota
                End If

                'Situação
                objNotaFiscal.CodigoSituacao = eSituacao.Normal
                'Tipo Documento
                objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.Nota
                'Entrada Saida
                objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada
                'Eletronica
                objNotaFiscal.Eletronica = True
                'Emitente

                ' Verifica se a tabela "ICMSTot" existe no DataSet
                If DsXml.Tables.Contains("ICMSTot") Then
                    Dim icmsTotTable As DataTable = DsXml.Tables("ICMSTot")

                    ' Verifica se há pelo menos uma linha na tabela "ICMSTot"
                    If icmsTotTable.Rows.Count > 0 Then
                        Dim icmsTotRow As DataRow = icmsTotTable.Rows(0)

                        ' Extrai os valores das tags vProd e vNF
                        Dim vProd As String = If(icmsTotRow.Table.Columns.Contains("vProd") AndAlso Not IsDBNull(icmsTotRow("vProd")), icmsTotRow("vProd").ToString(), "")
                        Dim vNF As String = If(icmsTotRow.Table.Columns.Contains("vNF") AndAlso Not IsDBNull(icmsTotRow("vNF")), icmsTotRow("vNF").ToString(), "")

                        Dim vTotalBaseICMSXML As String = If(icmsTotRow.Table.Columns.Contains("vBC") AndAlso Not IsDBNull(icmsTotRow("vBC")), icmsTotRow("vBC").ToString(), "")
                        Dim vTotalICMSXML As String = If(icmsTotRow.Table.Columns.Contains("vICMS") AndAlso Not IsDBNull(icmsTotRow("vICMS")), icmsTotRow("vICMS").ToString(), "")

                        ' Exemplo: Atribuir os valores a outras variáveis
                        Dim valorProduto As Decimal = Decimal.Parse(vProd.Replace(".", ",")) ' Converte para Decimal
                        Dim valorNotaFiscal As Decimal = Decimal.Parse(vNF.Replace(".", ",")) ' Converte para Decimal
                        Dim valorBase As Decimal = Decimal.Parse(vTotalBaseICMSXML.Replace(".", ",")) ' Converte para Decimal
                        Dim valorICMS As Decimal = Decimal.Parse(vTotalICMSXML.Replace(".", ",")) ' Converte para Decimal

                        objNotaFiscal.TotalBaseICMSXML = valorBase
                        objNotaFiscal.TotalICMSXML = valorICMS
                        objNotaFiscal.DiferencaValorNFXProdutoXML = valorNotaFiscal - valorProduto

                    End If
                End If

                If bImportarProdutoUnico Then
                    ImportarProdutoUnico(objNotaFiscal, DsXml, usarProdutoXML)
                ElseIf bAGruparNCM Then
                    AgruparNCM(objNotaFiscal, DsXml, usarProdutoXML)
                ElseIf objNotaFiscal.NotaDeTerceiro Then
                    ImportarProdutoUnico(objNotaFiscal, DsXml, usarProdutoXML)
                Else
                    ImportarProdutoXML(objNotaFiscal, DsXml, usarProdutoXML)
                End If

                '****OBS
                If Not DsXml.Tables("infAdic") Is Nothing Then
                    If DsXml.Tables("infAdic").Columns.Contains("infAdFisco") Then
                        objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciaisNF(DsXml.Tables("infAdic").Rows(0)("infAdFisco"))
                    End If

                    If DsXml.Tables("infAdic").Columns.Contains("infCpl") Then
                        objNotaFiscal.ObservacoesDeEmbarque = Funcoes.EliminarCaracteresEspeciaisNF(DsXml.Tables("infAdic").Rows(0)("infCpl"))
                        objNotaFiscal.ObservacoesDeEmbarque = objNotaFiscal.ObservacoesDeEmbarque.Replace("'", "")
                    End If

                    If DsXml.Tables.Contains("infRespTec") Then
                        Dim infRespTecTable As DataTable = DsXml.Tables("infRespTec")

                        ' Itera sobre as linhas da tabela infRespTec
                        For Each infRespTecRow As DataRow In infRespTecTable.Rows
                            ' Lê os valores das tags dentro de infRespTec
                            Dim cnpj As String = infRespTecRow("CNPJ").ToString()

                            If Not cnpj Is Nothing AndAlso cnpj = "11342233000199" And Left(objNotaFiscal.CodigoEmpresa, 8) = "24450490" Then
                                objNotaFiscal.MicSistemas = True
                                Exit For
                            ElseIf Not cnpj Is Nothing AndAlso cnpj = "02439327000190" And Left(objNotaFiscal.CodigoEmpresa, 8) = "62780383" Then
                                objNotaFiscal.MicSistemas = True
                                Exit For
                            ElseIf Not cnpj Is Nothing AndAlso cnpj = "02439327000190" And Left(objNotaFiscal.CodigoEmpresa, 8) = "63358210" Then
                                objNotaFiscal.MicSistemas = True
                                Exit For
                            End If

                        Next
                    End If

                    Dim infAdicRow As DataRow = DsXml.Tables("infAdic").Rows(0)

                    ' Itera sobre os nós filhos de infAdic para encontrar obsCont
                    If infAdicRow.Table.ChildRelations.Count > 0 Then
                        Dim obsContTable As DataTable = DsXml.Tables("obsCont")

                        If objNotaFiscal.MicSistemas AndAlso Left(objNotaFiscal.CodigoEmpresa, 8) = "62780383" AndAlso CInt(pedido) > 0 Then
                            objNotaFiscal.CodigoPedidoMIC = pedido
                        Else
                            If Not obsContTable Is Nothing Then
                                For Each obsContRow As DataRow In obsContTable.Rows
                                    ' Verifica se o valor de xCampo é "NR_PEDIDO"
                                    If obsContRow("xCampo").ToString() = "NR_PEDIDO" Then
                                        Dim nrPedido As String = obsContRow("xTexto").ToString()

                                        objNotaFiscal.CodigoPedidoMIC = nrPedido

                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If
                '********

                If Not DsXml.Tables("transp") Is Nothing Then
                    '*** Transporte
                    If CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 0 OrElse CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 3 Then
                        objNotaFiscal.CIFFOB = eTiposFrete.CIF
                    ElseIf CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 1 OrElse CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 4 Then
                        objNotaFiscal.CIFFOB = eTiposFrete.FOB
                    ElseIf CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 2 Then
                        objNotaFiscal.CIFFOB = eTiposFrete.TER
                    ElseIf CInt(DsXml.Tables("transp").Rows(0)("modfrete")) = 9 Then
                        objNotaFiscal.CIFFOB = eTiposFrete.NEN
                    End If
                End If

                If Not DsXml.Tables("transporta") Is Nothing Then

                    Dim temTransporte As Boolean = False

                    If DsXml.Tables("transporta").Columns.Contains("CNPJ") Then
                        temTransporte = True
                        objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CNPJ").ToString()
                    ElseIf DsXml.Tables("transporta").Columns.Contains("CPF") Then
                        temTransporte = True
                        objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CPF").ToString()
                    End If

                    If temTransporte Then

                        Dim cTransportador As String = String.Empty
                        If DsXml.Tables("transporta").Columns.Contains("CNPJ") Then
                            cTransportador = DsXml.Tables("transporta").Rows(0)("CNPJ").ToString()
                            objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CNPJ").ToString()
                        ElseIf DsXml.Tables("transporta").Columns.Contains("CPF") Then
                            cTransportador = DsXml.Tables("transporta").Rows(0)("CPF").ToString()
                            objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CPF").ToString()
                        End If

                        Dim TranspCodigo As New ListCliente("", Nothing, cTransportador)

                        Dim TranpCidadeTemp As String = String.Empty
                        If DsXml.Tables("transporta").Columns.Contains("xMun") Then
                            TranpCidadeTemp = DsXml.Tables("transporta").Rows(0)("xMun")
                        End If

                        Dim TranpEstadoTemp As String = String.Empty
                        If DsXml.Tables("transporta").Columns.Contains("UF") Then
                            TranpEstadoTemp = DsXml.Tables("transporta").Rows(0)("UF")
                        End If

                        Dim TranspEnderecoTemp As String = String.Empty
                        If DsXml.Tables("transporta").Columns.Contains("xEnder") Then
                            TranspEnderecoTemp = DsXml.Tables("transporta").Rows(0)("xEnder")
                        End If

                        Dim TranspEnd As Integer = 0
                        Dim temTransportador As Boolean = False

                        For Each c In TranspCodigo
                            If TranpEstadoTemp.Length > 0 AndAlso TranpCidadeTemp.Length > 0 AndAlso TranspEnderecoTemp.Length > 0 Then
                                'AndAlso c.Endereco.Contains(TranspEnderecoTemp) - Removido - Furlan - 11/01/2024
                                If c.CodigoEstado = TranpEstadoTemp AndAlso c.Cidade.ToUpper() = TranpCidadeTemp.ToUpper() Then
                                    TranspEnd = c.CodigoEndereco

                                    temTransportador = True
                                End If
                            End If
                        Next

                        If temTransportador Then
                            'If DsXml.Tables("transporta").Columns.Contains("CNPJ") Then
                            '    objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CNPJ").ToString()
                            'ElseIf DsXml.Tables("transporta").Columns.Contains("CPF") Then
                            '    objNotaFiscal.CodigoTransportador = DsXml.Tables("transporta").Rows(0)("CPF").ToString()
                            'End If
                            objNotaFiscal.EnderecoTransportador = TranspEnd
                        End If

                        'If objNotaFiscal.Transportador Is Nothing Then
                        '    Throw New Exception("Transportador não foi encontrado, verifique!")
                        '    Exit Sub
                        'ElseIf objNotaFiscal.Transportador.Nome Is Nothing Then
                        '    Throw New Exception("Transportador não foi encontrado, verifique!")
                        '    Exit Sub
                        'End If
                    End If
                End If

                If Not DsXml.Tables("veicTransp") Is Nothing AndAlso Not DsXml.Tables("veicTransp").Rows(0)("placa") Is Nothing Then

                    objNotaFiscal.PlacaTransportador = DsXml.Tables("veicTransp").Rows(0)("placa")
                    If objNotaFiscal.PlacaDetalhes Is Nothing OrElse objNotaFiscal.PlacaDetalhes.Placa01.Length = 0 Then
                        objNotaFiscal.PlacaDetalhes = New [Lib].Negocio.Placa
                    End If

                    objNotaFiscal.PlacaDetalhes.Placa01 = DsXml.Tables("veicTransp").Rows(0)("placa")

                    If DsXml.Tables("veicTransp").Columns.Contains("UF") Then
                        objNotaFiscal.PlacaDetalhes.EstadoPlaca01 = DsXml.Tables("veicTransp").Rows(0)("UF")
                    Else
                        objNotaFiscal.PlacaDetalhes.EstadoPlaca01 = ""
                    End If

                    If Not DsXml.Tables("reboque") Is Nothing Then
                        objNotaFiscal.PlacaDetalhes.Placa02 = DsXml.Tables("reboque").Rows(0)("placa")

                        If DsXml.Tables("reboque").Columns.Contains("UF") AndAlso Not IsDBNull(DsXml.Tables("reboque").Rows(0)("UF")) Then
                            objNotaFiscal.PlacaDetalhes.EstadoPlaca02 = DsXml.Tables("reboque").Rows(0)("UF")
                        Else
                            objNotaFiscal.PlacaDetalhes.EstadoPlaca02 = ""
                        End If

                    End If

                    objNotaFiscal.PlacaDetalhes.Motorista = New [Lib].Negocio.Cliente
                    objNotaFiscal.PlacaDetalhes.Motorista.Codigo = ""
                    objNotaFiscal.PlacaDetalhes.Motorista.CodigoEndereco = 0

                End If

                If Not DsXml.Tables("prod") Is Nothing AndAlso Not DsXml.Tables("prod").Rows(0)("xProd") Is Nothing Then
                    objNotaFiscal.ParteNomeProdNCMXML = Left(DsXml.Tables("prod").Rows(0)("xProd").ToString(), 5)
                End If

                'objNotaFiscal.CarregandoItens = False

                'If objNotaFiscal.Itens.Count > 0 Then

                '    If objNotaFiscal.Itens.Count = 1 Then
                '        For Each item In objNotaFiscal.Itens
                '            If Not objNotaFiscal.Itens(0).Produto Is Nothing AndAlso Not objNotaFiscal.Itens(0).Produto.NCM = DsXml.Tables("prod").Rows(0)("NCM").ToString() Then

                '            End If
                '        Next
                '    End If

                'End If

            Catch ex As Exception
                Throw New Exception(ex.ToString)
            End Try
        End If
    End Sub

    Private Shared Function NormalizarIE(ie As String) As String
        If String.IsNullOrWhiteSpace(ie) Then Return ""

        ' Tratar valores especiais
        Dim tmp = ie.Trim()
        If tmp.Equals("ISENTO", StringComparison.OrdinalIgnoreCase) _
       OrElse tmp.Equals("ISENTA", StringComparison.OrdinalIgnoreCase) Then
            Return ""
        End If

        ' Manter apenas dígitos
        Dim soDigitos As String = New String(tmp.Where(AddressOf Char.IsDigit).ToArray())

        ' Remove zeros à esquerda
        soDigitos = soDigitos.TrimStart("0"c)

        ' Se ficar vazio depois de tirar zeros, volta "0"
        If soDigitos = "" Then soDigitos = "0"

        Return soDigitos
    End Function

    Private Shared Sub ImportarProdutoXML(ByVal objNotaFiscal As NotaFiscal, ByVal DsXml As DataSet, ByVal bUsarProdutoXML As Boolean)

        'Produto
        ' Obter a tabela "det"
        Dim detTable As DataTable = DsXml.Tables("det")

        objNotaFiscal.CarregandoItens = True

        Dim iSequencia As Integer = 1
        If detTable IsNot Nothing Then
            For Each detRow As DataRow In detTable.Rows
                ' Obter a tabela "prod" relacionada
                If detRow.Table.ChildRelations.Count > 0 Then
                    Dim prodRow As DataRow = detRow.GetChildRows(detRow.Table.ChildRelations(0)).First()

                    Dim cProd As String = ""
                    Dim cEAN As String = ""
                    Dim xProd As String = ""
                    Dim NCM As String = ""
                    Dim cBenef As String = ""
                    Dim EXTIPI As String = ""
                    Dim CFOP As String = ""
                    Dim uCom As String = ""
                    Dim qCom As String = ""
                    Dim vUnCom As String = ""
                    Dim vProd As String = ""
                    Dim cEANTrib As String = ""
                    Dim uTrib As String = ""
                    Dim qTrib As String = ""
                    Dim vUnTrib As String = ""
                    Dim indTot As String = ""
                    Dim infAdProd As String = ""
                    Dim xPed As String = ""

                    If prodRow.Table.Columns.Contains("cProd") AndAlso Not IsDBNull(prodRow("cProd")) Then
                        cProd = prodRow("cProd").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("cEAN") AndAlso Not IsDBNull(prodRow("cEAN")) Then
                        cEAN = prodRow("cEAN").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("xProd") AndAlso Not IsDBNull(prodRow("xProd")) Then
                        xProd = prodRow("xProd").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("NCM") AndAlso Not IsDBNull(prodRow("NCM")) Then
                        NCM = prodRow("NCM").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("cBenef") AndAlso Not IsDBNull(prodRow("cBenef")) Then
                        cBenef = prodRow("cBenef").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("EXTIPI") AndAlso Not IsDBNull(prodRow("EXTIPI")) Then
                        EXTIPI = prodRow("EXTIPI").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("CFOP") AndAlso Not IsDBNull(prodRow("CFOP")) Then
                        CFOP = prodRow("CFOP").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("uCom") AndAlso Not IsDBNull(prodRow("uCom")) Then
                        uCom = prodRow("uCom").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("qCom") AndAlso Not IsDBNull(prodRow("qCom")) Then
                        qCom = prodRow("qCom").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("vUnCom") AndAlso Not IsDBNull(prodRow("vUnCom")) Then
                        vUnCom = prodRow("vUnCom").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("vProd") AndAlso Not IsDBNull(prodRow("vProd")) Then
                        vProd = prodRow("vProd").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("cEANTrib") AndAlso Not IsDBNull(prodRow("cEANTrib")) Then
                        cEANTrib = prodRow("cEANTrib").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("uTrib") AndAlso Not IsDBNull(prodRow("uTrib")) Then
                        uTrib = prodRow("uTrib").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("qTrib") AndAlso Not IsDBNull(prodRow("qTrib")) Then
                        qTrib = prodRow("qTrib").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("vUnTrib") AndAlso Not IsDBNull(prodRow("vUnTrib")) Then
                        vUnTrib = prodRow("vUnTrib").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("indTot") AndAlso Not IsDBNull(prodRow("indTot")) Then
                        indTot = prodRow("indTot").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("infAdProd") AndAlso Not IsDBNull(prodRow("infAdProd")) Then
                        infAdProd = prodRow("infAdProd").ToString()
                    End If
                    If prodRow.Table.Columns.Contains("xPed") AndAlso Not IsDBNull(prodRow("xPed")) Then
                        xPed = prodRow("xPed").ToString()
                    End If

                    Dim pFatorConversao = 1

                    Dim objProdutoNota As New NotaFiscalXItem

                    objProdutoNota.NCMProdutoXML = NCM
                    objNotaFiscal.NCMXML = NCM
                    objProdutoNota.ProdutoXML = cProd

                    'Se o cliente for da mesma empresa, podemos usar o cadastro de produto do sistema
                    If bUsarProdutoXML Then
                        objProdutoNota.CodigoProduto = cProd
                    End If

                    objProdutoNota.NomeProdutoXML = Funcoes.EliminarCaracteresEspeciais(xProd)
                    objProdutoNota.DescricaoProdutoXML = Funcoes.EliminarCaracteresEspeciais(xProd)
                    objProdutoNota.InfAdicionalProdutoXML = infAdProd
                    objProdutoNota.UnidadeProdutoXML = uCom

                    objProdutoNota.NotaFiscal = objNotaFiscal
                    objProdutoNota.Sequencia = iSequencia

                    objProdutoNota.PesoBruto = If(prodRow("qCom").ToString.Length > 0, prodRow("qCom").Replace(".", ","), "0").ToString.Replace(".", ",").Replace(",,", ",")
                    objProdutoNota.PesoFiscal = If(prodRow("qCom").ToString.Length > 0, prodRow("qCom").Replace(".", ","), "0").Replace(",,", ",")
                    objProdutoNota.QuantidadeFiscal = If(prodRow("qCom").ToString.Length > 0, prodRow("qCom").Replace(".", ","), "0").Replace(",,", ",")
                    objProdutoNota.PesoLiquido = If(prodRow("qCom").ToString.Length > 0, prodRow("qCom").Replace(".", ","), "0").Replace(",,", ",")
                    objProdutoNota.Unitario = If(prodRow("vUnCom").ToString.Length > 0, prodRow("vUnCom").Replace(".", ","), "0").Replace(",,", ",")
                    objProdutoNota.Volumes = IIf(iSequencia = 1, 1, 0)
                    objProdutoNota.Numeracao = 1  '???
                    objProdutoNota.PesoQuantidade = 0
                    objProdutoNota.ValorLiquido = If(prodRow("vProd").ToString.Length > 0, prodRow("vProd").Replace(".", ","), "0").Replace(",,", ",")
                    objProdutoNota.ValorTotal = If(prodRow("vProd").ToString.Length > 0, prodRow("vProd").Replace(".", ","), "0").Replace(",,", ",")

                    Dim converterTons As String() = {"TO", "TON", "TN", "TNF", "BG", "T"}

                    If converterTons.Contains(uCom.ToString().ToUpper()) Then
                        objProdutoNota.Unitario = objProdutoNota.Unitario / 1000
                        objProdutoNota.PesoBruto = objProdutoNota.PesoBruto * 1000
                        objProdutoNota.PesoFiscal = objProdutoNota.PesoFiscal * 1000
                        objProdutoNota.PesoLiquido = objProdutoNota.PesoLiquido * 1000
                        objProdutoNota.QuantidadeFiscal = objProdutoNota.QuantidadeFiscal * 1000
                    End If

                    If iSequencia = 1 Then

                        If DsXml.Tables.Contains("vol") Then
                            Dim volTable As DataTable = DsXml.Tables("vol")

                            ' Lê os valores das tags dentro de vol
                            Dim qVol As String = 1
                            Dim esp As String = ""
                            Dim nVol As String = ""
                            Dim pesoL As Decimal = 0
                            Dim pesoB As Decimal = 0

                            ' Itera sobre as linhas da tabela vol
                            For Each volRow As DataRow In volTable.Rows

                                If volRow.Table.Columns.Contains("qVol") AndAlso Not IsDBNull(volRow("qVol")) Then
                                    qVol = volRow("qVol").ToString()
                                End If
                                If volRow.Table.Columns.Contains("esp") AndAlso Not IsDBNull(volRow("esp")) Then
                                    esp = volRow("esp").ToString()
                                End If
                                If volRow.Table.Columns.Contains("nVol") AndAlso Not IsDBNull(volRow("nVol")) Then
                                    nVol = volRow("nVol").ToString()
                                End If
                                If volRow.Table.Columns.Contains("pesoL") AndAlso Not IsDBNull(volRow("pesoL")) Then
                                    pesoL = volRow("pesoL").Replace(".", ",").Replace(",,", ",")
                                End If
                                If volRow.Table.Columns.Contains("pesoB") AndAlso Not IsDBNull(volRow("pesoB")) Then
                                    pesoB = volRow("pesoB").Replace(".", ",").Replace(",,", ",")
                                End If

                            Next

                            objProdutoNota.Volumes = qVol
                            objNotaFiscal.Especie = esp
                            objNotaFiscal.PesoBruto = Math.Round(pesoB, 0)
                            objNotaFiscal.PesoLiquido = Math.Round(pesoL, 0)

                        End If

                    End If

                    If prodRow.Table.ChildRelations.Count > 0 Then

                        Dim rastroRows As DataRow() = prodRow.GetChildRows(prodRow.Table.ChildRelations(0))
                        If rastroRows IsNot Nothing AndAlso rastroRows.Length > 0 Then

                            Dim rastroTable As DataTable = prodRow.GetChildRows(prodRow.Table.ChildRelations(0)).CopyToDataTable()
                            If rastroTable IsNot Nothing AndAlso rastroTable.Rows.Count > 0 Then
                                For Each rastroRow As DataRow In rastroTable.Rows

                                    Dim nLote As String = If(rastroRow.Table.Columns.Contains("nLote") AndAlso Not IsDBNull(rastroRow("nLote")), rastroRow("nLote").ToString(), "")
                                    Dim qLote As String = If(rastroRow.Table.Columns.Contains("qLote") AndAlso Not IsDBNull(rastroRow("qLote")), rastroRow("qLote").ToString().Replace(".", ","), "")
                                    Dim dFab As String = If(rastroRow.Table.Columns.Contains("dFab") AndAlso Not IsDBNull(rastroRow("dFab")), rastroRow("dFab").ToString(), "")
                                    Dim dVal As String = If(rastroRow.Table.Columns.Contains("dVal") AndAlso Not IsDBNull(rastroRow("dVal")), rastroRow("dVal").ToString(), "")

                                    If nLote.Length > 0 Then

                                        If objProdutoNota.Lotes.Where(Function(x) x.Lote = nLote).Count = 0 Then

                                            Dim lote As New NotaFiscalXLote(objProdutoNota)
                                            lote.Lote = nLote
                                            lote.Quantidade = IIf(qLote.Length = 0, 0, qLote)
                                            lote.Fabricado = IIf(dFab.Length = 0, "", CDate(dFab))
                                            lote.Validade = IIf(dFab.Length = 0, "", CDate(dVal))
                                            lote.QuantidadeDeConsumo = 0
                                            objProdutoNota.Lotes.Add(lote)

                                        Else

                                            For Each loteAdicionado In objProdutoNota.Lotes.Where(Function(x) x.Lote = nLote)
                                                loteAdicionado.Quantidade += IIf(qLote.Length = 0, 0, qLote)
                                            Next

                                        End If

                                    End If

                                Next

                            End If

                        End If

                    End If

                    objNotaFiscal.Itens.Add(objProdutoNota)

                    iSequencia += 1
                End If
            Next
        End If


    End Sub

    Private Shared Sub ImportarProdutoUnico(objNotaFiscal As NotaFiscal, ByVal DsXml As DataSet, ByVal bUsarProdutoXML As Boolean)

        ImportarProdutoXML(objNotaFiscal, DsXml, bUsarProdutoXML)

        Dim produto As New NotaFiscalXItem(objNotaFiscal)

        For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens

            produto.QuantidadeFiscal += item.QuantidadeFiscal

            If Not item.SubOperacao Is Nothing AndAlso item.SubOperacao.QuantidadeFisico Then
                produto.QuantidadeFisica += item.QuantidadeFiscal
                item.QuantidadeFisica = item.QuantidadeFiscal
            End If

            produto.PesoFiscal += item.PesoFiscal

            produto.PesoBruto += item.PesoBruto
            produto.PesoLiquido += item.PesoLiquido
            produto.ValorLiquido += item.ValorLiquido
            produto.ValorTotal += item.ValorTotal

        Next

        If produto.ValorTotal > 0 And produto.QuantidadeFiscal > 0 Then
            produto.Unitario = produto.ValorTotal / produto.QuantidadeFiscal
        End If

        objNotaFiscal.Itens.Clear()
        objNotaFiscal.Itens.Add(produto)

    End Sub

    Private Shared Sub AgruparNCM(objNotaFiscal As NotaFiscal, ByVal DsXml As DataSet, ByVal bUsarProdutoXML As Boolean)

        ImportarProdutoXML(objNotaFiscal, DsXml, bUsarProdutoXML)

        Dim itensAgrupadoPorNCMProduto = From item In objNotaFiscal.Itens
                                         Group item By item.NCMProdutoXML Into Grupo = Group
                                         Select New With {
                                          .NCMProdutoXML = NCMProdutoXML,
                                           .Itens = Grupo.ToList()
                                      }

        Dim listaDePordutos As New ListNotaFiscalXItem(objNotaFiscal)

        For Each grupo In itensAgrupadoPorNCMProduto

            ' Acessando o valor de NCMProdutoXML
            Dim ncmProduto As String = grupo.NCMProdutoXML

            Dim produto As New NotaFiscalXItem(objNotaFiscal)

            ' Iterando sobre os itens dentro do grupo
            For Each item In grupo.Itens

                ' Aqui você pode acessar as propriedades do item
                produto.QuantidadeFiscal += item.QuantidadeFiscal

                If Not item.SubOperacao Is Nothing AndAlso item.SubOperacao.QuantidadeFisico Then
                    produto.QuantidadeFisica += item.QuantidadeFiscal
                    item.QuantidadeFisica = item.QuantidadeFiscal
                End If

                produto.PesoFiscal += item.PesoFiscal

                produto.PesoBruto += item.PesoBruto
                produto.PesoLiquido += item.PesoLiquido
                produto.ValorLiquido += item.ValorLiquido
                produto.ValorTotal += item.ValorTotal
            Next

            If produto.ValorTotal > 0 And produto.QuantidadeFiscal > 0 Then
                produto.Unitario = produto.ValorTotal / produto.QuantidadeFiscal
            End If

            listaDePordutos.Add(produto)

        Next

        objNotaFiscal.Itens.Clear()
        objNotaFiscal.Itens = listaDePordutos

    End Sub

    Private Shared Sub IncluirDadosNovoClienteEmitente(objNotaFiscal As NotaFiscal, DsXml As DataSet, ClienteTemp As ListCliente, verCliente As String)

        Dim novoCliente As New Cliente

        If verCliente.Length = 14 Then

            novoCliente = New Cliente(verCliente)
            novoCliente.CodigoEndereco = ClienteTemp.Count()
            novoCliente = DadosUnicoNovoCliente(novoCliente, objNotaFiscal)

            If DsXml.Tables("emit").Columns.Contains("IE") Then
                novoCliente.InscricaoEstadual = DsXml.Tables("emit").Rows(0)("IE").ToString()
            Else
                novoCliente.InscricaoEstadual = ""
            End If

            If novoCliente.Nome.Length = 0 Then
                Throw New Exception("A Receita Federal não disponibilizou as informações desse CNPJ.")
                Exit Sub
            End If

        ElseIf verCliente.Length = 11 Then

            novoCliente = New Cliente()

            novoCliente.Codigo = objNotaFiscal.CodigoCliente
            novoCliente.CodigoEndereco = ClienteTemp.Count()
            novoCliente.Nome = DsXml.Tables("emit").Rows(0)("xNome").ToString()
            novoCliente.Fantasia = DsXml.Tables("emit").Rows(0)("xNome").ToString()

            If DsXml.Tables("emit").Columns.Contains("IE") Then
                novoCliente.InscricaoEstadual = DsXml.Tables("emit").Rows(0)("IE").ToString()
            Else
                novoCliente.InscricaoEstadual = ""
            End If

            novoCliente.CEP = DsXml.Tables("enderEmit").Rows(0)("CEP").ToString()
            novoCliente.Endereco = DsXml.Tables("enderEmit").Rows(0)("xLgr").ToString()
            Dim nroObj = DsXml.Tables("enderEmit").Rows(0)("nro")
            novoCliente.Numero = ParseNumeroEndereco(nroObj)

            If DsXml.Tables("enderEmit").Columns.Contains("xCpl") Then
                novoCliente.Complemento = DsXml.Tables("enderEmit").Rows(0)("xCpl").ToString()
            End If

            novoCliente.Bairro = DsXml.Tables("enderEmit").Rows(0)("xBairro").ToString()
            novoCliente.CodigoEstado = DsXml.Tables("enderEmit").Rows(0)("UF").ToString()

            Dim cMunCompleto As String = DsXml.Tables("enderEmit").Rows(0)("cMun").ToString()
            Dim codigoMunicipio As String = cMunCompleto.Substring(2)  ' remove os 2 primeiros dígitos (UF)

            novoCliente.CodigoMunicipio = codigoMunicipio
            novoCliente.Cidade = DsXml.Tables("enderEmit").Rows(0)("xMun").ToString()

            If DsXml.Tables("enderEmit").Columns.Contains("cPais") Then
                novoCliente.CodigoPais = DsXml.Tables("enderEmit").Rows(0)("cPais").ToString()
            Else
                Dim objEstado As New Estado(novoCliente.CodigoEstado)
                If objEstado Is Nothing OrElse String.IsNullOrEmpty(objEstado.Codigo) OrElse objEstado.Regiao = "EX" Then
                    novoCliente.CodigoPais = "5860"
                Else
                    novoCliente.CodigoPais = "1058"
                End If
            End If

            If DsXml.Tables("enderEmit").Columns.Contains("fone") Then
                novoCliente.Telefone = DsXml.Tables("enderEmit").Rows(0)("fone").ToString()
            End If

            If DsXml.Tables("enderEmit").Columns.Contains("email") Then
                novoCliente.Email = DsXml.Tables("emit").Rows(0)("email").ToString()
                novoCliente.EmailNFE = DsXml.Tables("emit").Rows(0)("email").ToString()
            End If

            novoCliente = DadosUnicoNovoCliente(novoCliente, objNotaFiscal)

            ClienteTemp.Add(novoCliente)

        Else

            Throw New Exception("Não foi possível cadastrar o cliente!")
            Exit Sub

        End If

        'Remover mascara CEP
        novoCliente.CEP = New String(novoCliente.CEP.Where(AddressOf Char.IsDigit).ToArray())
        novoCliente.NascimentoConstituicao = Now()
        novoCliente.NaturalidadeEstado = novoCliente.CodigoEstado
        novoCliente.NaturalidadeCidade = novoCliente.Cidade

        If Not novoCliente.Salvar Then
            Throw New Exception("Não foi possível cadastrar o cliente!")
        End If

    End Sub

    Private Shared Sub IncluirDadosNovoClienteDestinatario(objNotaFiscal As NotaFiscal, DsXml As DataSet, ClienteTemp As ListCliente, verCliente As String)

        Dim novoCliente As New Cliente

        If verCliente.Length = 14 Then

            novoCliente = New Cliente(verCliente)
            novoCliente.CodigoEndereco = ClienteTemp.Count()
            novoCliente = DadosUnicoNovoCliente(novoCliente, objNotaFiscal)

            If DsXml.Tables("dest").Columns.Contains("IE") Then
                novoCliente.InscricaoEstadual = DsXml.Tables("dest").Rows(0)("IE").ToString()
            Else
                novoCliente.InscricaoEstadual = ""
            End If

            If novoCliente.Nome.Length = 0 Then
                Throw New Exception("A Receita Federal não disponibilizou as informações desse CNPJ.")
                Exit Sub
            End If

        ElseIf verCliente.Length = 11 Then

            novoCliente = New Cliente()

            novoCliente.Codigo = objNotaFiscal.CodigoCliente
            novoCliente.CodigoEndereco = ClienteTemp.Count()
            novoCliente.Nome = DsXml.Tables("dest").Rows(0)("xNome").ToString()
            novoCliente.Fantasia = DsXml.Tables("dest").Rows(0)("xNome").ToString()

            If DsXml.Tables("dest").Columns.Contains("IE") Then
                novoCliente.InscricaoEstadual = DsXml.Tables("dest").Rows(0)("IE").ToString()
            Else
                novoCliente.InscricaoEstadual = ""
            End If

            novoCliente.CEP = DsXml.Tables("enderDest").Rows(0)("CEP").ToString()
            novoCliente.Endereco = DsXml.Tables("enderDest").Rows(0)("xLgr").ToString()
            Dim nroObj = DsXml.Tables("enderDest").Rows(0)("nro")
            novoCliente.Numero = ParseNumeroEndereco(nroObj)

            If DsXml.Tables("enderDest").Columns.Contains("xCpl") Then
                novoCliente.Complemento = DsXml.Tables("enderDest").Rows(0)("xCpl").ToString()
            End If

            novoCliente.Bairro = DsXml.Tables("enderDest").Rows(0)("xBairro").ToString()
            novoCliente.CodigoEstado = DsXml.Tables("enderDest").Rows(0)("UF").ToString()

            Dim cMunCompleto As String = DsXml.Tables("enderDest").Rows(0)("cMun").ToString()
            Dim codigoMunicipio As String = cMunCompleto.Substring(2)  ' remove os 2 primeiros dígitos (UF)

            novoCliente.CodigoMunicipio = codigoMunicipio
            novoCliente.Cidade = DsXml.Tables("enderDest").Rows(0)("xMun").ToString()

            If DsXml.Tables("enderDest").Columns.Contains("cPais") Then
                novoCliente.CodigoPais = DsXml.Tables("enderDest").Rows(0)("cPais").ToString()
            Else
                Dim objEstado As New Estado(novoCliente.CodigoEstado)
                If objEstado Is Nothing OrElse String.IsNullOrEmpty(objEstado.Codigo) OrElse objEstado.Regiao = "EX" Then
                    novoCliente.CodigoPais = "5860"
                Else
                    novoCliente.CodigoPais = "1058"
                End If
            End If

            If DsXml.Tables("enderDest").Columns.Contains("fone") Then
                novoCliente.Telefone = DsXml.Tables("enderDest").Rows(0)("fone").ToString()
            End If

            If DsXml.Tables("enderDest").Columns.Contains("email") Then
                novoCliente.Email = DsXml.Tables("dest").Rows(0)("email").ToString()
                novoCliente.EmailNFE = DsXml.Tables("dest").Rows(0)("email").ToString()
            End If

            novoCliente = DadosUnicoNovoCliente(novoCliente, objNotaFiscal)

            ClienteTemp.Add(novoCliente)

        Else

            Throw New Exception("Não foi possível cadastrar o cliente!")
            Exit Sub

        End If

        'Remover mascara CEP
        novoCliente.CEP = New String(novoCliente.CEP.Where(AddressOf Char.IsDigit).ToArray())
        novoCliente.NascimentoConstituicao = Now()
        novoCliente.NaturalidadeEstado = novoCliente.CodigoEstado
        novoCliente.NaturalidadeCidade = novoCliente.Cidade

        If Not novoCliente.Salvar Then
            Throw New Exception("Não foi possível cadastrar o cliente!")
        End If

    End Sub

    Private Shared Function ParseNumeroEndereco(nro As Object) As Integer

        ' Converte nulo/DBNull para string vazia
        Dim s As String = If(nro Is Nothing OrElse nro Is DBNull.Value, "", nro.ToString().Trim())

        If s = "" Then Return 0

        Dim up = s.ToUpperInvariant()

        ' Casos típicos de "sem número"
        If up = "SN" OrElse up = "S/N" OrElse up = "SEM NUMERO" OrElse up = "SEM NÚMERO" Then
            Return 0
        End If

        ' Tenta converter direto
        Dim n As Integer
        If Integer.TryParse(s, n) Then
            Return n
        End If

        ' Extrai o primeiro bloco de dígitos (ex.: "123-A" -> 123, "123B" -> 123)
        Dim m = Regex.Match(s, "\d+")
        If m.Success Then
            Return Integer.Parse(m.Value)
        End If

        ' Se não achar nada numérico, considera 0
        Return 0

    End Function

    Private Shared Function DadosUnicoNovoCliente(novoCliente As Cliente, objNotaFiscal As NotaFiscal) As Cliente

        novoCliente.IUD = "I"

        novoCliente.ClienteDesde = String.Format("{0}/{1}/{2}", Now.Day.ToString.PadLeft(2, "0"), Now.Month.ToString.PadLeft(2, "0"), Now.Year)

        novoCliente.CodigoSituacao = 1
        novoCliente.CodigoCategoria = 4

        Dim objTipo As New [Lib].Negocio.ClientexTipo(novoCliente)
        objTipo.IUD = "I"
        objTipo.CodigoTipo = 4
        novoCliente.Tipos.Add(objTipo)

        objTipo = New [Lib].Negocio.ClientexTipo(novoCliente)
        objTipo.IUD = "I"
        objTipo.CodigoTipo = 5
        novoCliente.Tipos.Add(objTipo)

        novoCliente.UsuarioInclusao = objNotaFiscal.UsuarioInclusao
        novoCliente.UsuarioInclusaoData = objNotaFiscal.DataInclusao

        Return novoCliente
    End Function

    'Retorna dados de campos predefinidos do dataset que contém o xml a ser importado
    Public Shared Function CarregarDadosDoCampo(ByRef DsXml As DataSet, ByVal Tabela As String, ByVal IndiceProduto As Integer, ByVal Campo As String, Optional ByVal ValorMonetario As Boolean = False) As String
        Dim Dados As String = String.Empty
        If DsXml.Tables(Tabela) IsNot Nothing Then
            If DsXml.Tables(Tabela).Columns.Contains(Campo) Then
                Try
                    If ValorMonetario Then
                        Dados = CType(DsXml.Tables(Tabela).Rows(IndiceProduto)(Campo), String).Replace(".", ",")
                    Else
                        Dados = DsXml.Tables(Tabela).Rows(IndiceProduto)(Campo)
                    End If
                Catch ex As Exception
                    'Throw New Exception("Este campo não existe no Xml", ex)
                    Dim t = ex.Message
                End Try
            End If
        End If
        Return Dados
    End Function

    Public Shared Function ValidaNFEntradaComXML(ByRef ObjNotafiscal As NotaFiscal, ByRef DsXml As DataSet) As String
        'Xml
        Dim XmlQtd As Decimal = 0
        Dim XmlValorTotalItem As Decimal = 0
        Dim XmlValorDoImposto As Decimal = 0
        Dim XmlUnidade As String
        Dim XmlUnitario As Decimal = 0

        'NF
        Dim NFValorTotalItem As Decimal = 0
        Dim NFQtd As Decimal = 0
        Dim NFValorDoImposto As Decimal = 0
        Dim NFUnidade As String
        Dim NFUnitario As Decimal = 0

        Dim objUnidadeConversao As UnidadeConversao

        For i As Integer = 0 To ObjNotafiscal.Itens.Count - 1
            XmlQtd = CarregarDadosDoCampo(DsXml, "Prod", i, "qCom", True)
            XmlUnidade = CarregarDadosDoCampo(DsXml, "Prod", i, "uCom")
            XmlUnitario = CarregarDadosDoCampo(DsXml, "Prod", i, "vUnCom", True)

            NFQtd = ObjNotafiscal.Itens(i).QuantidadeFiscal
            NFUnidade = ObjNotafiscal.Itens(i).Produto.Unidade
            NFUnitario = ObjNotafiscal.Itens(i).Unitario

            objUnidadeConversao = New UnidadeConversao(XmlUnidade, NFUnidade)

            If objUnidadeConversao IsNot Nothing AndAlso objUnidadeConversao.Fator > 0 Then

                XmlQtd = XmlQtd * objUnidadeConversao.Fator
                If XmlQtd <> NFQtd Then
                    Return "Produto: " & ObjNotafiscal.Itens(i).Produto.Descricao & "Quantidade do item no Xml: " & XmlQtd & " Quantidade do item na NF: " & NFQtd
                End If

                XmlUnitario = XmlUnitario / objUnidadeConversao.Fator
                If XmlUnitario <> NFUnitario Then
                    Return "Produto: " & ObjNotafiscal.Itens(i).Produto.Descricao & "Unitário do item no Xml: " & XmlUnitario & " Unitário do item na NF: " & NFUnitario
                End If
            End If

            XmlValorTotalItem = CarregarDadosDoCampo(DsXml, "Prod", i, "vProd", True)
            NFValorTotalItem = ObjNotafiscal.Itens(i).ValorTotal

            If XmlValorTotalItem <> NFValorTotalItem Then
                Return "Produto: " & ObjNotafiscal.Itens(i).Produto.Descricao & " Valor do item no Xml: " & XmlValorTotalItem & " Valor do item na NF: " & NFValorTotalItem
            End If


            'ICMS
            If DsXml.Tables.Contains("ICMS00") Then
                XmlValorDoImposto = CarregarDadosDoCampo(DsXml, "ICMS00", 0, "vICMS", True)
                NFValorDoImposto = ObjNotafiscal.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "ICMS").Sum(Function(s) s.Valor)
                If XmlValorDoImposto <> NFValorDoImposto Then
                    Return "Produto: " & ObjNotafiscal.Itens(i).Produto.Descricao & " Valor do ICMS no Xml: " & XmlValorDoImposto & " Valor do ICMS na NF: " & NFValorDoImposto
                End If
            End If

            'IPI
            If DsXml.Tables.Contains("IPI") Then
                XmlValorDoImposto = CarregarDadosDoCampo(DsXml, "IPI", 0, "vIPI", True)
                NFValorDoImposto = ObjNotafiscal.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "IPI").Sum(Function(s) s.Valor)
                If XmlValorDoImposto <> NFValorDoImposto Then
                    Return "Produto: " & ObjNotafiscal.Itens(i).Produto.Descricao & " Valor do IPI no Xml: " & XmlValorDoImposto & " Valor do IPI na NF: " & NFValorDoImposto
                End If
            End If

            'PIS
            If DsXml.Tables.Contains("PIS") Then
                XmlValorDoImposto = CarregarDadosDoCampo(DsXml, "PIS", 0, "vPIS", True)
                NFValorDoImposto = ObjNotafiscal.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "PIS").Sum(Function(s) s.Valor)
                If XmlValorDoImposto <> NFValorDoImposto Then
                    Return "Produto: " & ObjNotafiscal.Itens(i).Produto.Descricao & " Valor do PIS no Xml: " & XmlValorDoImposto & " Valor do PIS na NF: " & NFValorDoImposto
                End If
            End If

            'COFINS
            If DsXml.Tables.Contains("COFINS") Then
                XmlValorDoImposto = CarregarDadosDoCampo(DsXml, "COFINS", 0, "vCOFINS", True)
                NFValorDoImposto = ObjNotafiscal.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "COFINS").Sum(Function(s) s.Valor)
                If XmlValorDoImposto <> NFValorDoImposto Then
                    Return "Produto: " & ObjNotafiscal.Itens(i).Produto.Descricao & " Valor do COFINS no Xml: " & XmlValorDoImposto & " Valor do COFINS na NF: " & NFValorDoImposto
                End If
            End If
        Next

        If DsXml.Tables("ICMStot").Columns.Contains("vIPI") Then
            Dim XmlValorTotalDoIPI As Decimal = CarregarDadosDoCampo(DsXml, "ICMStot", 0, "vIPI", True)
            Dim NFValorTotalDoIPI = ObjNotafiscal.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "IPI").Sum(Function(s) s.Valor)

            If XmlValorTotalDoIPI <> NFValorTotalDoIPI Then
                Return "O valor total do IPI NF no XML " & XmlValorTotalDoIPI & " está diferente do valor na NF " & NFValorTotalDoIPI
            End If
        End If

        Dim NFValorTotalNota As Decimal = ObjNotafiscal.TotalNota
        Dim XmlValorTotalNota As Decimal = CarregarDadosDoCampo(DsXml, "ICMStot", 0, "vNF", True)
        If XmlValorTotalNota <> NFValorTotalNota Then
            Return "O valor total da NF no XML " & XmlValorTotalNota & " está diferente do valor da NF " & NFValorTotalNota
        End If

        Return ""
    End Function

#End Region

#Region "CT-e"
    Private Function getTextoCTe(ByVal cte As [Lib].Negocio.NotaFiscal, ByVal HomologAlfasig As Boolean) As String
        Return ""
    End Function
#End Region

#Region "MDF-e"

#End Region

End Class
