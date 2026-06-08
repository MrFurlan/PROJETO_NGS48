Imports System.IO
Imports System.Linq
Imports System.Xml
Imports BoletoNet.Util
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ConhecimentoDeTransporte
    Inherits BasePage

#Region "Variáveis"

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objConhecimento As [Lib].Negocio.NotaFiscal
    Private chaveXMLautomatico As String

    Private Property lstNotaFiscal As List(Of [Lib].Negocio.NotaFiscal)
        Get
            If Session("_lstNotaFiscal" & HID.Value) IsNot Nothing Then
                Return CType(Session("_lstNotaFiscal" & HID.Value), List(Of [Lib].Negocio.NotaFiscal))
            End If
            Return Nothing
        End Get
        Set(ByVal value As List(Of [Lib].Negocio.NotaFiscal))
            Session("_lstNotaFiscal" & HID.Value) = value
        End Set
    End Property

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ConhecimentoDeTransporte", "ACESSAR") Then

                CargaEmpresas()
                CargaSituacao()
                WebHelpers.BindDropDownWithEnum(ddlTipoConhecimento, GetType(eTipoDeDocumentoFrete), True, "-- [SELECIONE] --")

                If Not Session("chaveXMLautomacao" & HID.Value) Is Nothing Then
                    chaveXMLautomatico = Session("chaveXMLautomacao" & HID.Value).ToString()
                    Session.Remove("chaveXMLautomacao" & HID.Value)
                End If

                LimparCampos()

                txtObservacao.Height = Unit.Pixel(105)
                tcPrincipal.ActiveTab = tabCTRC

                carregarNotaXMLautomatico(sender, e)

            Else

                MsgBox(Me.Page, "Usuário sem permissão para acessar a página!", "~/Expedicao.aspx")
                Exit Sub

            End If
        End If
    End Sub

    Protected Sub lnkConsultarNota_Click(sender As Object, e As EventArgs) Handles lnkConsultarNota.Click
        If Funcoes.VerificaPermissao("ConhecimentoDeTransporte", "LEITURA") Then
            If DdlEmpresa.SelectedIndex > 0 Then
                ConsultarNotas()
            Else
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
        End If
    End Sub

    Protected Sub lnkConsultarSefaz_Click(sender As Object, e As EventArgs) Handles lnkConsultarSefaz.Click

        If Funcoes.VerificaPermissao("ConhecimentoDeTransporte", "GRAVAR") Then

            RecuperarSessaoConhecimento()

            Dim Sqls As New ArrayList
            Dim Sql As String = String.Empty
            Dim msgNFE As String = String.Empty

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

                If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "100" OrElse strCodigo = "101" OrElse strCodigo = "110" OrElse strCodigo = "124" OrElse strCodigo = "302") _
                    And strChave.Length = 44 Then
                    'And strRecibo.Length = 15 _

                    Dim strOperacao = String.Empty
                    Dim strSituacao = String.Empty
                    Sql = String.Empty
                    Sqls.Clear()

                    If strCodigo = "101" Then
                        strOperacao = "Incluir"
                        If strCodigo = "101" Then
                            strSituacao = CInt(eSituacao.Cancelado)
                        Else
                            strSituacao = CInt(eSituacao.Normal)
                        End If

                        Sql = "UPDATE NotasFiscais " & vbCrLf &
                              "   SET Situacao = " & strSituacao & " " & vbCrLf &
                              " WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "' AND EndEmpresa_Id = '" & objConhecimento.EnderecoEmpresa & "' AND " & vbCrLf &
                              "       Cliente_Id = '" & objConhecimento.CodigoCliente & "' AND EndCliente_Id = '" & objConhecimento.EnderecoCliente & "' AND " & vbCrLf &
                              "       EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND Serie_Id = '" & objConhecimento.Serie & "' AND " & vbCrLf &
                              "       Nota_Id = '" & objConhecimento.Codigo & "'; " & vbCrLf
                        Sqls.Add(Sql)

                    End If

                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-consultacte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)

                    If strSituacao = "1" Then
                        ContabilizarNota(Sqls, True)
                    Else

                        Sql = "UPDATE DocumentoXML set Situacao = 2" & vbCrLf &
                              " WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "'" & vbCrLf &
                              "   AND Chave_id   = '" & objConhecimento.ChaveNFE & "'"
                        Sqls.Add(Sql)

                        Sql = "Delete NotasXNotas " & vbCrLf &
                              " WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "' AND EndEmpresa_Id = '" & objConhecimento.EnderecoEmpresa & "' AND " & vbCrLf &
                              "       Cliente_Id = '" & objConhecimento.CodigoCliente & "' AND EndCliente_Id = '" & objConhecimento.EnderecoCliente & "' AND " & vbCrLf &
                              "       EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND Serie_Id = '" & objConhecimento.Serie & "' AND " & vbCrLf &
                              "       Nota_Id = '" & objConhecimento.Codigo & "'; " & vbCrLf
                        Sqls.Add(Sql)


                        Dim lst As New [Lib].Negocio.ListFaturasDeFretesXItens(objConhecimento)
                        For Each item As [Lib].Negocio.FaturaDeFreteXItens In lst
                            item.IUD = "D"
                            item.SalvarSql(Sqls)
                        Next

                        Dim faturas As New [Lib].Negocio.ListFaturaDeFrete(objConhecimento, lst)
                        For Each item As [Lib].Negocio.FaturaDeFrete In faturas
                            item.IUD = "D"
                            item.SalvarSql(Sqls)
                        Next


                        Dim titulos As New [Lib].Negocio.ListTitulo
                        For Each item As [Lib].Negocio.FaturaDeFrete In faturas
                            For Each titfat In item.ListTituloFatura
                                titulos.Add(New [Lib].Negocio.Titulo(titfat.CodigoTitulo, titfat.Titulo.ReceberPagar))
                            Next
                        Next

                        If Not ValidarExcluir(titulos) Then
                            MsgBox(Me.Page, "Cancelamento homologado com Titulo Relacionado. Entre em contato com o Suporte.")
                            Exit Sub
                        End If

                        For Each item As [Lib].Negocio.Titulo In titulos
                            item.IUD = "C"
                            item.ObservacoesControleInterno = "CONHECIMENTO DE TRANSPORTE CANCELADO"
                            item.SalvarSql(Sqls)
                        Next

                        Dim razao As New [Lib].Negocio.Razao(objConhecimento)
                        razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                        Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                        Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                        razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)

                    End If

                    If Not Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Exit Sub
                    ElseIf Not String.IsNullOrEmpty(msgNFE) Then
                        MsgBox(Me.Page, msgNFE)
                        Exit Sub
                    End If

                End If

                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                Session("NfConsultaCTRC" & HID.Value) = objConhecimento
                CarregarNotaFiscal()

            Else
                MsgBox(Me.Page, "Sefaz não retornou nenhuma resposta, consulte novamente.")
            End If

        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
        End If

    End Sub

    Private Function getTextoConsultar(ByVal CTe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVECTE=" & CTe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NRECIBO =" & CTe.ReciboNota & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Protected Sub imgExcluirNF_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)

        RecuperarSessaoConhecimento()

        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)
        objConhecimento.NotasTrocaOrigem(row.RowIndex).IUD = "D"
        grdNotasFretes.DataSource = objConhecimento.NotasTrocaOrigem.Where(Function(s) s.IUD <> "D")
        grdNotasFretes.DataBind()
        objConhecimento.Itens.ForEach(Function(nxi)
                                          nxi.QuantidadeFiscal = objConhecimento.NotasTrocaOrigem.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFiscal)
                                          nxi.QuantidadeFisica = objConhecimento.NotasTrocaOrigem.Where(Function(s) s.IUD <> "D").SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFisica)
                                          Return True
                                      End Function)
        SalvarSessaoConhecimento()

    End Sub

    Protected Sub imgDelNFReferencial_Click(sender As Object, e As ImageClickEventArgs)

        Try

            Dim cancRe As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(cancRe.NamingContainer, GridViewRow)

            Dim lblNota As Label = row.FindControl("lblNota")

            RecuperarSessaoConhecimento()

            If objConhecimento.IUD = "I" Then
                objConhecimento.NotasReferenciais.RemoveAt(row.RowIndex)
            Else
                objConhecimento.NotasReferenciais.Where(Function(s) s.IUD <> "D" And s.Nota_Id = lblNota.Text).Single.IUD = "D"

                For i As Integer = 0 To objConhecimento.NotasReferenciais.Count - 1
                    If String.IsNullOrWhiteSpace(objConhecimento.NotasReferenciais(i).IUD) OrElse objConhecimento.NotasReferenciais(i).IUD <> "D" Then
                        objConhecimento.NotasReferenciais(i).IUD = "X"
                    End If
                Next
            End If
            If objConhecimento.NotasReferenciais IsNot Nothing Then
                If objConhecimento.NotasReferenciais.Where(Function(s) s.IUD <> "D").Count = 0 Then
                    txtPedidoNFReferencial.ReadOnly = False
                    txtPedidoNFReferencial.Text = String.Empty
                    btnPedidoNFReferencial.Enabled = True
                End If
            End If

            SalvarSessaoConhecimento()

            grdNotasReferenciais.DataSource = objConhecimento.NotasReferenciais.Where(Function(s) s.IUD <> "D")
            grdNotasReferenciais.DataBind()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Protected Sub lnkSelecionarNota_Click(sender As Object, e As EventArgs) Handles lnkSelecionarNota.Click
        If lstNotaFiscal IsNot Nothing Then
            lstNotaFiscal.Clear()
        End If
        Selecionar()
    End Sub

    Protected Sub lnkRelatorioNota_Click(sender As Object, e As EventArgs) Handles lnkRelatorioNota.Click

        If ValidarCampos() Then
            Dim parameters As New Dictionary(Of String, Object)
            parameters.Add("ParametrosConsulta", getParametros())
            parameters.Add("empresa", DdlEmpresa.SelectedValue.Split("-")(0))
            parameters.Add("enderecoEmpresa", DdlEmpresa.SelectedValue.Split("-")(1))
            parameters.Add("es", IIf(rdoSaida.Checked, "S", "E"))
            parameters.Add("ciffob", IIf(rdoSaida.Checked, "CIF", "FOB"))
            parameters.Add("codcliente", txtCodigoCliente.Value)
            parameters.Add("numconhecimento", txtNumConhecimento.Text)
            parameters.Add("data1", txtPeriodoInicial.Text)
            parameters.Add("data2", txtPeriodoFinal.Text)
            parameters.Add("viaDeTransporte", IIf(rdoRodoviario.Checked, "R", "F"))
            Session("MyParameters" & HID.Value) = parameters
            ucConsultaRelatorio.pageLoad()
            Popup.ConsultaRelatorio(Me.Page, "objFerroviario" & HID.Value)
        End If
    End Sub

    Protected Sub lnkLimparNota_Click(sender As Object, e As EventArgs) Handles lnkLimparNota.Click
        LimparCamposDaConsulta()
    End Sub

    Public Sub MetodoCriarFaturaDeFrete()

        Dim sql As String = "   SELECT DISTINCT NF.Movimento,
		                                   NF.Empresa_id,
		                                   NF.EndEmpresa_id,
		                                   NF.Cliente_Id AS Cliente_Id,
		                                   NF.EndCliente_Id AS EndCliente_Id,
		                                   NF.EntradaSaida_Id AS EntradaSaida_Id,
		                                   NF.Serie_Id AS Serie_Id,
		                                   NF.Nota_Id AS Nota_Id,
		                                   NF.Operacao,
		                                   NF.SubOperacao,
		                                   NFxI.QuantidadeFiscal AS Quantidade,
		                                   NFxI.Unitario,
		                                   NFxI.Valor,
		                                   FFXT.Titulo_Id
                                FROM NotasFiscais NF
                                INNER JOIN NotasFiscaisxItens NFxI
	                                ON NF.Empresa_Id			= NFxI.Empresa_Id
	                                AND NF.EndEmpresa_Id		= NFxI.EndEmpresa_Id
	                                AND NF.Cliente_Id			= NFxI.Cliente_Id
	                                AND NF.EndCliente_Id		= NFxI.EndCliente_Id
	                                AND NF.EntradaSaida_Id		= NFxI.EntradaSaida_Id
	                                AND NF.Serie_Id				= NFxI.Serie_Id
	                                AND NF.Nota_Id				= NFxI.Nota_Id
                                LEFT JOIN FaturasDeFretesXItens FFXI
	                                ON NFxI.Empresa_Id			= FFXI.Empresa_Id
	                                AND NFxI.EndEmpresa_Id		= FFXI.EndEmpresa_Id
	                                AND NFxI.Cliente_Id			= FFXI.Cliente_Id
	                                AND NFxI.EndCliente_Id		= FFXI.EndCliente_Id
	                                AND NFxI.EntradaSaida_Id	= FFXI.EntradaSaida_Id
	                                AND NFxI.Serie_Id			= FFXI.Serie_Id
	                                AND NFxI.Nota_Id			= FFXI.Nota_Id
                                LEFT JOIN FaturaDeFreteXTitulo FFXT
	                                ON FFXI.Empresa_Id			= FFXT.Empresa_Id
	                                AND FFXI.EndEmpresa_Id		= FFXT.EndEmpresa_Id
	                                AND FFXI.Conveniado_Id      = FFXT.Conveniado_Id
	                                AND FFXI.EndConveniado_Id   = FFXT.EndConveniado_Id
	                                AND FFXI.Fatura_Id			= FFXT.Fatura_Id
                                WHERE NF.Situacao				IN (1) 
	                                AND NF.TipoDeDocumento		IN (57)
	                                --AND NF.Nota_Id				= 2825
	                                AND FFXT.Titulo_Id IS NULL
                                ORDER BY NF.Movimento desc, NF.Nota_Id, NF.Serie_Id, NF.EntradaSaida_Id; "


        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "FaturaDeFrete")

        For Each row As DataRow In ds.Tables(0).Rows

            Dim obj As New [Lib].Negocio.NotaFiscal()
            obj.CodigoEmpresa = row("Empresa_Id")
            obj.EnderecoEmpresa = row("EndEmpresa_Id")
            obj.CodigoCliente = row("Cliente_Id")
            obj.EnderecoCliente = row("EndCliente_Id")
            obj.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", 0, 1)
            obj.Serie = row("Serie_Id")
            obj.Codigo = row("Nota_Id")

            objConhecimento = New [Lib].Negocio.NotaFiscal(obj)
            objConhecimento.IUD = "U"
            SalvarSessaoConhecimento()

            Dim Sqls As New ArrayList

            Dim sop As New [Lib].Negocio.SubOperacao(objConhecimento.CodigoOperacao, objConhecimento.CodigoSubOperacao)
            If sop IsNot Nothing AndAlso sop.Financeiro Then
                SalvarFaturaDeFrete(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
            End If

        Next

    End Sub

    Public Sub MetodoExcluirFaturaDeFrete()

        Dim sql As String = "   SELECT DISTINCT NF.Movimento,
		                                   NF.Empresa_id,
		                                   NF.EndEmpresa_id,
		                                   NF.Cliente_Id AS Cliente_Id,
		                                   NF.EndCliente_Id AS EndCliente_Id,
		                                   NF.EntradaSaida_Id AS EntradaSaida_Id,
		                                   NF.Serie_Id AS Serie_Id,
		                                   NF.Nota_Id AS Nota_Id,
		                                   NF.Operacao,
		                                   NF.SubOperacao,
		                                   NFxI.QuantidadeFiscal AS Quantidade,
		                                   NFxI.Unitario,
		                                   NFxI.Valor,
		                                   FFXT.Titulo_Id
                                FROM NotasFiscais NF
                                INNER JOIN NotasFiscaisxItens NFxI
	                                ON NF.Empresa_Id			= NFxI.Empresa_Id
	                                AND NF.EndEmpresa_Id		= NFxI.EndEmpresa_Id
	                                AND NF.Cliente_Id			= NFxI.Cliente_Id
	                                AND NF.EndCliente_Id		= NFxI.EndCliente_Id
	                                AND NF.EntradaSaida_Id		= NFxI.EntradaSaida_Id
	                                AND NF.Serie_Id				= NFxI.Serie_Id
	                                AND NF.Nota_Id				= NFxI.Nota_Id
                                LEFT JOIN FaturasDeFretesXItens FFXI
	                                ON NFxI.Empresa_Id			= FFXI.Empresa_Id
	                                AND NFxI.EndEmpresa_Id		= FFXI.EndEmpresa_Id
	                                AND NFxI.Cliente_Id			= FFXI.Cliente_Id
	                                AND NFxI.EndCliente_Id		= FFXI.EndCliente_Id
	                                AND NFxI.EntradaSaida_Id	= FFXI.EntradaSaida_Id
	                                AND NFxI.Serie_Id			= FFXI.Serie_Id
	                                AND NFxI.Nota_Id			= FFXI.Nota_Id
                                LEFT JOIN FaturaDeFreteXTitulo FFXT
	                                ON FFXI.Empresa_Id			= FFXT.Empresa_Id
	                                AND FFXI.EndEmpresa_Id		= FFXT.EndEmpresa_Id
	                                AND FFXI.Conveniado_Id      = FFXT.Conveniado_Id
	                                AND FFXI.EndConveniado_Id   = FFXT.EndConveniado_Id
	                                AND FFXI.Fatura_Id			= FFXT.Fatura_Id
                                WHERE NF.Situacao				IN (1) 
	                                AND NF.TipoDeDocumento		IN (57)
	                                --AND FFXT.Fatura_Id          = -1
                                    --AND NOT FFXT.Titulo_Id IS NULL
                                    AND NF.Situacao = 2
                                ORDER BY NF.Movimento desc, NF.Nota_Id, NF.Serie_Id, NF.EntradaSaida_Id; "


        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "FaturaDeFrete")

        For Each row As DataRow In ds.Tables(0).Rows

            Dim obj As New [Lib].Negocio.NotaFiscal()
            obj.CodigoEmpresa = row("Empresa_Id")
            obj.EnderecoEmpresa = row("EndEmpresa_Id")
            obj.CodigoCliente = row("Cliente_Id")
            obj.EnderecoCliente = row("EndCliente_Id")
            obj.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", 0, 1)
            obj.Serie = row("Serie_Id")
            obj.Codigo = row("Nota_Id")

            objConhecimento = New [Lib].Negocio.NotaFiscal(obj)
            objConhecimento.IUD = "U"
            SalvarSessaoConhecimento()

            Dim lst As New [Lib].Negocio.ListFaturasDeFretesXItens(objConhecimento)
            Dim faturas As New [Lib].Negocio.ListFaturaDeFrete(objConhecimento, lst)
            Dim Sqls As New ArrayList

            For Each itens As FaturaDeFreteXItens In lst

                itens.IUD = "D"
                itens.SalvarSql(Sqls)

            Next

            For Each fatura As FaturaDeFrete In faturas

                fatura.IUD = "D"
                fatura.SalvarSql(Sqls)

            Next

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If

        Next

    End Sub

    Protected Sub lnkAjudaNota_Click(sender As Object, e As EventArgs) Handles lnkAjudaNota.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConhecimentoDeTransporte")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkVerificarChaveNFE_Click(sender As Object, e As EventArgs) Handles lnkVerificarChaveNFE.Click

        Dim ChaveNfe As String = txtChaveNFe.Text.Replace(".", "")

        If String.IsNullOrWhiteSpace(ChaveNfe) Then
            MsgBox(Me.Page, "Chave da Nota Fiscal não foi informada.")
        ElseIf ChaveNfe.Trim.Length <> 44 Then
            txtChaveNFe.Text = ""
            MsgBox(Me.Page, "Chave da nota eletrônica deve ter 44 dígitos.")
        Else

            RecuperarSessaoConhecimento()

            objConhecimento.ChaveNFE = txtChaveNFe.Text.Replace(".", "")

            objConhecimento.Codigo = CInt(Mid(objConhecimento.ChaveNFE, 26, 9))
            txtNumero.Text = CInt(Mid(objConhecimento.ChaveNFE, 26, 9))

            objConhecimento.Serie = CInt(Mid(objConhecimento.ChaveNFE, 23, 3))
            txtSerie.Text = objConhecimento.Serie

            Dim ModeloNFe As String = Mid(objConhecimento.ChaveNFE, 21, 2)
            Dim msgResultado As String = String.Empty

            If Not ModeloNFe.Trim.Equals("57") Then
                MsgBox(Me.Page, "o XML informado não é um CTE.")
                Exit Sub
            End If

            'If ddlTipoConhecimento.SelectedValue = 0 Then
            '    MsgBox(Me.Page, "Selecione o tipo documento de conhecimento.")
            '    tcPrincipal.ActiveTab = tabNotaFiscal
            '    Exit Sub
            'End If

            'Realiza o manifesto da NFe
            If ucFile.Parent.Visible AndAlso (ModeloNFe.Equals("55") Or ModeloNFe.Equals("57")) Then '(Modelo: 55 NFe, 57 CTe)

                Dim tpExt As String = String.Empty

                If ModeloNFe.Equals("55") Then tpExt = "-nfe.xml"
                If ModeloNFe.Equals("57") Then tpExt = "-CTe.xml"

                'Download do Arquivo.
                Dim bytes As Byte() = New FilesManager().getFileXml(String.Format("{0}{1}", objConhecimento.ChaveNFE, tpExt))
                If bytes Is Nothing Then
                    If ModeloNFe.Equals("55") Then
                        MsgBox(Me.Page, "XML da NFe não foi encontrado, favor inserir manualmente.")
                    Else
                        MsgBox(Me.Page, "XML do CTe não foi encontrado, favor inserir manualmente.")
                    End If

                    Exit Sub
                Else
                    Dim caminhoArquivoFile As String = Server.MapPath(String.Format("~/Files/{0}{1}", objConhecimento.ChaveNFE, tpExt))
                    If Not File.Exists(caminhoArquivoFile) Then
                        System.IO.File.WriteAllBytes(caminhoArquivoFile, bytes)
                    End If
                End If

                If bytes IsNot Nothing Then
                    Try
                        If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.ChaveNFE) Then
                            Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}{1}", objConhecimento.ChaveNFE, tpExt))
                            If System.IO.File.Exists(caminhoArquivo) Then
                                System.IO.File.Delete(caminhoArquivo)
                            End If
                            System.IO.File.WriteAllBytes(caminhoArquivo, bytes)

                            Dim arquivo As String = objConhecimento.ChaveNFE & tpExt

                            If File.Exists(caminhoArquivo) Then
                                Dim sourceDir As String = Server.MapPath("~/Files/")
                                Dim backupDir As String = "C:/Alfasig/LeituraNFe/-download/"

                                If Not File.Exists(backupDir & arquivo) Then
                                    File.Copy(Path.Combine(sourceDir, arquivo), Path.Combine(backupDir, arquivo))
                                End If
                            End If
                        End If

                        Dim temArquivo As Boolean = False

                        If objConhecimento.Arquivos IsNot Nothing AndAlso objConhecimento.Arquivos.Count > 0 Then
                            For Each arq In objConhecimento.Arquivos
                                If arq.Descricao = String.Format("{0}{1}", objConhecimento.ChaveNFE, tpExt) Then
                                    temArquivo = True
                                End If
                            Next
                        End If

                        If Not temArquivo Then
                            objConhecimento.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                             .IUD = "I",
                             .Codigo = String.Empty,
                             .Descricao = String.Format("{0}{1}", objConhecimento.ChaveNFE, tpExt),
                             .Arquivo = bytes})
                        End If

                        ucFile.Bind(objConhecimento.Arquivos)

                        PreencherNFeXML(objConhecimento, String.Format("{0}{1}", objConhecimento.ChaveNFE, tpExt), False)
                        '****************************** (FIM) colocar em produção quando estiver configurado o BD de Arquivo ******************************

                        If msgResultado.Length > 0 Then
                            lnkNovo.Parent.Visible = False
                            txtChaveNFe.Enabled = True
                            MsgBox(Me.Page, msgResultado, eTitulo.Info, False)
                        Else
                            lnkNovo.Parent.Visible = True
                            txtChaveNFe.Enabled = False
                        End If

                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message.ToString())
                        If ex.Message = "Nota de origem não encontrada!" Then
                            ucNFOrigem.Limpar()
                            Popup.ConsultaNotaOrigem(Me.Page, "objOrigemNFG" & HID.Value)
                        End If
                        Exit Sub
                    End Try
                Else
                    MsgBox(Me.Page, "XML não foi encontrado, favor inserir o manualmente.")
                    Exit Sub
                End If

                If objConhecimento.Empresa.Empresa.NotaFiscalEletronica AndAlso ModeloNFe.Equals("55") AndAlso Not objConhecimento.TemConfirmacaoDaOperacao Then

                    Dim verStatusNFe As String = DocumentoEletronico.ConsultaNFEFornecedor(objConhecimento)
                    Dim statusNFE As String() = verStatusNFe.Split(";")

                    If statusNFE(0) = "100" Then
                        If Not DocumentoEletronico.ManifestoNFe(objConhecimento, eTipoManifesto.ConfirmacaoDaOperacao, msgResultado) Then
                            MsgBox(Me.Page, msgResultado)
                            Exit Sub
                        End If
                    ElseIf statusNFE(0) = "101" Then
                        MsgBox(Me.Page, "Nota Fiscal Cancelada pelo Fornecedor não pode ser utilizada.")
                        Exit Sub
                    Else
                        MsgBox(Me.Page, statusNFE(1))
                        Exit Sub
                    End If
                End If

                SalvarSessaoConhecimento()
                GerarCTRCAutomatico()
                'Chamamos o encargo
                getEncargos()

            End If

        End If
    End Sub

    Private Sub carregarNotaXMLautomatico(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not chaveXMLautomatico Is Nothing Then
            txtChaveNFe.Text = chaveXMLautomatico
            tcPrincipal.ActiveTab = tabCTRC
            lnkVerificarChaveNFE_Click(sender, e)
        End If
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        ucConsultaEmpresas.Limpar()
        Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCTRC" & HID.Value)
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Session("ssCampo") = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteCTRC" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnPedido_Click(sender As Object, e As EventArgs) Handles btnPedido.Click
        Try
            If DdlEmpresa.SelectedIndex = 0 Then
                MsgBox(Me.Page, "Informe uma empresa.")
                DdlEmpresa.Focus()
            ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "Informe um cliente.")
                btnCliente.Focus()
            Else
                HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
                Dim strEmpresa() As String = DdlEmpresa.SelectedValue.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")

                Dim parameters As New Dictionary(Of String, Object)
                parameters("empresa") = strEmpresa(0)
                parameters("enderecoEmpresa") = strEmpresa(1)
                parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
                parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
                parameters("situacao") = eSituacao.Normal
                Popup.ConsultaDePedidos(Me.Page, "objPedidoRPA" & HID.Value, "txtNome")
                ucConsultaPedidos.BindGridView(parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtUnitario_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUnitario.TextChanged
        getEncargos()
    End Sub

    Protected Sub ddlTipoConhecimento_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTipoConhecimento.SelectedIndexChanged
        BindGridView()
    End Sub

    Protected Sub btnCalcular_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnCalcular.Click
        Calcular()
    End Sub

    Protected Sub rdoPamcard_CheckedChanged(sender As Object, e As EventArgs) Handles rdoPamcard.CheckedChanged
        txtCartao.Text = String.Empty
        txtCartao.Enabled = rdoPamcard.Checked
        btnFavorecidoCartao.Enabled = False
    End Sub

    Protected Sub rdoBanco_CheckedChanged(sender As Object, e As EventArgs) Handles rdoBanco.CheckedChanged
        txtCartao.Text = String.Empty
        txtCartao.Enabled = rdoPamcard.Checked
        btnFavorecidoCartao.Enabled = False
    End Sub

    Protected Sub btnFavorecidoCartao_Click(sender As Object, e As EventArgs) Handles btnFavorecidoCartao.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objFavorecidoCTRC" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnTransportador_Click(sender As Object, e As EventArgs) Handles btnTransportador.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objTransportadorCTRC" & HID.Value, "txtNome")
    End Sub

    Protected Sub txtCartao_TextChanged(sender As Object, e As EventArgs) Handles txtCartao.TextChanged
        RecuperarSessaoConhecimento()
        Dim ws As New [Lib].Negocio.Pamcard()
        Dim cartao As String = txtCartao.Text.OnlyNumbers().Trim()
        If String.IsNullOrWhiteSpace(cartao) OrElse Not cartao.Length = 16 Then
            MsgBox(Me.Page, "Informe um número válido de cartão!")
            'lnkNovo.Parent.Visible = False
            Exit Sub
        Else
            Dim ds As DataSet = ws.ConsultarCartao("04854422000185", cartao)
            If ds Is Nothing OrElse ds.Tables Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Erro ao consultar o cartão, entre em contato com o suporte!")
                'lnkNovo.Parent.Visible = False
                Exit Sub
            Else
                If ds.Tables(0).Rows(0).Item("erro") <> 0 AndAlso ds.Tables(0).Rows(0).Item("erro") <> 4 Then
                    MsgBox(Me.Page, "Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                    'lnkNovo.Parent.Visible = False
                    Exit Sub
                End If
                objConhecimento.CartaoNovo = ds.Tables(0).Rows(0).Item("erro") = 4
                btnFavorecidoCartao.Enabled = objConhecimento.CartaoNovo
                If Not String.IsNullOrWhiteSpace(ds.Tables(0).Rows(0).Item("numero")) Then
                    Dim objCliente As New [Lib].Negocio.Cliente(ds.Tables(0).Rows(0).Item("numero"), 0)
                    If objCliente IsNot Nothing Then
                        txtCodigoFavorecidoCartao.Value = objCliente.Codigo & "-" & objCliente.CodigoEndereco
                        txtFavorecidoCartao.Text = objCliente.Nome
                        txtEnderecoFavorecidoCartao.Text = objCliente.CodigoEndereco
                        txtCidadeFavorecidoCartao.Text = objCliente.Cidade
                        txtEstadoFavorecidoCartao.Text = objCliente.CodigoEstado
                        txtCPFFavorecidoCartao.Text = objCliente.Codigo
                        btnFavorecidoCartao.Enabled = False
                    End If
                End If
                objConhecimento.CartaoPgtoFrete = cartao
                SalvarSessaoConhecimento()
            End If
        End If
    End Sub

    Protected Sub btnInutilizar_Click(sender As Object, e As EventArgs) Handles btnInutilizar.Click
        RecuperarSessaoConhecimento()
        ucInutilizacao.CarregarInutilizacao(objConhecimento)
        Popup.ConsultaDeInutilizacao(Me.Page, "objInutilizacaoCTe" & HID.Value)
    End Sub

    Protected Sub btnCadastro_Click(sender As Object, e As EventArgs) Handles btnCadastro.Click
        Dim txtCpf As TextBox = CType(ucConsultaCadastro.FindControlRecursive("txtCpf"), TextBox)
        ucConsultaCadastro.Limpar()
        Popup.ConsultaCadastro(Me.Page, "objCadastro" & HID.Value, txtCpf.ClientID, True, 100)
    End Sub

    Protected Sub btnRelatorio_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            RecuperarSessaoConhecimento()
            ImprimeContrato(objConhecimento)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnRecontabilizar_Click(sender As Object, e As EventArgs) Handles btnRecontabilizar.Click
        Recontabilizar()
    End Sub

    Protected Sub lnkEmitir_Click(sender As Object, e As EventArgs) Handles lnkEmitir.Click
        RecuperarSessaoConhecimento()

        ucEmitirCTe.Limpar()

        Popup.ConsultaEmitirCTe(Me, "objEmitirCTe" & HID.Value)
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        RecuperarSessaoConhecimento()
        If objConhecimento.NossaEmissao Then
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                Salvar()
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
            End If
        Else
            Salvar()
        End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ConhecimentoDeTransporte", "EXCLUIR") Then
                RecuperarSessaoConhecimento()

                Dim temBaixa As Boolean = False

                If grdVencimentos.Rows.Count > 0 Then
                    For n As Integer = 0 To grdVencimentos.Rows.Count - 1

                        Dim t As String = CType(grdVencimentos.Rows(n).FindControl("hpTitulo"), HyperLink).Text
                        Dim t22 As String = CType(grdVencimentos.Rows(n).FindControl("lblReceberPagar"), Label).Text

                        Dim titulo As [Lib].Negocio.Titulo = New [Lib].Negocio.Titulo(CType(grdVencimentos.Rows(n).FindControl("hpTitulo"), HyperLink).Text, CType(grdVencimentos.Rows(n).FindControl("lblReceberPagar"), Label).Text)

                        If titulo.Codigo > 0 AndAlso titulo.CodigoProvisao = 1 Then
                            temBaixa = True
                        End If
                    Next
                End If

                If temBaixa Then
                    MsgBox(Me.Page, "Conhecimento com Financeiro Baixado não pode ser excluído.", eTitulo.Info)
                    Exit Sub
                End If

                If objConhecimento.NossaEmissao Then
                    Dim fm As New FilesManager()
                    If fm.IsConnect() Then
                        Excluir()
                    Else
                        MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                    End If
                Else
                    Excluir()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If Funcoes.VerificaPermissao("ConhecimentoDeTransporte", "LEITURA") Then
            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            If DdlEmpresa.SelectedIndex > 0 Then
                Dim strEmpresa() As String = DdlEmpresa.SelectedValue.ToString.Split("-")
                objNotaFiscal.CodigoEmpresa = (strEmpresa(0))
            Else
                objNotaFiscal.CodigoEmpresa = Session("ssEmpresa")
            End If
            objNotaFiscal.DataNota = txtDataDeEmissao.Text.Trim()
            objNotaFiscal.Movimento = txtDataDeSaida.Text.Trim()
            objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ManifestoEletronico)
            objNotaFiscal.CodigoSituacao = CInt(eSituacao.Normal)
            If Not String.IsNullOrWhiteSpace(txtSerie.Text) Then objNotaFiscal.Serie = txtSerie.Text.Trim
            If Not String.IsNullOrWhiteSpace(txtNumero.Text) Then objNotaFiscal.Codigo = Convert.ToInt32(txtNumero.Text.Trim)
            Session("objNotaFiscal" & HID.Value) = objNotaFiscal
            Session("ssCampo" & HID.Value) = "ConhecimentoDeTransporte"
            Popup.ConsultaDeCTeXNotas(Me.Page, "NfConsultaCTRC" & HID.Value)
            objNotaFiscal.CarregandoNota = True
            ucConsultaCTeXNotas.BindGridView()
            objNotaFiscal.CarregandoNota = False
        Else
            MsgBox(Me.Page, "Usuário sem permissão para consultar registro!")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        LimparCampos()
    End Sub

    Protected Sub lnkEspelho_Click(sender As Object, e As EventArgs) Handles lnkEspelho.Click
        Try
            ImprimirEspelho()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            viewReport()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEnviarSEFAZ_Click(sender As Object, e As EventArgs) Handles lnkEnviarSEFAZ.Click
        Dim fm As New FilesManager()
        If fm.IsConnect() Then
            EnviarSEFAZ()
        Else
            MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
        End If
    End Sub

    Protected Sub lnkEnviarEmail_Click(sender As Object, e As EventArgs) Handles lnkEnviarEmail.Click
        Dim fm As New FilesManager()
        If fm.IsConnect() Then
            ucEmailNFe.Limpar()
            Dim txtDestinatario As TextBox = CType(ucEmailNFe.FindControlRecursive("txtDestinatario"), TextBox)
            Popup.ConsultaDeEmailNFe(Me.Page, "objEmailCTe" & HID.Value, txtDestinatario.ClientID, 100)
        Else
            MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
        End If
    End Sub

    Protected Sub lnkDACTE_Click(sender As Object, e As EventArgs) Handles lnkDACTE.Click
        Dim fm As New FilesManager()
        If fm.IsConnect() Then
            ImprimirCTe()
        Else
            MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
        End If
    End Sub

    Protected Sub lnkVisualizar_Click(sender As Object, e As EventArgs) Handles lnkVisualizar.Click
        Try
            RecuperarSessaoConhecimento()
            If objConhecimento IsNot Nothing Then
                Dim fileName As String = String.Empty

                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    fileName = Server.MapPath(String.Format("~/Files/nota{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa))
                Else
                    fileName = Server.MapPath(String.Format("~/Files/cte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa))
                End If

                If System.IO.File.Exists(fileName) Then
                    System.IO.File.Delete(fileName)
                End If

                Dim str As String = String.Empty

                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    'str = [Lib].Negocio.DocumentoEletronico.getTextoNFe(objConhecimento, False, 1)
                    'str = [Lib].Negocio.DocumentoEletronico.getTextoNFe3G(objConhecimento, False, 1)
                    'Como ainda não estamos usando troquei porque 3G está descontinuada - Furlan - 12/02/2025
                    str = [Lib].Negocio.DocumentoEletronico.getTextoNFe4G(objConhecimento, False, 1)
                Else
                    str = getTextoCTe(objConhecimento)
                End If

                Using sw As New StreamWriter(fileName)
                    sw.WriteLine(str)
                    sw.Close()
                End Using

                Response.Clear()
                Response.ClearHeaders()
                Response.AddHeader("content-length", str.Length.ToString())
                Response.ContentType = "text/plain"
                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    Response.AppendHeader("content-disposition", "attachment;filename=" & String.Format("nota{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa))
                Else
                    Response.AppendHeader("content-disposition", "attachment;filename=" & String.Format("cte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa))
                End If
                Response.Write(str)
                Response.End()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkDuplicar_Click(sender As Object, e As EventArgs) Handles lnkDuplicar.Click
        Try
            RecuperarSessaoConhecimento()
            Duplicar(objConhecimento)
            LimparCampos()
            MsgBox(Me.Page, "CT-e " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " duplicado com Sucesso.", eTitulo.Sucess)
            tcPrincipal.ActiveTab = tabCTRC
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Private Sub Salvar()
        Try
            If Funcoes.VerificaPermissao("ConhecimentoDeTransporte", "GRAVAR") Then
                Dim Sqls As New ArrayList
                If String.IsNullOrWhiteSpace(txtNumero.Text) OrElse CInt(txtNumero.Text) = 0 OrElse String.IsNullOrWhiteSpace(txtSerie.Text) Then
                    MsgBox(Me.Page, "É necessário informar os campos de número e série!")
                    Exit Sub
                End If

                If String.IsNullOrWhiteSpace(txtDataDeEmissao.Text) OrElse String.IsNullOrWhiteSpace(txtDataDeSaida.Text) Then
                    MsgBox(Me.Page, "É necessário informar os campos de data de emissão e data de saída!")
                    Exit Sub
                End If

                If Not gridEncargos.Rows.Count > 0 Then
                    MsgBox(Me.Page, "É necessário informar os encargos do conhecimento de transporte!")
                    Exit Sub
                End If

                If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA _
                    AndAlso (String.IsNullOrWhiteSpace(txtVencimentoFatura.Text) OrElse Not IsDate(txtVencimentoFatura.Text)) Then
                    MsgBox(Me.Page, "É necessário uma data válida no campo vencimento da fatura!")
                    txtVencimentoFatura.Focus()
                    Exit Sub
                End If

                If Not String.IsNullOrWhiteSpace(txtUnitario.Text) Then
                    Calcular()
                    RecuperarSessaoConhecimento()

                    If rdoPamcard.Checked Then
                        objConhecimento.Transportador = New [Lib].Negocio.Cliente(objConhecimento.CodigoTransportador, objConhecimento.EnderecoTransportador)
                        If objConhecimento.Transportador Is Nothing OrElse objConhecimento.Transportador.ContasBancarias Is Nothing OrElse Not objConhecimento.Transportador.ContasBancarias.Count > 0 Then
                            MsgBox(Me.Page, "É necessário que o conveniado/transportador possua uma conta bancária cadastrada!")
                            Exit Sub
                        End If
                    End If

                    objConhecimento.CarregandoNota = True
                    objConhecimento.NossaEmissao = getNossaEmissao()
                    objConhecimento.Eletronica = getEletronica()
                    objConhecimento.CodigoSituacao = getSituacao()
                    objConhecimento.CodigoTipoDeDocumentoFrete = CInt(CType(hdfTipo.Value, eTipoDeDocumentoFrete))
                    objConhecimento.TipoDeDocumentoFrete = CType(hdfTipo.Value, eTipoDeDocumentoFrete)
                    objConhecimento.CIFFOB = IIf(getEntradaSaida() = eEntradaSaida.Saida, eTiposFrete.CIF, eTiposFrete.FOB)
                    objConhecimento.DataNota = CDate(txtDataDeEmissao.Text)
                    objConhecimento.Movimento = CDate(txtDataDeSaida.Text)
                    objConhecimento.ObservacoesDeEmbarque = Funcoes.EliminarCaracteresEspeciaisNF(txtObservacao.Text.ToUpper().Trim())
                    objConhecimento.CarregandoNota = False

                    If String.IsNullOrWhiteSpace(objConhecimento.CodigoTransportador) Then
                        MsgBox(Me.Page, "É necessário informar o transportador!")
                        Exit Sub
                    End If

                    If objConhecimento.PlacaDetalhes Is Nothing OrElse objConhecimento.PlacaDetalhes.Placa01 Is Nothing OrElse objConhecimento.PlacaDetalhes.Placa01.Length = 0 Then
                        MsgBox(Me.Page, "É necessário informar a Placa do veiculo!")
                        Exit Sub
                    End If

                    If String.IsNullOrWhiteSpace(objConhecimento.CodigoMotorista) Then
                        MsgBox(Me.Page, "É necessário informar o motorista, faça o cadastro corretamenta da placa! Placa informada: " & objConhecimento.PlacaDetalhes.Placa01)
                        Exit Sub
                    End If

                    If CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Circulacao _
                            AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Anulacao Then
                        objConhecimento.Codigo = txtNumero.Text.Trim()
                        objConhecimento.Serie = txtSerie.Text.Trim()
                        objConhecimento.ChaveNFE = txtChaveNFe.Text.OnlyNumbers().Trim()
                    End If

                    If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA Then
                        objConhecimento.NotaProdutor = txtNumero.Text.Trim()
                        objConhecimento.SerieNotaProdutor = txtSerie.Text.Trim()
                    End If

                    If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
                        For index = 0 To objConhecimento.Itens.Count - 1
                            Dim strNotaFiscal() As String = hdfCodigo.Value.Split(New Char() {"-"c}, StringSplitOptions.RemoveEmptyEntries)
                            Dim nfd As New [Lib].Negocio.NotaFiscalDevolucaoXNotaFiscal(objConhecimento.Itens(index))
                            nfd.IUD = "I"
                            nfd.NumeroNota = strNotaFiscal(0)
                            nfd.Serie = strNotaFiscal(1)
                            nfd.EntradaSaida = strNotaFiscal(2)
                            nfd.Sequencia = objConhecimento.NotaTrocaOrigem.Itens(index).Sequencia
                            nfd.CodigoCFOP = objConhecimento.NotaTrocaOrigem.Itens(index).CFOP
                            nfd.QuantidadeDevolucao = objConhecimento.Itens(0).QuantidadeFiscal
                            nfd.ValorDevolucao = objConhecimento.Itens(0).ValorTotal
                            objConhecimento.Itens(index).NotasDevolucao = New [Lib].Negocio.ListNotaFiscalDevolucaoXNotaFiscal(objConhecimento.Itens(index))
                            objConhecimento.Itens(index).NotasDevolucao.Add(nfd)
                        Next
                    End If

                    If objConhecimento IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.IUD) AndAlso objConhecimento.IUD = "U" Then
                        Dim sql = "UPDATE NOTASFISCAIS " & vbCrLf &
                                  "   SET DataDaNota        = '" & CDate(objConhecimento.DataNota).ToString("yyyy-MM-dd") & "'," & vbCrLf &
                                  "       Movimento         = '" & CDate(objConhecimento.Movimento).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                                  " WHERE Empresa_Id      = '" & objConhecimento.CodigoEmpresa & "'" & vbCrLf &
                                  "   AND EndEmpresa_Id   = " & objConhecimento.EnderecoEmpresa & vbCrLf &
                                  "	  AND Cliente_Id      = '" & objConhecimento.CodigoCliente & "'" & vbCrLf &
                                  "   AND EndCliente_Id   = " & objConhecimento.EnderecoCliente & vbCrLf &
                                  "   AND EntradaSaida_Id = '" & objConhecimento.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                  "   AND Nota_Id         = " & objConhecimento.Codigo & vbCrLf &
                                  "   AND Serie_Id        = '" & objConhecimento.Serie & "';" & vbCrLf
                        Sqls.Add(sql)

                        ContabilizarNota(Sqls, True)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            Exit Sub
                        End If

                        'lnkNovo.Parent.Visible = False
                        MsgBox(Me.Page, "Conhecimento " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " atualizado com Sucesso.", eTitulo.Sucess)
                        LimparCampos()
                        tcPrincipal.ActiveTab = tabCTRC
                        Exit Sub
                    End If

                    If rdoBanco.Checked Then
                        SalvarSessaoConhecimento()
                        If objConhecimento.Salvar(Sqls) AndAlso ContabilizarNota(Sqls) Then
                            'lnkNovo.Parent.Visible = False
                            Dim sop As New [Lib].Negocio.SubOperacao(objConhecimento.CodigoOperacao, objConhecimento.CodigoSubOperacao)
                            If sop IsNot Nothing AndAlso sop.Financeiro Then

                                If objConhecimento.CodigoEmpresa = "63358210000176" AndAlso
                                    objConhecimento.Movimento.Year = 2026 AndAlso
                                    (objConhecimento.Movimento.Month = 1 OrElse
                                    objConhecimento.Movimento.Month = 2 OrElse
                                    objConhecimento.Movimento.Month = 3 OrElse
                                    objConhecimento.Movimento.Month = 4) Then
                                    'NÃO GERAR FATURA E FINANCEIRO, APENAS FISCAL
                                Else
                                    SalvarFaturaDeFrete(Sqls)
                                End If
                            End If

                            If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao AndAlso sop.Financeiro Then
                                For Each itemFrete In objConhecimento.Itens
                                    For Each itemFreteDev In itemFrete.NotasDevolucao
                                        For Each tit In itemFreteDev.Nota.VencimentosNota
                                            If Not tit.CodigoProvisao = 1 Then
                                                MsgBox(Me.Page, "Conhecimento de anulação está com Financeiro em aberto, como não foi pago troque para operação Sem Financeiro.")
                                                Exit Sub
                                            End If
                                        Next
                                    Next
                                Next

                                SalvarAnulacao(Sqls)
                            End If

                            If objConhecimento.ChaveNFE.Length > 0 Then

                                Dim sqlDocumentoXML = "     UPDATE DocumentoXML Set Situacao = 104, RecusaNFE = 0"

                                If objConhecimento.EmpresaEmissor Then
                                    sqlDocumentoXML &= "    ,Tipo = 'Mic' "
                                End If

                                sqlDocumentoXML &= "        WHERE chave_id = '" & objConhecimento.ChaveNFE & "';"

                                Sqls.Add(sqlDocumentoXML)

                            End If

                            If Not Banco.GravaBanco(Sqls) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                                Exit Sub
                            End If

                            If objConhecimento IsNot Nothing AndAlso objConhecimento.NossaEmissao Then
                                Dim fm As New FilesManager()
                                If fm.IsConnect() AndAlso VerificarCTe() Then
                                    EnviarSEFAZ()
                                Else
                                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                                End If
                            Else

                                Dim Mensagem As String = "Conhecimento " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " incluído com sucesso! \n "
                                Mensagem &= MensagemFinal
                                MsgBox(Me.Page, Mensagem, eTitulo.Info, False)

                                ImprimirEspelho()
                                'ConsultarNotas()
                                LimparCampos()

                                tcPrincipal.ActiveTab = tabCTRC
                            End If
                        Else
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If
                    ElseIf rdoPamcard.Checked AndAlso Validar() Then
                        If String.IsNullOrWhiteSpace(txtCodigoFavorecidoCartao.Value) Then
                            MsgBox(Me.Page, "É necessário informar o favorecido do cartão pamcard!")
                            Exit Sub
                        End If

                        Dim strFavorecido = txtCodigoFavorecidoCartao.Value.Split("-")
                        Dim objFavorecido As New [Lib].Negocio.Cliente(strFavorecido(0), strFavorecido(1))
                        objConhecimento.CarregandoNota = True
                        objConhecimento.CodigoFavorecido = strFavorecido(0)
                        objConhecimento.EnderecoFavorecido = strFavorecido(1)
                        objConhecimento.CarregandoNota = False

                        If Not objFavorecido.NascimentoConstituicao > DateTime.MinValue Then
                            MsgBox(Me.Page, "É necessário informar o campo de 'Nascimento' no cadastro do favorecido do cartão!")
                            Exit Sub
                        End If

                        If String.IsNullOrWhiteSpace(objFavorecido.Telefone) OrElse objFavorecido.Telefone.Length < 10 Then
                            MsgBox(Me.Page, "É necessário informar o campo 'DDD + Telefone' no cadastro do favorecido do cartão!")
                            Exit Sub
                        End If

                        SalvarSessaoConhecimento()

                        If objConhecimento.Salvar(Sqls) AndAlso ContabilizarNota(Sqls) AndAlso SalvarPamcard(Sqls) Then
                            If Not Banco.GravaBanco(Sqls) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                                Exit Sub
                            End If

                            If objConhecimento IsNot Nothing AndAlso objConhecimento.NossaEmissao Then
                                Dim fm As New FilesManager()
                                If fm.IsConnect() AndAlso VerificarCTe() Then
                                    EnviarSEFAZ()
                                Else
                                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                                End If
                            Else
                                MsgBox(Me.Page, "Conhecimento " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " incluído com Sucesso.", eTitulo.Sucess)
                                ImprimirEspelho()
                                'ConsultarNotas()
                                LimparCampos()
                                tcPrincipal.ActiveTab = tabCTRC
                            End If
                        Else
                            If Not String.IsNullOrWhiteSpace(HttpContext.Current.Session("ssMessage")) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            End If
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Campo valor frete é obrigatório!")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Excluir(Optional ByVal IUD As String = "D")
        Dim Sqls As New ArrayList
        RecuperarSessaoConhecimento()
        objConhecimento.IUD = IUD

        If Not ValidarNotasXNotas() Then
            MsgBox(Me.Page, "Não é possível excluir este conhecimento de transporte, pois ele é uma origem no relacionamento notas x notas!")
            Exit Sub
        End If

        If IUD = "D" OrElse IUD = "C" Then
            For Each item As [Lib].Negocio.NotaFiscalXItem In objConhecimento.Itens
                item.IUD = IUD
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In item.Encargos
                    enc.IUD = IUD
                Next
            Next

            Dim lst As New [Lib].Negocio.ListFaturasDeFretesXItens(objConhecimento)
            For Each item As [Lib].Negocio.FaturaDeFreteXItens In lst
                item.IUD = IUD
                item.SalvarSql(Sqls)
            Next

            Dim faturas As New [Lib].Negocio.ListFaturaDeFrete(objConhecimento, lst)
            For Each item As [Lib].Negocio.FaturaDeFrete In faturas
                item.IUD = IUD
                item.SalvarSql(Sqls)
            Next

            If Not FinanceiroNovo Then
                Dim titulos As New [Lib].Negocio.ListTitulo
                For Each item As [Lib].Negocio.FaturaDeFrete In faturas
                    For Each titfat In item.ListTituloFatura
                        titulos.Add(New [Lib].Negocio.Titulo(titfat.CodigoTitulo, titfat.Titulo.ReceberPagar))
                    Next
                Next

                If Not ValidarExcluir(titulos) Then
                    Exit Sub
                End If

                For Each item As [Lib].Negocio.Titulo In titulos
                    item.IUD = IUD
                    item.SalvarSql(Sqls)
                Next
            Else
                Dim titulos As New [Lib].Negocio.Novo.ListTituloNovo
                For Each item As FaturaDeFrete In faturas
                    For Each titfat In item.ListTituloFatura
                        titulos.Add(New [Lib].Negocio.Novo.TituloNovo(titfat.CodigoTitulo))
                    Next
                Next

                If Not ValidarExcluir(titulos) Then
                    Exit Sub
                End If

                For Each item As [Lib].Negocio.Novo.TituloNovo In titulos
                    item.IUD = IUD
                    item.SalvarSql(Sqls)
                Next
            End If
        End If

        If Not (objConhecimento.NotasReferenciais IsNot Nothing AndAlso objConhecimento.NotasReferenciais.Count > 0) Then
            objConhecimento.NotasReferenciais = New [Lib].Negocio.ListNotaFiscalReferencial(objConhecimento)
        End If

        If objConhecimento.ChaveNFE.Length > 0 Then

            Dim sqlDocumentoXML = "     UPDATE DocumentoXML Set Situacao = 1, Tipo = 'CTe' "

            sqlDocumentoXML &= "        WHERE chave_id = '" & objConhecimento.ChaveNFE & "';"

            Sqls.Add(sqlDocumentoXML)

        End If

        If objConhecimento.NossaEmissao Then
            If CancelarCTe() Then
                objConhecimento.SalvarSql(Sqls)
                Dim sql = " UPDATE NotasFiscais " & vbCrLf &
                          "    SET Situacao = 2 " & vbCrLf &
                          "  WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "'" & vbCrLf &
                          "    AND EndEmpresa_Id = " & objConhecimento.EnderecoEmpresa & vbCrLf &
                          "    AND Cliente_Id = '" & objConhecimento.CodigoCliente & "'" & vbCrLf &
                          "    AND EndCliente_Id = " & objConhecimento.EnderecoCliente & vbCrLf &
                          "    AND EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf &
                          "    AND Serie_Id = '" & objConhecimento.Serie & "'" & vbCrLf &
                          "    AND Nota_Id = " & objConhecimento.Codigo & ";" & vbCrLf
                Sqls.Add(sql)

                sql = " DELETE FROM NFEPendencias " & vbCrLf &
                      "  WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "' " & vbCrLf &
                      "    AND EndEmpresa_Id = " & objConhecimento.EnderecoEmpresa & vbCrLf &
                      "    AND Cliente_Id = '" & objConhecimento.CodigoCliente & "' " & vbCrLf &
                      "    AND EndCliente_Id = " & objConhecimento.EnderecoCliente & vbCrLf &
                      "    AND EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                      "    AND Serie_Id = '" & objConhecimento.Serie & "' " & vbCrLf &
                      "    AND Nota_Id = " & objConhecimento.Codigo & ";" & vbCrLf
                Sqls.Add(sql)

                sql = "UPDATE NFERealizadas " & vbCrLf &
                      "   SET Retorno = '101', MsgRetorno = 'Cancelamento de CT-e homologado' " & vbCrLf &
                      " WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "' " & vbCrLf &
                      "   AND EndEmpresa_Id = " & objConhecimento.EnderecoEmpresa & vbCrLf &
                      "   AND Cliente_Id = '" & objConhecimento.CodigoCliente & "' " & vbCrLf &
                      "   AND EndCliente_Id = " & objConhecimento.EnderecoCliente & vbCrLf &
                      "   AND EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                      "   AND Serie_Id = '" & objConhecimento.Serie & "' " & vbCrLf &
                      "   AND Nota_Id = " & objConhecimento.Codigo & "; " & vbCrLf
                Sqls.Add(sql)

                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                    sql = "DELETE R "
                    sql &= " FROM Razao R" & vbCrLf
                    sql &= "INNER JOIN NotasFiscais NF" & vbCrLf
                    sql &= "   ON NF.Cliente_Id = R.Cliente_Nf" & vbCrLf
                    sql &= "  AND NF.EndCliente_Id = R.EndCliente_Nf" & vbCrLf
                    sql &= "  AND NF.EntradaSaida_Id = R.EntradaSaida_Nf" & vbCrLf
                    sql &= "  AND NF.Serie_Id = R.Serie_Nf" & vbCrLf
                    sql &= "  AND NF.Nota_Id = R.Numero_Nf" & vbCrLf
                    sql &= "  AND NF.TipoDeDocumento = 2" & vbCrLf
                    sql &= "  AND NF.TipoDeDocumentoFrete = 5" & vbCrLf
                    sql &= "WHERE 1=1 " & vbCrLf
                    sql &= "  AND R.Numero_Nf = '" & objConhecimento.Codigo & "' " & vbCrLf
                    sql &= "  AND R.EntradaSaida_Nf = 'E' " & vbCrLf
                    sql &= "  AND R.Serie_Nf = '" & objConhecimento.Serie & "'"
                    Sqls.Add(sql)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
            End If
        Else
            objConhecimento.SalvarSql(Sqls)
            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If
        End If

        Dim msg As String = String.Empty
        If IUD = "C" Then
            msg = "Cancelamento de CT-e " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " homologado com sucesso!"
        Else
            msg = "Conhecimento " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " excluído com sucesso!"
        End If

        MsgBox(Me.Page, msg)
        LimparCampos()
    End Sub

    Private Function ValidarNotasXNotas() As Boolean
        Dim aux As Boolean = True
        RecuperarSessaoConhecimento()
        Dim sql As String = "SELECT 1" & vbCrLf &
                            "  FROM NotasXNotas" & vbCrLf &
                            " WHERE OrigemEmpresa_Id      ='" & objConhecimento.CodigoEmpresa & "' " & vbCrLf &
                            "   AND OrigemEndEmpresa_Id   ='" & objConhecimento.EnderecoEmpresa & "' " & vbCrLf &
                            "   AND OrigemCliente_Id      ='" & objConhecimento.CodigoCliente & "' " & vbCrLf &
                            "   AND OrigemEndCliente_Id   ='" & objConhecimento.EnderecoCliente & "' " & vbCrLf &
                            "   AND OrigemEntradaSaida_Id ='" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                            "   AND OrigemSerie_Id        ='" & objConhecimento.Serie & "' " & vbCrLf &
                            "   AND OrigemNota_Id         ='" & objConhecimento.Codigo & "' "

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NotasXNotas")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            aux = False
        End If
        Return aux
    End Function

    Private Function ValidarExcluir(ByVal titulos As [Lib].Negocio.ListTitulo) As Boolean
        Dim aux As Boolean = True
        RecuperarSessaoConhecimento()
        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
            Dim objClienteXEmpresa As New [Lib].Negocio.ClienteXEmpresa(objConhecimento.CodigoEmpresa, objConhecimento.EnderecoEmpresa)
            If (objConhecimento.DataHoraNFE.HasValue AndAlso DateTime.Now > CDate(objConhecimento.DataHoraNFE).AddDays(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2)) Then
                MsgBox(Me.Page, "Este conhecimento não pode ser cancelado, pois o prazo para cancelamento terminou em: " & CDate(objConhecimento.DataHoraNFE).AddDays(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2).ToString("dd/MM/yyyy HH:mm:ss"))
                aux = False
            End If

            If objConhecimento.NossaEmissao AndAlso String.IsNullOrWhiteSpace(objConhecimento.ObservacaoCancelamento) Then
                Popup.ConsultaDeCancelamento(Me.Page, "objCancelamento" & HID.Value)
                aux = False
            End If
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao AndAlso titulos.Any(Function(s) s.Provisao.Codigo = CInt(eProvisao.Baixa)) Then
            MsgBox(Me.Page, "Não é possível excluir o conhecimento de transporte, pois existe um título baixado!")
            aux = False
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao AndAlso titulos.Any(Function(s) s.Provisao.Codigo = CInt(eProvisao.Provisao)) Then
            MsgBox(Me.Page, "Não é possível excluir o conhecimento de transporte, pois existe um título em provisão!")
            aux = False
        End If

        If (CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro _
                OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA) _
                AndAlso titulos.Any(Function(s) s.Provisao IsNot Nothing AndAlso s.Provisao.Codigo = CInt(eProvisao.Baixa)) Then
            MsgBox(Me.Page, "Não é possível excluir o conhecimento de transporte, pois existe um título baixado!")
            aux = False
        End If
        Return aux
    End Function

    Private Function ValidarExcluir(ByVal titulos As [Lib].Negocio.Novo.ListTituloNovo) As Boolean
        Dim aux As Boolean = True
        RecuperarSessaoConhecimento()
        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
            Dim objClienteXEmpresa As New [Lib].Negocio.ClienteXEmpresa(objConhecimento.CodigoEmpresa, objConhecimento.EnderecoEmpresa)
            If (objConhecimento.DataHoraNFE.HasValue AndAlso DateTime.Now > CDate(objConhecimento.DataHoraNFE).AddDays(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2)) Then
                MsgBox(Me.Page, "Este conhecimento não pode ser cancelado, pois o prazo para cancelamento terminou em: " & CDate(objConhecimento.DataHoraNFE).AddDays(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2).ToString("dd/MM/yyyy HH:mm:ss"))
                aux &= False
            End If

            If String.IsNullOrWhiteSpace(objConhecimento.ObservacaoCancelamento) Then
                Popup.ConsultaDeCancelamento(Me.Page, "objCancelamento" & HID.Value)
                aux &= False
            End If
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao AndAlso titulos.Any(Function(s) s.Provisao.Codigo = CInt(eProvisao.Baixa)) Then
            MsgBox(Me.Page, "Não é possível excluir o conhecimento de transporte, pois existe um título baixado!")
            aux &= False
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao AndAlso titulos.Any(Function(s) s.Provisao.Codigo = CInt(eProvisao.Provisao)) Then
            MsgBox(Me.Page, "Não é possível excluir o conhecimento de transporte, pois existe um título em provisão!")
            aux &= False
        End If

        If (CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro _
                OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA) _
                AndAlso titulos.Any(Function(s) s.Provisao IsNot Nothing AndAlso s.Provisao.Codigo = CInt(eProvisao.Baixa)) Then
            MsgBox(Me.Page, "Não é possível excluir o conhecimento de transporte, pois existe um título baixado!")
            aux &= False
        End If
        Return aux
    End Function

    Private Sub Cancelar(ByVal observacao As String)
        RecuperarSessaoConhecimento()
        If objConhecimento IsNot Nothing Then
            objConhecimento.IUD = "C"
            objConhecimento.ObservacaoCancelamento = observacao
            Excluir(objConhecimento.IUD)
            SalvarSessaoConhecimento()
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal str As String)
        If Session("objCancelamento" & HID.Value) IsNot Nothing Then
            Cancelar(str)
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objEmpresaCTRC" & HID.Value) Is Nothing Then
            Dim objEmpresa = CType(Session("objEmpresaCTRC" & HID.Value), [Lib].Negocio.Cliente)
            txtNomeDaEmpresa.Text = objEmpresa.Nome
            txtEnderecoDaEmpresa.Text = objEmpresa.Endereco & IIf(objEmpresa.Numero.ToString.Length > 0, ", " & objEmpresa.Numero.ToString, "")
            txtComplementoDaEmpresa.Text = objEmpresa.Complemento
            txtBairroDaEmpresa.Text = objEmpresa.Bairro
            txtCepDaEmpresa.Text = objEmpresa.CEP
            txtInscricaoDaEmpresa.Text = objEmpresa.InscricaoEstadual
            txtCnpjDaEmpresa.Text = objEmpresa.CodigoFormatado
            RecuperarSessaoConhecimento()
            objConhecimento.CodigoEmpresa = objEmpresa.Codigo
            objConhecimento.EnderecoEmpresa = objEmpresa.CodigoEndereco
            objConhecimento.Empresa = objEmpresa
            SalvarSessaoConhecimento()
            Session.Remove("objEmpresaCTRC" & HID.Value)
        ElseIf Session("objClienteCTRC" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtCodigoCliente.Value = Funcoes.FormatarListItemCliente(obj).Value
            txtNomeCliente.Text = Funcoes.FormatarListItemCliente(obj).Text
            Session.Remove("objClienteCTRC" & HID.Value)
        ElseIf Session("objPedidoRPA" & HID.Value) IsNot Nothing Then 'Pedido para as Notas fiscais referenciais de origem do RPA
            Dim objPedido As [Lib].Negocio.Pedido = CType(Session("objPedidoRPA" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = objPedido.Codigo
            Session.Remove("objPedidoRPA" & HID.Value)

        ElseIf Session("objClienteRPA" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtCodigoClienteRPA.Value = Funcoes.FormatarListItemCliente(obj).Value
            txtNomeClienteRPA.Text = Funcoes.FormatarListItemCliente(obj).Text
            Session.Remove("objClienteRPA" & HID.Value)
        ElseIf Session("objFavorecidoCTRC" & HID.Value) IsNot Nothing Then
            RecuperarSessaoConhecimento()
            Dim objFavorecido As [Lib].Negocio.Cliente = CType(Session("objFavorecidoCTRC" & HID.Value), [Lib].Negocio.Cliente)
            If Not objConhecimento.CartaoNovo Then
                Dim ws As New [Lib].Negocio.Pamcard()
                Dim ds As DataSet = ws.ConsultarFavorecido("04854422000185", objFavorecido.Codigo)
                If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Erro ao consultar favorecido, entre em contato com o suporte!")
                    Session.Remove("objFavorecidoCTRC" & HID.Value)
                    Exit Sub
                Else
                    If ds.Tables(0).Rows(0).Item("erro") = 0 Then
                        If Not ds.Tables(0).Rows(0).Item("favStatus").ToString.ToUpper() = "ATIVO" Then
                            MsgBox(Me.Page, "Favorecido não está ativo, verifique: " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                            Session.Remove("objFavorecidoCTRC" & HID.Value)
                            Exit Sub
                        End If
                    Else
                        MsgBox(Me.Page, "Erro consultando favorecido " & objConhecimento.Transportador.Codigo & " - " & objConhecimento.Transportador.Nome & " : Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                        Session.Remove("objFavorecidoCTRC" & HID.Value)
                        Exit Sub
                    End If
                End If
            Else
                txtCodigoFavorecidoCartao.Value = objFavorecido.Codigo & "-" & objFavorecido.CodigoEndereco
                txtFavorecidoCartao.Text = objFavorecido.Nome
                txtEnderecoFavorecidoCartao.Text = objFavorecido.CodigoEndereco
                txtCidadeFavorecidoCartao.Text = objFavorecido.Cidade
                txtEstadoFavorecidoCartao.Text = objFavorecido.CodigoEstado
                txtCPFFavorecidoCartao.Text = objFavorecido.Codigo
                Session.Remove("objFavorecidoCTRC" & HID.Value)
            End If
        ElseIf Session("objTransportadorCTRC" & HID.Value) IsNot Nothing Then
            Dim objTransportador As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtNomeDoTransportador.Text = objTransportador.Nome
            txtRNTRCTransportador.Text = objTransportador.RNTRCTransportador
            txtEnderecoDoTransportador.Text = objTransportador.Endereco & IIf(objTransportador.Numero.ToString.Length > 0, ", " & objTransportador.Numero.ToString, "")
            txtCidadeDoTransportador.Text = objTransportador.Cidade
            txtEstadoDoTransportador.Text = objTransportador.CodigoEstado
            txtCnpjDoTransportador.Text = objTransportador.CodigoFormatado
            txtInscricaoDoTransportador.Text = objTransportador.InscricaoEstadual

            RecuperarSessaoConhecimento()
            objConhecimento.CarregandoNota = True
            'objConhecimento.CodigoCliente = objTransportador.Codigo
            'objConhecimento.EnderecoCliente = objTransportador.CodigoEndereco
            objConhecimento.CodigoTransportador = objTransportador.Codigo
            objConhecimento.EnderecoTransportador = objTransportador.CodigoEndereco
            objConhecimento.Transportador = objTransportador
            objConhecimento.NotaTrocaOrigem.CodigoTransportador = objTransportador.Codigo
            objConhecimento.NotaTrocaOrigem.EnderecoTransportador = objTransportador.CodigoEndereco
            objConhecimento.NotaTrocaOrigem.Transportador = objTransportador

            If objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso Not objConhecimento.NotaTrocaOrigem.NossaEmissao AndAlso Not rdoFerroviario.Checked Then
                Dim Sqls As New ArrayList
                Dim sql As String = "UPDATE NotasFiscaisXTransportadores SET Proprietario = '" & objTransportador.Codigo & "', EndProprietario = '" & objTransportador.CodigoEndereco & "' " & vbCrLf &
                         "WHERE 1=1 and Empresa_Id = '" & objConhecimento.NotaTrocaOrigem.CodigoEmpresa & "' and EndEmpresa_Id = '" & objConhecimento.NotaTrocaOrigem.EnderecoEmpresa & "' " & vbCrLf &
                         "AND Cliente_Id = '" & objConhecimento.NotaTrocaOrigem.CodigoCliente & "' and EndCliente_Id = '" & objConhecimento.NotaTrocaOrigem.EnderecoCliente & "' " & vbCrLf &
                         "AND EntradaSaida_Id = '" & IIf(objConhecimento.NotaTrocaOrigem.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                         "AND Nota_Id = '" & objConhecimento.NotaTrocaOrigem.Codigo & "' and Serie_Id = '" & objConhecimento.NotaTrocaOrigem.Serie & "'"
                Sqls.Add(sql)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
            End If

            'cruzamento da operação e suboperação atráves do tipo de pessoa do cliente da nota, da classe da operação e se a suboperação é entrada ou saída
            Dim classe As String = IIf(objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso objConhecimento.NotaTrocaOrigem.SubOperacao IsNot Nothing,
                                    objConhecimento.NotaTrocaOrigem.SubOperacao.Classe.ToString(), String.Empty)
            Dim entradaSaida As String = IIf(getEntradaSaida() = eEntradaSaida.Saida, "S", "E")
            Dim tipo As Nullable(Of eTipoOperacao) = Nothing
            Dim fisicaJuridica As String = String.Empty
            Dim financeiro As String = String.Empty

            If classe.ToUpper().Trim().Contains("CONTAEORDEM") OrElse classe.ToUpper().Trim().Contains("REMESSAS") OrElse classe.ToUpper().Trim().Contains("AFIXAR") Then
                If objConhecimento.NotaTrocaOrigem.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                    classe = eClassesOperacoes.VENDAS.ToString
                Else
                    classe = eClassesOperacoes.COMPRAS.ToString
                End If
            End If

            If objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.NotaTrocaOrigem.CodigoTransportador) Then
                fisicaJuridica = IIf(objConhecimento.NotaTrocaOrigem.CodigoTransportador.IsCnpj(), "J", "F")
            End If

            If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro Then
                financeiro = "S"
            ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro Then
                financeiro = "N"
            End If

            If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
                classe = eClassesOperacoes.FRETES.ToString
                tipo = eTipoOperacao.Estadia
            End If

            Dim objOperacaoDeFrete As [Lib].Negocio.OperacoesDeFretes = Nothing
            If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
                objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(objConhecimento.NotaTrocaOrigem.CodigoOperacao, objConhecimento.NotaTrocaOrigem.CodigoSubOperacao)
            Else
                Dim Fin As Integer = 0
                If objConhecimento.NotaTrocaOrigem.Pedido.CodigoFinalidade = 7 Or objConhecimento.NotaTrocaOrigem.Pedido.CodigoFinalidade = 3 Then
                    Fin = objConhecimento.NotaTrocaOrigem.Pedido.CodigoFinalidade
                End If

                'objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(classe, entradaSaida, fisicaJuridica, financeiro, tipo, Fin)

                objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(
                    entradaSaida,                                              ' Entrada/Saída do documento
                    objConhecimento.EmpresaEmissor,                            ' Emissor do CTe é a propria Empresa
                    objConhecimento.Tomador.Estado.Codigo = objConhecimento.Empresa.Estado.Codigo, ' Tomador e Empresa emissor na mesma UF?
                    IIf(objConhecimento.NotaDeTroca.EntradaSaida = eEntradaSaida.Entrada, objConhecimento.NotaDeTroca.Cliente.Estado.Codigo, objConhecimento.NotaDeTroca.Empresa.Estado.Codigo) = objConhecimento.Empresa.Estado.Codigo, ' UF da nota de origem é o mesmo da empresa emissor                   
                    objConhecimento.NotaDeTroca.Empresa.Estado.Codigo = objConhecimento.Destino.Estado.Codigo ' UF da nota de origem é o mesmo do destino da nota origem
                )

            End If

            If (objOperacaoDeFrete IsNot Nothing AndAlso objOperacaoDeFrete.OperacaoId > 0 AndAlso objOperacaoDeFrete.SubOperacaoId > 0) Then
                If ((CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                            OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro _
                            OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia) _
                        AndAlso objOperacaoDeFrete.OpDestinoId IsNot Nothing AndAlso objOperacaoDeFrete.SubOpDestinoId IsNot Nothing) Then
                    objConhecimento.CodigoOperacao = objOperacaoDeFrete.OpDestinoId
                    objConhecimento.CodigoSubOperacao = objOperacaoDeFrete.SubOpDestinoId
                ElseIf (CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao _
                        AndAlso objOperacaoDeFrete.OpAnulacao IsNot Nothing AndAlso objOperacaoDeFrete.SubOpAnulacao IsNot Nothing) Then
                    objConhecimento.CodigoOperacao = objOperacaoDeFrete.OpAnulacao
                    objConhecimento.CodigoSubOperacao = objOperacaoDeFrete.SubOpAnulacao
                Else
                    objConhecimento.CodigoOperacao = objOperacaoDeFrete.OperacaoId
                    objConhecimento.CodigoSubOperacao = objOperacaoDeFrete.SubOperacaoId
                End If
            Else
                Dim msg As String = String.Format("Não foi possível encontrar uma operação de frete para a nota fiscal selecionada, com os seguintes parâmetros: CLASSE: {0} - CPF/CNPJ: {1} - E/S: {2} - FINANCEIRO: {3}!", classe, fisicaJuridica, entradaSaida, financeiro)
                If (CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                        OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro) _
                        AndAlso Not objConhecimento.NotaTrocaOrigem.CodigoTransportador.IsCnpj() Then
                    MsgBox(Me.Page, msg)
                    btnTransportador.Enabled = True
                Else
                    MsgBox(Me.Page, msg)
                    LimparCampos()
                    Exit Sub
                End If
            End If

            If objConhecimento.CodigoOperacao > 0 AndAlso objConhecimento.CodigoSubOperacao > 0 Then
                Dim op As New SubOperacao(objConhecimento.CodigoOperacao, objConhecimento.CodigoSubOperacao)
                txtOperacao.Text = String.Format("{0}-{1} - {2}", objConhecimento.CodigoOperacao, objConhecimento.CodigoSubOperacao, IIf(op IsNot Nothing, op.Descricao, String.Empty))

                txtEntSai.Text = String.Format("{0} - {1}", IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA"), "ID OperaçãoXEncargos: " & objConhecimento.Itens.Where(Function(s) s.IUD <> "D")(0).CodigoOperacaoEstado)

                rdoBanco.Enabled = op.Financeiro
                rdoPamcard.Enabled = op.Financeiro
                getEncargos()
            End If

            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlacaCTRC" & HID.Value)
            objConhecimento.CarregandoNota = False
            Session.Remove("objTransportadorCTRC" & HID.Value)
            SalvarSessaoConhecimento()
        ElseIf Session("objPlacaCTRC" & HID.Value) IsNot Nothing Then
            RecuperarSessaoConhecimento()
            objConhecimento.PlacaDetalhes = CType(Session("objPlacaCTRC" & HID.Value), [Lib].Negocio.Placa)

            If objConhecimento.PlacaDetalhes.Motorista IsNot Nothing Then
                txtNomeDoMotorista.Text = objConhecimento.PlacaDetalhes.Motorista.Nome
                txtEnderecoDoMotorista.Text = objConhecimento.PlacaDetalhes.Motorista.Endereco & IIf(objConhecimento.PlacaDetalhes.Motorista.Numero.ToString.Length > 0, ", " & objConhecimento.PlacaDetalhes.Motorista.Numero.ToString, "")
                txtCidadeDoMotorista.Text = objConhecimento.PlacaDetalhes.Motorista.Cidade
                txtEstadoDoMotorista.Text = objConhecimento.PlacaDetalhes.Motorista.CodigoEstado
                txtCPFDoMotorista.Text = objConhecimento.PlacaDetalhes.Motorista.CodigoFormatado
                txtHabilitacaoDoMotorista.Text = objConhecimento.PlacaDetalhes.Motorista.Habilitacao

                objConhecimento.CodigoMotorista = objConhecimento.PlacaDetalhes.Motorista.Codigo
                objConhecimento.EnderecoMotorista = objConhecimento.PlacaDetalhes.Motorista.CodigoEndereco
            End If

            If objConhecimento.PlacaDetalhes IsNot Nothing Then

                objConhecimento.PlacaTransportador = objConhecimento.PlacaDetalhes.Placa01

                txtCPlaca1.Text = objConhecimento.PlacaDetalhes.Placa01
                txtCidadePlaca1.Text = objConhecimento.PlacaDetalhes.CidadePlaca01
                txtEstadoPlaca1.Text = objConhecimento.PlacaDetalhes.EstadoPlaca01
                txtRNTRCPlaca1.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca01
                txtProprietarioPlaca1.Text = objConhecimento.PlacaDetalhes.CodigoProprietario01

                If Not String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.Placa02) Then
                    txtCPlaca2.Text = objConhecimento.PlacaDetalhes.Placa02
                    txtCidadePlaca2.Text = objConhecimento.PlacaDetalhes.CidadePlaca02
                    txtEstadoPlaca2.Text = objConhecimento.PlacaDetalhes.EstadoPlaca02
                    txtRNTRCPlaca2.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca02
                    txtProprietarioPlaca2.Text = objConhecimento.PlacaDetalhes.CodigoProprietario02
                End If

                If Not String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.Placa03) Then
                    txtCPlaca3.Text = objConhecimento.PlacaDetalhes.Placa03
                    txtCidadePlaca3.Text = objConhecimento.PlacaDetalhes.CidadePlaca03
                    txtEstadoPlaca3.Text = objConhecimento.PlacaDetalhes.EstadoPlaca03
                    txtRNTRCPlaca3.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca03
                    txtProprietarioPlaca3.Text = objConhecimento.PlacaDetalhes.CodigoProprietario03
                End If
            End If
            SalvarSessaoConhecimento()
        ElseIf Session("objClienteNP" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(Session("objClienteNP" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtNomeCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteNP" & HID.Value)



        ElseIf Session("objEmailCTe" & HID.Value) IsNot Nothing Then
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                Dim Sqls As New ArrayList
                Try
                    RecuperarSessaoConhecimento()
                    objConhecimento = New [Lib].Negocio.NotaFiscal(objConhecimento)
                    Dim lst As [Lib].Negocio.ListCliente = CType(Session("objEmailCTe" & HID.Value), [Lib].Negocio.ListCliente)
                    If Not String.IsNullOrWhiteSpace(objConhecimento.Empresa.EmailNFE) Then
                        Dim objFil As New [Lib].Negocio.Fil()
                        objFil.IUD = "I"
                        If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                            objFil.NomeArquivo = String.Format("email{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                        Else
                            objFil.NomeArquivo = String.Format("emailcte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                        End If
                        objFil.Texto = getTextoEmail(objConhecimento, lst, Session("strAssunto" & HID.Value), Session("strMensagem" & HID.Value), CBool(Session("strCompactado" & HID.Value)))
                        objFil.SalvarSql(Sqls)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If

                        Dim resp As [Lib].Negocio.Resp = Nothing
                        Dim fileName As String = String.Empty
                        If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                            fileName = String.Format("resp-email{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                        Else
                            fileName = String.Format("resp-emailcte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                        End If

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

                            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4014" Then
                                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                                Exit Sub
                            Else
                                MsgBox(Me.Page, strMsg)
                            End If
                        End If
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
            End If
        End If
    End Sub

    Public Sub Recontabilizar()
        Try
            RecuperarSessaoConhecimento()
            Dim Sqls As New ArrayList
            If objConhecimento IsNot Nothing AndAlso ContabilizarNota(Sqls, True) Then
                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
                MsgBox(Me.Page, "Conhecimento " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " recontalizado com Sucesso.", eTitulo.Sucess)
                ImprimirEspelho()
                LimparCampos()
                tcPrincipal.ActiveTab = tabCTRC
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ContabilizarNota(ByRef Sqls As ArrayList, Optional ByVal recontabilizar As Boolean = False) As Boolean
        RecuperarSessaoConhecimento()
        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao AndAlso recontabilizar AndAlso objConhecimento.CodigoOperacao <> 80 Then
            MsgBox(Me.Page, "Este conhecimento de anulação foi emitido pela página de notas gerais, por favor recontabilize por lá!")
            Return False
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao AndAlso objConhecimento.DataNota.Date < New DateTime(2014, 5, 14).Date Then
            Return True
        End If

        Dim razao As New [Lib].Negocio.Razao(objConhecimento)
        Dim pEmpresa = String.Empty
        Dim pEndereco = String.Empty

        If Left(Session("ssEmpresa"), 8) = Left(objConhecimento.Empresa.Codigo, 8) Then
            pEmpresa = objConhecimento.CodigoEmpresa
            pEndereco = objConhecimento.EnderecoEmpresa
        ElseIf CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao Then
            pEmpresa = objConhecimento.CodigoEmpresa
            pEndereco = objConhecimento.EnderecoEmpresa
        ElseIf CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Comprovacao AndAlso objConhecimento.NotaTrocaOrigem IsNot Nothing Then
            pEmpresa = objConhecimento.NotaTrocaOrigem.CodigoEmpresa
            pEndereco = objConhecimento.NotaTrocaOrigem.EnderecoEmpresa
        Else
            Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
            nf.CodigoEmpresa = objConhecimento.NotaTrocaOrigem.CodigoEmpresa
            nf.EnderecoEmpresa = objConhecimento.NotaTrocaOrigem.EnderecoEmpresa
            nf.CodigoCliente = objConhecimento.NotaTrocaOrigem.CodigoCliente
            nf.EnderecoCliente = objConhecimento.NotaTrocaOrigem.EnderecoCliente
            nf.EntradaSaida = objConhecimento.NotaTrocaOrigem.EntradaSaida
            nf.Codigo = objConhecimento.NotaTrocaOrigem.Codigo
            nf.Serie = objConhecimento.NotaTrocaOrigem.Serie
            nf = New [Lib].Negocio.NotaFiscal(nf)

            Dim origem As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
            origem.CodigoEmpresa = nf.NotasXNotas.OrigemEmpresaCnpj
            origem.EnderecoEmpresa = nf.NotasXNotas.OrigemEndEmpresa
            origem.CodigoCliente = nf.NotasXNotas.OrigemClienteCnpj
            origem.EnderecoCliente = nf.NotasXNotas.OrigemEndCliente
            origem.EntradaSaida = IIf(nf.NotasXNotas.OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            origem.Codigo = nf.NotasXNotas.OrigemNota
            origem.Serie = nf.NotasXNotas.OrigemSerie
            origem = New [Lib].Negocio.NotaFiscal(origem)
            pEmpresa = origem.CodigoEmpresa
            pEndereco = origem.EnderecoEmpresa
        End If

        'If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
        '    Sqls.Add("DELETE R " & vbCrLf &
        '             "  FROM Razao R " & vbCrLf &
        '             " INNER JOIN NotasFiscais nf " & vbCrLf &
        '             "    ON NF.Cliente_Id = R.Cliente_Nf " & vbCrLf &
        '             "   AND NF.EndCliente_Id = R.EndCliente_Nf " & vbCrLf &
        '             "   AND NF.EntradaSaida_Id = R.EntradaSaida_Nf " & vbCrLf &
        '             "   AND NF.Serie_Id = R.Serie_Nf " & vbCrLf &
        '             "   AND NF.Nota_Id = R.Numero_Nf " & vbCrLf &
        '             "   AND NF.TipoDeDocumento = 2 " & vbCrLf &
        '             "   AND NF.TipoDeDocumentoFrete = 1 " & vbCrLf &
        '             " WHERE 1=1 " & vbCrLf &
        '             "   AND R.Cliente_Nf = '" & objConhecimento.CodigoCliente & "'  " & vbCrLf &
        '             "   AND R.EndCliente_Nf = '" & objConhecimento.EnderecoCliente & "'  " & vbCrLf &
        '             "   AND R.EntradaSaida_Nf = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'  " & vbCrLf &
        '             "   AND R.Serie_Nf = '" & objConhecimento.Serie & "'  " & vbCrLf &
        '             "   AND R.Numero_Nf = '" & objConhecimento.Codigo & "'")
        'ElseIf objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Comprovacao Then
        '    Sqls.Add("DELETE R " & vbCrLf &
        '             "  FROM Razao R " & vbCrLf &
        '             " INNER JOIN NotasFiscais nf " & vbCrLf &
        '             "    ON NF.Cliente_Id = R.Cliente_Nf " & vbCrLf &
        '             "   AND NF.EndCliente_Id = R.EndCliente_Nf " & vbCrLf &
        '             "   AND NF.EntradaSaida_Id = R.EntradaSaida_Nf " & vbCrLf &
        '             "   AND NF.Serie_Id = R.Serie_Nf " & vbCrLf &
        '             "   AND NF.Nota_Id = R.Numero_Nf " & vbCrLf &
        '             "   AND NF.TipoDeDocumento = 2 " & vbCrLf &
        '             "   AND NF.TipoDeDocumentoFrete = 2 " & vbCrLf &
        '             "WHERE 1=1 " & vbCrLf &
        '             "  AND R.Cliente_Nf = '" & objConhecimento.CodigoCliente & "'  " & vbCrLf &
        '             "  AND R.EndCliente_Nf = '" & objConhecimento.EnderecoCliente & "'  " & vbCrLf &
        '             "  AND R.EntradaSaida_Nf = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'  " & vbCrLf &
        '             "  AND R.Serie_Nf = '" & objConhecimento.Serie & "'  " & vbCrLf &
        '             "  AND R.Numero_Nf = '" & objConhecimento.Codigo & "'")
        'Else
        '    razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
        '    razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
        '    razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
        '    razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
        'End If

        razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
        razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
        razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
        razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)

        razao.Contabilizador(Sqls, "", IIf(objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.RPA, 9, 21), pEmpresa, pEndereco)
        Return True
    End Function

    Private Function ContabilizarNota(ByVal cte As [Lib].Negocio.NotaFiscal, ByRef Sqls As ArrayList, Optional ByVal recontabilizar As Boolean = False) As Boolean
        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao AndAlso recontabilizar AndAlso cte.CodigoOperacao <> 80 Then
            MsgBox(Me.Page, "Este conhecimento de anulação foi emitido pela página de notas gerais, por favor recontabilize por lá!")
            Return False
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao AndAlso cte.DataNota.Date < New DateTime(2014, 4, 1).Date Then
            Return True
        End If

        Dim razao As New [Lib].Negocio.Razao(cte)
        Dim pEmpresa = String.Empty
        Dim pEndereco = String.Empty

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao Then
            pEmpresa = cte.CodigoEmpresa
            pEndereco = cte.EnderecoEmpresa
        ElseIf CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Comprovacao AndAlso cte.NotaTrocaOrigem IsNot Nothing Then
            pEmpresa = cte.NotaTrocaOrigem.CodigoEmpresa
            pEndereco = cte.NotaTrocaOrigem.EnderecoEmpresa
        Else
            Dim nf As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
            nf.CodigoEmpresa = cte.NotaTrocaOrigem.CodigoEmpresa
            nf.EnderecoEmpresa = cte.NotaTrocaOrigem.EnderecoEmpresa
            nf.CodigoCliente = cte.NotaTrocaOrigem.CodigoCliente
            nf.EnderecoCliente = cte.NotaTrocaOrigem.EnderecoCliente
            nf.EntradaSaida = cte.NotaTrocaOrigem.EntradaSaida
            nf.Codigo = cte.NotaTrocaOrigem.Codigo
            nf.Serie = cte.NotaTrocaOrigem.Serie
            nf = New [Lib].Negocio.NotaFiscal(nf)

            Dim origem As [Lib].Negocio.NotaFiscal = New [Lib].Negocio.NotaFiscal()
            origem.CodigoEmpresa = nf.NotasXNotas.OrigemEmpresaCnpj
            origem.EnderecoEmpresa = nf.NotasXNotas.OrigemEndEmpresa
            origem.CodigoCliente = nf.NotasXNotas.OrigemClienteCnpj
            origem.EnderecoCliente = nf.NotasXNotas.OrigemEndCliente
            origem.EntradaSaida = IIf(nf.NotasXNotas.OrigemEntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            origem.Codigo = nf.NotasXNotas.OrigemNota
            origem.Serie = nf.NotasXNotas.OrigemSerie
            origem = New [Lib].Negocio.NotaFiscal(origem)
            pEmpresa = origem.CodigoEmpresa
            pEndereco = origem.EnderecoEmpresa
        End If

        razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
        razao.Contabilizador(Sqls, "", 21, pEmpresa, pEndereco)
        Return True
    End Function


    Private Sub SalvarFaturaDeFrete(ByRef Sqls As ArrayList)

        If objConhecimento.EmpresaEmissor And objConhecimento.CodigoEmpresa = objConhecimento.CodigoTomador Then
            Exit Sub
        End If

        RecuperarSessaoConhecimento()
        If objConhecimento IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete IsNot Nothing _
                AndAlso objConhecimento.TipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Comprovacao _
                AndAlso objConhecimento.TipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Anulacao Then

            Dim hasAdto As Boolean = False
            Dim objEncargo As New [Lib].Negocio.Encargo("LIQUIDOAPAGAR")

            Dim objTomador As [Lib].Negocio.Cliente

            If objConhecimento.EntradaSaida = eEntradaSaida.Entrada Then
                objTomador = New [Lib].Negocio.Cliente(objConhecimento.CodigoCliente, objConhecimento.EnderecoCliente)
            Else
                objTomador = New [Lib].Negocio.Cliente(objConhecimento.CodigoTomador, objConhecimento.EnderecoTomador)
            End If

            If Not FinanceiroNovo Then
                Dim objAdto As New [Lib].Negocio.Titulo
                If objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo = "ADTODEFRETE" OrElse s.Codigo = "TARIFA PEDAGIO") _
                        AndAlso objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo = "ADTODEFRETE" OrElse s.Codigo = "TARIFA PEDAGIO").Sum(Function(s) s.Valor) > 0 Then

                    'GERAR PRÉ FATURA DE ADIANTAMENTO
                    Dim pf As New [Lib].Negocio.FaturaDeFrete()
                    pf.IUD = "I"
                    pf.CodigoEmpresa = objConhecimento.CodigoEmpresa
                    pf.EnderecoEmpresa = objConhecimento.EnderecoEmpresa
                    pf.CodigoConveniado = objTomador.Codigo
                    pf.EnderecoConveniado = objTomador.CodigoEndereco
                    pf.CodigoFatura = [Lib].Negocio.Numerador.PegarNumero(HttpContext.Current.Session("ssNomeServidor"), eTiposNumerador.FaturaDeFrete)
                    If rdoPamcard.Checked Then
                        pf.ValorDaFatura = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE" OrElse s.Codigo.Trim = "TARIFA PEDAGIO").Sum(Function(s) s.Valor)
                    Else
                        pf.ValorDaFatura = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE").Sum(Function(s) s.Valor)
                    End If
                    pf.Movimento = DateTime.Now

                    'GERAR ITENS DA FATURA DE BAIXA DE ADIANTAMENTO
                    Dim pfi As New [Lib].Negocio.FaturaDeFreteXItens(pf)
                    pfi.IUD = "I"
                    pfi.EmpresaCnpj = objConhecimento.CodigoEmpresa
                    pfi.EmpresaEnd = objConhecimento.EnderecoEmpresa
                    pfi.ClienteCnpj = objTomador.Codigo
                    pfi.ClienteEnd = objTomador.CodigoEndereco
                    pfi.NumeroNota = objConhecimento.Codigo
                    pfi.EntradaSaida = IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                    pfi.Serie = objConhecimento.Serie
                    pfi.CodigoEncargo = "BAIXAADTO"
                    pfi.ValorAdiantamento = pf.ValorDaFatura
                    pfi.ViaAdiantamento = True
                    pfi.PesoDeChegada = IIf(pfi.EntradaSaida = "E" AndAlso objConhecimento.NotasXNotas.PesoBrutoRomaneio > 0, objConhecimento.NotasXNotas.PesoBrutoRomaneio, objConhecimento.NotasXNotas.PesoFiscal)
                    pf.Itens.Add(pfi)

                    If Not pf.JaFaturada Then
                        'GERAR TÍTULO DE ADIANTAMENTO EM PROVISÃO
                        objAdto.IUD = "I"

                        objAdto.CodigoMoeda = 1 'REAIS

                        If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.RPA Then
                            objAdto.Vencimento = CDate(txtVencimentoFatura.Text).ToString("yyyy-MM-dd")
                        Else
                            If rdoPamcard.Checked Then
                                objAdto.CodigoProvisao = CInt(eProvisao.Baixa)
                            Else
                                objAdto.CodigoProvisao = CInt(eProvisao.Provisao)
                            End If
                            objAdto.Vencimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        End If

                        If objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo.Trim = "TARIFA PEDAGIO") _
                            AndAlso Not objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE") Then
                            objAdto.CodigoCarteira = "001002012" 'PAGAMENTO SOMENTE PEDAGIO
                        Else
                            objAdto.CodigoCarteira = "001002011" 'PAGAMENTO DE ADTO DE FRETES E PEDAGIO
                        End If

                        If objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                            objAdto.ReceberPagar = "R"
                        ElseIf objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro Then
                            objAdto.ReceberPagar = "P"
                        Else
                            objAdto.ReceberPagar = "P"
                        End If

                        objAdto.CodigoTipoPgto = 1
                        objAdto.CodigoSituacao = 1
                        objAdto.Movimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        objAdto.DataMoeda = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        objAdto.Prorrogacao = objAdto.Vencimento
                        objAdto.Baixa = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        objAdto.CodigoUnidadeDeNegocio = objConhecimento.Empresa.CodigoUnidadeDeNegocio
                        objAdto.CodigoEmpresa = pf.CodigoEmpresa
                        objAdto.EnderecoEmpresa = pf.EnderecoEmpresa
                        objAdto.CodigoEmpresaPagadora = pf.CodigoEmpresa
                        objAdto.EndEmpresaPagadora = pf.EnderecoEmpresa

                        objAdto.CodigoIndexador = 3
                        objAdto.ContaContabilCliente = IIf(objEncargo Is Nothing OrElse String.IsNullOrWhiteSpace(objEncargo.ContaCredito), "2010103", objEncargo.ContaCredito)

                        objAdto.CodigoCliente = objTomador.Codigo
                        objAdto.EndCliente = objTomador.CodigoEndereco

                        If objAdto.ReceberPagar = "P" AndAlso objTomador.ContasBancarias IsNot Nothing AndAlso objTomador.ContasBancarias.Count > 0 Then
                            objAdto.CodigoBancoCliente = objTomador.ContasBancarias(0).CodigoBanco
                            objAdto.CodigoAgenciaCliente = objTomador.ContasBancarias(0).CodigoAgencia
                            objAdto.DigitoAgenciaCliente = objTomador.ContasBancarias(0).DigitoAgencia
                            objAdto.ContaCliente = objTomador.ContasBancarias(0).ContaCorrente
                            objAdto.DigitoContaCliente = objTomador.ContasBancarias(0).DigitoConta
                        End If

                        objAdto.CodigoDestinatario = pf.CodigoConveniado
                        objAdto.EndDestinatario = pf.EnderecoConveniado
                        objAdto.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                        objAdto.UsuarioInclusaoData = DateTime.Now
                        If rdoPamcard.Checked Then
                            objAdto.Acrescimos = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "TARIFA PEDAGIO").Sum(Function(s) s.Valor)
                        End If
                        objAdto.ValorDoDocumento = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE").Sum(Function(s) s.Valor)
                        objAdto.Historico = "FATURA FRETE NÚMERO " & pf.CodigoFatura & " - PGTO ADTO CONFORME CONHECIMENTO NÚMERO " & objConhecimento.Codigo

                        Dim n As New [Lib].Negocio.Numerador(1)
                        objAdto.Codigo = n.Sequencia + 1
                        objAdto.SalvarSql(Sqls, False)

                        If objAdto IsNot Nothing Then
                            Sqls.Add(New [Lib].Negocio.Numerador(1).IncrementarNumeradorSql(True, 2))
                        End If

                        If rdoPamcard.Checked Then
                            Dim razao As New [Lib].Negocio.Razao()
                            If objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE") AndAlso
                                    objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "TARIFA PEDAGIO") Then
                                Dim objCarteira = New [Lib].Negocio.CarteiraFinanceira(objAdto.CodigoCarteira)
                                razao.ContabilizarPAMCARD(objAdto, objCarteira.CodigoContaCliente, "D", objAdto.ValorDoDocumento, objAdto.MoedaValorDoDocumento, "PGTO ADTO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objCarteira.CodigoContaAcrescimo, "D", objAdto.Acrescimos, Funcoes.ConverteParaMoedaExtrangeira(objAdto.Acrescimos, 3, objAdto.Movimento), "PGTO PEDÁGIO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objAdto.ContaContabilPagadora, "C", objAdto.ValorLiquido, objAdto.MoedaValorLiquido, "PGTO ADTO", Sqls)
                            ElseIf objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE") AndAlso
                                    Not objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "TARIFA PEDAGIO") Then
                                Dim objCarteira = New [Lib].Negocio.CarteiraFinanceira(objAdto.CodigoCarteira)
                                razao.ContabilizarPAMCARD(objAdto, objCarteira.CodigoContaCliente, "D", objAdto.ValorDoDocumento, objAdto.MoedaValorDoDocumento, "PGTO ADTO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objAdto.ContaContabilPagadora, "C", objAdto.ValorLiquido, objAdto.MoedaValorLiquido, "PGTO ADTO", Sqls)
                            ElseIf objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "TARIFA PEDAGIO") AndAlso
                                    Not objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE") Then
                                Dim objCarteira = New [Lib].Negocio.CarteiraFinanceira(objAdto.CodigoCarteira)
                                razao.ContabilizarPAMCARD(objAdto, objCarteira.CodigoContaCliente, "D", objAdto.ValorDoDocumento, objAdto.MoedaValorDoDocumento, "PGTO PEDÁGIO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objAdto.ContaContabilPagadora, "C", objAdto.ValorLiquido, objAdto.MoedaValorLiquido, "PGTO PEDÁGIO", Sqls)
                            End If
                        End If

                        hasAdto = True
                        pf.ListTituloFatura.Add(New FaturaDeFretexTitulo(pf) With {.CodigoTitulo = objAdto.Codigo})
                        pf.SalvarSql(Sqls)
                    Else
                        MsgBox(Me.Page, "Pré-fatura já cadastrada!")
                    End If
                End If
                SalvarSaldo(objEncargo, objTomador, Sqls, hasAdto, objAdto.Codigo)
            Else
                Dim objAdto As New [Lib].Negocio.Novo.TituloNovo
                If objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo = "ADTODEFRETE" OrElse s.Codigo = "TARIFA PEDAGIO") _
                        AndAlso objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo = "ADTODEFRETE" OrElse s.Codigo = "TARIFA PEDAGIO").Sum(Function(s) s.Valor) > 0 Then

                    'GERAR PRÉ FATURA DE ADIANTAMENTO
                    Dim pf As New [Lib].Negocio.FaturaDeFrete()
                    pf.IUD = "I"
                    pf.CodigoEmpresa = objConhecimento.CodigoEmpresa
                    pf.EnderecoEmpresa = objConhecimento.EnderecoEmpresa
                    pf.CodigoConveniado = objTomador.Codigo
                    pf.EnderecoConveniado = objTomador.CodigoEndereco
                    pf.CodigoFatura = [Lib].Negocio.Numerador.PegarNumero(HttpContext.Current.Session("ssNomeServidor"), eTiposNumerador.FaturaDeFrete)
                    If rdoPamcard.Checked Then
                        pf.ValorDaFatura = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE" OrElse s.Codigo.Trim = "TARIFA PEDAGIO").Sum(Function(s) s.Valor)
                    Else
                        pf.ValorDaFatura = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE").Sum(Function(s) s.Valor)
                    End If
                    pf.Movimento = DateTime.Now

                    'GERAR ITENS DA FATURA DE BAIXA DE ADIANTAMENTO
                    Dim pfi As New [Lib].Negocio.FaturaDeFreteXItens(pf)
                    pfi.IUD = "I"
                    pfi.EmpresaCnpj = objConhecimento.CodigoEmpresa
                    pfi.EmpresaEnd = objConhecimento.EnderecoEmpresa
                    pfi.ClienteCnpj = objTomador.Codigo
                    pfi.ClienteEnd = objTomador.CodigoEndereco
                    pfi.NumeroNota = objConhecimento.Codigo
                    pfi.EntradaSaida = IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                    pfi.Serie = objConhecimento.Serie
                    pfi.CodigoEncargo = "BAIXAADTO"
                    pfi.ValorAdiantamento = pf.ValorDaFatura
                    pfi.ViaAdiantamento = True
                    pfi.PesoDeChegada = IIf(pfi.EntradaSaida = "E" AndAlso objConhecimento.NotasXNotas.PesoBrutoRomaneio > 0, objConhecimento.NotasXNotas.PesoBrutoRomaneio, objConhecimento.NotasXNotas.PesoFiscal)
                    pf.Itens.Add(pfi)

                    If Not pf.JaFaturada Then
                        'GERAR TÍTULO DE ADIANTAMENTO EM PROVISÃO
                        objAdto.IUD = "I"

                        objAdto.CodigoMoeda = 1

                        If objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                            objAdto.ReceberPagar = "R"
                        ElseIf objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro Then
                            objAdto.ReceberPagar = "P"
                        Else
                            objAdto.ReceberPagar = "P"
                        End If

                        If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.RPA Then
                            objAdto.CodigoProvisao = CInt(eProvisao.Previsao)
                            objAdto.Vencimento = CDate(txtVencimentoFatura.Text).ToString("yyyy-MM-dd")
                        Else
                            If rdoPamcard.Checked Then
                                objAdto.CodigoProvisao = CInt(eProvisao.Baixa)
                            Else
                                objAdto.CodigoProvisao = CInt(eProvisao.Provisao)
                            End If
                            objAdto.Vencimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        End If

                        objAdto.Movimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        objAdto.Reprogramacao = objAdto.Vencimento
                        objAdto.DataMoeda = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        objAdto.DataBaixa = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                        objAdto.CodigoIndexador = 3
                        objAdto.CodigoTipoPgto = 1
                        objAdto.CodigoSituacao = 1
                        objAdto.CodigoUnidadeDeNegocio = objConhecimento.Empresa.CodigoUnidadeDeNegocio
                        objAdto.CodigoEmpresa = pf.CodigoEmpresa
                        objAdto.EnderecoEmpresa = pf.EnderecoEmpresa
                        objAdto.CodigoEmpresaRecPag = pf.CodigoEmpresa
                        objAdto.EndEmpresaRecPag = pf.EnderecoEmpresa

                        objAdto.CodigoIndexador = 3
                        objAdto.CodigoCliFor = objTomador.Codigo
                        objAdto.EnderecoCliFor = objTomador.CodigoEndereco

                        'CONTA - VALOR DO DOCUMENTO
                        objAdto.CodigoContaContabilCliFor = objConhecimento.SubOperacao.CodigoGrupoContas
                        'CONTA - VALOR LÍQUIDO
                        objAdto.CodigoContaContabilRecPag = objAdto.Empresa.Empresa.CodigoContaGrupoBanco
                        'VALOR DO DOCUMENTO
                        objAdto.Valores.EncargoValorDocumento.ValorOficial = pf.ValorDaFatura
                        objAdto.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(pf.ValorDaFatura, 3, objAdto.Movimento)
                        'VALOR LÍQUIDO
                        objAdto.Valores.AtualizaLiquido()

                        If objAdto.ReceberPagar = "P" AndAlso objConhecimento.Tomador.ContasBancarias IsNot Nothing AndAlso objConhecimento.Tomador.ContasBancarias.Count > 0 Then
                            objAdto.CodigoBancoCliFor = objTomador.ContasBancarias(0).CodigoBanco
                            objAdto.CodigoAgenciaCliFor = objTomador.ContasBancarias(0).CodigoAgencia
                            objAdto.DigitoAgenciaCliFor = objTomador.ContasBancarias(0).DigitoAgencia
                            objAdto.ContaCliFor = objTomador.ContasBancarias(0).ContaCorrente
                            objAdto.DigitoContaCliFor = objTomador.ContasBancarias(0).DigitoConta
                        End If

                        objAdto.TituloDestinacao.CodigoDestinatario = pf.CodigoConveniado
                        objAdto.TituloDestinacao.EndDestinatario = pf.EnderecoConveniado

                        'VALOR DE DEDUÇÃO (ADTO) / VALOR DE ACRÉSCIMO (PEDÁGIO)
                        Dim enc As New [Lib].Negocio.Novo.TituloXContaContabil(objAdto)
                        enc.CodigoContaEncargo = objAdto.Empresa.Empresa.CodigoContaPedagioDeFrete
                        enc.ValorOficial = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo = "TARIFA PEDAGIO").Sum(Function(s) s.Valor)
                        objAdto.Valores.Add(enc)
                        objAdto.Valores.AtualizaLiquido()

                        objAdto.Valores.EncargoValorDocumento.ValorOficial = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE").Sum(Function(s) s.Valor)
                        objAdto.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(objAdto.Valores.EncargoValorDocumento.ValorOficial, 3, objAdto.Movimento)

                        objAdto.Historico = "FATURA FRETE NÚMERO " & pf.CodigoFatura & " - PGTO ADTO CONFORME CONHECIMENTO NÚMERO " & objConhecimento.Codigo

                        Dim n As New [Lib].Negocio.Numerador(1)
                        objAdto.Codigo = n.Sequencia + 1
                        objAdto.SalvarSql(Sqls, False)

                        If objAdto IsNot Nothing Then
                            Sqls.Add(New [Lib].Negocio.Numerador(1).IncrementarNumeradorSql(True, 2))
                        End If

                        If rdoPamcard.Checked Then
                            Dim razao As New [Lib].Negocio.Razao()
                            If objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE") AndAlso
                                    objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "TARIFA PEDAGIO") Then
                                razao.ContabilizarPAMCARD(objAdto, objAdto.CodigoContaContabilCliFor, "D", objAdto.Valores.EncargoValorDocumento.ValorOficial, objAdto.Valores.EncargoValorDocumento.ValorMoeda, "PGTO ADTO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objAdto.Empresa.Empresa.CodigoContaPedagioDeFrete, "D", enc.ValorOficial, Funcoes.ConverteParaMoedaExtrangeira(enc.ValorOficial, 3, objAdto.Movimento), "PGTO PEDÁGIO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objAdto.CodigoContaContabilRecPag, "C", objAdto.Valores.EncargoValorLiquido.ValorOficial, objAdto.Valores.EncargoValorLiquido.ValorMoeda, "PGTO ADTO", Sqls)
                            ElseIf objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE") AndAlso
                                    Not objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "TARIFA PEDAGIO") Then
                                razao.ContabilizarPAMCARD(objAdto, objAdto.CodigoContaContabilCliFor, "D", objAdto.Valores.EncargoValorDocumento.ValorOficial, objAdto.Valores.EncargoValorDocumento.ValorMoeda, "PGTO ADTO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objAdto.CodigoContaContabilRecPag, "C", objAdto.Valores.EncargoValorLiquido.ValorOficial, objAdto.Valores.EncargoValorLiquido.ValorMoeda, "PGTO ADTO", Sqls)
                            ElseIf objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "TARIFA PEDAGIO") AndAlso
                                    Not objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Any(Function(s) s.Codigo = "ADIANTAMENTO" OrElse s.Codigo.Trim = "ADTODEFRETE") Then
                                razao.ContabilizarPAMCARD(objAdto, objAdto.CodigoContaContabilCliFor, "D", objAdto.Valores.EncargoValorDocumento.ValorOficial, objAdto.Valores.EncargoValorDocumento.ValorMoeda, "PGTO PEDÁGIO", Sqls)
                                razao.ContabilizarPAMCARD(objAdto, objAdto.CodigoContaContabilRecPag, "C", objAdto.Valores.EncargoValorLiquido.ValorOficial, objAdto.Valores.EncargoValorLiquido.ValorMoeda, "PGTO PEDÁGIO", Sqls)
                            End If
                        End If

                        hasAdto = True
                        pf.ListTituloFatura.Add(New FaturaDeFretexTitulo(pf) With {.CodigoTitulo = objAdto.Codigo})
                        pf.SalvarSql(Sqls)
                    Else
                        MsgBox(Me.Page, "Pré-fatura já cadastrada!")
                    End If
                End If
                SalvarSaldo(objEncargo, objTomador, Sqls, hasAdto, objAdto.Codigo)
            End If
        End If
    End Sub

    Private Function SalvarPamcard(ByRef Sqls As ArrayList) As Boolean
        RecuperarSessaoConhecimento()
        Dim strFavorecido() As String = txtCodigoFavorecidoCartao.Value.ToString.Split("-")
        Dim objFavorecido As New [Lib].Negocio.Cliente(strFavorecido(0), strFavorecido(1))

        objNotaFiscal = New [Lib].Negocio.NotaFiscal()
        objNotaFiscal.CodigoEmpresa = objConhecimento.NotaTrocaOrigem.CodigoEmpresa
        objNotaFiscal.EnderecoEmpresa = objConhecimento.NotaTrocaOrigem.EnderecoEmpresa
        objNotaFiscal.CodigoCliente = objConhecimento.NotaTrocaOrigem.CodigoCliente
        objNotaFiscal.EnderecoCliente = objConhecimento.NotaTrocaOrigem.EnderecoCliente
        objNotaFiscal.Serie = objConhecimento.NotaTrocaOrigem.Serie
        objNotaFiscal.Codigo = objConhecimento.NotaTrocaOrigem.Codigo
        objNotaFiscal.EntradaSaida = objConhecimento.NotaTrocaOrigem.EntradaSaida
        objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)

        Dim ws As New [Lib].Negocio.Pamcard()
        Dim ds As DataSet = ws.InclusaoContratoDeFrete("04854422000185", objConhecimento, objNotaFiscal, objFavorecido)
        If ds Is Nothing OrElse ds.Tables Is Nothing OrElse Not ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Return False
        Else
            If ds.Tables(0).Rows(0).Item("erro") = 0 Then
                objConhecimento.IUD = "U"
                objConhecimento.CodigoSituacao = getSituacao()
                objConhecimento.ContratoANTT = ds.Tables(0).Rows(0).Item("anttNumero")
                objConhecimento.ProtocoloANTT = ds.Tables(0).Rows(0).Item("anttProtocolo")
                objConhecimento.idPamcard = ds.Tables(0).Rows(0).Item("viagemId")
                SalvarSessaoConhecimento()
                Dim sop As New [Lib].Negocio.SubOperacao(objConhecimento.CodigoOperacao, objConhecimento.CodigoSubOperacao)
                If sop IsNot Nothing AndAlso sop.Financeiro Then
                    SalvarFaturaDeFrete(Sqls)
                End If
                ImprimeContrato(objConhecimento)
            Else
                HttpContext.Current.Session("ssMessage") = String.Empty
                MsgBox(Me.Page, ds.Tables(0).Rows(0).Item("errodescricao"))
                Return False
            End If
        End If
        Return True
    End Function

    Private Sub SalvarSaldo(ByVal objEncargo As [Lib].Negocio.Encargo, ByVal objTomador As [Lib].Negocio.Cliente, ByRef Sqls As ArrayList, Optional ByVal hasAdto As Boolean = False, Optional numerador As Integer = 0)
        RecuperarSessaoConhecimento()
        If objConhecimento IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete IsNot Nothing _
            AndAlso objConhecimento.TipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Comprovacao _
            AndAlso objConhecimento.TipoDeDocumentoFrete <> eTipoDeDocumentoFrete.Anulacao Then

            'GERAR PRÉ FATURA DE SALDO
            Dim pfSaldo As New [Lib].Negocio.FaturaDeFrete()
            pfSaldo.IUD = "I"
            pfSaldo.CodigoEmpresa = objConhecimento.CodigoEmpresa
            pfSaldo.EnderecoEmpresa = objConhecimento.EnderecoEmpresa
            pfSaldo.CodigoConveniado = objTomador.Codigo
            pfSaldo.EnderecoConveniado = objTomador.CodigoEndereco
            pfSaldo.CodigoFatura = [Lib].Negocio.Numerador.PegarNumero(HttpContext.Current.Session("ssNomeServidor"), eTiposNumerador.FaturaDeFrete)
            If objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                Dim vlrEncargo As Decimal = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "TAXA CADASTRO" OrElse s.Codigo.Trim = "TARIFA SEGURO").Sum(Function(s) s.Valor)
                pfSaldo.ValorDaFatura = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor) - vlrEncargo
            Else
                pfSaldo.ValorDaFatura = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor)
            End If
            pfSaldo.Movimento = DateTime.Now

            'GERAR ITENS DA FATURA DE LÍQUIDO A PAGAR
            Dim pfiSaldo As New [Lib].Negocio.FaturaDeFreteXItens(pfSaldo)
            pfiSaldo.IUD = "I"
            pfiSaldo.EmpresaCnpj = objConhecimento.CodigoEmpresa
            pfiSaldo.EmpresaEnd = objConhecimento.EnderecoEmpresa
            pfiSaldo.ClienteCnpj = objTomador.Codigo
            pfiSaldo.ClienteEnd = objTomador.CodigoEndereco
            pfiSaldo.NumeroNota = objConhecimento.Codigo
            pfiSaldo.EntradaSaida = IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
            pfiSaldo.Serie = objConhecimento.Serie
            pfiSaldo.CodigoEncargo = "LIQUIDOAPAGAR"
            pfiSaldo.ValorAdiantamento = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO").Sum(Function(s) s.Valor)
            pfiSaldo.ViaCartaFrete = True
            pfiSaldo.PesoDeChegada = IIf(pfiSaldo.EntradaSaida = "E" AndAlso objConhecimento.NotasXNotas.PesoBrutoRomaneio > 0, objConhecimento.NotasXNotas.PesoBrutoRomaneio, objConhecimento.NotasXNotas.PesoFiscal)
            pfSaldo.Itens.Add(pfiSaldo)

            If Not pfSaldo.JaFaturada Then
                'GERAR TÍTULO DE SALDO EM PROVISÃO / PREVISÃO
                If Not FinanceiroNovo Then
                    Dim objSaldo As New [Lib].Negocio.Titulo
                    objSaldo.IUD = "I"

                    If objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                        objSaldo.ReceberPagar = "R"
                        'objSaldo.CodigoCarteira = "001002007" 'TRANSPORTADORES / CARRETEIROS
                        objSaldo.CodigoCarteira = objConhecimento.Itens(0).Produto.CodigoCarteiraVenda
                    ElseIf objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro Then
                        objSaldo.ReceberPagar = "P"
                        'objSaldo.CodigoCarteira = "001002007" 'TRANSPORTADORES / CARRETEIROS
                        objSaldo.CodigoCarteira = objConhecimento.Itens(0).Produto.CodigoCarteiraCompra
                    ElseIf pfiSaldo.EntradaSaida = "S" Then
                        objSaldo.ReceberPagar = "R"
                        objSaldo.CodigoCarteira = objConhecimento.Itens(0).Produto.CodigoCarteiraVenda
                    Else
                        objSaldo.ReceberPagar = "P"
                        'objSaldo.CodigoCarteira = "001002007" 'TRANSPORTADORES / CARRETEIROS
                        objSaldo.CodigoCarteira = objConhecimento.Itens(0).Produto.CodigoCarteiraCompra
                    End If

                    objSaldo.CodigoMoeda = 1

                    If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.RPA Then
                        objSaldo.CodigoProvisao = CInt(eProvisao.Previsao)
                        objSaldo.Vencimento = CDate(txtVencimentoFatura.Text).ToString("yyyy-MM-dd")
                    Else
                        objSaldo.CodigoProvisao = CInt(eProvisao.Provisao)
                        objSaldo.Vencimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    End If

                    objSaldo.CodigoTipoPgto = 1
                    objSaldo.CodigoSituacao = 1

                    objSaldo.Movimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    objSaldo.Prorrogacao = objSaldo.Vencimento
                    objSaldo.DataMoeda = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    objSaldo.Baixa = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    objSaldo.CodigoUnidadeDeNegocio = objConhecimento.Empresa.CodigoUnidadeDeNegocio
                    objSaldo.CodigoEmpresa = pfSaldo.CodigoEmpresa
                    objSaldo.EnderecoEmpresa = pfSaldo.EnderecoEmpresa
                    objSaldo.CodigoEmpresaPagadora = pfSaldo.CodigoEmpresa
                    objSaldo.EndEmpresaPagadora = pfSaldo.EnderecoEmpresa

                    objSaldo.CodigoIndexador = 3
                    objSaldo.ContaContabilCliente = IIf(objEncargo Is Nothing OrElse String.IsNullOrWhiteSpace(objEncargo.ContaCredito), "2010103", objEncargo.ContaCredito)
                    objSaldo.CodigoCliente = pfSaldo.CodigoConveniado
                    objSaldo.EndCliente = pfSaldo.EnderecoConveniado

                    If objSaldo.ReceberPagar = "R" Then
                        objSaldo.CodigoCliente = objTomador.Codigo
                        objSaldo.EndCliente = objTomador.CodigoEndereco
                    End If

                    If Not objConhecimento.Tomador Is Nothing AndAlso objConhecimento.Tomador.ContasBancarias IsNot Nothing AndAlso objConhecimento.Tomador.ContasBancarias.Count > 0 Then
                        objSaldo.CodigoBancoCliente = objTomador.ContasBancarias(0).CodigoBanco
                        objSaldo.CodigoAgenciaCliente = objTomador.ContasBancarias(0).CodigoAgencia
                        objSaldo.DigitoAgenciaCliente = objTomador.ContasBancarias(0).DigitoAgencia
                        objSaldo.ContaCliente = objTomador.ContasBancarias(0).ContaCorrente
                        objSaldo.DigitoContaCliente = objTomador.ContasBancarias(0).DigitoConta
                    Else
                        objSaldo.CodigoBancoCliente = 0
                        objSaldo.CodigoAgenciaCliente = ""
                        objSaldo.DigitoAgenciaCliente = ""
                        objSaldo.ContaCliente = ""
                        objSaldo.DigitoContaCliente = ""
                    End If

                    objSaldo.CodigoDestinatario = pfSaldo.CodigoConveniado
                    objSaldo.EndDestinatario = pfSaldo.EnderecoConveniado
                    objSaldo.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                    objSaldo.UsuarioInclusaoData = DateTime.Now

                    If objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                        Dim vlrEncargo As Decimal = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO" OrElse s.Codigo.Trim = "TAXA CADASTRO" OrElse s.Codigo.Trim = "TARIFA SEGURO").Sum(Function(s) s.Valor)
                        objSaldo.ValorDoDocumento = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor) - vlrEncargo
                    Else
                        objSaldo.ValorDoDocumento = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor)
                    End If

                    objSaldo.Historico = "FATURA FRETE NÚMERO " & pfSaldo.CodigoFatura & " - SALDO CONFORME CONHECIMENTO NÚMERO " & objConhecimento.Codigo

                    If hasAdto Then
                        objSaldo.Codigo = numerador + 1
                        objSaldo.SalvarSql(Sqls, False)
                    Else
                        Dim n As New [Lib].Negocio.Numerador(1)
                        objSaldo.SalvarSql(Sqls, True)
                    End If
                    pfSaldo.ListTituloFatura.Add(New FaturaDeFretexTitulo(pfSaldo) With {.CodigoTitulo = objSaldo.Codigo})

                Else
                    Dim objSaldo As New [Lib].Negocio.Novo.TituloNovo
                    objSaldo.IUD = "I"

                    If objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                        objSaldo.ReceberPagar = "R"
                    ElseIf objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro Then
                        objSaldo.ReceberPagar = "P"
                    Else
                        objSaldo.ReceberPagar = "P"
                    End If

                    objSaldo.Movimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    objSaldo.Vencimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    objSaldo.Reprogramacao = objSaldo.Vencimento
                    objSaldo.DataMoeda = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    objSaldo.DataBaixa = CDate(DateTime.Now).ToString("yyyy-MM-dd")
                    objSaldo.CodigoIndexador = 3
                    objSaldo.CodigoMoeda = 1
                    objSaldo.CodigoProvisao = CInt(eProvisao.Previsao)
                    objSaldo.CodigoTipoPgto = 1
                    objSaldo.CodigoSituacao = 1

                    objSaldo.CodigoUnidadeDeNegocio = objConhecimento.Empresa.Empresa.CodigoUnidadeNegocio
                    objSaldo.CodigoEmpresa = pfSaldo.CodigoEmpresa
                    objSaldo.EnderecoEmpresa = pfSaldo.EnderecoEmpresa
                    objSaldo.CodigoUnidadeDeNegocioRecPag = objConhecimento.Empresa.Empresa.CodigoUnidadeNegocio
                    objSaldo.CodigoEmpresaRecPag = pfSaldo.CodigoEmpresa
                    objSaldo.EndEmpresaRecPag = pfSaldo.EnderecoEmpresa

                    objSaldo.CodigoCliFor = pfSaldo.CodigoConveniado
                    objSaldo.EnderecoCliFor = pfSaldo.EnderecoConveniado
                    objSaldo.DesligarControles()
                    'CONTA - VALOR DO DOCUMENTO
                    objSaldo.CodigoContaContabilCliFor = objConhecimento.SubOperacao.CodigoGrupoContas
                    'CONTA - VALOR LÍQUIDO
                    objSaldo.CodigoContaContabilRecPag = objSaldo.Empresa.Empresa.CodigoContaGrupoBanco

                    objSaldo.Valores.EncargoValorDocumento.ValorOficial = pfSaldo.ValorDaFatura
                    objSaldo.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(pfSaldo.ValorDaFatura, 3, objSaldo.Movimento)
                    objSaldo.Valores.AtualizaLiquido()

                    If objConhecimento.Tomador.ContasBancarias IsNot Nothing AndAlso objConhecimento.Tomador.ContasBancarias.Count > 0 Then
                        objSaldo.CodigoBancoCliFor = objTomador.ContasBancarias(0).CodigoBanco
                        objSaldo.CodigoAgenciaCliFor = objTomador.ContasBancarias(0).CodigoAgencia
                        objSaldo.DigitoAgenciaCliFor = objTomador.ContasBancarias(0).DigitoAgencia
                        objSaldo.ContaCliFor = objTomador.ContasBancarias(0).ContaCorrente
                        objSaldo.DigitoContaCliFor = objTomador.ContasBancarias(0).DigitoConta
                    End If

                    objSaldo.TituloDestinacao.CodigoDestinatario = pfSaldo.CodigoConveniado
                    objSaldo.TituloDestinacao.EndDestinatario = pfSaldo.EnderecoConveniado

                    If objConhecimento.TipoDeDocumentoFrete IsNot Nothing AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao Then
                        Dim vlrEncargo As Decimal = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "ADIANTAMENTO").Sum(Function(s) s.Valor)
                        objSaldo.Valores.EncargoValorDocumento.ValorOficial = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor) - vlrEncargo
                        objSaldo.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(objSaldo.Valores.EncargoValorDocumento.ValorOficial, 3, objSaldo.Movimento)
                        objSaldo.Valores.AtualizaLiquido()
                    Else
                        objSaldo.Valores.EncargoValorDocumento.ValorOficial = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor)
                        objSaldo.Valores.EncargoValorDocumento.ValorMoeda = Funcoes.ConverteParaMoedaExtrangeira(objSaldo.Valores.EncargoValorDocumento.ValorOficial, 3, objSaldo.Movimento)
                        objSaldo.Valores.AtualizaLiquido()
                    End If

                    objSaldo.Historico = "FATURA FRETE NÚMERO " & pfSaldo.CodigoFatura & " - SALDO CONFORME CONHECIMENTO NÚMERO " & objConhecimento.Codigo

                    If hasAdto Then
                        objSaldo.Codigo = numerador + 1
                        objSaldo.SalvarSql(Sqls, False)
                    Else
                        Dim n As New [Lib].Negocio.Numerador(1)
                        objSaldo.SalvarSql(Sqls, True)
                    End If
                    pfSaldo.ListTituloFatura.Add(New FaturaDeFretexTitulo(pfSaldo) With {.CodigoTitulo = objSaldo.Codigo})
                End If

                pfSaldo.SalvarSql(Sqls)
                MensagemFinal &= "Título " & pfSaldo.ListTituloFatura(0).CodigoTitulo
            Else
                MsgBox(Me.Page, "Pré-fatura já cadastrada!")
            End If
        End If
    End Sub

    Private Sub SalvarAnulacao(ByRef Sqls As ArrayList)
        RecuperarSessaoConhecimento()
        Dim obj As New [Lib].Negocio.Titulo
        obj.IUD = "I"
        obj.ReceberPagar = "R"
        obj.CodigoProvisao = CInt(eProvisao.Previsao) 'PREVISÃO
        obj.CodigoMoeda = 1
        obj.CodigoTipoPgto = 1
        obj.CodigoSituacao = 1
        'obj.CodigoCarteira = "001002007" 'TRANSPORTADORES / CARRETEIROS
        obj.CodigoCarteira = objConhecimento.Itens(0).Produto.CodigoCarteiraVenda
        obj.Movimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
        obj.Vencimento = CDate(DateTime.Now).ToString("yyyy-MM-dd")
        obj.Prorrogacao = CDate(DateTime.Now).ToString("yyyy-MM-dd")
        obj.DataMoeda = CDate(DateTime.Now).ToString("yyyy-MM-dd")
        obj.CodigoUnidadeDeNegocio = objConhecimento.Empresa.Empresa.CodigoUnidadeNegocio
        obj.CodigoEmpresa = objConhecimento.CodigoEmpresa
        obj.EnderecoEmpresa = objConhecimento.EnderecoEmpresa
        obj.CodigoEmpresaPagadora = objConhecimento.CodigoEmpresa
        obj.EndEmpresaPagadora = objConhecimento.EnderecoEmpresa
        obj.CodigoIndexador = 3

        Dim objEncargo As New [Lib].Negocio.Encargo("LIQUIDOAPAGAR")
        Dim objTomador As New [Lib].Negocio.Cliente(objConhecimento.CodigoTomador, objConhecimento.EnderecoTomador)

        obj.ContaContabilCliente = IIf(objEncargo Is Nothing OrElse String.IsNullOrWhiteSpace(objEncargo.ContaCredito), "2010103", objEncargo.ContaCredito)
        obj.CodigoCliente = objConhecimento.CodigoTomador
        obj.EndCliente = objConhecimento.EnderecoTomador
        obj.ValorDoDocumento = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor)
        obj.MoedaValorDoDocumento = Funcoes.ConverteParaMoedaExtrangeira(obj.ValorDoDocumento, 3, obj.Movimento)

        If objConhecimento.Tomador.ContasBancarias IsNot Nothing AndAlso objConhecimento.Tomador.ContasBancarias.Count > 0 Then
            obj.CodigoBancoCliente = objTomador.ContasBancarias(0).CodigoBanco
            obj.CodigoAgenciaCliente = objTomador.ContasBancarias(0).CodigoAgencia
            obj.DigitoAgenciaCliente = objTomador.ContasBancarias(0).DigitoAgencia
            obj.ContaCliente = objTomador.ContasBancarias(0).ContaCorrente
            obj.DigitoContaCliente = objTomador.ContasBancarias(0).DigitoConta
        End If

        obj.CodigoDestinatario = objConhecimento.CodigoTomador
        obj.EndDestinatario = objConhecimento.EnderecoTomador
        obj.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
        obj.UsuarioInclusaoData = DateTime.Now
        obj.ValorDoDocumento = objConhecimento.Itens.SelectMany(Function(s) s.Encargos).Where(Function(s) s.Codigo.Trim = "LIQUIDO").Sum(Function(s) s.Valor)
        obj.Historico = "PGTO ANULAÇÃO CONHECIMENTO DE FRETE NÚMERO: " & objConhecimento.Codigo & " - " & objConhecimento.Serie
        obj.SalvarSql(Sqls, True)
    End Sub

    Private Function Validar() As Boolean
        RecuperarSessaoConhecimento()
        Dim ws As New [Lib].Negocio.Pamcard()
        If String.IsNullOrWhiteSpace(objConhecimento.Transportador.RNTRCTransportador) Then
            MsgBox(Me.Page, "RNTRC do transportador não foi informando, verifique no cadastro de clientes!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.Placa01) Then
            MsgBox(Me.Page, "A placa do CTRC não encontrada!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.CodigoProprietario01) Then
            MsgBox(Me.Page, "Proprietário do CTRC não encontrado!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.RNTRCPlaca01) Then
            MsgBox(Me.Page, "RNTRC da primeira placa do CTRC não encontrado!")
            Return False
        ElseIf String.IsNullOrWhiteSpace(objConhecimento.Motorista.RG) Then
            MsgBox(Me.Page, "RG deve ser informado no cadastro do cliente!")
            Return False
        ElseIf objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso objConhecimento.NotaTrocaOrigem.Itens IsNot Nothing _
                AndAlso objConhecimento.NotaTrocaOrigem.Itens.Count > 0 AndAlso objConhecimento.NotaTrocaOrigem.Itens(0).Produto.DetalhesGenero Is Nothing Then
            MsgBox(Me.Page, "É necessário informar sub gênero no cadastro do produto!")
            Return False
        Else
            Dim ds As DataSet
            ds = ws.ConsultarRNTRC("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario01, objConhecimento.PlacaDetalhes.RNTRCPlaca01)
            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            Else
                If ds.Tables(0).Rows(0).Item("erro") = 0 Then
                    If CDate(ds.Tables(0).Rows(0).Item("validade")) < Today() Then
                        MsgBox(Me.Page, "Validade " & ds.Tables(0).Rows(0).Item("validade") & " do RNTRC da primeira placa na ANTT está menor que a data atual!")
                        Return False
                    End If

                    If Not ds.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() = "ATIVO" Then
                        MsgBox(Me.Page, "Situação do proprietário da primeira placa é " & ds.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                        Return False
                    End If
                Else
                    MsgBox(Me.Page, "Erro primeira placa: Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                    Return False
                End If
            End If

            If objConhecimento.PlacaDetalhes.Placa02.Length > 0 AndAlso Not objConhecimento.PlacaDetalhes.CodigoProprietario01 = objConhecimento.PlacaDetalhes.CodigoProprietario02 Then
                Dim ds1 As DataSet = ws.ConsultarRNTRC("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario02, objConhecimento.PlacaDetalhes.RNTRCPlaca02)
                If ds1 Is Nothing OrElse ds1.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Erro inesperado consultando RNTRC proprietário da segunda placa, entre em contato com o suporte!")
                    Return False
                Else
                    If ds1.Tables(0).Rows(0).Item("erro") = 0 Then
                        If CDate(ds1.Tables(0).Rows(0).Item("validade")) < Today() Then
                            MsgBox(Me.Page, "Validade do RNTRC da segunda placa na ANTT está menor que a data atual!")
                            Return False
                        End If

                        If Not ds1.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() = "ATIVO" Then
                            MsgBox(Me.Page, "Situação do proprietário da segunda placa é " & ds1.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() & " - " & ds1.Tables(0).Rows(0).Item("errodescricao"))
                            Return False
                        End If
                    Else
                        MsgBox(Me.Page, "Erro segunda placa: Código " & ds1.Tables(0).Rows(0).Item("erro") & " - " & ds1.Tables(0).Rows(0).Item("errodescricao"))
                        Return False
                    End If
                End If
            End If

            If objConhecimento.PlacaDetalhes.Placa03.Length > 0 AndAlso Not objConhecimento.PlacaDetalhes.CodigoProprietario01 = objConhecimento.PlacaDetalhes.CodigoProprietario03 Then
                Dim ds2 As DataSet = ws.ConsultarRNTRC("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario03, objConhecimento.PlacaDetalhes.RNTRCPlaca03)
                If ds2 Is Nothing OrElse ds2.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Erro inesperado consultando RNTRC proprietário da terceira placa, entre em contato com o suporte!")
                    Return False
                Else
                    If ds2.Tables(0).Rows(0).Item("erro") = 0 Then
                        If CDate(ds2.Tables(0).Rows(0).Item("validade")) < Today() Then
                            MsgBox(Me.Page, "Validade do RNTRC da terceira placa na ANTT está menor que a data atual!")
                            Return False
                        End If

                        If Not ds2.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() = "ATIVO" Then
                            MsgBox(Me.Page, "Situação do proprietário da terceira placa é " & ds2.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() & " - " & ds2.Tables(0).Rows(0).Item("errodescricao"))
                            Return False
                        End If
                    Else
                        MsgBox(Me.Page, "Erro terceira placa: Código " & ds2.Tables(0).Rows(0).Item("erro") & " - " & ds2.Tables(0).Rows(0).Item("errodescricao"))
                        Return False
                    End If
                End If
            End If

            If objConhecimento.PlacaDetalhes.Placa04.Length > 0 AndAlso Not objConhecimento.PlacaDetalhes.CodigoProprietario01 = objConhecimento.PlacaDetalhes.CodigoProprietario04 Then
                Dim ds3 As DataSet = ws.ConsultarRNTRC("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario04, objConhecimento.PlacaDetalhes.RNTRCPlaca04)
                If ds3 Is Nothing OrElse ds3.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Erro inesperado consultando RNTRC proprietário da quarta placa, entre em contato com o suporte!")
                    Return False
                Else
                    If ds3.Tables(0).Rows(0).Item("erro") = 0 Then
                        If CDate(ds3.Tables(0).Rows(0).Item("validade")) < Today() Then
                            MsgBox(Me.Page, "Validade do RNTRC da quarta placa na ANTT está menor que a data atual!")
                            Return False
                        End If

                        If Not ds3.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() = "ATIVO" Then
                            MsgBox(Me.Page, "Situação do proprietário da quarta placa é " & ds3.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() & ". " & ds3.Tables(0).Rows(0).Item("errodescricao"))
                            Return False
                        End If
                    Else
                        MsgBox(Me.Page, "Erro quarta placa: Código " & ds3.Tables(0).Rows(0).Item("erro") & " - " & ds3.Tables(0).Rows(0).Item("errodescricao"))
                        Return False
                    End If
                End If
            End If

            Dim dsp As DataSet = ws.ConsultarFrota("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario01, objConhecimento.PlacaDetalhes.RNTRCPlaca01, objConhecimento.PlacaDetalhes)
            If dsp Is Nothing OrElse dsp.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Erro inesperado consultando frota, entre em contato com o suporte!")
                Return False
            Else
                If dsp.Tables(0).Rows(0).Item("erro") = 0 Then
                    If Not dsp.Tables(0).Rows(0).Item("situacaoplaca1") = "S" Then
                        MsgBox(Me.Page, "Placa " & objConhecimento.PlacaDetalhes.Placa01 & " não faz parte da frota do proprietário informado!")
                        Return False
                    End If

                    If objConhecimento.PlacaDetalhes.Placa02.Length > 0 AndAlso
                       Not dsp.Tables(0).Rows(0).Item("situacaoplaca2") = "S" Then
                        If objConhecimento.PlacaDetalhes.CodigoProprietario01 = objConhecimento.PlacaDetalhes.CodigoProprietario02 Then
                            MsgBox(Me.Page, "Segunda placa " & objConhecimento.PlacaDetalhes.Placa02 & " não faz parte da frota do proprietário informado!")
                            Return False
                        Else
                            Dim objPlaca As New [Lib].Negocio.Placa()
                            objPlaca.Placa01 = objConhecimento.PlacaDetalhes.Placa02

                            Dim dsp1 As DataSet = ws.ConsultarFrota("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario02, objConhecimento.PlacaDetalhes.RNTRCPlaca02, objPlaca)
                            If dsp1 Is Nothing OrElse dsp1.Tables(0).Rows.Count = 0 Then
                                MsgBox(Me.Page, "Erro inesperado consultando frota do proprietário da segunda placa!")
                                Return False
                            Else
                                If dsp1.Tables(0).Rows(0).Item("erro") = 0 Then
                                    If Not dsp1.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() = "ATIVO" Then
                                        MsgBox(Me.Page, "Situação do proprietário da segunda placa é " & dsp1.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() & " - " & dsp1.Tables(0).Rows(0).Item("errodescricao"))
                                        Return False
                                    End If

                                    If Not dsp1.Tables(0).Rows(0).Item("situacaoplaca1") = "S" Then
                                        MsgBox(Me.Page, "Segunda placa " & objConhecimento.PlacaDetalhes.Placa02 & " não faz parte da frota do proprietário " & objConhecimento.PlacaDetalhes.Proprietario02.Nome & " informado!")
                                        Return False
                                    End If
                                Else
                                    MsgBox(Me.Page, "Erro: Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                                    Return False
                                End If
                            End If
                        End If
                    End If

                    If objConhecimento.PlacaDetalhes.Placa03.Length > 0 AndAlso
                       Not dsp.Tables(0).Rows(0).Item("situacaoplaca3") = "S" Then
                        If objConhecimento.PlacaDetalhes.CodigoProprietario01 = objConhecimento.PlacaDetalhes.CodigoProprietario03 Then
                            MsgBox(Me.Page, "Terceira Placa " & objConhecimento.PlacaDetalhes.Placa03 & " não faz parte da Frota do Proprietário informado, verifique.")
                            Return False
                        Else
                            Dim objPlaca As New [Lib].Negocio.Placa()
                            objPlaca.Placa01 = objConhecimento.PlacaDetalhes.Placa03

                            Dim dsp2 As DataSet = ws.ConsultarFrota("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario03, objConhecimento.PlacaDetalhes.RNTRCPlaca03, objPlaca)
                            If dsp2 Is Nothing OrElse dsp2.Tables(0).Rows.Count = 0 Then
                                MsgBox(Me.Page, "Erro inesperado consultando frota do proprietário da terceira placa!")
                                Return False
                            Else
                                If dsp2.Tables(0).Rows(0).Item("erro") = 0 Then
                                    If Not dsp2.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() = "ATIVO" Then
                                        MsgBox(Me.Page, "Situação do proprietário da terceira placa é " & dsp2.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() & " - " & dsp2.Tables(0).Rows(0).Item("errodescricao"))
                                        Return False
                                    End If

                                    If Not dsp2.Tables(0).Rows(0).Item("situacaoplaca1") = "S" Then
                                        MsgBox(Me.Page, "Terceira placa " & objConhecimento.PlacaDetalhes.Placa03 & " não faz parte da frota do proprietário " & objConhecimento.PlacaDetalhes.Proprietario03.Nome & " informado!")
                                        Return False
                                    End If
                                Else
                                    MsgBox(Me.Page, "Erro: Código " & ds.Tables(0).Rows(0).Item("erro") & " - " & ds.Tables(0).Rows(0).Item("errodescricao"))
                                    Return False
                                End If
                            End If
                        End If
                    End If

                    If objConhecimento.PlacaDetalhes.Placa04.Length > 0 AndAlso
                       Not ds.Tables(0).Rows(0).Item("situacaoplaca4") = "S" Then
                        If objConhecimento.PlacaDetalhes.CodigoProprietario01 = objConhecimento.PlacaDetalhes.CodigoProprietario04 Then
                            MsgBox(Me.Page, "Quarta placa " & objConhecimento.PlacaDetalhes.Placa04 & " não faz parte da frota do proprietário informado!")
                            Return False
                        Else
                            Dim objPlaca As New [Lib].Negocio.Placa()
                            objPlaca.Placa01 = objConhecimento.PlacaDetalhes.Placa04

                            Dim dsp3 As DataSet = ws.ConsultarFrota("04854422000185", objConhecimento.PlacaDetalhes.CodigoProprietario04, objConhecimento.PlacaDetalhes.RNTRCPlaca04, objPlaca)
                            If dsp3 Is Nothing OrElse dsp3.Tables(0).Rows.Count = 0 Then
                                MsgBox(Me.Page, "Erro inesperado consultando frota do proprietário da quarta placa!")
                                Return False
                            Else
                                If dsp3.Tables(0).Rows(0).Item("erro") = 0 Then
                                    If Not dsp3.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() = "ATIVO" Then
                                        MsgBox(Me.Page, "Situação do proprietário da quarta placa é " & dsp3.Tables(0).Rows(0).Item("situacao").ToString.ToUpper() & " - " & dsp3.Tables(0).Rows(0).Item("errodescricao"))
                                        Return False
                                    End If

                                    If Not dsp3.Tables(0).Rows(0).Item("situacaoplaca1") = "S" Then
                                        MsgBox(Me.Page, "Quarta placa " & objConhecimento.PlacaDetalhes.Placa04 & " não faz parte da frota do proprietário " & objConhecimento.PlacaDetalhes.Proprietario04.Nome & " informado!")
                                        Return False
                                    End If
                                Else
                                    MsgBox(Me.Page, "Erro: Código " & dsp3.Tables(0).Rows(0).Item("erro") & " - " & dsp3.Tables(0).Rows(0).Item("errodescricao"))
                                    Return False
                                End If
                            End If
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Erro: Código " & dsp.Tables(0).Rows(0).Item("erro") & " - " & dsp.Tables(0).Rows(0).Item("errodescricao"))
                    Return False
                End If
            End If
        End If
        Return True
    End Function

    Private Property PesoTotalRPA() As Decimal
        Get
            Return CDec(Session("PesoTotalRPA" & HID.Value))
        End Get
        Set(ByVal value As Decimal)
            Session("PesoTotalRPA" & HID.Value) = CDec(value)
        End Set
    End Property

    Private Property MensagemFinal As String
        Get
            Return Session("MensagemFinal" & HID.Value)
        End Get
        Set(ByVal value As String)
            Session("MensagemFinal" & HID.Value) = value
        End Set
    End Property

    Private Sub SalvarSessaoConhecimento()
        Session("ssConhecimentoDeTransporte" & HID.Value) = objConhecimento
    End Sub

    Private Sub RecuperarSessaoConhecimento()
        objConhecimento = CType(Session("ssConhecimentoDeTransporte" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CargaSituacao()
        ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "", True)
    End Sub

    Public Property SessaoDsXML() As DataSet
        Get
            Return CType(Session("dsXML" & HID.Value), DataSet)
        End Get
        Set(ByVal value As DataSet)
            If value Is Nothing Then
                Session.Remove("dsXml" & HID.Value)
            Else
                Session("dsXml" & HID.Value) = value
            End If
        End Set
    End Property

    Public Sub CarregarNotaFiscal()
        If Session("NfConsultaCTRC" & HID.Value) IsNot Nothing Then
            objConhecimento = New [Lib].Negocio.NotaFiscal(CType(Session("NfConsultaCTRC" & HID.Value), [Lib].Negocio.NotaFiscal))
            objConhecimento.IUD = "U"
            txtNumero.Text = objConhecimento.Codigo
            txtSerie.Text = objConhecimento.Serie
            txtChaveNFe.Text = objConhecimento.ChaveNFE
            If objConhecimento.EntradaSaida = eEntradaSaida.Entrada Then
                txtQuantidade.Text = objConhecimento.Itens(0).QuantidadeFiscal.ToString("N0")
            Else
                txtQuantidade.Text = objConhecimento.Itens(0).QuantidadeFisica.ToString("N0")
            End If
            txtQuantidade.Enabled = False
            txtUnitario.Text = objConhecimento.Itens(0).Unitario.ToString("N4")
            txtUnitario.Enabled = False
            txtValor.Text = objConhecimento.Itens(0).ValorTotal.ToString("N2")
            SalvarSessaoConhecimento()
            CarregarCTE(False, True)
            txtChaveNFe.Enabled = False
            txtNumero.Enabled = False
            txtSerie.Enabled = False
            rdoBanco.Enabled = False
            rdoPamcard.Enabled = False
            If objConhecimento.NossaEmissao Then
                txtDataDeEmissao.Enabled = False
                txtDataDeSaida.Enabled = False
                HideCalendar(Me.Page, txtDataDeEmissao)
                HideCalendar(Me.Page, txtDataDeSaida)
            End If
            btnRelatorio.Enabled = True
            btnCadastro.Enabled = False
            lnkExcluir.Parent.Visible = Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) AndAlso CInt(ddlSituacao.SelectedValue) = CInt(eSituacao.Normal) AndAlso objConhecimento.CodigoOperacao <> 81 'CONTRAPARTIDA
            btnRecontabilizar.Enabled = Not String.IsNullOrWhiteSpace(ddlSituacao.SelectedValue) AndAlso CInt(ddlSituacao.SelectedValue) = CInt(eSituacao.Normal)
            'lnkNovo.Parent.Visible = Not objConhecimento.NossaEmissao
            TcDados.ActiveTabIndex = 0
            Session.Remove("NfConsultaCTRC" & HID.Value)
            SalvarSessaoConhecimento()
        End If
    End Sub

    Protected Sub PreencherNFeXML(ByRef objNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal pNomeArquivo As String, ByVal pOrigem As Boolean)

        Dim DsXml As New DataSet

        DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}", pNomeArquivo)))
        SessaoDsXML = DsXml

        objNotaFiscal.XMLImportado = True
        objNotaFiscal.NossaEmissao = False
        objNotaFiscal.NFG = True

        Dim xmlDoc As New XmlDocument
        xmlDoc.Load(Server.MapPath(String.Format("~/Files/{0}", pNomeArquivo)))

        'Chave NFe ou CTe
        objNotaFiscal.ChaveNFE = xmlDoc.GetElementsByTagName("protCTe").GetNode("protCTe").ChildNodes.GetNode("infProt").ChildNodes.GetNode("chCTe").InnerText

        Dim Empresa As New [Lib].Negocio.ListCliente
        Dim listaTomador As New [Lib].Negocio.ListCliente
        Dim empresaEmissor As String
        Dim inscricaoEmissor As String

        Dim codigoEmpresa As String = ""
        Dim inscricaoEmpresa As String = ""

        Dim codigoFornecedor As String = ""
        Dim inscricaoFornecedor As String = ""

        Dim iTomador As Integer = -1
        Dim sCodigoTomador As String = ""
        Dim iEnderecoTomador As Integer = 0
        Dim inscricaoTomador As String = ""


        Dim sCodigoRemetente As String = ""
        Dim iEnderecoRemetente As Integer = 0
        Dim inscricaoRemetente As String = ""

        Dim sCodigoDestinatario As String = ""
        Dim iEnderecoDestinatario As Integer = 0
        Dim inscricaoDestinatario As String = ""

        Dim tomaValue As String = ""
        ' Verifica a existência de qualquer tag <toma>
        Dim tomaNodes = xmlDoc.GetElementsByTagName("toma")

        If tomaNodes.Count > 0 Then
            iTomador = tomaNodes(0).InnerText
        End If

        'Regra para o tomador
        '********************************************************************************
        '0 - Remetente
        '1 - Expedidor
        '2 - Recebedor
        '3 - Destinatario
        '4 - Cliente é o Tomador
        '********************************************************************************

        objNotaFiscal.EmpresaEmissor = False

        empresaEmissor = xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("CNPJ").InnerText
        inscricaoEmissor = xmlDoc.GetElementsByTagName("emit").GetNode("emit").ChildNodes.GetNode("IE").InnerText

        If Left(Session("ssEmpresa"), 8) = Left(empresaEmissor, 8) Then
            codigoEmpresa = empresaEmissor
            objNotaFiscal.EmpresaEmissor = True
        End If

        If objNotaFiscal.EmpresaEmissor = True Then
            ddlTipoConhecimento.SelectedValue = eTipoDeDocumentoFrete.Circulacao
        Else
            ddlTipoConhecimento.SelectedValue = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro
        End If

        Dim emitNode As XmlNode
        Dim cnpjNode As XmlNode
        Dim cpfNode As XmlNode


        Dim destNodesDest As XmlNodeList = xmlDoc.GetElementsByTagName("dest")
        If destNodesDest.Count > 0 Then
            Dim destNode As XmlNode = destNodesDest(0)
            For Each childNode As XmlNode In destNode.ChildNodes
                If childNode.Name = "CNPJ" OrElse childNode.Name = "CPF" Then
                    sCodigoDestinatario = childNode.InnerText
                ElseIf childNode.Name = "IE" Then
                    inscricaoDestinatario = childNode.InnerText
                End If
            Next
        End If

        Dim destNodesRem As XmlNodeList = xmlDoc.GetElementsByTagName("rem")
        If destNodesRem.Count > 0 Then
            Dim destNode As XmlNode = destNodesRem(0)
            For Each childNode As XmlNode In destNode.ChildNodes
                If childNode.Name = "CNPJ" OrElse childNode.Name = "CPF" Then
                    sCodigoRemetente = childNode.InnerText
                ElseIf childNode.Name = "IE" Then
                    inscricaoRemetente = childNode.InnerText
                End If
            Next
        End If

        If iTomador = 0 Then
            Dim destNodes As XmlNodeList = xmlDoc.GetElementsByTagName("rem")
            If destNodes.Count > 0 Then
                Dim destNode As XmlNode = destNodes(0)
                For Each childNode As XmlNode In destNode.ChildNodes
                    If childNode.Name = "CNPJ" OrElse childNode.Name = "CPF" Then
                        sCodigoTomador = childNode.InnerText
                    ElseIf childNode.Name = "IE" Then
                        inscricaoTomador = childNode.InnerText
                    End If
                Next
            End If
        ElseIf iTomador = 1 Then
            Dim destNodes As XmlNodeList = xmlDoc.GetElementsByTagName("exped")
            If destNodes.Count > 0 Then
                Dim destNode As XmlNode = destNodes(0)
                For Each childNode As XmlNode In destNode.ChildNodes
                    If childNode.Name = "CNPJ" OrElse childNode.Name = "CPF" Then
                        sCodigoTomador = childNode.InnerText
                    ElseIf childNode.Name = "IE" Then
                        inscricaoTomador = childNode.InnerText
                    End If
                Next
            End If
        ElseIf iTomador = 2 Then
            Dim destNodes As XmlNodeList = xmlDoc.GetElementsByTagName("receb")
            If destNodes.Count > 0 Then
                Dim destNode As XmlNode = destNodes(0)
                For Each childNode As XmlNode In destNode.ChildNodes
                    If childNode.Name = "CNPJ" OrElse childNode.Name = "CPF" Then
                        sCodigoTomador = childNode.InnerText
                    ElseIf childNode.Name = "IE" Then
                        inscricaoTomador = childNode.InnerText
                    End If
                Next
            End If
        ElseIf iTomador = 3 Then
            Dim destNodes As XmlNodeList = xmlDoc.GetElementsByTagName("dest")
            If destNodes.Count > 0 Then
                Dim destNode As XmlNode = destNodes(0)
                For Each childNode As XmlNode In destNode.ChildNodes
                    If childNode.Name = "CNPJ" OrElse childNode.Name = "CPF" Then
                        sCodigoTomador = childNode.InnerText
                    ElseIf childNode.Name = "IE" Then
                        inscricaoTomador = childNode.InnerText
                    End If
                Next
            End If
        ElseIf iTomador = 4 Then
            Dim destNodes As XmlNodeList = xmlDoc.GetElementsByTagName("toma4")
            If destNodes.Count > 0 Then
                Dim destNode As XmlNode = destNodes(0)
                For Each childNode As XmlNode In destNode.ChildNodes
                    If childNode.Name = "CNPJ" OrElse childNode.Name = "CPF" Then
                        sCodigoTomador = childNode.InnerText
                    ElseIf childNode.Name = "IE" Then
                        inscricaoTomador = childNode.InnerText
                    End If
                Next
            End If
        End If

        objNotaFiscal.CodigoTomador = ""
        objNotaFiscal.EnderecoTomador = 0

        If iTomador >= 0 Then

            listaTomador = listaTomador.FindCliente(sCodigoTomador)

            If listaTomador.Count > 0 Then
                For Each clienteTomador As Cliente In listaTomador
                    If clienteTomador.InscricaoEstadual = inscricaoTomador Then
                        objNotaFiscal.CodigoTomador = clienteTomador.Codigo
                        objNotaFiscal.EnderecoTomador = clienteTomador.CodigoEndereco
                        Exit For
                    End If
                Next
            Else
                objNotaFiscal.CodigoTomador = sCodigoTomador
            End If

        End If

        If objNotaFiscal.EmpresaEmissor Then
            codigoEmpresa = empresaEmissor
            inscricaoEmpresa = inscricaoEmissor

            'If codigoEmpresa = objNotaFiscal.CodigoTomador Then
            '    codigoFornecedor = sCodigoDestinatario
            '    inscricaoFornecedor = inscricaoDestinatario
            'Else
            '    codigoFornecedor = objNotaFiscal.CodigoTomador
            '    inscricaoFornecedor = objNotaFiscal.EnderecoTomador
            'End If

            If objNotaFiscal.CodigoTomador.Length = 0 Then
                MsgBox(Me.Page, "Verifique a Inscrição Estadual do Tomador " & sCodigoTomador & " no Cadastro de Clientes.", eTitulo.Info, False)
                Exit Sub
            End If

            codigoFornecedor = objNotaFiscal.CodigoTomador
            inscricaoFornecedor = objNotaFiscal.Tomador.InscricaoEstadual

        Else
            codigoEmpresa = objNotaFiscal.CodigoTomador
            inscricaoEmpresa = objNotaFiscal.EnderecoTomador

            'If codigoEmpresa = objNotaFiscal.CodigoTomador Then
            '    codigoFornecedor = sCodigoDestinatario
            '    inscricaoFornecedor = inscricaoDestinatario
            'Else
            codigoFornecedor = empresaEmissor
            inscricaoFornecedor = inscricaoEmissor
            'End If

        End If

        Empresa = Empresa.FindCliente(codigoEmpresa)

        If Empresa.Count = 1 Then
            objNotaFiscal.CodigoEmpresa = Empresa(0).Codigo
            objNotaFiscal.EnderecoEmpresa = Empresa(0).CodigoEndereco
            objNotaFiscal.CodigoUnidadeDeNegocio = Empresa(0).CodigoUnidadeDeNegocio
            objNotaFiscal.Empresa = Empresa(0)
        ElseIf Empresa.Count > 1 Then
            For Each clienteEmpresa As Cliente In Empresa
                If clienteEmpresa.InscricaoEstadual = inscricaoEmpresa Then
                    objNotaFiscal.CodigoEmpresa = clienteEmpresa.Codigo
                    objNotaFiscal.EnderecoEmpresa = clienteEmpresa.CodigoEndereco
                    objNotaFiscal.CodigoUnidadeDeNegocio = clienteEmpresa.CodigoUnidadeDeNegocio
                    objNotaFiscal.Empresa = clienteEmpresa
                End If
            Next
        End If

        Dim Fornecedor As New [Lib].Negocio.ListCliente
        Fornecedor = Fornecedor.FindCliente(codigoFornecedor)

        If Fornecedor.Count = 1 Then
            objNotaFiscal.CodigoCliente = Fornecedor(0).Codigo
            objNotaFiscal.EnderecoCliente = Fornecedor(0).CodigoEndereco
        ElseIf Fornecedor.Count > 1 Then
            For Each clienteFornecedor As Cliente In Fornecedor
                If clienteFornecedor.InscricaoEstadual = inscricaoFornecedor Then
                    objNotaFiscal.CodigoCliente = clienteFornecedor.Codigo
                    objNotaFiscal.EnderecoCliente = clienteFornecedor.CodigoEndereco
                    objNotaFiscal.CodigoUnidadeDeNegocio = clienteFornecedor.CodigoUnidadeDeNegocio
                    objNotaFiscal.Cliente = clienteFornecedor
                End If
            Next
        ElseIf Fornecedor.Count = 0 Then

            IncluirDadosNovoCliente(objNotaFiscal, DsXml, Fornecedor, codigoFornecedor)

            Fornecedor = Fornecedor.FindCliente(codigoFornecedor)

            If Fornecedor.Count = 0 Then
                MsgBox(Me.Page, "Cliente não encontrado! CNPJ: " & codigoFornecedor, eTitulo.Info, False)
                txtNomeCliente.Text = String.Empty
                Exit Sub
            Else
                For Each clienteFornecedor As Cliente In Fornecedor
                    If clienteFornecedor.InscricaoEstadual = inscricaoFornecedor Then
                        objNotaFiscal.CodigoCliente = clienteFornecedor.Codigo
                        objNotaFiscal.EnderecoCliente = clienteFornecedor.CodigoEndereco
                        objNotaFiscal.CodigoUnidadeDeNegocio = clienteFornecedor.CodigoUnidadeDeNegocio
                        objNotaFiscal.Cliente = clienteFornecedor
                    End If
                Next
            End If
        Else
            ucConsultaClientes.Limpar()
            ucConsultaClientes.CarregarCNPJ(codigoFornecedor)
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
        End If

        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
            'Codigo NFe
            objNotaFiscal.Codigo = xmlDoc.GetElementsByTagName("nNF").GetNode("nNF").InnerText
            objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.Nota
        Else

            'Codigo CTe
            objNotaFiscal.Codigo = xmlDoc.GetElementsByTagName("nCT").GetNode("nCT").InnerText

            Dim chaveNFEOrigem As String = ""

            ' --- Coleta e processa TODAS as notas referenciadas (NF-e e CT-e complemento) ---
            Dim chavesReferenciadas As New List(Of String)
            Dim houveComplementoCTe As Boolean = False

            If xmlDoc.GetElementsByTagName("infNFe").Count > 0 OrElse
                xmlDoc.GetElementsByTagName("infCteComp").Count > 0 OrElse
                xmlDoc.GetElementsByTagName("infCteSub").Count > 0 OrElse
                xmlDoc.GetElementsByTagName("idDocAntEle").Count > 0 Then

                ' --- Coleta TODAS as chaves de NF-e referenciadas ---
                For Each infNFe As XmlElement In xmlDoc.GetElementsByTagName("infNFe")
                    Dim chaveNodes = infNFe.GetElementsByTagName("chave")
                    If chaveNodes IsNot Nothing AndAlso chaveNodes.Count > 0 Then
                        Dim chave As String = (chaveNodes(0).InnerText & "").Trim()
                        If IsChaveValida(chave) Then chavesReferenciadas.Add(chave)
                    End If
                Next

                ' --- Coleta TODAS as chaves de CT-e complementar ---
                For Each infCteComp As XmlElement In xmlDoc.GetElementsByTagName("infCteComp")
                    Dim chCTeNodes = infCteComp.GetElementsByTagName("chCTe")
                    If chCTeNodes IsNot Nothing AndAlso chCTeNodes.Count > 0 Then
                        Dim chaveCTe As String = (chCTeNodes(0).InnerText & "").Trim()
                        If IsChaveValida(chaveCTe) Then
                            chavesReferenciadas.Add(chaveCTe)
                            houveComplementoCTe = True
                        End If
                    End If
                Next

                ' --- Coleta TODAS as chaves de CT-e Substituto ---
                For Each infCteSub As XmlElement In xmlDoc.GetElementsByTagName("infCteSub")
                    Dim chCTeNodes = infCteSub.GetElementsByTagName("chCte")
                    If chCTeNodes IsNot Nothing AndAlso chCTeNodes.Count > 0 Then
                        Dim chaveCTe As String = (chCTeNodes(0).InnerText & "").Trim()
                        If IsChaveValida(chaveCTe) Then
                            chavesReferenciadas.Add(chaveCTe)
                        End If
                    End If
                Next

                ' --- Coleta TODAS as chaves de Documento anterior ---
                For Each idDocAntEle As XmlElement In xmlDoc.GetElementsByTagName("idDocAntEle")
                    Dim chCTeNodes = idDocAntEle.GetElementsByTagName("chCTe")
                    If chCTeNodes IsNot Nothing AndAlso chCTeNodes.Count > 0 Then
                        Dim chaveCTe As String = (chCTeNodes(0).InnerText & "").Trim()
                        If IsChaveValida(chaveCTe) Then
                            chavesReferenciadas.Add(chaveCTe)
                        End If
                    End If
                Next

                ' --- Remove duplicadas (sem LINQ) ---
                Dim unicas As New HashSet(Of String)(StringComparer.Ordinal)
                For Each c In chavesReferenciadas
                    If Not unicas.Contains(c) Then unicas.Add(c)
                Next

                If unicas.Count > 0 Then
                    If houveComplementoCTe Then
                        ddlTipoConhecimento.SelectedValue = eTipoDeDocumentoFrete.Complemento
                    End If

                    objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E
                    objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E

                    If objNotaFiscal.NotasTrocaOrigem Is Nothing Then
                        objNotaFiscal.NotasTrocaOrigem = New List(Of [Lib].Negocio.NotaFiscal)
                    End If

                    Dim adicionouAlguma As Boolean = False

                    For Each chave In unicas
                        Dim nf As New [Lib].Negocio.NotaFiscal()
                        nf.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                        nf.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa

                        nf.CarregarNotaComChaveXML(chave)

                        If nf.Codigo > 0 Then
                            ' Normaliza/clona conforme sua regra
                            nf = New [Lib].Negocio.NotaFiscal(nf)

                            ' Evita duplicidade na lista de origem (sem LINQ)
                            Dim jaExiste As Boolean = False
                            For Each x In objNotaFiscal.NotasTrocaOrigem
                                If x.Empresa.Codigo = nf.Empresa.Codigo AndAlso
                                   x.Cliente.Codigo = nf.Cliente.Codigo AndAlso
                                   x.Serie = nf.Serie AndAlso
                                   x.Codigo = nf.Codigo Then
                                    jaExiste = True
                                    Exit For
                                End If
                            Next

                            If Not jaExiste Then
                                objNotaFiscal.NotasTrocaOrigem.Add(nf)
                                adicionouAlguma = True
                            End If
                        End If
                    Next

                    ' Define uma principal (se aplicável)
                    If objNotaFiscal.NotasTrocaOrigem.Count > 0 Then
                        objNotaFiscal.NotaTrocaOrigem = objNotaFiscal.NotasTrocaOrigem(0)
                    End If

                    If adicionouAlguma Then
                        CarregarNotasDeOrigem(objNotaFiscal)
                    Else
                        If objNotaFiscal.NotasTrocaOrigem Is Nothing OrElse objNotaFiscal.NotasTrocaOrigem.Count = 0 Then
                            MsgBox(Me.Page, "Notas de origem não encontradas pelas chaves informadas.", eTitulo.Info, False)
                            Exit Sub
                        End If
                    End If

                Else
                    MsgBox(Me.Page, "Nenhuma chave referenciada válida encontrada no XML.", eTitulo.Info, False)
                    Exit Sub
                End If

            End If

            objNotaFiscal.CodigoTipoDeDocumento = eTipoDeDocumento.CT_E
            objNotaFiscal.TipoDeDocumento.Codigo = eTipoDeDocumento.CT_E

        End If

        'Série
        objNotaFiscal.Serie = xmlDoc.GetElementsByTagName("serie").GetNode("serie").InnerText

        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
            'Emissão
            If xmlDoc.GetElementsByTagName("dEmi").GetNode("dEmi") IsNot Nothing Then
                objNotaFiscal.DataNota = xmlDoc.GetElementsByTagName("dEmi").GetNode("dEmi").InnerText
            Else
                Dim strData As String = xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi").InnerText
                objNotaFiscal.DataNota = strData.Remove(strData.Length - 6)
            End If
        Else
            'Emissão
            If xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi") IsNot Nothing Then
                objNotaFiscal.DataNota = xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi").InnerText
            Else
                Dim strData As String = xmlDoc.GetElementsByTagName("dhEmi").GetNode("dhEmi").InnerText
                objNotaFiscal.DataNota = strData.Remove(strData.Length - 6)
            End If
        End If

        'VerificarProdutos se a nota existe

        Dim strSQL As String

        strSQL = "  SELECT 1 " & vbCrLf &
                 "  FROM NotasFiscais NF" & vbCrLf &
                 "  INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
                 "      ON NF.Empresa_Id        = NFxI.Empresa_Id" & vbCrLf &
                 "      AND NF.EndEmpresa_Id    = NFxI.EndEmpresa_Id" & vbCrLf &
                 "      AND NF.Cliente_Id       = NFxI.Cliente_Id" & vbCrLf &
                 "      AND NF.EndCliente_Id    = NFxI.EndCliente_Id" & vbCrLf &
                 "      AND NF.EntradaSaida_Id  = NFxI.EntradaSaida_Id" & vbCrLf &
                 "      AND NF.Serie_Id         = NFxI.Serie_Id" & vbCrLf &
                 "      AND NF.Nota_Id          = NFxI.Nota_Id" & vbCrLf &
                 "  WHERE NF.Empresa_Id         = '" & objNotaFiscal.Empresa.Codigo & "'" & vbCrLf &
                 "      AND NF.Cliente_Id       = '" & objNotaFiscal.Cliente.Codigo & "'" & vbCrLf &
                 "      AND NF.Nota_Id          = " & objNotaFiscal.Codigo & vbCrLf &
                 "      AND NF.Serie_Id         = '" & objNotaFiscal.Serie & "'" & vbCrLf &
                 "      AND NF.EntradaSaida_Id  = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "NotasFiscaisGerais")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count() > 0 Then
            MsgBox(Me.Page, "XML já possui uma nota fiscal cadastrada, não é possivel importar!", eTitulo.Info, False)
            Exit Sub
        End If

        'Situação
        objNotaFiscal.CodigoSituacao = eSituacao.Normal

        'Safra
        objNotaFiscal.CodigoSafra = "NENHUMA"

        'Xml Obsoleto
        'objNotaFiscal.XML = xmlDoc.InnerXml

        'bloqueia os campos
        txtChaveNFe.Enabled = False
        txtNumero.Enabled = False
        txtSerie.Enabled = False
        txtDataDeEmissao.Enabled = False

        'Manifesta a NFe quando a origem for ucFile
        Dim msgResult As String = String.Empty
        If pOrigem Then
            If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
                If Not DocumentoEletronico.ManifestoNFe(objNotaFiscal, eTipoManifesto.ConfirmacaoDaOperacao, msgResult) Then
                    MsgBox(Me.Page, msgResult, eTitulo.Info, False)
                    'LimparCampos()
                    'Exit Sub
                Else
                    MsgBox(Me.Page, msgResult, eTitulo.Info, False)
                End If
            End If
        End If

        'A verificação tem que ser antes de adicionar os produtos
        ' Verifica se a tabela "ICMSTot" existe no DataSet
        If DsXml.Tables.Contains("ICMSTot") Then
            Dim icmsTotTable As DataTable = DsXml.Tables("ICMSTot")

            ' Verifica se há pelo menos uma linha na tabela "ICMSTot"
            If icmsTotTable.Rows.Count > 0 Then
                Dim icmsTotRow As DataRow = icmsTotTable.Rows(0)

                ' Extrai os valores das tags vProd e vNF
                Dim vProd As String = If(icmsTotRow.Table.Columns.Contains("vProd") AndAlso Not IsDBNull(icmsTotRow("vProd")), icmsTotRow("vProd").ToString(), "")
                Dim vNF As String = If(icmsTotRow.Table.Columns.Contains("vNF") AndAlso Not IsDBNull(icmsTotRow("vNF")), icmsTotRow("vNF").ToString(), "")

                ' Exemplo: Atribuir os valores a outras variáveis
                Dim valorProduto As Decimal = Decimal.Parse(vProd.Replace(".", ",")) ' Converte para Decimal
                Dim valorNotaFiscal As Decimal = Decimal.Parse(vNF.Replace(".", ",")) ' Converte para Decimal

                objNotaFiscal.DiferencaValorNFXProdutoXML = valorNotaFiscal - valorProduto

            End If
        End If

        objNotaFiscal.Itens.Clear()

        Dim dValorTotal As Decimal
        Dim QuantidadeNota As Decimal = 0
        Dim unitarioFrete As Decimal = 0

        If Not DsXml.Tables("vPrest") Is Nothing Then
            dValorTotal = DsXml.Tables("vPrest").Rows(0).Item("vRec").ToString.Replace(".", ",")
        End If

        'txtQuantidade.Text = objNotaFiscal.NotaTrocaOrigem.Itens.Sum(Function(x) x.PesoFiscal).ToString("N0")

        For Each nfOrigem In objNotaFiscal.NotasTrocaOrigem
            QuantidadeNota += nfOrigem.Itens.Sum(Function(x) x.PesoFiscal).ToString("N0")
        Next

        'UTILIZADO PARA LANÇAR CONHECIMENTO COM NOTA JÁ CANCELADA - FURLAN 28/05/2026
        If txtChaveNFe.Text = "41260563358210000176570010000014491297848963" Then QuantidadeNota = 33630
        If txtChaveNFe.Text = "41260563358210000176570010000015681721149797" Then QuantidadeNota = 31520
        If txtChaveNFe.Text = "41260363358210000176570010000008161782360780" Then QuantidadeNota = 41140
        If txtChaveNFe.Text = "41260363358210000176570010000008171992786330" Then QuantidadeNota = 41460
        If txtChaveNFe.Text = "41260363358210000176570010000008181510306119" Then QuantidadeNota = 41140


        txtQuantidade.Text = QuantidadeNota

        txtValor.Text = dValorTotal.ToString("N2")

        If objNotaFiscal.NotaTrocaOrigem.Itens.Sum(Function(x) x.PesoFiscal) > 0 Then
            txtUnitario.Text = (dValorTotal / (objNotaFiscal.NotaTrocaOrigem.Itens.Sum(Function(x) x.PesoFiscal) / 1000)).ToString("N2")
        ElseIf objNotaFiscal.NotaTrocaOrigem.Itens.Sum(Function(x) x.QuantidadeFiscal) > 0 Then
            txtUnitario.Text = (dValorTotal / (objNotaFiscal.NotaTrocaOrigem.Itens.Sum(Function(x) x.QuantidadeFiscal) / 1000)).ToString("N2")
        Else
            txtUnitario.Text = (dValorTotal / (QuantidadeNota / 1000)).ToString("N2")
        End If

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
                    objNotaFiscal.EnderecoTransportador = TranspEnd
                End If

            End If

        Else

            objNotaFiscal.CodigoTransportador = objNotaFiscal.CodigoEmpresa
            objNotaFiscal.EnderecoTransportador = objNotaFiscal.EnderecoEmpresa

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

                If DsXml.Tables("reboque").Columns.Contains("UF") Then
                    objNotaFiscal.PlacaDetalhes.EstadoPlaca02 = DsXml.Tables("reboque").Rows(0)("UF")
                Else
                    objNotaFiscal.PlacaDetalhes.EstadoPlaca02 = ""
                End If

            End If

            objNotaFiscal.PlacaDetalhes.Motorista = New [Lib].Negocio.Cliente
            objNotaFiscal.PlacaDetalhes.Motorista.Codigo = ""
            objNotaFiscal.PlacaDetalhes.Motorista.CodigoEndereco = 0

        End If

        If Not DsXml.Tables("compl") Is Nothing Then
            If DsXml.Tables("compl").Columns.Contains("xObs") Then
                objNotaFiscal.Observacoes = DsXml.Tables("compl").Rows(0)("xObs")
            End If

            If Not DsXml.Tables("ObsCont") Is Nothing AndAlso DsXml.Tables("ObsCont").Columns.Contains("xCampo") Then
                If objNotaFiscal.Observacoes.Length > 0 Then
                    objNotaFiscal.Observacoes += ". " & DsXml.Tables("ObsCont").Rows(0)("xTexto")
                Else
                    objNotaFiscal.Observacoes = DsXml.Tables("ObsCont").Rows(0)("xTexto")
                End If
            End If
        End If

    End Sub

    ' --- Helper: valida chave (44 dígitos numéricos) ---
    Private Function IsChaveValida(ByVal s As String) As Boolean
        If String.IsNullOrEmpty(s) OrElse s.Length <> 44 Then Return False
        For Each ch As Char In s
            If Not Char.IsDigit(ch) Then Return False
        Next
        Return True
    End Function

    Public Sub IncluirDadosNovoCliente(objNotaFiscal As NotaFiscal, DsXml As DataSet, ClienteTemp As ListCliente, verCliente As String)

        Dim novoCliente As New Cliente

        If verCliente.Length = 14 Then

            novoCliente = New Cliente(verCliente)

            novoCliente = DadosUnicoNovoCliente(novoCliente, objNotaFiscal)

            If novoCliente.Nome.Length = 0 Then
                Throw New Exception("A Receita Federal não disponibilizou as informações desse CNPJ.")
                Exit Sub
            End If

        ElseIf verCliente.Length = 11 Then

            novoCliente = New Cliente()

            novoCliente.Codigo = objNotaFiscal.CodigoCliente
            novoCliente.CodigoEndereco = 0
            novoCliente.Nome = DsXml.Tables("dest").Rows(0)("xNome").ToString()
            novoCliente.Fantasia = DsXml.Tables("dest").Rows(0)("xNome").ToString()
            novoCliente.CEP = DsXml.Tables("enderDest").Rows(0)("CEP").ToString()
            novoCliente.Endereco = DsXml.Tables("enderDest").Rows(0)("xLgr").ToString()
            novoCliente.Numero = DsXml.Tables("enderDest").Rows(0)("nro").ToString()
            novoCliente.Complemento = DsXml.Tables("enderDest").Rows(0)("xCpl").ToString()
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

            novoCliente.Email = DsXml.Tables("dest").Rows(0)("email").ToString()
            novoCliente.EmailNFE = DsXml.Tables("dest").Rows(0)("email").ToString()
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

    Private Function DadosUnicoNovoCliente(novoCliente As Cliente, objNotaFiscal As NotaFiscal) As Cliente

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

    Public Sub CarregarNotasDeOrigem(ByVal objNotaFiscal As [Lib].Negocio.NotaFiscal)

        If objNotaFiscal.NotasTrocaOrigem IsNot Nothing AndAlso objNotaFiscal.NotasTrocaOrigem.Count > 0 Then

            grdNotasFretes.DataSource = objNotaFiscal.NotasTrocaOrigem
            grdNotasFretes.DataBind()

            Dim obs As String = String.Empty

            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) Then

                obs = "FRETE REF. NF(S). "
            End If

            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
                obs = "ESTADIA REF. NF(S). "
            End If

            If objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) Then
                obs = "COMPLEMENTO DE FRETE REF. NF(S). "
            End If

            Dim separador As String = ""
            For Each item In objNotaFiscal.NotasTrocaOrigem
                obs += item.Codigo & "-" & item.Cliente.Nome & separador
                separador = ","
            Next

            txtObservacao.Text = objNotaFiscal.Observacoes + " - " + obs
            TabNFOrigem.Visible = True

        End If
    End Sub

    Private Function getSql(ByVal strEmpresa() As String, ByVal strCliente() As String) As String
        Dim Sql As String
        Sql = "SELECT NF.Cliente_Id, NF.EndCliente_Id, Cli.Nome, NF.EntradaSaida_Id, " & vbCrLf &
              "       NF.Serie_Id, NF.Nota_Id, NF.Operacao, NF.SubOperacao, " & vbCrLf &
              "       NFE.ChaveNfe, 'False' AS chkConhecimento, NFE.Data, NFI.PesoFiscal, NFI.ValorOficial, " & vbCrLf &
              "       ISNULL(NxD.PesoLiquido,0) AS PesoDeChegada, ISNULL(NxD.TarifaFrete,0) AS TarifaFrete, ISNULL(NxD.PesoBruto,0) * ISNULL(NxD.TarifaFrete,0) /1000 AS ValorFrete, " & vbCrLf &
              "       CONVERT(INT, nfi.PesoFiscal - ISNULL(JaExisteCTE.QtdeJaUsada,0)) Saldo " & vbCrLf &
              "  FROM NotasFiscais NF " & vbCrLf &
              " INNER JOIN (SELECT Empresa_Id,EndEmpresa_Id,Cliente_Id,EndCliente_Id,EntradaSaida_Id,Serie_Id,Nota_Id," & vbCrLf &
              "                    SUM(QuantidadeFiscal) AS PesoFiscal, " & vbCrLf &
              "                    SUM(Valor) AS ValorOficial " & vbCrLf &
              "	              FROM NotasFiscaisXItens  " & vbCrLf &
              "	             GROUP BY Empresa_Id,EndEmpresa_Id,Cliente_Id,EndCliente_Id,EntradaSaida_Id,Serie_Id,Nota_Id " & vbCrLf &
              "	            ) nfi" & vbCrLf &
              "    ON nfi.Empresa_Id       = NF.Empresa_Id " & vbCrLf &
              "   AND nfi.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND nfi.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND nfi.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND nfi.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND nfi.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND nfi.Nota_Id         = NF.Nota_Id  " & vbCrLf &
              "  LEFT JOIN  NFERealizadas NFe " & vbCrLf &
              "	   ON NF.Empresa_Id      = NFe.Empresa_Id " & vbCrLf &
              "	  AND NF.EndEmpresa_Id   = NFe.EndEmpresa_Id " & vbCrLf &
              "	  AND NF.Cliente_Id      = NFe.Cliente_Id " & vbCrLf &
              "	  AND NF.EndCliente_Id   = NFe.EndCliente_Id " & vbCrLf &
              "	  AND NF.EntradaSaida_Id = NFe.EntradaSaida_Id " & vbCrLf &
              "	  AND NF.Serie_Id        = NFe.Serie_Id " & vbCrLf &
              "	  AND NF.Nota_Id         = NFe.Nota_Id " & vbCrLf &
              "	INNER JOIN  Clientes Cli " & vbCrLf &
              "	   ON NF.Cliente_Id     = Cli.Cliente_Id " & vbCrLf &
              "	  AND NF.EndCliente_Id = Cli.Endereco_Id " & vbCrLf &
              "  LEFT JOIN (SELECT nxn.OrigemEmpresa_Id," & vbCrLf &
              "                    nxn.OrigemEndEmpresa_Id," & vbCrLf &
              "                    nxn.OrigemCliente_Id," & vbCrLf &
              "                    nxn.OrigemEndCliente_Id," & vbCrLf &
              "                    nxn.OrigemEntradaSaida_Id," & vbCrLf &
              "                    nxn.OrigemSerie_Id," & vbCrLf &
              "                    nxn.OrigemNota_Id," & vbCrLf &
              "                    COUNT(*) as qtde," & vbCrLf &
              "                    SUM(ISNULL(NFxI.QuantidadeFiscal,0)) AS QtdeJaUsada " & vbCrLf &
              "               FROM NotasXNotas nxn " & vbCrLf &
              "               JOIN NotasFiscaisXItens NFxI " & vbCrLf &
              "                 ON NFxI.Empresa_Id = nxn.Empresa_Id" & vbCrLf &
              " 	           AND NFxI.EndEmpresa_Id = nxn.EndEmpresa_Id" & vbCrLf &
              "	               AND NFxI.Cliente_Id = nxn.Cliente_Id" & vbCrLf &
              "	               AND NFxI.EndCliente_Id = nxn.EndCliente_Id" & vbCrLf &
              "	               AND NFxI.EntradaSaida_Id = nxn.EntradaSaida_Id" & vbCrLf &
              " 	           AND NFxI.Serie_Id = nxn.Serie_Id" & vbCrLf &
              "	               AND NFxI.Nota_Id = nxn.Nota_Id" & vbCrLf &
              "               LEFT JOIN NotaFiscalReferencial ref" & vbCrLf &
              "                 ON ref.EmpresaReferencial_Id      = nxn.Empresa_Id" & vbCrLf &
              "                AND ref.EndEmpresaReferencial_Id   = nxn.EndEmpresa_Id" & vbCrLf &
              "                AND ref.ClienteReferencial_Id      = nxn.Cliente_Id" & vbCrLf &
              "                AND ref.EndClienteReferencial_Id   = nxn.EndCliente_Id" & vbCrLf &
              "                AND ref.EntradaSaidaReferencial_Id = nxn.EntradaSaida_Id" & vbCrLf &
              "                AND ref.SerieReferencial_Id        = nxn.Serie_Id" & vbCrLf &
              "                AND ref.NotaReferencial_Id         = nxn.Nota_Id" & vbCrLf &
              "               LEFT JOIN NotasFiscais CRTC " & vbCrLf &
              "                 ON CRTC.Empresa_Id      =  nxn.Empresa_Id " & vbCrLf &
              "                AND CRTC.EndEmpresa_Id   =  nxn.EndEmpresa_Id  " & vbCrLf &
              "                AND CRTC.Cliente_Id      =  nxn.Cliente_Id  " & vbCrLf &
              "                AND CRTC.EndCliente_Id   =  nxn.EndCliente_Id  " & vbCrLf &
              "                AND CRTC.EntradaSaida_Id =  nxn.EntradaSaida_Id  " & vbCrLf &
              "                AND CRTC.Serie_Id        =  nxn.Serie_Id  " & vbCrLf &
              "                AND CRTC.Nota_Id         =  nxn.Nota_Id  " & vbCrLf
        If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento Then
            Sql &= "                AND CRTC.TipoDeDocumento = 14 " & vbCrLf
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
            Sql &= "                AND CRTC.TipoDeDocumento = 10 " & vbCrLf
        ElseIf getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
            Sql &= "                AND CRTC.TipoDeDocumento = 13 " & vbCrLf
        Else
            Sql &= "                AND CRTC.TipoDeDocumento = 2 " & vbCrLf
        End If
        Sql &= "              WHERE ref.NotaReferencial_Id Is null" & vbCrLf &
        "                AND CRTC.Situacao = 1" & vbCrLf &
        "              GROUP BY nxn.OrigemEmpresa_Id, nxn.OrigemEndEmpresa_Id, nxn.OrigemCliente_Id, nxn.OrigemEndCliente_Id," & vbCrLf &
        "              nxn.OrigemEntradaSaida_Id, nxn.OrigemSerie_Id, nxn.OrigemNota_Id" & vbCrLf &
        "            ) as JaExisteCTE " & vbCrLf &
        "      ON  NF.Empresa_Id      = JaExisteCTE.OrigemEmpresa_Id          " & vbCrLf &
        "	  AND NF.EndEmpresa_Id   = JaExisteCTE.OrigemEndEmpresa_Id       " & vbCrLf &
        "	  AND NF.Cliente_Id      = JaExisteCTE.OrigemCliente_Id          " & vbCrLf &
        "	  AND NF.EndCliente_Id   = JaExisteCTE.OrigemEndCliente_Id       " & vbCrLf &
        "	  AND NF.EntradaSaida_Id = JaExisteCTE.OrigemEntradaSaida_Id     " & vbCrLf &
        "	  AND NF.Serie_Id        = JaExisteCTE.OrigemSerie_Id            " & vbCrLf &
        "     AND NF.Nota_Id         = JaExisteCTE.OrigemNota_Id             " & vbCrLf &
        "    LEFT JOIN NotasXDestinos NxD" & vbCrLf &
        "      ON NF.Empresa_Id     = NxD.Empresa_Id " & vbCrLf &
        "     AND NF.EndEmpresa_Id   = NxD.EndEmpresa_Id " & vbCrLf &
        "     AND NF.Cliente_Id      = NxD.Cliente_Id" & vbCrLf &
        "     AND NF.EndCliente_Id   = NxD.EndCliente_Id" & vbCrLf &
        "     AND NF.EntradaSaida_Id = NxD.EntradaSaida_Id" & vbCrLf &
        "     AND NF.Serie_Id        = NxD.Serie_Id" & vbCrLf &
        "     AND NF.Nota_Id         = NxD.Nota_Id" & vbCrLf &
        "   WHERE 1=1 " & vbCrLf &
        "     AND (NF.Empresa_Id = '" & strEmpresa(0) & "') " & vbCrLf &
        "     AND (NF.EndEmpresa_Id = " & strEmpresa(1) & ")" & vbCrLf &
        "     AND (NF.EntradaSaida_Id = '" & IIf(rdoSaida.Checked, "S", "E") & "') " & vbCrLf &
        "     AND (NF.Situacao = 1) " & vbCrLf &
        "     AND (NF.CIFFOB = '" & IIf(rdoSaida.Checked, "CIF", "FOB") & "') " & vbCrLf

        If strCliente IsNot Nothing AndAlso strCliente.Length > 0 Then
            Sql &= "     AND (NF.Cliente_Id = '" & strCliente(0) & "') " & vbCrLf &
                   "     AND (NF.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtPedido.Text) Then
            Sql &= "     AND NF.Pedido = " & txtPedido.Text.Trim & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtNumConhecimento.Text) Then
            Dim strNota = txtNumConhecimento.Text.Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            Sql &= "     AND (NF.Nota_Id in (" & String.Join(",", strNota) & ")) " & vbCrLf
        End If

        If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro Then
            Sql &= "     AND (NF.TipoDeDocumento = 1) " & vbCrLf &
                   "     AND Nfi.PesoFiscal - ISNULL(JaExisteCTE.QtdeJaUsada,0) >0 " & vbCrLf
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao _
                OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento _
                OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
            Sql &= "     AND (NF.TipoDeDocumento = 2) " & vbCrLf &
                   "     AND isnull(JaExisteCTE.qtde,0) = 0 " & vbCrLf
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            Sql &= "     AND (NF.TipoDeDocumento in (2,8)) " & vbCrLf &
                   "     AND NOT EXISTS (" & getSqlNFDevolucao() & ")" & vbCrLf

        ElseIf getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
            Sql &= "     AND (NF.TipoDeDocumento = 1) " & vbCrLf &
                   "     AND ISNULL(JaExisteCTE.qtde,0) = 0 " & vbCrLf &
                   "     --AND ISNULL(NxD.PesoBruto,0) > 0 " & vbCrLf &
                   "     --AND ISNULL(NxD.TarifaFrete,0) > 0 " & vbCrLf
        End If

        If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA AndAlso String.IsNullOrWhiteSpace(txtPedido.Text) Then
            Sql &= "  AND (NF.Movimento between '" & CDate(txtPeriodoInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtPeriodoFinal.Text).ToString("yyyy-MM-dd") & "') " & vbCrLf
        Else
            Sql &= "  AND (NF.Movimento between '" & CDate(txtPeriodoInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtPeriodoFinal.Text).ToString("yyyy-MM-dd") & "') " & vbCrLf
        End If

        Sql &= "ORDER BY NF.Nota_Id, NF.Serie_Id, NFE.Data DESC " & vbCrLf

        Return Sql
    End Function

    Private Function getSqlNFDevolucao() As String
        Dim sql As String
        sql = "SELECT NotaDevolucao_id " & vbCrLf &
              "  FROM NotaFiscalDevolucaoXNotaFiscal dev" & vbCrLf &
              " WHERE Dev.EmpresaDevolucao_Id      = NF.Empresa_Id" & vbCrLf &
              "	  AND Dev.EndEmpresaDevolucao_Id   = NF.EndEmpresa_Id " & vbCrLf &
              "	  AND Dev.ClienteDevolucao_Id	   = NF.Cliente_Id " & vbCrLf &
              "	  AND Dev.EndClienteDevolucao_Id   = NF.EndCliente_Id " & vbCrLf &
              "	  AND Dev.Nota_Id			       = NF.Nota_Id  " & vbCrLf &
              "	  AND Dev.Serie_Id		           = NF.Serie_Id  " & vbCrLf &
              "	  AND Dev.EntradaSaida_Id          = NF.EntradaSaida_Id " & vbCrLf
        Return sql
    End Function

    Private Function getSqlSaldo(ByVal strEmpresa() As String, ByVal strCliente() As String) As String
        Dim Sql As String
        Sql = "SELECT NF.Cliente_Id, " & vbCrLf &
              "       NF.EndCliente_Id, " & vbCrLf &
              "       Clientes.Nome, " & vbCrLf &
              "       NF.EntradaSaida_Id, " & vbCrLf &
              "       NF.Serie_Id, " & vbCrLf &
              "       NF.Nota_Id, " & vbCrLf &
              "       NF.Operacao, " & vbCrLf &
              "       NF.SubOperacao, " & vbCrLf &
              "       NFER.ChaveNfe, " & vbCrLf &
              "      'False' as chkConhecimento, " & vbCrLf &
              "       NFER.Data, " & vbCrLf &
              "       ISNULL(NFI.QuantidadeFiscal,0) as PesoFiscal, " & vbCrLf &
              "       ISNULL(consumido.QuantidadeFiscalConsumido ,0) as ConsumidoPeso, " & vbCrLf &
              "        " & vbCrLf &
              "       ISNULL(NFI.Valor, 0) as ValorOficial, " & vbCrLf &
              "       ISNULL(consumido.ValorConsumido,0) as ConsumidoValor, " & vbCrLf &
              "        " & vbCrLf &
              "       ISNULL(NFI.QuantidadeFiscal,0) - ISNULL(consumido.QuantidadeFiscalConsumido ,0) as PesoSaldo, " & vbCrLf &
              "       ISNULL(NFI.Valor, 0)           - ISNULL(consumido.ValorConsumido,0)             as ValorSaldo, " & vbCrLf &
              "        " & vbCrLf &
              "       '' as CRTC_Empresa_Id,  " & vbCrLf &
              "       '' as CRTC_EndEmpresa_Id,  " & vbCrLf &
              "       '' as CRTC_Cliente_Id,  " & vbCrLf &
              "       '' as CRTC_EndCliente_Id,  " & vbCrLf &
              "       '' as CRTC_EntradaSaida_Id,  " & vbCrLf &
              "       '' as CRTC_Serie_Id,  " & vbCrLf &
              "       '' as CRTC_Nota_Id,  " & vbCrLf &
              "       '' as CRTC_TipoDeDocumento  " & vbCrLf &
              "  FROM NotasFiscais NF  " & vbCrLf &
              " INNER JOIN (SELECT Empresa_Id," & vbCrLf &
              "                    EndEmpresa_Id," & vbCrLf &
              "                    Cliente_Id," & vbCrLf &
              "                    EndCliente_Id," & vbCrLf &
              "                    EntradaSaida_Id," & vbCrLf &
              "                    Serie_Id," & vbCrLf &
              "                    Nota_Id," & vbCrLf &
              "                    SUM(QuantidadeFiscal) AS QuantidadeFiscal," & vbCrLf &
              "                    SUM(Valor) AS Valor " & vbCrLf &
              "	              FROM NotasFiscaisXItens " & vbCrLf &
              "	             GROUP BY Empresa_id," & vbCrLf &
              "                    EndEmpresa_id," & vbCrLf &
              "                    Cliente_id," & vbCrLf &
              "                    EndCliente_id," & vbCrLf &
              "                    EntradaSaida_id," & vbCrLf &
              "                    Serie_id," & vbCrLf &
              "                    Nota_id " & vbCrLf &
              "             ) As NFI  " & vbCrLf &
              "   on NFI.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "  AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "  AND NFI.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "  AND NFI.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "  AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "  AND NFI.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "  AND NFI.Nota_Id         = NF.Nota_Id             " & vbCrLf &
              "     INNER JOIN Clientes " & vbCrLf &
              "   ON NF.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf &
              "  AND NF.EndCliente_Id = Clientes.Endereco_Id  " & vbCrLf &
              "     LEFT JOIN NFERealizadas NFER " & vbCrLf &
              "   ON NF.Empresa_Id      = NFER.Empresa_Id " & vbCrLf &
              "  AND NF.EndEmpresa_Id   = NFER.EndEmpresa_Id " & vbCrLf &
              "  AND NF.Cliente_Id      = NFER.Cliente_Id " & vbCrLf &
              "  AND NF.EndCliente_Id   = NFER.EndCliente_Id " & vbCrLf &
              "  AND NF.EntradaSaida_Id = NFER.EntradaSaida_Id " & vbCrLf &
              "  AND NF.Serie_Id        = NFER.Serie_Id " & vbCrLf &
              "  AND NF.Nota_Id         = NFER.Nota_Id " & vbCrLf &
              " LEFT JOIN (SELECT nxi.Empresa_Id, " & vbCrLf &
              "                   nxi.EndEmpresa_Id, " & vbCrLf &
              "                   nxi.Cliente_Id, " & vbCrLf &
              "                   nxi.EndCliente_Id, " & vbCrLf &
              "                   nxi.EntradaSaida_Id, " & vbCrLf &
              "                   nxi.Serie_Id, " & vbCrLf &
              "                   nxi.Nota_Id, " & vbCrLf &
              "                   SUM(nfr.Quantidade) AS QuantidadeFiscalConsumido, " & vbCrLf &
              "                   SUM(nfr.Valor) AS ValorConsumido " & vbCrLf &
              "              FROM NotasFiscaisXItens nxi " & vbCrLf &
              "             INNER JOIN NotaFiscalReferencial nfr " & vbCrLf &
              "                ON nxi.Empresa_Id = nfr.Empresa_Id " & vbCrLf &
              "               AND nxi.EndEmpresa_Id = nfr.EndEmpresa_Id " & vbCrLf &
              "               AND nxi.Cliente_Id = nfr.Cliente_Id " & vbCrLf &
              "               AND nxi.EndCliente_Id = nfr.EndCliente_Id " & vbCrLf &
              "               AND nxi.EntradaSaida_Id = nfr.EntradaSaida_Id " & vbCrLf &
              "               AND nxi.Nota_Id = nfr.Nota_Id " & vbCrLf &
              "               AND nxi.Serie_Id = nfr.Serie_Id " & vbCrLf &
              "               AND nxi.Produto_Id = nfr.Produto_Id " & vbCrLf &
              "               AND nxi.CFOP_Id = nfr.CFOP_Id " & vbCrLf &
              "               AND nxi.Sequencia_Id = nfr.Sequencia_Id " & vbCrLf &
              "             GROUP BY nxi.Empresa_id, " & vbCrLf &
              "                   nxi.EndEmpresa_id, " & vbCrLf &
              "                   nxi.Cliente_id, " & vbCrLf &
              "                   nxi.EndCliente_id, " & vbCrLf &
              "                   nxi.EntradaSaida_id, " & vbCrLf &
              "                   nxi.Serie_id, " & vbCrLf &
              "                   nxi.Nota_id " & vbCrLf &
              "            ) As Consumido   " & vbCrLf &
              "   on Consumido.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
              "  AND Consumido.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
              "  AND Consumido.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
              "  AND Consumido.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
              "  AND Consumido.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
              "  AND Consumido.Serie_Id        = NF.Serie_Id " & vbCrLf &
              "  AND Consumido.Nota_Id         = NF.Nota_Id " & vbCrLf &
              "WHERE 1=1 " & vbCrLf &
              "  AND (NF.Empresa_Id = '" & strEmpresa(0) & "') " & vbCrLf &
              "  AND (NF.EndEmpresa_Id = " & strEmpresa(1) & ")" & vbCrLf &
              "  AND (NF.EntradaSaida_Id = '" & IIf(rdoSaida.Checked, "S", "E") & "') " & vbCrLf &
              "  AND (NF.Situacao = 1) " & vbCrLf &
              "  AND (NF.CIFFOB = '" & IIf(rdoSaida.Checked, "CIF", "FOB") & "') " & vbCrLf

        If strCliente IsNot Nothing AndAlso strCliente.Length > 0 Then
            Sql &= "  AND (NF.Cliente_Id = '" & strCliente(0) & "') " & vbCrLf &
                   "  AND (NF.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtNumConhecimento.Text) Then
            Dim strNota = txtNumConhecimento.Text.Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            Sql &= "  AND (NF.Nota_Id in (" & String.Join(",", strNota) & ")) " & vbCrLf
        End If

        If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro _
                OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
            Sql &= "  AND (NF.TipoDeDocumento = 1) " & vbCrLf
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao Then
            Sql &= "  AND (NF.TipoDeDocumento = 2) " & vbCrLf
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            Sql &= "  AND (NF.TipoDeDocumento in (2,8)) " & vbCrLf &
                   "  AND NOT EXISTS (" & getSqlNFDevolucao() & ")" & vbCrLf
        End If

        If rdoFerroviario.Checked Then
            Sql &= "  AND (ISNULL(NFI.QuantidadeFiscal,0) - ISNULL(consumido.QuantidadeFiscalConsumido ,0)) > 0 " & vbCrLf
        End If

        Sql &= "  AND (NF.Movimento between '" & CDate(txtPeriodoInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtPeriodoFinal.Text).ToString("yyyy-MM-dd") & "') " & vbCrLf &
               "ORDER BY NF.Nota_Id, NF.Serie_Id, NFER.Data DESC " & vbCrLf
        Return Sql
    End Function

    Private Sub ConsultarNotas()
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "É necessário selecionar a empresa!")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(ddlTipoConhecimento.SelectedValue) OrElse Not ddlTipoConhecimento.SelectedValue <> "0" Then
            MsgBox(Me.Page, "É necessário selecionar o tipo de conhecimento!")
            Exit Sub
        End If

        If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA AndAlso String.IsNullOrWhiteSpace(txtPedido.Text) Then
            MsgBox(Me.Page, "No tipo de Conhecimento  " & getTipoDeFrete.ToString & " é necessário escolher um pedido!")
            txtPedido.Focus()
            Exit Sub
        End If

        Dim strEmpresa() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim strCliente() As String = Nothing

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            strCliente = txtCodigoCliente.Value.Split("-")
        End If

        Dim sql As String = IIf(rdoRodoviario.Checked, getSql(strEmpresa, strCliente), getSqlSaldo(strEmpresa, strCliente))
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NotasFiscais")
        If rdoRodoviario.Checked Then
            grdNFeSaldo.Visible = False
            grdNFeSaldo.DataSource = New List(Of Object)
            grdNFeSaldo.DataBind()
            grdNFe.Visible = True
            grdNFe.DataSource = ds
            If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
                grdNFe.Columns(12).Visible = True
                grdNFe.Columns(13).Visible = True
                grdNFe.Columns(14).Visible = True
            Else
                grdNFe.Columns(12).Visible = False
                grdNFe.Columns(13).Visible = False
                grdNFe.Columns(14).Visible = False
            End If
            grdNFe.DataBind()

        Else
            grdNFe.Visible = False
            grdNFe.DataSource = New List(Of Object)
            grdNFe.DataBind()
            grdNFeSaldo.Visible = True
            grdNFeSaldo.DataSource = ds
            grdNFeSaldo.DataBind()
        End If
    End Sub

    Private Sub LimparCamposDaConsulta()
        DdlEmpresa.SelectedValue = ""
        txtCodigoCliente.Value = ""
        txtNomeCliente.Text = ""
        txtNumConhecimento.Text = ""
        txtPedido.Text = ""
        rdoSaida.Checked = True
        rdoEntrada.Checked = False
        rdoRodoviario.Checked = True
        rdoFerroviario.Checked = False
        ddlTipoConhecimento.SelectedIndex = "0"
        ddlTipoConhecimento_SelectedIndexChanged(ddlTipoConhecimento, New EventArgs)
        txtCodigoClienteRPA.Value = ""
        txtNomeClienteRPA.Text = ""
        txtPesoTotalSelecionado.Text = ""
        txtValorFreteSelecionado.Text = ""
        txtPeriodoInicial.Text = String.Format("01/01/{0}", DateTime.Now.Year)
        txtPeriodoFinal.Text = DateTime.Now.ToShortDateString()
        grdNFe.DataSource = New List(Of Object)
        grdNFe.DataBind()
        grdNFeSaldo.DataSource = New List(Of Object)
        grdNFeSaldo.DataBind()
    End Sub

    Private Sub LimparCampos()

        HID.Value = Guid.NewGuid().ToString

        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaCTeXNotas.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucConsultaEstados.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)
        ucConsultaRelatorio.SetarHID(HID.Value)
        ucConsultaCadastro.SetarHID(HID.Value)
        ucInutilizacao.SetarHID(HID.Value)
        ucEmailNFe.SetarHID(HID.Value)
        ucEmitirCTe.SetarHID(HID.Value)

        'Só irá limpar os campos da consulta caso não esteja marcado o checkbox de Mais de um CTE para a NF.
        If Not chkNFxCtes.Checked Then
            LimparCamposDaConsulta()
        End If

        txtNomeDaEmpresa.Text = ""
        txtEnderecoDaEmpresa.Text = ""
        txtComplementoDaEmpresa.Text = ""
        txtBairroDaEmpresa.Text = ""
        txtCepDaEmpresa.Text = ""
        txtInscricaoDaEmpresa.Text = ""
        txtCnpjDaEmpresa.Text = ""

        txtNomeDoTomador.Text = ""
        txtEnderecoDoTomador.Text = ""
        txtComplementoDoTomador.Text = ""
        txtBairroDoTomador.Text = ""
        txtCepDoTomador.Text = ""
        txtInscricaoDoTomador.Text = ""
        txtCnpjDoTomador.Text = ""

        txtNomeDoEmitente.Text = ""
        txtEnderecoDoEmitente.Text = ""
        txtComplementoDoEmitente.Text = ""
        txtBairroDoEmitente.Text = ""
        txtCepDoEmitente.Text = ""
        txtInscricaoDoEmitente.Text = ""
        txtCnpjDoEmitente.Text = ""

        txtNomeDoRemetente.Text = ""
        txtEnderecoDoRemetente.Text = ""
        txtComplementoDoRemetente.Text = ""
        txtBairroDoRemetente.Text = ""
        txtCepDoRemetente.Text = ""
        txtInscricaoDoRemetente.Text = ""
        txtCnpjDoRemetente.Text = ""

        txtNomeDoCliente.Text = ""
        txtCnpjDoCliente.Text = ""
        txtDataDeEmissao.Text = DateTime.Now.ToShortDateString()
        txtDataDeSaida.Text = DateTime.Now.ToShortDateString()

        txtEnderecoDoCliente.Text = ""
        txtComplementoDoCliente.Text = ""
        txtBairroDoCliente.Text = ""
        txtCepDoCliente.Text = ""
        txtDataDeSaida.Text = DateTime.Now.ToShortDateString()

        txtVencimentoFatura.Text = DateTime.Now.ToShortDateString()

        txtCidadeDoCliente.Text = ""
        txtTelefoneDoCliente.Text = ""
        txtEstadoDoCliente.Text = ""
        txtInscricaoDoCliente.Text = ""
        txtNomeDoTransportador.Text = ""
        txtRNTRCTransportador.Text = ""
        txtEnderecoDoTransportador.Text = ""
        txtCidadeDoTransportador.Text = ""
        txtEstadoDoTransportador.Text = ""
        txtCnpjDoTransportador.Text = ""
        txtInscricaoDoTransportador.Text = ""

        txtNumero.Text = ""
        txtSerie.Text = ""

        txtNumero.Enabled = True
        txtSerie.Enabled = True
        txtChaveNFe.Enabled = True

        DdlEmpresa.Enabled = True
        ddlTipoConhecimento.Enabled = True
        ddlSituacao.SelectedIndex = 0
        lblUsuario.Text = HttpContext.Current.Session("ssNomeUsuario")

        txtNomeDoMotorista.Text = ""
        txtEnderecoDoMotorista.Text = ""
        txtCidadeDoMotorista.Text = ""
        txtEstadoDoMotorista.Text = ""
        txtCPFDoMotorista.Text = ""
        txtHabilitacaoDoMotorista.Text = ""

        txtCPlaca1.Text = ""
        txtCidadePlaca1.Text = ""
        txtEstadoPlaca1.Text = ""
        txtRNTRCPlaca1.Text = ""
        txtProprietarioPlaca1.Text = ""

        txtCPlaca2.Text = ""
        txtCidadePlaca2.Text = ""
        txtEstadoPlaca2.Text = ""
        txtRNTRCPlaca2.Text = ""
        txtProprietarioPlaca2.Text = ""

        txtCPlaca3.Text = ""
        txtCidadePlaca3.Text = ""
        txtEstadoPlaca3.Text = ""
        txtRNTRCPlaca3.Text = ""
        txtProprietarioPlaca3.Text = ""

        txtOperacao.Text = ""
        txtEntSai.Text = ""
        txtCfop.Text = ""

        txtQuantidade.Text = "0"
        txtQuantidade.Enabled = False

        txtUnitario.Text = ""
        txtUnitario.Enabled = True

        txtValor.Text = "0,00"
        txtObservacao.Height = Unit.Pixel(105)
        txtObservacao.Text = ""

        gridEncargos.DataSource = Nothing
        gridEncargos.DataBind()

        gridNotasFiscais.DataSource = Nothing
        gridNotasFiscais.DataBind()

        grdFaturas.DataSource = Nothing
        grdFaturas.DataBind()

        grdVencimentos.DataSource = Nothing
        grdVencimentos.DataBind()

        grdContabilizacao.DataSource = Nothing
        grdContabilizacao.DataBind()

        grdNotasFretes.DataSource = Nothing
        grdNotasFretes.DataBind()

        txtEmbarque.Text = ""
        txtMunicipioEmbarque.Text = ""
        txtEstadoEmbarque.Text = ""
        txtNomeOrigemDestino.Text = ""
        txtMunicipioOrigemDestino.Text = ""
        txtEstadoOrigemDestino.Text = ""

        txtChaveNFe.Text = ""
        txtCodigoFavorecidoCartao.Value = ""
        txtFavorecidoCartao.Text = ""
        txtEnderecoFavorecidoCartao.Text = ""
        txtCidadeFavorecidoCartao.Text = ""
        txtEstadoFavorecidoCartao.Text = ""
        txtCPFFavorecidoCartao.Text = ""

        ShowCalendar(Me.Page, txtDataDeEmissao)
        ShowCalendar(Me.Page, txtDataDeSaida)

        txtDataDeEmissao.Enabled = True
        txtDataDeSaida.Enabled = True

        ucFile.Clear()

        btnTransportador.Enabled = False
        btnRecontabilizar.Enabled = False

        divEmitente.Visible = False

        If Session("ssNomeUsuario") = "FURLAN" Then
            lnkEmitir.Parent.Visible = True
        Else
            lnkEmitir.Parent.Visible = False
        End If

        lnkConsultarSefaz.Parent.Visible = False
        lnkNovo.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True
        lnkEspelho.Parent.Visible = True
        lnkEnviarSEFAZ.Parent.Visible = False
        lnkEnviarEmail.Parent.Visible = False
        lnkDACTE.Parent.Visible = False
        lnkVisualizar.Parent.Parent.Parent.Visible = False
        lnkDuplicar.Parent.Visible = False
        lnkExcel.Parent.Visible = False
        hdfTipo.Value = getTipoDeFrete()

        Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
        SetarEmpresa(objEmpresa)
        Session("ssTransportadoraCTE" & HID.Value) = objEmpresa

        'If Left(HttpContext.Current.Session("ssNomeEmpresa"), 17) = "AGRICOLA ALVORADA" Then
        '    Dim objCliente = New [Lib].Negocio.Cliente("04854422000347", 0)
        '    Session("ssTransportadoraCTE" & HID.Value) = objCliente
        'Else
        '    Session("ssTransportadoraCTE" & HID.Value) = objEmpresa
        'End If

        If Session("ssTransportadoraCTE" & HID.Value) IsNot Nothing Then
            objConhecimento = New [Lib].Negocio.NotaFiscal()
            objConhecimento.IUD = "I"
            'If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
            '        OrElse CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao Then
            '    objConhecimento.CodigoEmpresa = CType(Session("ssTransportadoraCTE" & HID.Value), [Lib].Negocio.Cliente).Codigo
            '    objConhecimento.EnderecoEmpresa = CType(Session("ssTransportadoraCTE" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            'ElseIf Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            '    Dim strEmpresa() As String = DdlEmpresa.SelectedValue.ToString.Split("-")
            '    Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
            '    If objEmpresa IsNot Nothing Then
            '        objConhecimento.CodigoEmpresa = objEmpresa.Codigo
            '        objConhecimento.EnderecoEmpresa = objEmpresa.CodigoEndereco
            '    End If
            'End If
            objConhecimento.CodigoEmpresa = objEmpresa.Codigo
            objConhecimento.EnderecoEmpresa = objEmpresa.CodigoEndereco
            objConhecimento.CodigoTipoDeDocumento = getTipoDeDocumento()
            objConhecimento.TipoDeDocumentoFrete = getTipoDeFrete()

            If objEmpresa.Empresa.EmitirCTe Then
                objConhecimento.EntradaSaida = eEntradaSaida.Saida
                objConhecimento.Eletronica = True
            Else
                objConhecimento.EntradaSaida = getEntradaSaida()
                objConhecimento.Eletronica = getEletronica()
            End If

            objConhecimento.CIFFOB = IIf(getEntradaSaida() = eEntradaSaida.Saida, eTiposFrete.CIF, eTiposFrete.FOB)
            objConhecimento.CodigoSituacao = getSituacao()
            SalvarSessaoConhecimento()

            If Not objEmpresa.Empresa.EmitirCTe Then
                DdlEmpresa.Enabled = False
                ddlTipoConhecimento.Enabled = False
            End If

        End If
    End Sub

    Private Sub SetarEmpresa(ByVal pEmpresa As [Lib].Negocio.Cliente)

        RecuperarSessaoConhecimento()

        If objConhecimento Is Nothing Then
            objConhecimento = New [Lib].Negocio.NotaFiscal
        End If
        objConhecimento.Empresa = CType(pEmpresa, [Lib].Negocio.Cliente)

        With objConhecimento.Empresa
            objConhecimento.CodigoEmpresa = .Codigo
            objConhecimento.EnderecoEmpresa = .CodigoEndereco

            txtNomeDaEmpresa.Text = .Nome
            txtInscricaoDaEmpresa.Text = .InscricaoEstadual
            txtCnpjDaEmpresa.Text = .CodigoFormatado

        End With

        SalvarSessaoConhecimento()

        'Verifica se a empresa está habilitada para gravar arquivo
        trFile.Visible = Not String.IsNullOrWhiteSpace(pEmpresa.Empresa.PathDownloadNFe)

        LiberaEmpresa()

    End Sub

    Private Sub LiberaEmpresa()

        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
        Else
            btnEmpresa.Enabled = True
        End If

        RecuperarSessaoConhecimento()

        If objConhecimento.Empresa.Empresa.NotaFiscalEletronica AndAlso objConhecimento.Empresa.Empresa.NossaEmissao Then
            Dim fileEmpresa As String = "C:\NGS\xmlCopy\" & objConhecimento.Empresa.Codigo & ".xml"

            If Not File.Exists(fileEmpresa) Then
                MsgBox(Me.Page, "Arquivo de Configuracão da Empresa não foi encontrado.", "~/Expedicao.aspx")
                Exit Sub
            End If

            Dim xmlConf As New XmlDocument
            xmlConf.Load(fileEmpresa)

            Session("objHorarioServidor" & HID.Value) = xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("horarioServidor").InnerText
        End If
    End Sub

    Private Sub Selecionar()
        Try
            If String.IsNullOrWhiteSpace(ddlTipoConhecimento.SelectedValue) OrElse Not ddlTipoConhecimento.SelectedValue <> "0" Then
                MsgBox(Me.Page, "É necessário selecionar o tipo de conhecimento para emissão!")
                Exit Sub
            End If

            If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA AndAlso String.IsNullOrWhiteSpace(txtCodigoClienteRPA.Value) Then
                MsgBox(Me.Page, "É necessário selecionar o Cliente para o Recibo de Pagamento a Autônomo (RPA)!")
                Exit Sub
            End If

            Dim PesoDeChegadaPreenchido As Boolean = True
            Dim TarifaDoFretePreenchido As Boolean = True

            Dim aux As Integer = 0
            Dim grd As GridView = IIf(grdNFe.Visible, grdNFe, grdNFeSaldo)
            For Each row As GridViewRow In grd.Rows
                If CType(row.FindControl("chkConhecimento"), CheckBox).Checked Then
                    'Verificar se pelo menos uma das notas selecionadas têm ou peso de chegada ou valor de frete não preenchidos
                    If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
                        If CInt(CType(row.FindControl("txtPesoDeChegada"), TextBox).Text) <= 0 Then
                            PesoDeChegadaPreenchido = False
                        End If

                        If CInt(CType(row.FindControl("txtValorFrete"), TextBox).Text) <= 0 Then
                            TarifaDoFretePreenchido = False
                        End If
                    End If
                    aux += 1
                End If
            Next

            If Not aux > 0 Then
                For Each row As GridViewRow In grd.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkConhecimento"), CheckBox)
                    If (chk IsNot Nothing) Then
                        chk.Checked = False
                    End If
                Next
                MsgBox(Me.Page, "Selecione pelo menos um registro para emissão do conhecimento de transporte!")
            Else
                If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
                    If Not PesoDeChegadaPreenchido Then
                        MsgBox(Me.Page, "Uma ou mais notas estão sem o peso de chegada!")
                        Exit Sub
                    ElseIf Not TarifaDoFretePreenchido Then
                        MsgBox(Me.Page, "Uma ou mais notas estão sem tarifa de frete definido!")
                        Exit Sub
                    End If
                End If

                Dim auxMsg As Boolean = True
                Dim msg As String = ""
                Dim codigoCliente As String = ""
                Dim enderecoCliente As String = ""
                Dim operacao As String = ""
                Dim suboperacao As String = ""
                Dim codigoTransportador As String = ""
                Dim enderecoTransportador As String = ""
                Dim codigoMotorista As String = ""
                Dim enderecoMotorista As String = ""

                PesoTotalRPA = 0

                For Each row As GridViewRow In grd.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkConhecimento"), CheckBox)
                    If (chk IsNot Nothing AndAlso chk.Checked) Then
                        Dim strEmpresa() As String = DdlEmpresa.SelectedValue.ToString.Split("-")
                        objNotaFiscal = New [Lib].Negocio.NotaFiscal()
                        objNotaFiscal.CodigoEmpresa = strEmpresa(0)
                        objNotaFiscal.EnderecoEmpresa = strEmpresa(1)

                        objNotaFiscal.CodigoCliente = row.Cells(4).Text()
                        objNotaFiscal.EnderecoCliente = row.Cells(5).Text()

                        objNotaFiscal.Serie = row.Cells(3).Text()
                        objNotaFiscal.Codigo = row.Cells(2).Text()
                        objNotaFiscal.EntradaSaida = IIf(row.Cells(1).Text() = "S", [Lib].Negocio.eEntradaSaida.Saida, [Lib].Negocio.eEntradaSaida.Entrada)
                        objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)
                        hdfCodigo.Value = objNotaFiscal.Codigo & "-" & objNotaFiscal.Serie & "-" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Saida, "S", "E")
                        objNotaFiscal.IUD = "I"

                        If Not String.IsNullOrWhiteSpace(CType(row.FindControl("txtPesoDeChegada"), TextBox).Text) AndAlso IsNumeric(CType(row.FindControl("txtPesoDeChegada"), TextBox).Text) Then
                            PesoTotalRPA += CDec(CType(row.FindControl("txtPesoDeChegada"), TextBox).Text)
                        End If

                        If rdoFerroviario.Checked AndAlso grdNFeSaldo.Visible Then
                            Dim txtVlrPeso As TextBox = CType(row.FindControl("txtVlrPeso"), TextBox)
                            If txtVlrPeso IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txtVlrPeso.Text) Then
                                For Each item As [Lib].Negocio.NotaFiscalXItem In objNotaFiscal.Itens
                                    item.QuantidadeFisica = txtVlrPeso.Text
                                    item.QuantidadeFiscal = txtVlrPeso.Text
                                Next
                            End If
                        End If

                        If (lstNotaFiscal Is Nothing) Then
                            lstNotaFiscal = New List(Of [Lib].Negocio.NotaFiscal)
                        End If
                        lstNotaFiscal.Add(objNotaFiscal)

                        If (objNotaFiscal Is Nothing) Then
                            auxMsg = False
                            Continue For
                        End If

                        If (String.IsNullOrWhiteSpace(codigoCliente) AndAlso String.IsNullOrWhiteSpace(enderecoCliente)) Then
                            codigoCliente = row.Cells(4).Text()
                            enderecoCliente = row.Cells(5).Text()
                            operacao = row.Cells(7).Text()
                            suboperacao = row.Cells(8).Text()
                        End If

                        If (row.Cells(4).Text() <> codigoCliente OrElse row.Cells(5).Text() <> enderecoCliente) Then
                            auxMsg = False
                            If (Not String.IsNullOrWhiteSpace(msg)) Then
                                msg &= " - "
                            End If
                            msg &= "É necessário selecionar registros com o mesmo cliente! \n"
                        End If

                        If (row.Cells(7).Text() <> operacao OrElse row.Cells(8).Text() <> suboperacao) Then
                            auxMsg = False
                            If (Not String.IsNullOrWhiteSpace(msg)) Then
                                msg &= " - "
                            End If
                            msg &= "É necessário selecionar registros com a mesma operação e sub-operação! \n"
                        End If
                    End If
                Next

                If Not String.IsNullOrWhiteSpace(msg) Then
                    MsgBox(Me.Page, msg)
                    Exit Sub
                Else
                    GerarCTRC()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Private Function getSituacao() As Integer

        If objConhecimento.XMLImportado Then
            Return CInt(eSituacao.Normal)
        ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                OrElse CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            Return CInt(eSituacao.Bloqueado)
        End If
        Return CInt(eSituacao.Normal)
    End Function

    Private Function getTipoDeDocumento() As Integer
        If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento Then
            Return Convert.ToInt32(eTipoDeDocumento.ComplementoDeFrete)
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
            Return Convert.ToInt32(eTipoDeDocumento.Estadia)
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            Return Convert.ToInt32(eTipoDeDocumento.Anulacao)
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA Then
            Return Convert.ToInt32(eTipoDeDocumento.RPA)
        End If
        Return Convert.ToInt32(eTipoDeDocumento.CTRC)
    End Function

    Private Function getEletronica() As Boolean
        If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                OrElse CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            Return True
        End If
        Return False
    End Function

    Private Function getNossaEmissao() As Boolean

        If objConhecimento.XMLImportado Then
            Return False
        ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
            Return True
        ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            Return True
        Else
            Return False
        End If

    End Function

    Private Function getEntradaSaida() As eEntradaSaida
        If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                OrElse CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            Return eEntradaSaida.Saida
        ElseIf Not objConhecimento.NotaTrocaOrigem Is Nothing AndAlso CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento Then
            RecuperarSessaoConhecimento()
            Return objConhecimento.NotaTrocaOrigem.EntradaSaida
        End If
        Return eEntradaSaida.Entrada
    End Function

    Private Function getTipoDeFrete() As eTipoDeDocumentoFrete
        Return CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete)
    End Function

    Private Sub GerarCTRC()
        Dim index As Integer = 0
        Dim grd As GridView = IIf(grdNFe.Visible, grdNFe, grdNFeSaldo)
        Dim ValorTotalDoFrete As Decimal = 0

        For Each row As GridViewRow In grd.Rows
            Dim chk As CheckBox = CType(row.FindControl("chkConhecimento"), CheckBox)
            If (chk IsNot Nothing AndAlso chk.Checked) Then
                RecuperarSessaoConhecimento()
                Dim Empresa() As String = DdlEmpresa.SelectedValue.ToString.Split("-")
                objNotaFiscal = New [Lib].Negocio.NotaFiscal()
                objNotaFiscal.CodigoEmpresa = Empresa(0)
                objNotaFiscal.EnderecoEmpresa = Empresa(1)
                objNotaFiscal.CodigoCliente = row.Cells(4).Text()
                objNotaFiscal.EnderecoCliente = row.Cells(5).Text()
                objNotaFiscal.Serie = row.Cells(3).Text()
                objNotaFiscal.Codigo = row.Cells(2).Text()
                objNotaFiscal.EntradaSaida = IIf(row.Cells(1).Text() = "S", [Lib].Negocio.eEntradaSaida.Saida, [Lib].Negocio.eEntradaSaida.Entrada)
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(objNotaFiscal)
                objNotaFiscal.IUD = "I"

                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                        OrElse CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao Then
                    objConhecimento.CodigoEmpresa = CType(Session("ssTransportadoraCTE" & HID.Value), [Lib].Negocio.Cliente).Codigo
                    objConhecimento.EnderecoEmpresa = CType(Session("ssTransportadoraCTE" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
                ElseIf Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                    Dim strEmpresa() As String = DdlEmpresa.SelectedValue.ToString.Split("-")
                    Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
                    If objEmpresa IsNot Nothing Then
                        objConhecimento.CodigoEmpresa = objEmpresa.Codigo
                        objConhecimento.EnderecoEmpresa = objEmpresa.CodigoEndereco
                    End If
                End If

                objConhecimento = If(objConhecimento, New [Lib].Negocio.NotaFiscal())

                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA Then
                    objConhecimento.CodigoCliente = txtCodigoClienteRPA.Value.Split("-")(0)
                    objConhecimento.EnderecoCliente = txtCodigoClienteRPA.Value.Split("-")(1)
                Else
                    objConhecimento.CodigoCliente = objNotaFiscal.CodigoTransportador
                    objConhecimento.EnderecoCliente = objNotaFiscal.EnderecoTransportador
                End If


                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
                    objConhecimento.CodigoDeposito = objNotaFiscal.CodigoEmpresa
                    objConhecimento.EnderecoDeposito = objNotaFiscal.EnderecoEmpresa
                Else
                    objConhecimento.CodigoDeposito = objNotaFiscal.CodigoDeposito
                    objConhecimento.EnderecoDeposito = objNotaFiscal.EnderecoDeposito
                End If

                objConhecimento.CarregandoNota = True
                objConhecimento.EntradaSaida = getEntradaSaida()
                objConhecimento.NossaEmissao = getNossaEmissao()
                objConhecimento.Movimento = objNotaFiscal.Movimento
                objConhecimento.DataTermino = objNotaFiscal.DataTermino

                If row.Cells(1).Text() = "S" Then
                    objConhecimento.CodigoDestino = objNotaFiscal.CodigoDestino
                    objConhecimento.EnderecoDestino = objNotaFiscal.EnderecoDestino
                Else
                    objConhecimento.CodigoDestino = objNotaFiscal.CodigoDeposito
                    objConhecimento.EnderecoDestino = objNotaFiscal.EnderecoDeposito
                End If

                If objNotaFiscal.CodigoLocalEmbarque.Length > 0 Then
                    objConhecimento.CodigoLocalEmbarque = objNotaFiscal.CodigoLocalEmbarque
                    objConhecimento.EndLocalEmbarque = objNotaFiscal.EndLocalEmbarque
                Else
                    If row.Cells(1).Text() = "S" Then
                        objConhecimento.CodigoLocalEmbarque = objNotaFiscal.CodigoDeposito
                        objConhecimento.EndLocalEmbarque = objNotaFiscal.EnderecoDeposito
                    Else
                        objConhecimento.CodigoLocalEmbarque = objNotaFiscal.CodigoCliente
                        objConhecimento.EndLocalEmbarque = objNotaFiscal.EnderecoCliente
                    End If
                End If

                objConhecimento.Empresa.Empresa = New ClienteXEmpresa(objConhecimento.CodigoEmpresa, objConhecimento.EnderecoEmpresa)
                objConhecimento.TipoDeDocumentoFrete = CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete)
                objConhecimento.CodigoMotorista = objNotaFiscal.CodigoMotorista
                objConhecimento.EnderecoMotorista = objNotaFiscal.EnderecoMotorista

                'No RPA o Cliente é também o tranportador (Conveniado) para gerar os títulos para o pagamento
                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA Then
                    objConhecimento.CodigoTransportador = objConhecimento.CodigoCliente
                    objConhecimento.EnderecoTransportador = objConhecimento.EnderecoCliente
                Else
                    objConhecimento.CodigoTransportador = objNotaFiscal.CodigoTransportador
                    objConhecimento.EnderecoTransportador = objNotaFiscal.EnderecoTransportador
                End If

                objConhecimento.PlacaDetalhes = objNotaFiscal.PlacaDetalhes
                objConhecimento.ObservacoesDeEmbarque = String.Format("{0} {1}", objNotaFiscal.Observacoes, objNotaFiscal.ObservacoesDeEmbarque).Trim()
                objConhecimento.NotaTrocaOrigem = objNotaFiscal
                objConhecimento.NotasTrocaOrigem = lstNotaFiscal
                objConhecimento.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                objConhecimento.DataInclusao = Now()
                txtDataDeEmissao.Enabled = False
                txtDataDeSaida.Enabled = True

                If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA Then
                    ValorTotalDoFrete += CDec(CType(row.FindControl("txtValorFrete"), TextBox).Text())
                End If

                objConhecimento.ValorFrete = ValorTotalDoFrete
                SalvarSessaoConhecimento()
            End If
        Next

        CarregarCTE(True)
    End Sub

    Private Sub GerarCTRCAutomatico()

        If objConhecimento.NotaTrocaOrigem Is Nothing OrElse objConhecimento.NotaTrocaOrigem.Codigo = 0 Then
            Exit Sub
        End If

        RecuperarSessaoConhecimento()

        If Not objConhecimento.NotaTrocaOrigem Is Nothing Then
            objConhecimento.CodigoDeposito = objConhecimento.NotaTrocaOrigem.CodigoDeposito
            objConhecimento.EnderecoDeposito = objConhecimento.NotaTrocaOrigem.EnderecoDeposito
        End If

        hdfTipo.Value = ddlTipoConhecimento.SelectedValue
        objConhecimento.TipoDeDocumentoFrete = ddlTipoConhecimento.SelectedValue
        objConhecimento.NossaEmissao = getNossaEmissao()
        objConhecimento.EntradaSaida = getEntradaSaida()

        If objConhecimento.NotaTrocaOrigem.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida Then
            objConhecimento.CodigoDestino = objConhecimento.NotaTrocaOrigem.CodigoDestino
            objConhecimento.EnderecoDestino = objConhecimento.NotaTrocaOrigem.EnderecoDestino
        Else
            objConhecimento.CodigoDestino = objConhecimento.NotaTrocaOrigem.CodigoDeposito
            objConhecimento.EnderecoDestino = objConhecimento.NotaTrocaOrigem.EnderecoDeposito
        End If

        If objConhecimento.NotaTrocaOrigem.CodigoLocalEmbarque.Length > 0 Then
            objConhecimento.CodigoLocalEmbarque = objConhecimento.NotaTrocaOrigem.CodigoLocalEmbarque
            objConhecimento.EndLocalEmbarque = objConhecimento.NotaTrocaOrigem.EndLocalEmbarque
        Else
            If objConhecimento.NotaTrocaOrigem.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida Then
                objConhecimento.CodigoLocalEmbarque = objConhecimento.NotaTrocaOrigem.CodigoDeposito
                objConhecimento.EndLocalEmbarque = objConhecimento.NotaTrocaOrigem.EnderecoDeposito
            Else
                objConhecimento.CodigoLocalEmbarque = objConhecimento.NotaTrocaOrigem.CodigoCliente
                objConhecimento.EndLocalEmbarque = objConhecimento.NotaTrocaOrigem.EnderecoCliente
            End If
        End If

        objConhecimento.ObservacoesDeEmbarque = String.Format("{0} {1}", objConhecimento.NotaTrocaOrigem.Observacoes, objConhecimento.NotaTrocaOrigem.ObservacoesDeEmbarque).Trim()

        objConhecimento.CodigoMotorista = objConhecimento.NotaTrocaOrigem.CodigoMotorista
        objConhecimento.EnderecoMotorista = objConhecimento.NotaTrocaOrigem.EnderecoMotorista

        objConhecimento.CodigoTransportador = objConhecimento.NotaTrocaOrigem.CodigoTransportador
        objConhecimento.EnderecoTransportador = objConhecimento.NotaTrocaOrigem.EnderecoTransportador

        objConhecimento.PlacaDetalhes = objConhecimento.NotaTrocaOrigem.PlacaDetalhes

        If objConhecimento.PlacaDetalhes Is Nothing Then

            'Se não encontramos dados da placa, motorista... no XML, sistema vai extrair dados
            'da observação e encontrar uma placa cadastrada no sistema
            'caso não tenha o cadastro o CTe não será salvo e o usuario deve fazer o 
            'cadastro da placa
            Dim sPlaca As String = ExtrairPlacaVeiculo(objConhecimento.Observacoes)

            If sPlaca Is Nothing Then
                sPlaca = ExtrairPlacaVeiculo(objConhecimento.ObservacoesDeEmbarque)
            End If

            objConhecimento.PlacaDetalhes = New [Lib].Negocio.Placa(sPlaca)

            If String.IsNullOrEmpty(objConhecimento.PlacaDetalhes.Placa01) Then
                objConhecimento.PlacaDetalhes = New [Lib].Negocio.Placa
                objConhecimento.PlacaDetalhes.Placa01 = sPlaca
            Else
                objConhecimento.CodigoMotorista = objConhecimento.PlacaDetalhes.CpfMotorista
                objConhecimento.EnderecoMotorista = objConhecimento.PlacaDetalhes.EndCpfMotorista
            End If

        Else

            If String.IsNullOrEmpty(objConhecimento.CodigoMotorista) Then
                objConhecimento.CodigoMotorista = objConhecimento.PlacaDetalhes.CpfMotorista
            End If

        End If

        objConhecimento.PlacaTransportador = objConhecimento.PlacaDetalhes.Placa01

        objConhecimento.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
        objConhecimento.DataInclusao = Now()
        txtDataDeEmissao.Enabled = False
        txtDataDeSaida.Enabled = True

        SalvarSessaoConhecimento()

        CarregarCTE(True, False)

    End Sub

    Function ExtrairPlacaVeiculo(ByVal texto As String) As String
        Dim regexPlaca As New Regex("\b([A-Z]{3}[0-9][A-Z0-9][0-9]{2})\b|\b([A-Z]{3}-?[0-9]{4})\b", RegexOptions.IgnoreCase)
        Dim matches = regexPlaca.Matches(texto)

        For Each match As Match In matches
            If match.Success Then
                Return match.Value.ToUpper() ' <- retorna a primeira placa encontrada
            End If
        Next

        Return Nothing
    End Function

    Protected Sub CarregarCTE(ByVal ctrNovo As Boolean, Optional ByVal IsLoading As Boolean = False)
        RecuperarSessaoConhecimento()
        If ctrNovo Then

            If objConhecimento.CodigoTipoDeDocumento = 0 Then
                objConhecimento.CodigoTipoDeDocumento = getTipoDeDocumento()
            End If

            If objConhecimento.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida OrElse objConhecimento.EmpresaEmissor Then
                objConhecimento.Movimento = objConhecimento.DataNota
            Else
                objConhecimento.Movimento = DateTime.Now
            End If

            txtNumero.Enabled = Not String.IsNullOrWhiteSpace(hdfTipo.Value) AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Circulacao
            txtSerie.Enabled = Not String.IsNullOrWhiteSpace(hdfTipo.Value) AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Circulacao
            txtDataDeEmissao.Enabled = False
        End If

        If Not ctrNovo AndAlso objConhecimento.TipoDeDocumentoFrete IsNot Nothing Then
            hdfTipo.Value = objConhecimento.CodigoTipoDeDocumentoFrete
            If Not String.IsNullOrWhiteSpace(objConhecimento.CodigoFavorecido) AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.EnderecoFavorecido) Then
                Dim objFavorecido As New [Lib].Negocio.Cliente(objConhecimento.CodigoFavorecido, objConhecimento.EnderecoFavorecido)
                txtCodigoFavorecidoCartao.Value = objFavorecido.Codigo & "-" & objFavorecido.CodigoEndereco
                txtFavorecidoCartao.Text = objFavorecido.Nome
                txtEnderecoFavorecidoCartao.Text = objFavorecido.CodigoEndereco
                txtCidadeFavorecidoCartao.Text = objFavorecido.Cidade
                txtEstadoFavorecidoCartao.Text = objFavorecido.CodigoEstado
                txtCPFFavorecidoCartao.Text = objFavorecido.Codigo
            End If
        End If

        ucFile.Clear()
        If ucFile.Parent.Visible Then
            ucFile.Bind(objConhecimento.Arquivos)
        End If

        ddlSituacao.SelectedValue = objConhecimento.CodigoSituacao
        txtDataDeEmissao.Text = objConhecimento.DataNota.ToShortDateString()
        txtDataDeSaida.Text = objConhecimento.Movimento.ToShortDateString()
        txtNomeDaEmpresa.Text = objConhecimento.Empresa.Nome
        txtEnderecoDaEmpresa.Text = objConhecimento.Empresa.Endereco & IIf(objConhecimento.Empresa.Numero.ToString.Length > 0, ", " & objConhecimento.Empresa.Numero.ToString, "")
        txtComplementoDaEmpresa.Text = objConhecimento.Empresa.Complemento
        txtBairroDaEmpresa.Text = objConhecimento.Empresa.Bairro
        txtCepDaEmpresa.Text = objConhecimento.Empresa.CEP
        txtInscricaoDaEmpresa.Text = objConhecimento.Empresa.InscricaoEstadual
        txtCnpjDaEmpresa.Text = objConhecimento.Empresa.CodigoFormatado

        If Not IsLoading Then
            If CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Anulacao Then
                txtObservacao.Text = String.Format("{0} {1}", objConhecimento.Observacoes, objConhecimento.ObservacoesDeEmbarque).ToUpper().Trim()
            Else
                txtObservacao.Text = String.Format("ANULAÇÃO DO CONHECIMENTO DE TRANSPORTE N° {0}-{1}", objConhecimento.NotaTrocaOrigem.Codigo, objConhecimento.NotaTrocaOrigem.Serie)
            End If
        Else
            txtObservacao.Text = String.Format("{0} {1}", objConhecimento.Observacoes, objConhecimento.ObservacoesDeEmbarque).ToUpper().Trim()
        End If

        If objConhecimento.NotaTrocaOrigem Is Nothing AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Anulacao Then
            MsgBox(Me.Page, "Não foi possível encontrar a nota fiscal de origem do conhecimento de transporte selecionado!")
            If Not objConhecimento.IUD = "U" Then
                LimparCampos()
                Exit Sub
            End If
        End If

        If objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso objConhecimento.NotaTrocaOrigem.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida _
            AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.RPA Then

            txtNomeDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Empresa.Nome
            txtEnderecoDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Empresa.Endereco & IIf(objConhecimento.NotaTrocaOrigem.Empresa.Numero.ToString.Length > 0, ", " & objConhecimento.NotaTrocaOrigem.Empresa.Numero.ToString, "")
            txtComplementoDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Empresa.Complemento
            txtBairroDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Empresa.Bairro
            txtCepDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Empresa.CEP
            txtInscricaoDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Empresa.InscricaoEstadual
            txtCnpjDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Empresa.CodigoFormatado

        ElseIf objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.RPA Then

            txtNomeDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Nome
            txtEnderecoDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Endereco & IIf(objConhecimento.NotaTrocaOrigem.Cliente.Numero.ToString.Length > 0, ", " & objConhecimento.NotaTrocaOrigem.Cliente.Numero.ToString, "")
            txtComplementoDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Complemento
            txtBairroDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Bairro
            txtCepDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Cliente.CEP
            txtInscricaoDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Cliente.InscricaoEstadual
            txtCnpjDoRemetente.Text = objConhecimento.NotaTrocaOrigem.Cliente.CodigoFormatado

        Else
            txtNomeDoRemetente.Text = objConhecimento.Cliente.Nome
            txtEnderecoDoRemetente.Text = objConhecimento.Cliente.Endereco & IIf(objConhecimento.Cliente.Numero.ToString.Length > 0, ", " & objConhecimento.Cliente.Numero.ToString, "")
            txtComplementoDoRemetente.Text = objConhecimento.Cliente.Complemento
            txtBairroDoRemetente.Text = objConhecimento.Cliente.Bairro
            txtCepDoRemetente.Text = objConhecimento.Cliente.CEP
            txtInscricaoDoRemetente.Text = objConhecimento.Cliente.InscricaoEstadual
            txtCnpjDoRemetente.Text = objConhecimento.Cliente.CodigoFormatado
        End If

        If objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso objConhecimento.NotaTrocaOrigem.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida _
            AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.RPA Then
            txtNomeDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Nome
            txtCnpjDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.CodigoFormatado
            txtEnderecoDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Endereco & IIf(objConhecimento.NotaTrocaOrigem.Cliente.Numero.ToString.Length > 0, ", " & objConhecimento.NotaTrocaOrigem.Cliente.Numero.ToString, "")
            txtComplementoDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Complemento
            txtBairroDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Bairro
            txtCepDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.CEP
            txtCidadeDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Cidade
            txtTelefoneDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.Telefone
            txtEstadoDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.CodigoEstado
            txtInscricaoDoCliente.Text = objConhecimento.NotaTrocaOrigem.Cliente.InscricaoEstadual
        Else
            txtNomeDoCliente.Text = objConhecimento.Empresa.Nome
            txtCnpjDoCliente.Text = objConhecimento.Empresa.CodigoFormatado
            txtEnderecoDoCliente.Text = objConhecimento.Empresa.Endereco & IIf(objConhecimento.Empresa.Numero.ToString.Length > 0, ", " & objConhecimento.Empresa.Numero.ToString, "")
            txtComplementoDoCliente.Text = objConhecimento.Empresa.Complemento
            txtBairroDoCliente.Text = objConhecimento.Empresa.Bairro
            txtCepDoCliente.Text = objConhecimento.Empresa.CEP
            txtCidadeDoCliente.Text = objConhecimento.Empresa.Cidade
            txtTelefoneDoCliente.Text = objConhecimento.Empresa.Telefone
            txtEstadoDoCliente.Text = objConhecimento.Empresa.CodigoEstado
            txtInscricaoDoCliente.Text = objConhecimento.Empresa.InscricaoEstadual
        End If

        If objConhecimento.NotaTrocaOrigem IsNot Nothing Then
            grdNotasFretes.DataSource = objConhecimento.NotasTrocaOrigem
            grdNotasFretes.DataBind()
        End If

        If objConhecimento.LocalEmbarque IsNot Nothing Then
            txtEmbarque.Text = objConhecimento.LocalEmbarque.CodigoFormatado & " - " & objConhecimento.LocalEmbarque.Nome
            txtMunicipioEmbarque.Text = objConhecimento.LocalEmbarque.Cidade
            txtEstadoEmbarque.Text = objConhecimento.LocalEmbarque.CodigoEstado
        End If

        If objConhecimento.Destino IsNot Nothing Then
            txtNomeOrigemDestino.Text = objConhecimento.Destino.CodigoFormatado & " - " & objConhecimento.Destino.Nome
            txtMunicipioOrigemDestino.Text = objConhecimento.Destino.Cidade
            txtEstadoOrigemDestino.Text = objConhecimento.Destino.CodigoEstado
        End If

        If objConhecimento.Transportador IsNot Nothing Then
            txtNomeDoTransportador.Text = objConhecimento.Transportador.Nome
            txtRNTRCTransportador.Text = objConhecimento.Transportador.RNTRCTransportador
            txtEnderecoDoTransportador.Text = objConhecimento.Transportador.Endereco & IIf(objConhecimento.Transportador.Numero.ToString.Length > 0, ", " & objConhecimento.Transportador.Numero.ToString, "")
            txtCidadeDoTransportador.Text = objConhecimento.Transportador.Cidade
            txtEstadoDoTransportador.Text = objConhecimento.Transportador.CodigoEstado
            txtCnpjDoTransportador.Text = objConhecimento.Transportador.CodigoFormatado
            txtInscricaoDoTransportador.Text = objConhecimento.Transportador.InscricaoEstadual
        End If

        If objConhecimento.Tomador IsNot Nothing Then
            txtNomeDoTomador.Text = objConhecimento.Tomador.Nome
            txtEnderecoDoTomador.Text = objConhecimento.Tomador.Endereco & IIf(objConhecimento.Tomador.Numero.ToString.Length > 0, ", " & objConhecimento.Tomador.Numero.ToString, "")
            txtComplementoDoTomador.Text = objConhecimento.Tomador.Complemento
            txtBairroDoTomador.Text = objConhecimento.Tomador.Bairro
            txtCepDoTomador.Text = objConhecimento.Tomador.CEP
            txtInscricaoDoTomador.Text = objConhecimento.Tomador.InscricaoEstadual
            txtCnpjDoTomador.Text = objConhecimento.Tomador.CodigoFormatado
        End If

        If ctrNovo Then
            objConhecimento.NossaEmissao = getNossaEmissao()
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            txtNumero.Text = objConhecimento.Codigo
            txtSerie.Text = objConhecimento.Serie
        End If

        txtChaveNFe.Enabled = Not String.IsNullOrWhiteSpace(hdfTipo.Value) AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Circulacao AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Anulacao
        txtNumero.Enabled = Not String.IsNullOrWhiteSpace(hdfTipo.Value) AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Circulacao AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Anulacao
        txtSerie.Enabled = Not String.IsNullOrWhiteSpace(hdfTipo.Value) AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Circulacao AndAlso CType(hdfTipo.Value, eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Anulacao

        If ctrNovo Then
            lblUsuario.Text = objConhecimento.UsuarioInclusao
        Else
            lblUsuario.Text = IIf(objConhecimento.UsuarioAlteracao.Length = 0, objConhecimento.UsuarioInclusao, objConhecimento.UsuarioAlteracao)
        End If

        If objConhecimento.Motorista IsNot Nothing Then
            txtNomeDoMotorista.Text = objConhecimento.Motorista.Nome
            txtEnderecoDoMotorista.Text = objConhecimento.Motorista.Endereco & IIf(objConhecimento.Motorista.Numero.ToString.Length > 0, ", " & objConhecimento.Motorista.Numero.ToString, "")
            txtCidadeDoMotorista.Text = objConhecimento.Motorista.Cidade
            txtEstadoDoMotorista.Text = objConhecimento.Motorista.CodigoEstado
            txtCPFDoMotorista.Text = objConhecimento.Motorista.CodigoFormatado
            txtHabilitacaoDoMotorista.Text = objConhecimento.Motorista.Habilitacao
        End If

        If objConhecimento.PlacaDetalhes IsNot Nothing Then
            txtCPlaca1.Text = objConhecimento.PlacaDetalhes.Placa01
            txtCidadePlaca1.Text = objConhecimento.PlacaDetalhes.CidadePlaca01
            txtEstadoPlaca1.Text = objConhecimento.PlacaDetalhes.EstadoPlaca01
            txtRNTRCPlaca1.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca01
            txtProprietarioPlaca1.Text = objConhecimento.PlacaDetalhes.CodigoProprietario01

            If objConhecimento.PlacaDetalhes.Placa02.Length > 0 Then
                txtCPlaca2.Text = objConhecimento.PlacaDetalhes.Placa02
                txtCidadePlaca2.Text = objConhecimento.PlacaDetalhes.CidadePlaca02
                txtEstadoPlaca2.Text = objConhecimento.PlacaDetalhes.EstadoPlaca02
                txtRNTRCPlaca2.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca02
                txtProprietarioPlaca2.Text = objConhecimento.PlacaDetalhes.CodigoProprietario02
            End If

            If objConhecimento.PlacaDetalhes.Placa03.Length > 0 Then
                txtCPlaca3.Text = objConhecimento.PlacaDetalhes.Placa03
                txtCidadePlaca3.Text = objConhecimento.PlacaDetalhes.CidadePlaca03
                txtEstadoPlaca3.Text = objConhecimento.PlacaDetalhes.EstadoPlaca03
                txtRNTRCPlaca3.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca03
                txtProprietarioPlaca3.Text = objConhecimento.PlacaDetalhes.CodigoProprietario03
            End If
        End If

        Dim ds As New DataSet
        Dim dt As New DataTable("NotaFiscal")
        dt.Columns.Add("Nota", Type.GetType("System.String"))
        dt.Columns.Add("TipoDocumento", Type.GetType("System.String"))
        dt.Columns.Add("CodigoOperacao", Type.GetType("System.String"))
        dt.Columns.Add("Produto", Type.GetType("System.String"))
        dt.Columns.Add("CFOP", Type.GetType("System.String"))
        dt.Columns.Add("QuantidadeFisica", Type.GetType("System.String"))
        dt.Columns.Add("QuantidadeFiscal", Type.GetType("System.String"))
        dt.Columns.Add("ValorUnitario", Type.GetType("System.String"))
        dt.Columns.Add("ValorTotal", Type.GetType("System.String"))
        ds.Tables.Add(dt)

        Dim pesoFisico As Integer = 0
        Dim pesoFiscal As Integer = 0
        Dim valorTotal As Integer = 0

        If lstNotaFiscal Is Nothing Then
            lstNotaFiscal = New List(Of [Lib].Negocio.NotaFiscal)
        End If

        If IsLoading Then
            If objConhecimento.NotasOrigemDestino IsNot Nothing AndAlso objConhecimento.NotasOrigemDestino.Count > 0 Then
                lstNotaFiscal.Clear()
                For Each origem As [Lib].Negocio.NotaFiscal In objConhecimento.NotasOrigemDestino
                    lstNotaFiscal.Add(New [Lib].Negocio.NotaFiscal(origem))
                Next
            End If
        End If

        If Not SessaoDsXML Is Nothing And lstNotaFiscal.Count = 0 And Not objConhecimento.NotasTrocaOrigem Is Nothing Then
            For Each nfOrigem In objConhecimento.NotasTrocaOrigem
                lstNotaFiscal.Add(nfOrigem)
            Next
        End If

        'pesquisa a lista de notas selecionadas para exbir todos os produtos na gridview
        If lstNotaFiscal IsNot Nothing AndAlso lstNotaFiscal.Count > 0 Then
            For Each item As [Lib].Negocio.NotaFiscalXItem In lstNotaFiscal.SelectMany(Function(s) s.Itens).ToList
                Dim dr As DataRow = ds.Tables(0).NewRow()
                dr("Nota") = IIf(item.NotaFiscal.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida, "S", "E") & "-" & item.NotaFiscal.Codigo & "-" & item.NotaFiscal.Serie
                dr("TipoDocumento") = item.NotaFiscal.CodigoTipoDeDocumento & "-" & item.NotaFiscal.TipoDeDocumento.Descricao
                dr("CodigoOperacao") = item.CodigoOperacao & "-" & item.CodigoSubOperacao
                dr("Produto") = item.CodigoProduto & "-" & item.Produto.Nome
                dr("CFOP") = item.NotaFiscal.NaturezaDaOperacao
                dr("QuantidadeFisica") = item.QuantidadeFisica.ToString("N0")
                dr("QuantidadeFiscal") = item.QuantidadeFiscal.ToString("N0")
                pesoFisico += item.QuantidadeFisica
                pesoFiscal += item.QuantidadeFiscal
                dr("ValorUnitario") = item.Unitario.ToString("N2")
                dr("ValorTotal") = item.ValorTotal.ToString("N2")
                valorTotal += item.ValorTotal.ToString("N2")
                ds.Tables(0).Rows.Add(dr)
            Next
        End If


        '************
        'Soma as quantidades brutas de todos os Ctes já emitidos para essa NF
        Dim NFTotalConsumido As Decimal

        'Se não for importacao automatica
        If SessaoDsXML Is Nothing Then
            NFTotalConsumido = lstNotaFiscal.Sum(Function(s) s.NotasOrigemDestino.Sum(Function(t) t.PesoBruto AndAlso t.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC))
        End If

        'Calcula o saldo de quantidade bruta que pode ser consumido por este CTE

        Dim NFSaldoDisponivel As Decimal

        If getEntradaSaida() = eEntradaSaida.Saida Then
            NFSaldoDisponivel = pesoFisico - NFTotalConsumido
        Else
            NFSaldoDisponivel = pesoFiscal - NFTotalConsumido
        End If

        '***************

        gridNotasFiscais.DataSource = ds
        gridNotasFiscais.DataBind()

        'Carrega as Faturas de Frete'
        '-------------------------------------------------------------'
        Dim lst As New [Lib].Negocio.ListFaturasDeFretesXItens(objConhecimento)
        Dim faturas As New [Lib].Negocio.ListFaturaDeFrete(objConhecimento, lst)
        grdFaturas.DataSource = faturas
        grdFaturas.DataBind()


        If Not FinanceiroNovo Then
            grdVencimentosFNNovo.Visible = False
            grdVencimentos.Visible = True

            grdVencimentos.DataSource = From f In faturas.SelectMany(Function(s) s.ListTituloFatura).Where(Function(t) t.Titulo.CodigoSituacao = eSituacao.Normal)
                                        Select f.FaturaDeFrete.CodigoFatura, f.Titulo, CodigoCifrado = Funcoes.Cifrar(If(f.Titulo.ReceberPagar = "P", "ContasAPagar-" & f.Titulo.Codigo, "ContasAReceber-" & f.Titulo.Codigo))

            grdVencimentos.DataBind()
        Else
            grdVencimentosFNNovo.Visible = True
            grdVencimentos.Visible = False

            grdVencimentosFNNovo.DataSource = From f In faturas.SelectMany(Function(s) s.ListTituloFatura).Where(Function(t) t.TituloNovo.CodigoSituacao = eSituacao.Normal)
                                              Select f.FaturaDeFrete.CodigoFatura, f.TituloNovo
            grdVencimentosFNNovo.DataBind()
        End If
        '--------------------------------------------------------------'

        'Contabilização
        If objConhecimento.LancamentosContabeis IsNot Nothing Then
            'objConhecimento.LancamentosContabeis.CalcularSaldo()
            grdContabilizacao.DataSource = objConhecimento.LancamentosContabeis
            grdContabilizacao.DataBind()
        End If
        '--------------------------------------------------------------'

        If Not IsLoading And SessaoDsXML Is Nothing Then
            If objConhecimento.NotaTrocaOrigem.Itens.Any(Function(s) s.Produto.Agrupar = "N") _
                    OrElse objConhecimento.NotaTrocaOrigem.Itens.Any(Function(s) s.CodigoProduto = "502010001") _
                    OrElse objConhecimento.CodigoOperacao <> 80 Then

                If getEntradaSaida() = eEntradaSaida.Saida Then
                    txtQuantidade.Text = IIf(getTipoDeFrete() = eTipoDeDocumentoFrete.RPA, PesoTotalRPA, NFSaldoDisponivel.ToString("N0"))
                Else
                    txtQuantidade.Text = IIf(getTipoDeFrete() = eTipoDeDocumentoFrete.RPA, PesoTotalRPA, NFSaldoDisponivel.ToString("N0"))
                End If


            Else
                txtQuantidade.Text = "1"
            End If
        End If

        If Not ctrNovo Then
            gridEncargos.DataSource = objConhecimento.Itens(0).Encargos
            gridEncargos.DataBind()
        End If

        'cruzamento da operação e suboperação atráves do tipo de pessoa do cliente da nota, da classe da suboperação e se a suboperação é entrada ou saída
        objConhecimento.CarregandoNota = True
        If Not IsLoading Then
            Dim classe As String = IIf(objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso objConhecimento.NotaTrocaOrigem.SubOperacao IsNot Nothing,
                                        objConhecimento.NotaTrocaOrigem.SubOperacao.Classe.ToString(), String.Empty)
            Dim entradaSaida As String = IIf(getEntradaSaida() = eEntradaSaida.Saida, "S", "E")
            Dim tipo As Nullable(Of eTipoOperacao) = Nothing
            Dim fisicaJuridica As String = String.Empty
            Dim financeiro As String = String.Empty
            Dim devolucao As Boolean = False

            If classe.ToUpper().Trim().Contains("CONTAEORDEM") OrElse classe.ToUpper().Trim().Contains("REMESSAS") OrElse classe.ToUpper().Trim().Contains("AFIXAR") Then
                If objConhecimento.NotaTrocaOrigem.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                    classe = eClassesOperacoes.VENDAS.ToString
                Else
                    classe = eClassesOperacoes.COMPRAS.ToString
                End If
            End If

            If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
                fisicaJuridica = IIf(objConhecimento.CodigoCliente.IsCnpj(), "J", "F")
            Else
                If objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.NotaTrocaOrigem.CodigoTransportador) Then
                    fisicaJuridica = IIf(objConhecimento.NotaTrocaOrigem.CodigoTransportador.IsCnpj(), "J", "F")
                End If
            End If

            If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
                financeiro = "S"
            ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro Then
                financeiro = "S"
            ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro Then
                financeiro = "N"
            End If

            If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
                classe = eClassesOperacoes.FRETES.ToString
                tipo = eTipoOperacao.Estadia
            ElseIf getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
                classe = eClassesOperacoes.FRETES.ToString
                tipo = eTipoOperacao.PrestacaoServico
                financeiro = "S"
            End If

            If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
                devolucao = False
            Else
                devolucao = objConhecimento.NotaTrocaOrigem.SubOperacao.Devolucao
            End If

            Dim objOperacaoDeFrete As [Lib].Negocio.OperacoesDeFretes = Nothing
            If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
                objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(objConhecimento.NotaTrocaOrigem.CodigoOperacao, objConhecimento.NotaTrocaOrigem.CodigoSubOperacao)
            ElseIf CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento Then
                objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(objConhecimento.NotaTrocaOrigem.CodigoOperacao, objConhecimento.NotaTrocaOrigem.CodigoSubOperacao)
            Else

                Dim pFinalidade As Integer = 0

                'Quando for Complemento de Frete é necessário buscar o cte de origem, depois a nota de origem e finalmente o pedido.
                If CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento _
                    OrElse CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
                    pFinalidade = objConhecimento.NotaTrocaOrigem.NotaTrocaOrigem.CodigoFinalidade
                Else
                    pFinalidade = objConhecimento.NotaTrocaOrigem.CodigoFinalidade
                End If

                'Até o momento apenas as finalidades 3 e 7 são usadas na busca por operações de frete
                'Então o que for diferente disso a finalidade será 0
                If pFinalidade <> 3 OrElse pFinalidade <> 7 Then
                    pFinalidade = 0
                End If

                'If entradaSaida = "S" Then
                '    objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(
                '        entradaSaida,                                              ' Entrada/Saída do documento
                '        objConhecimento.EmpresaEmissor,                            ' Emissor do CTe é a propria Empresa
                '        objConhecimento.Tomador.Estado.Codigo = objConhecimento.Empresa.Estado.Codigo, ' Tomador e Empresa emissor na mesma UF?
                '        objConhecimento.NotaDeTroca.Empresa.Estado.Codigo = objConhecimento.Empresa.Estado.Codigo, ' UF da nota de origem é o mesmo da empresa emissor                   
                '        objConhecimento.NotaDeTroca.Empresa.Estado.Codigo = objConhecimento.Destino.Estado.Codigo ' UF da nota de origem é o mesmo do destino da nota origem
                '    )
                'Else
                '    objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(classe, entradaSaida, fisicaJuridica, financeiro, tipo, pFinalidade, devolucao, objConhecimento.Cliente.Estado.Codigo, objConhecimento.Cliente.Estado.Regiao)
                'End If

                objOperacaoDeFrete = New [Lib].Negocio.OperacoesDeFretes(
                        entradaSaida,                                              ' Entrada/Saída do documento
                        objConhecimento.EmpresaEmissor,                            ' Emissor do CTe é a propria Empresa
                        objConhecimento.Tomador.Estado.Codigo = objConhecimento.NotaDeTroca.Empresa.Estado.Codigo, ' Tomador e Empresa emissor na mesma UF?
                        IIf(objConhecimento.NotaDeTroca.EntradaSaida = eEntradaSaida.Entrada, objConhecimento.NotaDeTroca.Cliente.Estado.Codigo, objConhecimento.NotaDeTroca.Empresa.Estado.Codigo) = objConhecimento.Empresa.Estado.Codigo, ' UF da nota de origem é o mesmo da empresa emissor                   
                        objConhecimento.NotaDeTroca.Empresa.Estado.Codigo = objConhecimento.Destino.Estado.Codigo ' UF da nota de origem é o mesmo do destino da nota origem
                    )

            End If

            If (objOperacaoDeFrete IsNot Nothing AndAlso objOperacaoDeFrete.OperacaoId > 0 AndAlso objOperacaoDeFrete.SubOperacaoId > 0) Then
                If ((CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                            OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro _
                            OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento _
                            OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia) _
                        AndAlso objOperacaoDeFrete.OpDestinoId IsNot Nothing AndAlso objOperacaoDeFrete.SubOpDestinoId IsNot Nothing) Then
                    objConhecimento.CodigoOperacao = objOperacaoDeFrete.OpDestinoId
                    objConhecimento.CodigoSubOperacao = objOperacaoDeFrete.SubOpDestinoId
                ElseIf (CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao _
                        AndAlso objOperacaoDeFrete.OpAnulacao IsNot Nothing AndAlso objOperacaoDeFrete.SubOpAnulacao IsNot Nothing) Then
                    objConhecimento.CodigoOperacao = objOperacaoDeFrete.OpAnulacao
                    objConhecimento.CodigoSubOperacao = objOperacaoDeFrete.SubOpAnulacao
                Else
                    objConhecimento.CodigoOperacao = objOperacaoDeFrete.OperacaoId
                    objConhecimento.CodigoSubOperacao = objOperacaoDeFrete.SubOperacaoId
                End If

                If String.IsNullOrWhiteSpace(objConhecimento.CodigoTransportador) Then
                    btnTransportador.Enabled = True
                End If
            Else
                Dim msg As String = String.Format("Não foi possível encontrar uma operação de frete para a nota fiscal selecionada, com os seguintes parâmetros: CLASSE: {0} - CPF/CNPJ: {1} - E/S: {2} - FINANCEIRO: {3}!", classe, fisicaJuridica, entradaSaida, financeiro)
                If (CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                        OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento _
                        OrElse CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro) _
                        AndAlso Not objConhecimento.NotaTrocaOrigem.CodigoTransportador.IsCnpj() Then
                    MsgBox(Me.Page, msg)
                    btnTransportador.Enabled = True
                Else
                    MsgBox(Me.Page, msg)
                    LimparCampos()
                    Exit Sub
                End If
            End If
        End If

        If CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
            If Not IsLoading Then
                objConhecimento.Itens = New [Lib].Negocio.ListNotaFiscalXItem(objConhecimento.NotaTrocaOrigem)
                If objConhecimento.Itens IsNot Nothing AndAlso objConhecimento.Itens.Count > 0 Then
                    For Each item As [Lib].Negocio.NotaFiscalXItem In objConhecimento.Itens
                        item.NotaFiscal.Codigo = objConhecimento.Codigo
                        item.NotaFiscal.Serie = objConhecimento.Serie
                        item.NotaFiscal.EntradaSaida = objConhecimento.EntradaSaida
                        For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In item.Encargos
                            enc.ItemNotaFiscal.NotaFiscal.Codigo = objConhecimento.Codigo
                            enc.ItemNotaFiscal.NotaFiscal.Serie = objConhecimento.Serie
                            enc.ItemNotaFiscal.NotaFiscal.EntradaSaida = objConhecimento.EntradaSaida
                        Next
                    Next

                    If getEntradaSaida() = eEntradaSaida.Saida Then
                        txtQuantidade.Text = objConhecimento.Itens(0).QuantidadeFisica.ToString("N0")
                    Else
                        txtQuantidade.Text = objConhecimento.Itens(0).QuantidadeFiscal.ToString("N0")
                    End If

                    txtUnitario.Text = CDec(objConhecimento.Itens(0).Unitario * 1000).ToString("N4")
                    txtValor.Text = objConhecimento.Itens(0).ValorTotal.ToString("N2")
                    getEncargos()
                End If
            End If
            rdoBanco.Enabled = False
            rdoPamcard.Enabled = False
            txtUnitario.Enabled = False
            'lnkNovo.Parent.Visible = True
        ElseIf CType(CInt(ddlTipoConhecimento.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA AndAlso CDec(txtQuantidade.Text) > 0 Then
            txtValor.Text = objConhecimento.ValorFrete
            txtUnitario.Text = Math.Round(objConhecimento.ValorFrete / CDec(txtQuantidade.Text) * 1000, 6, MidpointRounding.AwayFromZero)
            txtUnitario_TextChanged(txtUnitario, New EventArgs())
            txtValor.Enabled = False
            txtUnitario.Enabled = False
        End If

        If objConhecimento.CFOP IsNot Nothing AndAlso objConhecimento.CFOP.Codigo > 0 Then
            txtCfop.Text = String.Format("{0} - {1}", objConhecimento.CFOP.Codigo, objConhecimento.CFOP.Descricao)
        End If

        If objConhecimento.CodigoOperacao > 0 AndAlso objConhecimento.CodigoSubOperacao > 0 Then
            Dim op As New SubOperacao(objConhecimento.CodigoOperacao, objConhecimento.CodigoSubOperacao)

            txtOperacao.Text = String.Format("{0}-{1} - {2}", objConhecimento.CodigoOperacao, objConhecimento.CodigoSubOperacao, IIf(op IsNot Nothing, op.Descricao, String.Empty))

            If objConhecimento.Itens.Count > 0 Then
                txtEntSai.Text = String.Format("{0} - {1}", IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA"), "ID OperaçãoXEncargos: " & objConhecimento.Itens.Where(Function(s) s.IUD <> "D")(0).CodigoOperacaoEstado)
            End If

            rdoBanco.Enabled = op.Financeiro
            rdoPamcard.Enabled = op.Financeiro
        End If

        If objConhecimento.NotaTrocaOrigem IsNot Nothing Then
            objConhecimento.CodigoFinalidade = objConhecimento.NotaTrocaOrigem.CodigoFinalidade
            'objConhecimento.CodigoPedido = objConhecimento.NotaTrocaOrigem.CodigoPedido
        End If

        objConhecimento.CarregandoNota = False
        SalvarSessaoConhecimento()

        If (rdoFerroviario.Checked AndAlso grdNFeSaldo.Visible) OrElse
            (objConhecimento.NotaTrocaOrigem IsNot Nothing AndAlso objConhecimento.NotaTrocaOrigem.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) _
                AndAlso objConhecimento.NotaTrocaOrigem.CodigoOperacao <> 80) Then
            btnTransportador.Enabled = True
        End If

        If (CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro OrElse
            CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro OrElse
            CType(hdfTipo.Value, eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento) AndAlso Not objConhecimento.IUD = "U" Then
            txtQuantidade.Enabled = True
        End If

        If objConhecimento.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
            divEmitente.Visible = True

            txtNomeDoEmitente.Text = objConhecimento.Cliente.Nome
            txtEnderecoDoEmitente.Text = objConhecimento.Cliente.Endereco & IIf(objConhecimento.Cliente.Numero.ToString.Length > 0, ", " & objConhecimento.Cliente.Numero.ToString, "")
            txtComplementoDoEmitente.Text = objConhecimento.Cliente.Complemento
            txtBairroDoEmitente.Text = objConhecimento.Cliente.Bairro
            txtCepDoEmitente.Text = objConhecimento.Cliente.CEP
            txtInscricaoDoEmitente.Text = objConhecimento.Cliente.InscricaoEstadual
            txtCnpjDoEmitente.Text = objConhecimento.Cliente.CodigoFormatado
        Else
            divEmitente.Visible = False
        End If

        txtObservacao.Height = Unit.Pixel(170)
        'lnkNovo.Parent.Visible = Not objConhecimento.IUD = "U"
        lnkConsultarSefaz.Parent.Visible = True

        'txtDataDeSaida.Enabled = Not objConhecimento.IUD = "U"
        lnkEspelho.Parent.Visible = True
        lnkEnviarSEFAZ.Parent.Visible = Not String.IsNullOrWhiteSpace(objConhecimento.IUD) AndAlso objConhecimento.IUD <> "I" AndAlso objConhecimento.NossaEmissao _
            AndAlso String.IsNullOrWhiteSpace(objConhecimento.ChaveNFE) AndAlso String.IsNullOrWhiteSpace(objConhecimento.ProtocoloNota) _
            AndAlso (String.IsNullOrWhiteSpace(objConhecimento.Retorno) OrElse (Not String.IsNullOrWhiteSpace(objConhecimento.Retorno) AndAlso objConhecimento.Retorno <> "100"))
        lnkDACTE.Parent.Visible = Not String.IsNullOrWhiteSpace(objConhecimento.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.ProtocoloNota) AndAlso objConhecimento.NossaEmissao
        lnkEnviarEmail.Parent.Visible = Not String.IsNullOrWhiteSpace(objConhecimento.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.ProtocoloNota) AndAlso objConhecimento.NossaEmissao
        lnkDuplicar.Parent.Visible = Not IsDuplicar(objConhecimento) AndAlso objConhecimento.IUD <> "I" AndAlso objConhecimento.NossaEmissao AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao
        lnkExcel.Parent.Visible = objConhecimento IsNot Nothing AndAlso objConhecimento.IUD <> "I" AndAlso objConhecimento.NossaEmissao AndAlso objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao
        lnkVisualizar.Parent.Parent.Parent.Visible = IsEnabled(objConhecimento) AndAlso objConhecimento.NossaEmissao
        btnCadastro.Enabled = False
        btnEmpresa.Enabled = True
        btnTransportador.Enabled = True
        tcPrincipal.ActiveTab = tabCTRC

    End Sub

    Public Sub RemoverArquivoBD()
        RecuperarSessaoConhecimento()

        If Not Funcoes.VerificaAcesso(objConhecimento.CodigoEmpresa, objConhecimento.EnderecoEmpresa, objConhecimento.Movimento, "NOTAS FISCAIS") Then
            MsgBox(Me.Page, "Movimento da Nota Fiscal já fechado para esta data", eTitulo.Info)
        Else
            If objConhecimento.IUD = "U" Then
                ucFile.Salvar(objConhecimento.Arquivos)

                Dim Sqls As New ArrayList

                For Each arq In objConhecimento.Arquivos
                    If arq.IUD = "D" Then
                        arq.CodigoEmpresa = objConhecimento.CodigoEmpresa
                        arq.EnderecoEmpresa = objConhecimento.EnderecoEmpresa
                        arq.CodigoCliente = objConhecimento.CodigoCliente
                        arq.EnderecoCliente = objConhecimento.EnderecoCliente
                        arq.CodigoNota = objConhecimento.Codigo
                        arq.Serie = objConhecimento.Serie
                        arq.CodigoPedido = objConhecimento.CodigoPedido
                        arq.SalvarSql(Sqls)
                    End If
                Next

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        SalvarSessaoConhecimento()
                        MsgBox(Me.Page, "Arquivo removido com sucesso. Caso tenha consultado a nota apenas para remover o arquivo, não precisa clicar em alterar pois o mesmo já foi removido.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub BindGridView()
        hdfTipo.Value = getTipoDeFrete()
        If getTipoDeFrete() = eTipoDeDocumentoFrete.RPA Then
            'divPaineltotais.Style.Clear()
            divClienteRPA.Style.Clear()
            divVencimentoFatura.Style.Clear()
            If String.IsNullOrWhiteSpace(txtPedido.Text) Then
                txtPedido.Focus()
            End If
        Else
            divPaineltotais.Style.Add("display", "none")
            divClienteRPA.Style.Add("display", "none")
            divVencimentoFatura.Style.Add("display", "none")
        End If
        grdNFe.DataSource = New List(Of Object)
        grdNFe.DataBind()
    End Sub

    Private Sub getEncargos()

        If (Not String.IsNullOrWhiteSpace(txtUnitario.Text)) AndAlso CDec(txtUnitario.Text) > 0 Then

            RecuperarSessaoConhecimento()

            If objConhecimento IsNot Nothing Then

                objConhecimento.NossaEmissao = getNossaEmissao()
                SalvarSessaoConhecimento()

                'Como o Cte pode ser feito a partir de mais de uma nota este servirá para somar suas quantidades já que será utilizado na validação em seguida.
                Dim PesoTotalNFsTrocaOrigem As Decimal = objConhecimento.NotasTrocaOrigem.Sum(Function(s) s.TotalQuantidadeFiscal)
                Dim NFTotalConsumido As Decimal

                'Se não for importacao automatica
                If SessaoDsXML Is Nothing Then
                    NFTotalConsumido = lstNotaFiscal.Sum(Function(s) s.NotasOrigemDestino.Sum(Function(t) t.PesoBruto AndAlso t.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC))
                End If

                lstNotaFiscal.Sum(Function(s) s.NotasOrigemDestino.Sum(Function(t) t.PesoBruto AndAlso t.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC))
                Dim NFSaldoDisponivel As Decimal = PesoTotalNFsTrocaOrigem - NFTotalConsumido


                If (CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro OrElse
                    CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro) AndAlso
                    CDec(txtQuantidade.Text) > NFSaldoDisponivel Then
                    MsgBox(Me.Page, "Quantidade informada não pode ser maior que a Quantidade Fiscal da Nota!")
                    Exit Sub
                End If

                If String.IsNullOrWhiteSpace(objConhecimento.CodigoCliente) AndAlso CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Anulacao Then
                    txtUnitario.Text = 0
                    MsgBox(Me.Page, "É necessário informar o transportador do conhecimento de transporte!")
                    Exit Sub
                End If

                Dim objProduto As [Lib].Negocio.Produto = Nothing
                Dim objClienteXEmpresa As New [Lib].Negocio.ClienteXEmpresa(objConhecimento.CodigoEmpresa, objConhecimento.EnderecoEmpresa)

                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Estadia Then
                    If objClienteXEmpresa Is Nothing OrElse String.IsNullOrWhiteSpace(objClienteXEmpresa.CodigoProdutoDeEstadia) Then
                        MsgBox(Me.Page, "É necessário informar o produto de estadia no cadastro de empresas!")
                        Exit Sub
                    End If
                    objProduto = New [Lib].Negocio.Produto(objClienteXEmpresa.CodigoProdutoDeEstadia) 'PRODUTO ESTADIA
                Else
                    If objClienteXEmpresa Is Nothing OrElse String.IsNullOrWhiteSpace(objClienteXEmpresa.CodigoProdutoDeFrete) Then
                        MsgBox(Me.Page, "É necessário informar o produto de frete no cadastro de empresas!")
                        Exit Sub
                    End If
                    objProduto = New [Lib].Negocio.Produto(objClienteXEmpresa.CodigoProdutoDeFrete) 'PRODUTO FRETE
                End If

                Dim cliente As String = objConhecimento.CodigoTransportador
                Dim endereco As Integer = objConhecimento.EnderecoTransportador

                If objConhecimento.XMLImportado = False Then
                    If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
                        objConhecimento.CodigoCliente = objConhecimento.NotaTrocaOrigem.CodigoEmpresa
                        objConhecimento.EnderecoCliente = objConhecimento.NotaTrocaOrigem.EnderecoEmpresa
                    ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Anulacao Then
                        objConhecimento.CodigoCliente = objConhecimento.NotaTrocaOrigem.CodigoCliente
                        objConhecimento.EnderecoCliente = objConhecimento.NotaTrocaOrigem.EnderecoCliente
                    End If
                End If

                If SessaoDsXML Is Nothing Then

                    If objConhecimento.NotaTrocaOrigem.Itens.Any(Function(s) s.CodigoProduto = objClienteXEmpresa.CodigoProdutoDeFrete) _
                                        OrElse objProduto.Codigo = objClienteXEmpresa.CodigoProdutoDeFrete _
                                        OrElse objProduto.Codigo = objClienteXEmpresa.CodigoProdutoDeEstadia _
                                        OrElse objConhecimento.CodigoOperacao <> 80 Then

                        txtValor.Text = ((CInt(txtQuantidade.Text) * CDec(txtUnitario.Text)) / 1000).ToString("N2")
                    Else
                        txtValor.Text = (CInt(txtQuantidade.Text) * CDec(txtUnitario.Text)).ToString("N2")
                    End If

                End If

                Dim objConhecimentoXItem As New [Lib].Negocio.NotaFiscalXItem(objConhecimento)
                Dim peso As Decimal = CDec(txtQuantidade.Text)
                Dim valor As Decimal = CDec(txtValor.Text)

                With objConhecimentoXItem
                    .CodigoProduto = objProduto.Codigo
                    '.CodigoPedido = objConhecimento.CodigoPedido
                    .PesoQuantidade = objProduto.PesoQuantidade
                    .QuantidadeFiscal = peso
                    .QuantidadeFisica = peso
                    .Unitario = CDec(txtUnitario.Text)
                    .NotaFiscal.CarregandoItens = True
                    .ValorTotal = valor
                End With

                objConhecimentoXItem.UsarRegiao = False
                objConhecimentoXItem.CodigoOperacao = objConhecimento.CodigoOperacao
                objConhecimentoXItem.CodigoSubOperacao = objConhecimento.CodigoSubOperacao

                Dim fCliente As String = objConhecimento.CodigoCliente
                Dim fEndCliente As Integer = objConhecimento.EnderecoCliente

                objConhecimento.CodigoCliente = objConhecimento.CodigoTomador
                objConhecimento.EnderecoCliente = objConhecimento.EnderecoTomador

                objConhecimentoXItem.CarregandoEncargos = True
                objConhecimentoXItem.Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objConhecimentoXItem, New List(Of eEtapaEncago) From {eEtapaEncago.Normal, eEtapaEncago.Adiantamento})

                'Se o encargo for zero, vamos tentar encontrar na regiao
                If objConhecimentoXItem.Encargos.Count() = 0 Then
                    objConhecimentoXItem.UsarRegiao = True
                    objConhecimentoXItem.Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objConhecimentoXItem, New List(Of eEtapaEncago) From {eEtapaEncago.Normal, eEtapaEncago.Adiantamento})
                End If

                objConhecimentoXItem.CarregandoEncargos = False

                objConhecimento.CodigoCliente = fCliente
                objConhecimento.EnderecoCliente = fEndCliente

                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
                    objConhecimento.CodigoTransportador = cliente
                    objConhecimento.EnderecoTransportador = endereco
                End If

                'If Not objConhecimentoXItem.Encargos IsNot Nothing AndAlso objConhecimentoXItem.Encargos.Count > 0 Then
                '    txtUnitario.Text = 0
                '    MsgBox(Me.Page, "Não foi possível encontrar encargos para a operação " & objConhecimento.CodigoOperacao & " e suboperação " & objConhecimento.CodigoSubOperacao & "!")
                '    Exit Sub
                'End If
                If objConhecimentoXItem.Encargos Is Nothing OrElse objConhecimentoXItem.Encargos.Count = 0 Then
                    txtUnitario.Text = 0
                    MsgBox(Me.Page, "Não foi possível encontrar encargos para a operação " & objConhecimento.CodigoOperacao & " e suboperação " & objConhecimento.CodigoSubOperacao & "!")
                    Exit Sub
                End If

                objConhecimento.CarregandoItens = True
                objConhecimento.Itens.Clear()
                objConhecimento.Itens.Add(objConhecimentoXItem)
                objConhecimento.AtualizaTotais()
                objConhecimento.CarregandoItens = False

                If objConhecimento.Itens IsNot Nothing AndAlso objConhecimento.Itens.Count > 0 Then
                    Dim cfop As New [Lib].Negocio.CFOP(objConhecimento.Itens(0).CFOP)
                    txtCfop.Text = String.Format("{0} - {1}", cfop.Codigo, cfop.Descricao)

                    If rdoFerroviario.Checked AndAlso grdNFeSaldo.Visible Then
                        If Not (objConhecimento.NotasReferenciais IsNot Nothing AndAlso objConhecimento.NotasReferenciais.Count > 0) Then
                            objConhecimento.NotasReferenciais = New ListNotaFiscalReferencial()
                        End If

                        For Each item As [Lib].Negocio.NotaFiscalXItem In lstNotaFiscal.SelectMany(Function(s) s.Itens).ToList
                            Dim nfr As New NotaFiscalReferencial(objConhecimento.Itens(0))
                            nfr.Empresa_Id = item.NotaFiscal.CodigoEmpresa
                            nfr.EndEmpresa_Id = item.NotaFiscal.EnderecoEmpresa
                            nfr.Cliente_Id = item.NotaFiscal.CodigoCliente
                            nfr.EndCliente_Id = item.NotaFiscal.EnderecoCliente
                            nfr.EntradaSaida_Id = item.NotaFiscal.EntradaSaida
                            nfr.Serie_Id = item.NotaFiscal.Serie
                            nfr.Nota_Id = item.NotaFiscal.Codigo
                            nfr.Produto_Id = item.CodigoProduto
                            nfr.CFOP_Id = item.CFOP
                            nfr.Sequencia_Id = item.Sequencia
                            nfr.TipoReferencial_Id = "CTE"
                            nfr.Quantidade = item.QuantidadeFiscal
                            nfr.Valor = item.QuantidadeFiscal * item.Unitario
                            objConhecimento.NotasReferenciais.Add(nfr)
                        Next
                    End If
                End If

                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                        OrElse CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro Then
                    'TOMADOR DA NOTA FISCAL (REMETENTE) X TRANSPORTADORA
                    If Not objConhecimento.Transportador Is Nothing AndAlso objConhecimento.NotaTrocaOrigem.Empresa.Estado.Codigo.Trim() = objConhecimento.Transportador.Estado.Codigo.Trim() Then
                        Dim prod As [Lib].Negocio.NotaFiscalXItemXEncargo = objConhecimentoXItem.Encargos.Where(Function(s) s.Codigo.Trim.Contains("PRODUTO")).FirstOrDefault()
                        If prod IsNot Nothing AndAlso prod.EstadoDestino.Trim() <> objConhecimento.NotaTrocaOrigem.Empresa.Estado.Codigo.Trim() Then
                            For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objConhecimentoXItem.Encargos
                                enc.EstadoDestino = objConhecimento.NotaTrocaOrigem.Empresa.Estado.Codigo.Trim()
                            Next
                            If objConhecimento.Itens IsNot Nothing AndAlso objConhecimento.Itens.Count > 0 Then
                                Dim numCfop As Integer = objConhecimento.Itens(0).CFOP - 1000
                                objConhecimento.Itens(0).CFOP = numCfop
                                Dim cfop As New [Lib].Negocio.CFOP(numCfop)
                                txtCfop.Text = String.Format("{0} - {1}", cfop.Codigo, cfop.Descricao)
                            End If
                        End If
                    End If
                End If

                If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao Then
                    Dim vlrAdto As Decimal = Decimal.Zero
                    Dim lstEncargos As New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objConhecimento.NotaTrocaOrigem.Itens(0), New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In lstEncargos.Where(Function(s) Not s.Codigo.Trim.Contains("PRODUTO") AndAlso Not s.Codigo.Trim.Contains("LIQUIDO"))
                        If enc.Codigo.Trim.Contains("ADIANTAMENTO") Then
                            vlrAdto = enc.Valor
                            Exit For
                        End If
                    Next

                    For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In objConhecimento.Itens(0).Encargos.Where(Function(s) Not s.Codigo.Trim.Contains("PRODUTO") AndAlso Not s.Codigo.Trim.Contains("LIQUIDO"))
                        If enc.Codigo.Trim.Contains("ADIANTAMENTO") Then
                            enc.Valor = Decimal.Zero
                        ElseIf enc.Codigo.Trim.Contains("AMORTIZAADTO") Then
                            enc.Valor = vlrAdto
                        End If

                        If Not enc.Codigo.Trim.Contains("ADIANTAMENTO") AndAlso Not enc.Codigo.Trim.Contains("ADTODEFRETE") Then
                            Dim x As [Lib].Negocio.NotaFiscalXItemXEncargo = enc
                            Dim objEncargo As [Lib].Negocio.NotaFiscalXItemXEncargo = lstEncargos.Where(Function(s) x.Codigo.Trim.Equals(s.Codigo.Trim)).FirstOrDefault()
                            If objEncargo IsNot Nothing Then
                                enc.Valor = objEncargo.Valor
                            End If
                        End If
                    Next
                End If

                gridEncargos.DataSource = objConhecimento.Itens(0).Encargos
                gridEncargos.DataBind()

                Dim i As Integer = 0
                While i < gridEncargos.Rows.Count
                    If gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("ADIANTAMENTO") _
                            OrElse gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("ADTODEFRETE") _
                            OrElse gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("PEDAGIO") _
                            OrElse gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("TARIFA") _
                            OrElse gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("TAXA") Then
                        If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao _
                                AndAlso Not gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("ADTODEFRETE") _
                                AndAlso Not gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("TAXA PEDAGIO") Then
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = True
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).ForeColor = Drawing.Color.Red
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text = String.Format("{0:N2}", Decimal.Zero)
                        ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.Comprovacao _
                                AndAlso CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) <> eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro Then
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text = String.Format("{0:N2}", Decimal.Zero)
                        ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Comprovacao _
                                AndAlso gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("ADTODEFRETE") Then
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text = String.Format("{0:N2}", Decimal.Zero)
                        ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro _
                                AndAlso gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("ADTODEFRETE") Then
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = True
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).ForeColor = Drawing.Color.Red
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text = String.Format("{0:N2}", Decimal.Zero)
                        End If
                    ElseIf CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.RPA _
                            AndAlso (gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("IRRF PF") _
                            OrElse gridEncargos.Rows(i).Cells(0).Text.Trim.Contains("INSS")) Then
                        CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = True
                        CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).ForeColor = Drawing.Color.Red
                        If String.IsNullOrWhiteSpace(CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text) Then
                            CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Text = String.Format("{0:N2}", Decimal.Zero)
                        End If
                    Else
                        CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).Enabled = False
                        CType(gridEncargos.Rows(i).FindControl("txtValorEncargo"), TextBox).ForeColor = Drawing.Color.Black
                    End If
                    i += 1
                End While
                Calcular()
                'lnkNovo.Parent.Visible = True
                SalvarSessaoConhecimento()

                txtEntSai.Text = String.Format("{0} - {1}", IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "ENTRADA", "SAÍDA"), "ID OperaçãoXEncargos: " & objConhecimento.Itens.Where(Function(s) s.IUD <> "D")(0).CodigoOperacaoEstado)

                If (CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosComFinanceiro OrElse
                    CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro) AndAlso
                    CDec(txtQuantidade.Text) < NFSaldoDisponivel AndAlso
                    Not chkNFxCtes.Checked Then 'objConhecimento.NotaTrocaOrigem.Itens(0).QuantidadeFiscal Then
                    MsgBox(Me.Page, "Atenção, Quantidade informada é menor que a Quantidade Fiscal da Nota, certifique-se que a Transportadora vai mandar um Complemento ou entre em contato com o responsável!")
                End If
            End If
        End If
    End Sub

    Private Sub Calcular()
        RecuperarSessaoConhecimento()
        Dim vlr As Decimal = Decimal.Zero
        Dim vlrLiquido As Decimal = Decimal.Zero
        Dim vlrPedagio As Decimal = Decimal.Zero

        For Each row As GridViewRow In gridEncargos.Rows
            If row.Cells(0).Text.Contains("PRODUTO") Then
                Dim txtValorEncargo As TextBox = CType(row.FindControl("txtValorEncargo"), TextBox)
                If Not String.IsNullOrWhiteSpace(txtValorEncargo.Text) Then
                    vlrLiquido = CDec(txtValorEncargo.Text)
                End If
            End If

            If row.Cells(0).Text.Contains("TARIFA PEDAGIO") Then
                Dim txtValorEncargo As TextBox = CType(row.FindControl("txtValorEncargo"), TextBox)
                If Not String.IsNullOrWhiteSpace(txtValorEncargo.Text) Then
                    vlrPedagio = CDec(txtValorEncargo.Text)
                End If
            End If
        Next

        For Each row As GridViewRow In gridEncargos.Rows
            If (Not String.IsNullOrWhiteSpace(row.Cells(0).Text)) AndAlso (row.Cells(0).Text.Contains("ADTODEFRETE") OrElse
                        row.Cells(0).Text.Contains("TARIFA PEDAGIO") OrElse row.Cells(0).Text.Contains("TARIFA SEGURO") OrElse row.Cells(0).Text.Contains("TAXA CADASTRO")) Then
                Dim txtValorEncargo As TextBox = CType(row.FindControl("txtValorEncargo"), TextBox)
                Dim lblSinal As Label = CType(row.FindControl("lblSinal"), Label)
                Dim vlrEncargo As Decimal = CDec(txtValorEncargo.Text)
                If (lblSinal.Text.Trim().Equals("+")) Then
                    vlrLiquido += vlrEncargo
                ElseIf (lblSinal.Text.Trim().Equals("-")) Then
                    vlrLiquido -= vlrEncargo
                End If
            End If
        Next

        For Each row As GridViewRow In gridEncargos.Rows
            If Not String.IsNullOrWhiteSpace(row.Cells(0).Text) AndAlso row.Cells(0).Text.Contains("LIQUIDO") Then
                Dim txtValorEncargo As TextBox = CType(row.FindControl("txtValorEncargo"), TextBox)
                txtValorEncargo.Text = String.Format("{0:N2}", vlrLiquido)
            ElseIf Not String.IsNullOrWhiteSpace(row.Cells(0).Text) AndAlso row.Cells(0).Text.Contains("TAXA PEDAGIO") Then
                Dim txtValorEncargo As TextBox = CType(row.FindControl("txtValorEncargo"), TextBox)
                txtValorEncargo.Text = String.Format("{0:N2}", vlrPedagio * (0.5 / 100))
            End If
        Next

        If objConhecimento IsNot Nothing AndAlso objConhecimento.Itens IsNot Nothing AndAlso objConhecimento.Itens.Count > 0 AndAlso
                                        objConhecimento.Itens(0).Encargos IsNot Nothing AndAlso objConhecimento.Itens(0).Encargos.Count > 0 Then
            For Each objEncargo As NotaFiscalXItemXEncargo In objConhecimento.Itens(0).Encargos
                For Each row As GridViewRow In gridEncargos.Rows
                    If Not String.IsNullOrWhiteSpace(row.Cells(0).Text) AndAlso objEncargo.Codigo.ToUpper() = row.Cells(0).Text.Trim().ToUpper() Then
                        Dim txtValor As TextBox = CType(row.Cells(1).FindControl("txtValorEncargo"), TextBox)
                        If Not String.IsNullOrWhiteSpace(txtValor.Text) Then
                            objEncargo.Valor = CDec(txtValor.Text.Trim())
                        End If
                    End If
                Next
            Next

            SalvarSessaoConhecimento()
        End If
    End Sub

    Private Sub ImprimirEspelho()
        Try
            RecuperarSessaoConhecimento()
            Dim espelho As New [Lib].Negocio.NotaFiscalEspelho
            espelho.ExibirEspelho(Me.Page, objConhecimento)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ImprimirEspelho(ByVal obj As [Lib].Negocio.NotaFiscal)
        Try
            Dim espelho As New [Lib].Negocio.NotaFiscalEspelho
            espelho.ExibirEspelho(Me.Page, obj)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ImprimeContrato(ByVal Contrato As [Lib].Negocio.NotaFiscal)
        Dim ds As New DataSet
        Dim dtImagem As DataTable = ds.Tables.Add("Images")
        dtImagem.Columns.Add("path", GetType(String))
        dtImagem.Columns.Add("image", GetType(System.Byte()))

        Dim drImagem As DataRow = dtImagem.NewRow()
        Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))

        drImagem("path") = strCaminhoImagem
        drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
        dtImagem.Rows.Add(drImagem)

        Dim dtImagemPamcard As DataTable = ds.Tables.Add("ImagePamcard")
        dtImagemPamcard.Columns.Add("pathPamcard", GetType(String))
        dtImagemPamcard.Columns.Add("imagePamcard", GetType(System.Byte()))

        Dim drImagemPamcard As DataRow = dtImagemPamcard.NewRow()
        Dim strCaminhoImagemPamcard As String = Server.MapPath("~/Images/pamcard.jpg")

        drImagemPamcard("pathPamcard") = strCaminhoImagemPamcard
        drImagemPamcard("imagePamcard") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagemPamcard)
        dtImagemPamcard.Rows.Add(drImagemPamcard)

        Dim tbPagamento As New DataTable("Pagamento")
        tbPagamento.Columns.Add("numero", Type.GetType("System.String"))
        tbPagamento.Columns.Add("tipo", Type.GetType("System.String"))
        tbPagamento.Columns.Add("efetivacao", Type.GetType("System.String"))
        tbPagamento.Columns.Add("vencimento", Type.GetType("System.String"))
        tbPagamento.Columns.Add("pagamento", Type.GetType("System.String"))
        tbPagamento.Columns.Add("situacao", Type.GetType("System.String"))
        tbPagamento.Columns.Add("valor", Type.GetType("System.String"))
        ds.Tables.Add(tbPagamento)

        Dim vlrProduto As Decimal = 0
        Dim vlrAdto As Decimal = 0
        Dim vlrPedagio As Decimal = 0
        Dim vlrliquido As Decimal = 0
        Dim vlrSeguro As Decimal = 0
        Dim vlrCadastro As Decimal = 0

        For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In Contrato.Itens(0).Encargos
            If enc.Codigo = "TARIFA PEDAGIO" Then
                vlrPedagio = enc.Valor
            ElseIf enc.Codigo = "ADTODEFRETE" OrElse enc.Codigo = "ADIANTAMENTO" Then
                vlrAdto += enc.Valor
            ElseIf enc.Codigo = "TARIFA SEGURO" Then
                vlrSeguro = enc.Valor
            ElseIf enc.Codigo = "TAXA CADASTRO" Then
                vlrCadastro = enc.Valor
            ElseIf enc.Codigo = "PRODUTO" Then
                vlrProduto = enc.Valor
            ElseIf enc.Codigo = "LIQUIDO" Then
                vlrliquido = enc.Valor
            End If
        Next

        Dim i As Integer = 0

        If vlrPedagio > 0 Then
            Dim drPagamento As DataRow = tbPagamento.NewRow()
            i += 1
            drPagamento("numero") = i.ToString
            drPagamento("tipo") = "PEDAGIO(+)"
            drPagamento("efetivacao") = "AUTO"
            drPagamento("vencimento") = Contrato.Movimento.ToString("dd/MM/yyyy")
            drPagamento("pagamento") = Contrato.Movimento.ToString("dd/MM/yyyy")
            drPagamento("situacao") = "EFETIVADA"
            drPagamento("valor") = vlrPedagio.ToString("N2")
            tbPagamento.Rows.Add(drPagamento)
        End If

        If vlrAdto > 0 Then
            Dim drPagamento As DataRow = tbPagamento.NewRow()
            i += 1
            drPagamento("numero") = i.ToString
            drPagamento("tipo") = "ADIANTAMENTO(-)"
            drPagamento("efetivacao") = "AUTO"
            drPagamento("vencimento") = Contrato.Movimento.ToString("dd/MM/yyyy")
            drPagamento("pagamento") = Contrato.Movimento.ToString("dd/MM/yyyy")
            drPagamento("situacao") = "EFETIVADA"
            drPagamento("valor") = vlrAdto.ToString("N2")
            tbPagamento.Rows.Add(drPagamento)
        End If

        If vlrSeguro > 0 Then
            Dim drPagamento As DataRow = tbPagamento.NewRow()
            i += 1
            drPagamento("numero") = i.ToString
            drPagamento("tipo") = "SEGURO(-)"
            drPagamento("efetivacao") = ""
            drPagamento("vencimento") = ""
            drPagamento("pagamento") = ""
            drPagamento("situacao") = ""
            drPagamento("valor") = vlrSeguro.ToString("N2")
            tbPagamento.Rows.Add(drPagamento)
        End If

        If vlrCadastro > 0 Then
            Dim drPagamento As DataRow = tbPagamento.NewRow()
            i += 1
            drPagamento("numero") = i.ToString
            drPagamento("tipo") = "CADASTRO PAMCARD(-)"
            drPagamento("efetivacao") = ""
            drPagamento("vencimento") = ""
            drPagamento("pagamento") = ""
            drPagamento("situacao") = ""
            drPagamento("valor") = vlrCadastro.ToString("N2")
            tbPagamento.Rows.Add(drPagamento)
        End If

        If vlrliquido > 0 Then
            vlrliquido = vlrliquido - vlrAdto
            Dim drPagamento As DataRow = tbPagamento.NewRow()
            i += 1
            drPagamento("numero") = i.ToString
            drPagamento("tipo") = "SALDO"
            drPagamento("efetivacao") = "AUTO"
            drPagamento("vencimento") = Contrato.DataTermino.ToString("dd/MM/yyyy")
            drPagamento("pagamento") = Contrato.DataTermino.ToString("dd/MM/yyyy")
            drPagamento("situacao") = "BLOQUEADO"
            drPagamento("valor") = vlrliquido.ToString("N2")
            tbPagamento.Rows.Add(drPagamento)
        End If

        Dim strCaminho As String = Server.MapPath("~/Reports/Cr_ContratoDeFrete.rpt")

        Dim crpt As New ReportDocument()

        Try
            crpt.Load(strCaminho)

            Dim strNomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim strArquivo As String = HttpContext.Current.Server.MapPath(strNomeArquivo)

            crpt.SetDataSource(ds)

            Dim param As New Dictionary(Of String, Object)()
            param.Add("codigoANTT", "CIOT " & Contrato.ContratoANTT & " / " & Contrato.ProtocoloANTT)
            param.Add("codigoViagem", "Viagem: " & Contrato.idPamcard)
            param.Add("cartaoPamcard", "Cartão Nr.: " & Left(Contrato.CartaoPgtoFrete, 4) & "." & Mid(Contrato.CartaoPgtoFrete, 5, 4) & "." & Mid(Contrato.CartaoPgtoFrete, 9, 4) & "." & Mid(Contrato.CartaoPgtoFrete, 13, 4))
            param.Add("nomeContratante", "Empresa: " & Contrato.NotaTrocaOrigem.Empresa.Nome)
            param.Add("cnpjContratante", "CNPJ: " & Contrato.NotaTrocaOrigem.Empresa.CodigoFormatado)
            param.Add("enderecoContratante", "Endereco: " & IIf(Contrato.NotaTrocaOrigem.Empresa.Numero > 0, Contrato.NotaTrocaOrigem.Empresa.Endereco & ", " & Contrato.NotaTrocaOrigem.Empresa.Numero, Contrato.NotaTrocaOrigem.Empresa.Endereco))
            param.Add("bairroContratante", "Bairro: " & Contrato.NotaTrocaOrigem.Empresa.Bairro)
            param.Add("cepContratante", "Cep: " & Contrato.NotaTrocaOrigem.Empresa.CEP)
            param.Add("telefoneContratante", "Telefone: " & Contrato.NotaTrocaOrigem.Empresa.Telefone)
            param.Add("cidadeContratante", "Cidade/UF.: " & Contrato.NotaTrocaOrigem.Empresa.Cidade & "/" & Contrato.NotaTrocaOrigem.Empresa.CodigoEstado)

            'Contratado/Transportador
            param.Add("nomeContratado", "Empresa: " & Contrato.Transportador.Nome)
            param.Add("cnpjContratado", "CPF/CNPJ: " & Funcoes.FormatarCpfCnpj(Contrato.Transportador.Codigo))
            param.Add("enderecoContratado", "Endereco: " & IIf(Contrato.Transportador.Numero > 0, Contrato.Transportador.Endereco & ", " & Contrato.Transportador.Numero, Contrato.Transportador.Endereco))
            param.Add("bairroContratado", "Bairro: " & Contrato.Transportador.Bairro)
            param.Add("cepContratado", "Cep: " & Contrato.Transportador.CEP)
            param.Add("telefoneContratado", "Telefone: " & Contrato.Transportador.Telefone)
            param.Add("cidadeContratado", "Cidade/UF.: " & Contrato.Transportador.Cidade & "/" & Contrato.Transportador.CodigoEstado)
            param.Add("rntrcContratado", "RNTRC: " & Contrato.Transportador.RNTRCTransportador)

            'Veículo
            Dim placas As String = Contrato.PlacaDetalhes.Placa01.Replace("-", "")
            If Contrato.PlacaDetalhes.Placa02.Length > 0 Then placas = placas & "/" & Contrato.PlacaDetalhes.Placa02.Replace("-", "")
            If Contrato.PlacaDetalhes.Placa03.Length > 0 Then placas = placas & "/" & Contrato.PlacaDetalhes.Placa03.Replace("-", "")
            If Contrato.PlacaDetalhes.Placa04.Length > 0 Then placas = placas & "/" & Contrato.PlacaDetalhes.Placa04.Replace("-", "")
            param.Add("veiculo", "Placa: " & placas & "  -  Categoria: " & Contrato.PlacaDetalhes.TipoDeVeiculoDetalhes.Descricao)

            'Motorista
            param.Add("dadosMotorista", "Motorista: " & Contrato.Motorista.Nome & " - CPF: " & Funcoes.FormatarCpfCnpj(Contrato.Motorista.Codigo) & " - " & Contrato.Motorista.Cidade & "/" & Contrato.Motorista.CodigoEstado)

            'Dados da Viagem
            param.Add("DataDeEmbarque", "Data do Embarque: " & Contrato.Movimento)
            param.Add("DataFimViagem", "Previsão Fim da Viagem: " & Contrato.Movimento.AddDays(15))
            param.Add("viagemOrigem", "Origem: " & Contrato.NotaTrocaOrigem.Empresa.Cidade & "/" & Contrato.NotaTrocaOrigem.Empresa.CodigoEstado)
            param.Add("viagemDestino", "Destino: " & Contrato.NotaTrocaOrigem.Cliente.Cidade & "/" & Contrato.NotaTrocaOrigem.Cliente.CodigoEstado)

            'Remetente
            param.Add("nomeRemetente", "Remetente: " & Contrato.NotaTrocaOrigem.Empresa.Nome)
            param.Add("cnpjRemetente", "CNPJ: " & Contrato.NotaTrocaOrigem.Empresa.CodigoFormatado & "  -  Cidade: " & Contrato.NotaTrocaOrigem.Empresa.Cidade & "/" & Contrato.NotaTrocaOrigem.Empresa.CodigoEstado)
            param.Add("enderecoRemetente", "Endereco: " & IIf(Contrato.NotaTrocaOrigem.Empresa.Numero > 0, Contrato.NotaTrocaOrigem.Empresa.Endereco & ", " & Contrato.NotaTrocaOrigem.Empresa.Numero, Contrato.NotaTrocaOrigem.Empresa.Endereco))
            param.Add("bairroRemetente", "Bairro: " & Contrato.NotaTrocaOrigem.Empresa.Bairro & "    CEP: " & Contrato.NotaTrocaOrigem.Empresa.CEP)

            'Destinatário
            param.Add("nomeDestinatario", "Destinatario: " & Contrato.NotaTrocaOrigem.Cliente.Nome)
            param.Add("cnpjDestinatario", "CNPJ: " & Contrato.NotaTrocaOrigem.Cliente.CodigoFormatado & "  -  Cidade: " & Contrato.NotaTrocaOrigem.Cliente.Cidade & "/" & Contrato.NotaTrocaOrigem.Cliente.CodigoEstado)
            param.Add("enderecoDestinatario", "Endereco: " & IIf(Contrato.NotaTrocaOrigem.Cliente.Numero > 0, Contrato.NotaTrocaOrigem.Cliente.Endereco & ", " & Contrato.NotaTrocaOrigem.Cliente.Numero, Contrato.NotaTrocaOrigem.Cliente.Endereco))
            param.Add("bairroDestinatario", "Bairro: " & Contrato.NotaTrocaOrigem.Cliente.Bairro & "    CEP: " & Contrato.NotaTrocaOrigem.Cliente.CEP)

            'Documento
            param.Add("numeroDocumento", "DOCUMENTO: CT-E Nr. " & Contrato.NotaTrocaOrigem.Codigo & " Série: " & Contrato.NotaTrocaOrigem.Serie & " - Nota Fiscal Nr. " & Contrato.NotaTrocaOrigem.Codigo & "-" & Contrato.NotaTrocaOrigem.Serie & "  Produto: " & Contrato.NotaTrocaOrigem.Itens(0).Produto.Descricao & "  Peso: " & Contrato.NotaTrocaOrigem.Itens(0).QuantidadeFiscal.ToString("N0") & Contrato.NotaTrocaOrigem.Itens(0).Produto.Unidade)

            'Valor do Frete
            param.Add("valorBruto", "- Unitário Frete: " & Contrato.Itens(0).Unitario.ToString("N10") & " - Valor Bruto: " & vlrProduto.ToString("N2"))

            'Assinatura Contratado
            param.Add("dataContrato", Contrato.Empresa.Cidade & "/" & Contrato.Empresa.CodigoEstado & ", " & Contrato.Movimento.ToLongDateString.ToUpper & ".")

            'Assinatura Contratado
            Dim strTexto As String = IIf(Contrato.Transportador.Codigo.Length = 11, "CPF: ", "CNPJ: ")
            param.Add("assinaturaContratado", Contrato.Transportador.Nome & " - " & strTexto & Funcoes.FormatarCpfCnpj(Contrato.Transportador.Codigo))

            'Assinatura Contratado
            param.Add("assinaturaContratante", Contrato.NotaTrocaOrigem.Empresa.Nome & " - CNPJ: " & Contrato.NotaTrocaOrigem.Empresa.CodigoFormatado)
            Funcoes.BindParameters(crpt, param)

            If Dir(strArquivo).Length > 0 Then Kill(strArquivo)

            crpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, strArquivo)

            If IO.File.Exists(strArquivo) Then
                Funcoes.AbrirArquivo(Page, strNomeArquivo)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtPeriodoInicial.Text) AndAlso Not String.IsNullOrWhiteSpace(txtPeriodoFinal.Text) Then
            MsgBox(Me.Page, "Datas são obrigatórias.")
            Return False
        End If
        Return True
    End Function

    Private Function getParametros() As String
        Dim param As String = "Empresa: " & DdlEmpresa.SelectedItem.Text & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtNomeCliente.Text) Then
            param &= "Cliente: " & txtNomeCliente.Text & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtNumConhecimento.Text) Then
            param &= "Nota: " & txtNumConhecimento.Text & vbCrLf
        End If

        param &= "Notas de: " & IIf(rdoEntrada.Checked, "Entrada", "Saída") & vbCrLf &
            "Período: " & txtPeriodoInicial.Text & " a " & txtPeriodoFinal.Text

        Return param
    End Function

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

    Private Function GravarNFeXpress(ByVal cte As [Lib].Negocio.NotaFiscal) As Boolean
        Dim aux As Boolean = True
        Try
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"

            If cte.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                obj.NomeArquivo = String.Format("nota{0:000000000}#{1}.txt", cte.Codigo, cte.CodigoEmpresa)
            Else
                obj.NomeArquivo = String.Format("cte{0:000000000}#{1}.txt", cte.Codigo, cte.CodigoEmpresa)
            End If

            If cte.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                'obj.Texto = [Lib].Negocio.DocumentoEletronico.getTextoNFe(cte, HomologAlfasig, 1)
                'obj.Texto = [Lib].Negocio.DocumentoEletronico.getTextoNFe3G(cte, HomologAlfasig, 1)
                'Como ainda não estamos usando troquei porque 3G está descontinuada - Furlan - 12/02/2025
                obj.Texto = [Lib].Negocio.DocumentoEletronico.getTextoNFe4G(objConhecimento, False, 1)
            Else
                obj.Texto = getTextoCTe(cte)
            End If

            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If
        Catch ex As Exception
            aux = False
            Throw New Exception(ex.Message)
        End Try
        Return aux
    End Function

    Private Sub EnviarSEFAZ()
        RecuperarSessaoConhecimento()
        If objConhecimento Is Nothing Then
            MsgBox(Me.Page, "É necessário consultar um documento CT-e para realizar o envio para a SEFAZ!")
            Exit Sub
        End If

        If GravarNFeXpress(objConhecimento) Then
            objConhecimento = New [Lib].Negocio.NotaFiscal(objConhecimento)
            Dim obj As [Lib].Negocio.Resp = Nothing

            Dim fileName As String = String.Empty

            If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                fileName = String.Format("resp-nota{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
            Else
                fileName = String.Format("resp-cte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
            End If

            Dim tempoLimite As DateTime
            tempoLimite = Now.AddSeconds(30)

            While obj Is Nothing AndAlso Now < tempoLimite
                obj = GetResp(fileName)
                System.Threading.Thread.Sleep(3000)
            End While

            If obj IsNot Nothing Then
                Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

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

                If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "100" AndAlso strCodigo <> "4016" Then
                    lnkEnviarSEFAZ.Parent.Visible = True
                    lnkVisualizar.Parent.Parent.Parent.Visible = True
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    Exit Sub
                End If

                'ATUALIZAR NFE PENDENCIAS COM OS DADOS DO RETORNO
                Dim Sqls As New ArrayList
                Dim sql As String = String.Empty

                sql = "INSERT INTO NFERealizadas "
                sql &= "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, "
                sql &= "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) "
                sql &= "VALUES ('" & objConhecimento.CodigoEmpresa & "', " & objConhecimento.EnderecoEmpresa & ", '" & objConhecimento.CodigoCliente & "', '"
                sql &= objConhecimento.EnderecoCliente & "', '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objConhecimento.Serie & "', '" & objConhecimento.Codigo & "', '"
                sql &= CDate(objConhecimento.DataInclusao).ToString("yyyy-MM-dd") & "', '" & Format(objConhecimento.DataInclusao, "HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', '"
                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    sql &= String.Format("nota{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "', 'INCLUIR', '"
                Else
                    sql &= String.Format("cte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "', 'INCLUIR', '"
                End If
                sql &= strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '1', '" & objConhecimento.Observacoes & "', '', '" & strProtocolo & "', ''); "
                Sqls.Add(sql)

                sql = "DELETE FROM NFEPendencias "
                sql &= "WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "' And EndEmpresa_Id = '" & objConhecimento.EnderecoEmpresa & "' And "
                sql &= "      Cliente_Id = '" & objConhecimento.CodigoCliente & "' And EndCliente_Id = '" & objConhecimento.EnderecoCliente & "' And "
                sql &= "      EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' And Serie_Id = '" & objConhecimento.Serie & "' And "
                sql &= "      Nota_Id = '" & objConhecimento.Codigo & "'; "
                Sqls.Add(sql)

                sql = "UPDATE NotasFiscais Set Situacao = 1 "
                sql &= "WHERE Empresa_Id = '" & objConhecimento.CodigoEmpresa & "' And EndEmpresa_Id = '" & objConhecimento.EnderecoEmpresa & "' And "
                sql &= "      Cliente_Id = '" & objConhecimento.CodigoCliente & "' And EndCliente_Id = '" & objConhecimento.EnderecoCliente & "' And "
                sql &= "      EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' And Serie_Id = '" & objConhecimento.Serie & "' And "
                sql &= "      Nota_Id = '" & objConhecimento.Codigo & "'; "
                Sqls.Add(sql)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Recontabilizar()
                    Exit Sub
                End If
                'FIM DA ROTINA

                SalvarSessaoConhecimento()
                If SendMailCTe() Then
                    ImprimirCTe(True)
                End If

                Sqls.Clear()

                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-nota{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                Else
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statuscte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-cte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-dacte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa) & "';"
                    Sqls.Add(sql)
                End If

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If

                ImprimirEspelho()
                MsgBox(Me.Page, "CT-e " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " incluído com Sucesso.", eTitulo.Sucess)
                Duplicar(objConhecimento)
                LimparCampos()
                tcPrincipal.ActiveTab = tabCTRC
            End If
        End If
    End Sub

    Private Function VerificarCTe() As Boolean
        RecuperarSessaoConhecimento()
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True
        Try
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                obj.NomeArquivo = String.Format("status{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
            Else
                obj.NomeArquivo = String.Format("statuscte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
            End If
            obj.Texto = String.Empty
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Empty
            If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                fileName = String.Format("resp-status{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
            Else
                fileName = String.Format("resp-statuscte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
            End If

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

                If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "107" Then
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    Return False
                End If
            End If
        Catch ex As Exception
            aux = False
            Throw New Exception(ex.Message)
        End Try
        Return aux
    End Function

    Private Sub ImprimirCTe(Optional ByVal IsFirst = False)
        Dim Sqls As New ArrayList
        RecuperarSessaoConhecimento()
        Dim pdf As New [Lib].Negocio.Fil()

        pdf.IUD = "I"
        objConhecimento = New [Lib].Negocio.NotaFiscal(objConhecimento)
        If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
            pdf.NomeArquivo = String.Format("danfe{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
        Else
            pdf.NomeArquivo = String.Format("dacte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
        End If
        pdf.Texto = getTextoDACTE(objConhecimento)
        pdf.SalvarSql(Sqls)

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Exit Sub
        End If

        Dim obj As [Lib].Negocio.Resp = Nothing

        Dim tempoLimite As DateTime
        tempoLimite = Now.AddSeconds(30)

        While obj Is Nothing AndAlso Now < tempoLimite
            obj = GetResp(String.Format("resp-{0}", pdf.NomeArquivo))
            System.Threading.Thread.Sleep(3000)
        End While

        If obj IsNot Nothing Then
            Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4012" Then
                MsgBox(Me.Page, strMsg)
                Exit Sub
            End If

            Try
                Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objConhecimento.ChaveNFE), IIf(objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao, eTipoDeDocumento.Nota, eTipoDeDocumento.CTRC))
                If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.ChaveNFE) Then
                    Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objConhecimento.ChaveNFE))
                    System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                    Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objConhecimento.ChaveNFE))
                End If
            Catch ex As Exception
                If IsFirst Then
                    MsgBox(Me.Page, "CT-e " & objConhecimento.Codigo & "-" & objConhecimento.Serie & " incluído com Sucesso.", eTitulo.Sucess)
                    ConsultarNotas()
                    LimparCampos()
                    tcPrincipal.ActiveTab = tabCTRC
                Else
                    Throw New Exception(ex.Message)
                End If
            End Try
        End If
    End Sub

    Private Function CancelarCTe() As Boolean
        RecuperarSessaoConhecimento()
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True
        Try
            If Not String.IsNullOrWhiteSpace(objConhecimento.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.ProtocoloNota) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    obj.NomeArquivo = String.Format("cancel{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                Else
                    obj.NomeArquivo = String.Format("cancelcte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                End If
                obj.Texto = getTextoCancelar(objConhecimento)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Return False
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Empty
                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    fileName = String.Format("resp-cancel{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                Else
                    fileName = String.Format("resp-cancelcte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                End If

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

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "101" Then
                        MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                        Return False
                    End If
                End If
            End If
        Catch ex As Exception
            aux = False
            Throw New Exception(ex.Message)
        End Try
        Return aux
    End Function

    Private Function SendMailCTe() As Boolean
        Dim aux As Boolean = True
        Dim Sqls As New ArrayList
        Try
            RecuperarSessaoConhecimento()
            objConhecimento = New [Lib].Negocio.NotaFiscal(objConhecimento)
            If Not String.IsNullOrWhiteSpace(objConhecimento.Empresa.EmailNFE) AndAlso Not String.IsNullOrWhiteSpace(objConhecimento.Cliente.EmailNFE) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                If objConhecimento.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
                    obj.NomeArquivo = String.Format("email{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                Else
                    obj.NomeArquivo = String.Format("emailcte{0:000000000}#{1}.txt", objConhecimento.Codigo, objConhecimento.CodigoEmpresa)
                End If
                obj.Texto = getTextoEmail(objConhecimento)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Return False
                End If
            End If
        Catch ex As Exception
            aux = False
            Throw New Exception(ex.Message)
        End Try
        Return aux
    End Function

    Private Function getDataSetEmpresa(ByVal cte As [Lib].Negocio.NotaFiscal)
        Dim sql As String = String.Empty
        sql &= "SELECT     Clientes.Nome, Clientes.Inscricao, Clientes.Endereco, ISNULL(Clientes.Numero, N'') AS Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, "
        sql &= "                      ISNULL(Clientes.Bairro, '') AS Bairro, Municipios.Codigo_id AS CodMunicipio, Clientes.Cidade, Clientes.Estado, Clientes.Cep, Clientes.Pais, "
        sql &= "                      Pais.Descricao AS NomePais, Clientes.Telefone, Municipios.Estadoibge, Clientes.EmailNFE, ClientesXEmpresas.Crt, Clientes.Email "
        sql &= "FROM         Clientes INNER JOIN "
        sql &= "                      ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id INNER JOIN "
        sql &= "                      Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.Cidade = Municipios.Municipio_id INNER JOIN "
        sql &= "                      Pais ON Clientes.Pais = Pais.Pais_Id INNER JOIN "
        sql &= "                      ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id "
        sql &= "WHERE     (Clientes.Cliente_Id = '" & cte.CodigoEmpresa & "') AND (Clientes.Endereco_Id = " & cte.EnderecoEmpresa & ") AND (ClientesXTipos.Tipo_Id = 1)"
        Return Banco.ConsultaDataSet(sql, "ConsultaEmpresa")
    End Function

    Private Function getDataSetCliente(ByVal cte As [Lib].Negocio.NotaFiscal)
        Dim sql As String = String.Empty
        sql &= "SELECT     TOP (1) Clientes.Cliente_Id AS Cliente, Clientes.Endereco_Id AS EndCliente, Clientes.Nome, Clientes.Inscricao, Clientes.Endereco, "
        sql &= " ISNULL(Clientes.Numero, N'') AS Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, "
        sql &= " ISNULL(Clientes.Bairro, '') AS Bairro, Municipios.Codigo_id AS CodMunicipio, Clientes.Cidade, Clientes.Estado, Clientes.Cep, Clientes.Pais, "
        sql &= " Pais.Descricao AS NomePais, Clientes.Telefone, Clientes.Email, Municipios.Estadoibge, Clientes.EmailNFE, Clientes.Email "
        sql &= " FROM Clientes "
        sql &= " INNER JOIN Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.Cidade = Municipios.Municipio_id "
        sql &= " INNER JOIN Pais ON Clientes.Pais = Pais.Pais_Id "
        sql &= " WHERE (Clientes.Cliente_Id = '" & cte.CodigoCliente & "') AND (Clientes.Endereco_Id = " & cte.EnderecoCliente & ")"
        Return Banco.ConsultaDataSet(sql, "ConsultaCliente")
    End Function

    Private Function getDataSetCliente(ByVal codigo As String, ByVal endereco As String) As DataSet
        Dim sql As String = "SELECT     TOP (1) Clientes.Cliente_Id AS Cliente, Clientes.Endereco_Id AS EndCliente, Clientes.Nome, Clientes.Inscricao, Clientes.Endereco, ISNULL(Clientes.Numero, N'') AS Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, "
        sql &= "                      ISNULL(Clientes.Bairro, '') AS Bairro, Municipios.Codigo_id AS CodMunicipio, Clientes.Cidade, Clientes.Estado, Clientes.Cep, Clientes.Pais, "
        sql &= "                      Pais.Descricao AS NomePais, Clientes.Telefone, Clientes.Email, Municipios.Estadoibge, Clientes.EmailNFE, Clientes.Email "
        sql &= "FROM         Clientes INNER JOIN "
        sql &= "                      Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.Cidade = Municipios.Municipio_id INNER JOIN "
        sql &= "                      Pais ON Clientes.Pais = Pais.Pais_Id "
        sql &= "WHERE     (Clientes.Cliente_Id = '" & codigo & "') AND (Clientes.Endereco_Id = " & endereco & ")"
        Return Banco.ConsultaDataSet(sql, "ConsultaCliente")
    End Function

    Private Function getDataSetConhecimento(ByVal cte As [Lib].Negocio.NotaFiscal)
        Dim sql As String = "SELECT Operacao, SubOperacao, Movimento, Deposito, EndDeposito, Destino, EndDestino, Transbordo, EndTransbordo, CIFFOB, "
        sql &= "UsuarioInclusao, isnull(LocalEmbarque,Deposito) AS LocalEmbarque, isnull(EndLocalEmbarque,EndDeposito) AS EndLocalEmbarque "
        sql &= " FROM  NotasFiscais "
        sql &= " WHERE (Empresa_Id = '" & cte.CodigoEmpresa & "') AND (EndEmpresa_Id = " & cte.EnderecoEmpresa & ") AND (Cliente_Id = '" & cte.CodigoCliente & "') "
        sql &= " AND (EndCliente_Id = " & cte.EnderecoCliente & ") AND (EntradaSaida_Id = '" & IIf(cte.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') "
        sql &= " AND (Serie_Id = '" & cte.Serie & "') AND (Nota_Id = " & cte.Codigo & ")"
        Return Banco.ConsultaDataSet(sql, "ConsultaConhecimento")
    End Function

    Private Function getDataSetItensConhecimento(ByVal cte As [Lib].Negocio.NotaFiscal)
        Dim sql As String = "SELECT NotasFiscaisXItens.Produto_Id AS Produto, NotasFiscaisXItens.CFOP_Id AS CFOP, NotasFiscaisXItens.QuantidadeFisica, NotasFiscaisXItens.QuantidadeFiscal, "
        sql &= " NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor, Cfop.Descricao AS DescricaoFiscal "
        sql &= " FROM NotasFiscaisXItens "
        sql &= " INNER JOIN Cfop ON NotasFiscaisXItens.CFOP_Id = Cfop.Cfop_Id "
        sql &= " WHERE     (NotasFiscaisXItens.Empresa_Id = '" & cte.CodigoEmpresa & "') AND (NotasFiscaisXItens.EndEmpresa_Id = " & cte.EnderecoEmpresa & ") "
        sql &= " AND (NotasFiscaisXItens.Cliente_Id = '" & cte.CodigoCliente & "') "
        sql &= " AND (NotasFiscaisXItens.EndCliente_Id = " & cte.EnderecoCliente & ") AND (NotasFiscaisXItens.EntradaSaida_Id = '" & IIf(cte.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') "
        sql &= " AND (NotasFiscaisXItens.Serie_Id = '" & cte.Serie & "') "
        sql &= " AND (NotasFiscaisXItens.Nota_Id = " & cte.Codigo & ")"
        Return Banco.ConsultaDataSet(sql, "ConsultaItemConhecimento")
    End Function

    Private Function getDataSetEncargosConhecimento(ByVal cte As [Lib].Negocio.NotaFiscal)
        Dim sql As String
        sql = "SELECT nfe.Encargo_Id AS Encargo," & vbCrLf &
              "       nfe.Base," & vbCrLf &
              "	      nfe.Percentual," & vbCrLf &
              "	      nfe.Valor," & vbCrLf &
              "       OE.StICMS as SituacaoTributaria," & vbCrLf &
              "		  isnull(OEE.AliquotaBase,0) AS PercentualBase" & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              " INNER JOiN NotasFiscaisXItens Nfi" & vbCrLf &
              "    ON NF.Empresa_Id      = Nfi.Empresa_Id" & vbCrLf &
              "   AND NF.EndEmpresa_Id   = Nfi.EndEmpresa_Id" & vbCrLf &
              "   AND NF.Cliente_Id      = Nfi.Cliente_Id" & vbCrLf &
              "   AND NF.EndCliente_Id   = Nfi.EndCliente_Id" & vbCrLf &
              "   AND NF.EntradaSaida_Id = Nfi.EntradaSaida_Id" & vbCrLf &
              "   AND NF.Serie_Id        = Nfi.Serie_Id" & vbCrLf &
              "   AND NF.Nota_Id         = Nfi.Nota_Id" & vbCrLf &
              " INNER JOIN NotasFiscaisXEncargos nfe" & vbCrLf &
              "    ON Nfi.Empresa_Id      = nfe.Empresa_Id" & vbCrLf &
              "   AND Nfi.EndEmpresa_Id   = nfe.EndEmpresa_Id" & vbCrLf &
              "   AND Nfi.Cliente_Id      = nfe.Cliente_Id" & vbCrLf &
              "   AND Nfi.EndCliente_Id   = nfe.EndCliente_Id" & vbCrLf &
              "   AND Nfi.EntradaSaida_Id = nfe.EntradaSaida_Id" & vbCrLf &
              "   AND Nfi.Serie_Id        = nfe.Serie_Id" & vbCrLf &
              "   AND Nfi.Nota_Id         = nfe.Nota_Id" & vbCrLf &
              "   AND Nfi.Produto_Id      = nfe.Produto_Id" & vbCrLf &
              "   AND Nfi.CFOP_Id         = nfe.CFOP_Id" & vbCrLf &
              "	  AND Nfi.Sequencia_Id    = nfe.Sequencia_id" & vbCrLf &
              " INNER JOIN OperacaoXEstadoXEncargo OEE" & vbCrLf &
              "    ON OEE.Codigo_Id  = Nfi.OperacaoXEstado" & vbCrLf &
              "	  AND OEE.Encargo_Id = nfe.Encargo_Id" & vbCrLf &
              "  Inner join OperacaoXEstado OE" & vbCrLf &
              "    on OEE.Codigo_Id = OE.Codigo_Id" & vbCrLf &
              "WHERE (Nfi.Empresa_Id      ='" & cte.CodigoEmpresa & "') " & vbCrLf &
              "  AND (Nfi.EndEmpresa_Id   = " & cte.EnderecoEmpresa & ") " & vbCrLf &
              "  AND (Nfi.Cliente_Id      ='" & cte.CodigoCliente & "') " & vbCrLf &
              "  AND (Nfi.EndCliente_Id   = " & cte.EnderecoCliente & ") " & vbCrLf &
              "  AND (Nfi.EntradaSaida_Id ='" & IIf(cte.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') " & vbCrLf &
              "  AND (Nfi.Serie_Id        ='" & cte.Serie & "') " & vbCrLf &
              "  AND (Nfi.Nota_Id         = " & cte.Codigo & ")" & vbCrLf
        Return Banco.ConsultaDataSet(sql, "ConsultaEncargosConhecimento")
    End Function

    Private Function getDataSetConhecimentoXNota(ByVal cte As [Lib].Negocio.NotaFiscal) As DataSet
        Dim sql As String = "SELECT  NotasXNotas.OrigemEmpresa_id, NotasXNotas.OrigemEndEmpresa_id, NotasXNotas.OrigemEntradaSaida_id, NotasFiscaisXItens.Produto_Id AS Produto, Produtos.Nome AS NomeProduto, Produtos.Unidade AS UnidadeProduto, NotasFiscaisXItens.Valor AS ValorMercadoria, NotasFiscaisXItens.QuantidadeFiscal, Embalagens.Descricao AS EmbalagemProduto, NFERealizadas.ChaveNfe "
        sql &= "FROM NotasXNotas INNER JOIN "
        sql &= "  NotasFiscaisXItens ON NotasXNotas.OrigemEmpresa_Id = NotasFiscaisXItens.Empresa_Id AND "
        sql &= "  NotasXNotas.OrigemEndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasXNotas.OrigemCliente_Id = NotasFiscaisXItens.Cliente_Id AND "
        sql &= "  NotasXNotas.OrigemEndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasXNotas.OrigemEntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND "
        sql &= "  NotasXNotas.OrigemSerie_Id = NotasFiscaisXItens.Serie_Id AND NotasXNotas.OrigemNota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN "
        sql &= "  Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id INNER JOIN "
        sql &= "  Embalagens ON Produtos.Embalagem = Embalagens.Embalagem_Id INNER JOIN "
        sql &= "  NFERealizadas ON NotasXNotas.OrigemEmpresa_Id = NFERealizadas.Empresa_Id AND NotasXNotas.OrigemEndEmpresa_Id = NFERealizadas.EndEmpresa_Id AND "
        sql &= "  NotasXNotas.OrigemCliente_Id = NFERealizadas.Cliente_Id AND NotasXNotas.OrigemEndCliente_Id = NFERealizadas.EndCliente_Id AND "
        sql &= "  NotasXNotas.OrigemEntradaSaida_Id = NFERealizadas.EntradaSaida_Id AND NotasXNotas.OrigemSerie_Id = NFERealizadas.Serie_Id AND "
        sql &= "  NotasXNotas.OrigemNota_Id = NFERealizadas.Nota_Id "
        sql &= " WHERE (NotasXNotas.Empresa_Id = '" & cte.CodigoEmpresa & "') "
        sql &= " AND (NotasXNotas.EndEmpresa_Id = " & cte.EnderecoEmpresa & ") "
        sql &= " AND (NotasXNotas.Cliente_Id = '" & cte.CodigoCliente & "') "
        sql &= " AND (NotasXNotas.EndCliente_Id = " & cte.EnderecoCliente & ") "
        sql &= " AND (NotasXNotas.EntradaSaida_Id = '" & IIf(cte.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') "
        sql &= " AND (NotasXNotas.Serie_Id = '" & cte.Serie & "') "
        sql &= " AND (NotasXNotas.Nota_Id = " & cte.Codigo & ") "
        Return Banco.ConsultaDataSet(sql, "ConsultaConhecimentoXNota")
    End Function

    Private Function getDataSetRemetente(ByVal dsConhecimentoXNota As DataSet) As DataSet
        Dim sql As String = "SELECT     Clientes.Cliente_Id AS Cliente, Clientes.Endereco_Id AS EndCliente, Clientes.Nome, Clientes.Inscricao, Clientes.Endereco, ISNULL(Clientes.Numero, N'') AS Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, "
        sql &= "                      ISNULL(Clientes.Bairro, '') AS Bairro, Municipios.Codigo_id AS CodMunicipio, Clientes.Cidade, Clientes.Estado, Clientes.Cep, Clientes.Pais, "
        sql &= "                      Pais.Descricao AS NomePais, Clientes.Telefone, Municipios.Estadoibge "
        sql &= "FROM         Clientes INNER JOIN "
        sql &= "                      ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id INNER JOIN "
        sql &= "                      Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.Cidade = Municipios.Municipio_id INNER JOIN "
        sql &= "                      Pais ON Clientes.Pais = Pais.Pais_Id "
        sql &= "WHERE     (Clientes.Cliente_Id = '" & dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEmpresa_id") & "') AND (Clientes.Endereco_Id = " & dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEndEmpresa_id") & ")"
        Return Banco.ConsultaDataSet(sql, "ConsultaRemetente")
    End Function

    Private Function getDataSetTransportador(ByVal cte As [Lib].Negocio.NotaFiscal) As DataSet
        Dim sql As String = "SELECT     NotasFiscaisXTransportadores.Proprietario, NotasFiscaisXTransportadores.EndProprietario, DadosDoTransportador.Nome AS NomeTransportador, ISNULL(DadosDoTransportador.Inscricao, '') As IETransportador, DadosDoTransportador.Estado AS UFTransportador, NotasFiscaisXTransportadores.Motorista, "
        sql &= "                      NotasFiscaisXTransportadores.EndMotorista, DadosDoMotorista.Nome AS NomeMotorista, Placas.Placa_Id AS Placa, isnull(DadosDoTransportador.RNTRCTransportador,'') AS RNTRCTransportador, Placas.CidadePlaca, Placas.EstadoPlaca, Placas.Placa01, Placas.CidadePlaca01, "
        sql &= "                      Placas.EstadoPlaca01, Placas.Placa02, Placas.CidadePlaca02, Placas.EstadoPlaca02, Placas.Placa03, Placas.CidadePlaca03, Placas.EstadoPlaca03, "
        sql &= "                      ISNULL(Placas.RNTRCPlaca,'') AS RNTRCPlaca "
        sql &= "FROM         NotasFiscaisXTransportadores INNER JOIN "
        sql &= "                      Placas ON NotasFiscaisXTransportadores.Placa = Placas.Placa_Id INNER JOIN "
        sql &= "                      Clientes AS DadosDoMotorista ON NotasFiscaisXTransportadores.Motorista = DadosDoMotorista.Cliente_Id AND "
        sql &= "                      NotasFiscaisXTransportadores.EndMotorista = DadosDoMotorista.Endereco_Id INNER JOIN "
        sql &= "                      Clientes AS DadosDoTransportador ON NotasFiscaisXTransportadores.Proprietario = DadosDoTransportador.Cliente_Id AND "
        sql &= "                      NotasFiscaisXTransportadores.EndProprietario = DadosDoTransportador.Endereco_Id "
        sql &= "WHERE     (NotasFiscaisXTransportadores.Empresa_Id = '" & cte.CodigoEmpresa & "') AND (NotasFiscaisXTransportadores.EndEmpresa_Id = " & cte.EnderecoEmpresa & ") AND "
        sql &= "                      (NotasFiscaisXTransportadores.Cliente_Id = '" & cte.CodigoCliente & "') AND (NotasFiscaisXTransportadores.EndCliente_Id = " & cte.EnderecoCliente & ") AND "
        sql &= "                      (NotasFiscaisXTransportadores.EntradaSaida_Id = '" & IIf(cte.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') AND (NotasFiscaisXTransportadores.Serie_Id = '" & cte.Serie & "') AND (NotasFiscaisXTransportadores.Nota_Id = " & cte.Codigo & ")"
        Return Banco.ConsultaDataSet(sql, "ConsultaTransportador")
    End Function

    Private Function getDataSetOrigemMercardoria(ByVal dsConhecimento As DataSet, ByVal dsConhecimentoXNota As DataSet) As DataSet
        Dim sql As String = "SELECT     Clientes.Cliente_Id AS Cliente, Clientes.Endereco_Id AS EndCliente, Clientes.Nome, Clientes.Inscricao, Clientes.Endereco, ISNULL(Clientes.Numero, N'') AS Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, "
        sql &= "                      ISNULL(Clientes.Bairro, '') AS Bairro, Municipios.Codigo_id AS CodMunicipio, Clientes.Cidade, Clientes.Estado, Clientes.Cep, Clientes.Pais, "
        sql &= "                      Pais.Descricao AS NomePais, Clientes.Telefone, Municipios.Estadoibge "
        sql &= "FROM         Clientes INNER JOIN "
        sql &= "                      ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id INNER JOIN "
        sql &= "                      Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.Cidade = Municipios.Municipio_id INNER JOIN "
        sql &= "                      Pais ON Clientes.Pais = Pais.Pais_Id "
        If dsConhecimento.Tables(0).Rows(0)("LocalEmbarque") = "" Then
            sql &= "WHERE     (Clientes.Cliente_Id = '" & dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEmpresa_id") & "') AND (Clientes.Endereco_Id = " & dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEndEmpresa_id") & ")"
        Else
            sql &= "WHERE     (Clientes.Cliente_Id = '" & dsConhecimento.Tables(0).Rows(0)("LocalEmbarque") & "') AND (Clientes.Endereco_Id = " & dsConhecimento.Tables(0).Rows(0)("EndLocalEmbarque") & ")"
        End If
        Return Banco.ConsultaDataSet(sql, "ConsultaOrigemMercadoria")
    End Function

    Private Function getDataSetPendencias(ByVal cte As [Lib].Negocio.NotaFiscal) As DataSet
        Dim sql As String = String.Empty
        sql = "SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Serie_Id, Nota_Id, EntradaSaida_Id, "
        sql &= " CONVERT(varchar, Data, 103) As Data, CONVERT(varchar(08),Hora,14) As Hora, Usuario, NfExpress, Operacao, ISNULL(Retorno, '') AS Retorno, MsgRetorno, Recibo, "
        sql &= " isnull(ChaveNfe,'') AS ChaveNfe, Lote, TpEmis, ISNULL(ObservacoesFiscais, '') AS ObservacoesFiscais, ISNULL(DadosAdicionais, '') AS DadosAdicionais, ISNULL(Protocolo, '') AS Protocolo, RegDPEC "
        sql &= " FROM NFEPendencias "
        sql &= " WHERE 1=1 "
        sql &= " AND (NFEPendencias.Empresa_Id = '" & cte.CodigoEmpresa & "') " & vbCrLf
        sql &= " AND (NFEPendencias.EndEmpresa_Id = " & cte.EnderecoEmpresa & ") " & vbCrLf
        sql &= " AND (NFEPendencias.Cliente_Id = '" & cte.CodigoCliente & "') " & vbCrLf
        sql &= " AND (NFEPendencias.EndCliente_Id = " & cte.EnderecoCliente & ") " & vbCrLf
        sql &= " AND (NFEPendencias.EntradaSaida_Id = '" & IIf(cte.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') " & vbCrLf
        sql &= " AND (NFEPendencias.Serie_Id = '" & cte.Serie & "') " & vbCrLf
        sql &= " AND (NFEPendencias.Nota_Id = " & cte.Codigo & ")" & vbCrLf
        sql &= " ORDER BY Data DESC, Hora DESC"
        Return Banco.ConsultaDataSet(sql, "NFEPendencias")
    End Function

    Private Function getTextoCTe(ByVal cte As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        Dim dsEmpresa As DataSet = getDataSetEmpresa(cte)
        Dim dsCliente As DataSet = getDataSetCliente(cte)
        Dim dsConhecimento As DataSet = getDataSetConhecimento(cte)
        Dim dsItensConhecimento As DataSet = getDataSetItensConhecimento(cte)
        Dim dsEncargosConhecimento As DataSet = getDataSetEncargosConhecimento(cte)
        Dim dsConhecimentoXNota As DataSet = getDataSetConhecimentoXNota(cte)
        Dim dsRemetente As DataSet = getDataSetRemetente(dsConhecimentoXNota)
        Dim dsTransportador As DataSet = getDataSetTransportador(cte)
        Dim dsOrigemMercadoria As DataSet = getDataSetOrigemMercardoria(dsConhecimento, dsConhecimentoXNota)
        Dim dsPendencias As DataSet = getDataSetPendencias(cte)
        Dim dsLocalEmbarque As New DataSet
        Dim dsDestino As New DataSet

        Dim situacaoTributaria As Integer = 0
        Dim valorMercadoria As Decimal = 0
        Dim valor As String = ""
        Dim pBaseICMS As String = ""
        Dim baseICMS As String = ""
        Dim aliqICMS As String = ""
        Dim valorICMS As String = ""

        Dim strEmpresa As String = dsPendencias.Tables(0).Rows(0).Item("Empresa_Id")
        Dim strSerie As String = dsPendencias.Tables(0).Rows(0).Item("Serie_Id")
        Dim strConhecimento As String = dsPendencias.Tables(0).Rows(0).Item("Nota_Id")
        Dim strData As String = dsPendencias.Tables(0).Rows(0).Item("Data")
        Dim strHora As String = dsPendencias.Tables(0).Rows(0).Item("Hora")
        Dim strObservacoesFiscais As String = dsPendencias.Tables(0).Rows(0).Item("ObservacoesFiscais")

        sb.Append("[IDE]" & ControlChars.CrLf)
        sb.Append("   CUF	= " & dsEmpresa.Tables(0).Rows(0).Item("Estadoibge") & ControlChars.CrLf)
        sb.Append("   CCT	= " & Format(CInt(CInt(Mid(strEmpresa, 1, 5)) * CInt(strConhecimento) + CDate(strHora).Hour + CDate(strHora).Minute - CDate(strHora).Second), "000000000") & ControlChars.CrLf)
        sb.Append("   CFOP  = " & dsItensConhecimento.Tables(0).Rows(0).Item("CFOP") & ControlChars.CrLf)
        sb.Append("   NATOP = " & Funcoes.EliminarCaracteresEspeciais(dsItensConhecimento.Tables(0).Rows(0).Item("DescricaoFiscal")) & ControlChars.CrLf)

        If dsConhecimento.Tables(0).Rows(0).Item("CIFFOB").ToString().Trim() = "CIF" Then
            sb.Append("   FORPAG = 0" & ControlChars.CrLf)
        Else
            sb.Append("   FORPAG = 1" & ControlChars.CrLf)
        End If

        sb.Append("   MOD	     = 57" & ControlChars.CrLf)

        sb.Append("   SERIE      = " & strSerie & ControlChars.CrLf)
        sb.Append("   NCT	     = " & strConhecimento & ControlChars.CrLf)
        sb.Append("   DHEMI      = " & CDate(dsConhecimento.Tables(0).Rows(0).Item("Movimento")).AddHours(-1).ToString("yyyy-MM-dd") & "T" & strHora & ControlChars.CrLf)
        sb.Append("   TPIMP      = 1" & ControlChars.CrLf)
        sb.Append("   TPEMIS     = 1" & ControlChars.CrLf)
        sb.Append("   TPCTE      = 0" & ControlChars.CrLf)
        sb.Append("   VERPROC    = v.2.3" & ControlChars.CrLf)
        sb.Append("   CMUNENV    = " & dsEmpresa.Tables(0).Rows(0).Item("Estadoibge") & Format(dsEmpresa.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
        sb.Append("   XMUNENV    = " & Funcoes.EliminarCaracteresEspeciais(dsEmpresa.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
        sb.Append("   UFENV      = " & dsEmpresa.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        sb.Append("   MODAL      = 01" & ControlChars.CrLf)
        sb.Append("   TPSERV     = 0" & ControlChars.CrLf)
        sb.Append("   CMUNINI    = " & dsOrigemMercadoria.Tables(0).Rows(0).Item("Estadoibge") & Format(dsOrigemMercadoria.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
        sb.Append("   XMUNINI    = " & Funcoes.EliminarCaracteresEspeciais(dsOrigemMercadoria.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
        sb.Append("   UFINI      = " & dsOrigemMercadoria.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        sb.Append("   CMUNFIM    = " & dsCliente.Tables(0).Rows(0).Item("Estadoibge") & Format(dsCliente.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
        sb.Append("   XMUNFIM    = " & Funcoes.EliminarCaracteresEspeciais(dsCliente.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
        sb.Append("   UFFIM      = " & dsCliente.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        sb.Append("   RETIRA     = 1" & ControlChars.CrLf)
        sb.Append("   XDETRETIRA = " & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[TOMA03]" & ControlChars.CrLf)
        If dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEntradaSaida_id") = "S" Then
            sb.Append("   TOMA = 0" & ControlChars.CrLf)
        Else
            sb.Append("   TOMA = 1" & ControlChars.CrLf)
        End If

        sb.Append(ControlChars.CrLf)
        sb.Append("[COMPL]" & ControlChars.CrLf)
        sb.Append("   XCARACAD  = ENTREGA" & ControlChars.CrLf)
        sb.Append("   XCARACSER = CONVENCIONAL" & ControlChars.CrLf)
        sb.Append("   XEMI      = " & Funcoes.EliminarCaracteresEspeciais(dsConhecimento.Tables(0).Rows(0).Item("UsuarioInclusao")) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[EMIT]" & ControlChars.CrLf)
        sb.Append("   CNPJ   = " & strEmpresa & ControlChars.CrLf)
        sb.Append("   IE     = " & dsEmpresa.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
        sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(dsEmpresa.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
        sb.Append("   XFANT  = " & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[ENDEREMIT]" & ControlChars.CrLf)
        sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(dsEmpresa.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
        sb.Append("   NRO	  = " & dsEmpresa.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
        sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(dsEmpresa.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
        sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(dsEmpresa.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
        sb.Append("   CMUN    = " & dsEmpresa.Tables(0).Rows(0).Item("Estadoibge") & Format(dsEmpresa.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
        sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(dsEmpresa.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
        sb.Append("   CEP	  = " & dsEmpresa.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        sb.Append("   UF	  = " & dsEmpresa.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[REM]" & ControlChars.CrLf)
        If dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEntradaSaida_id") = "S" Then
            If Not HomologAlfasig Then
                sb.Append("   CNPJ   = " & dsRemetente.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                sb.Append("   IE     = " & dsRemetente.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
            Else
                sb.Append("   CNPJ   = 99999999000191" & ControlChars.CrLf)
                sb.Append("   IE     = ISENTO" & ControlChars.CrLf)
                sb.Append("   XNOME  = CT-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
            End If
            sb.Append(ControlChars.CrLf)

            sb.Append("[ENDERREME]" & ControlChars.CrLf)
            sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
            sb.Append("   NRO	  = " & dsRemetente.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
            sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
            sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
            sb.Append("   CMUN    = " & dsRemetente.Tables(0).Rows(0).Item("Estadoibge") & Format(dsRemetente.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
            sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
            sb.Append("   CEP	  = " & dsRemetente.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            sb.Append("   UF	  = " & dsRemetente.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
            sb.Append("   CPAIS   = " & dsRemetente.Tables(0).Rows(0).Item("Pais") & ControlChars.CrLf)
            sb.Append("   XPAIS   = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("NomePais")) & ControlChars.CrLf)
        Else
            If dsConhecimento.Tables(0).Rows(0).Item("LocalEmbarque").ToString.Length > 0 Then
                dsLocalEmbarque = getDataSetCliente(dsConhecimento.Tables(0).Rows(0).Item("LocalEmbarque"), dsConhecimento.Tables(0).Rows(0).Item("EndLocalEmbarque"))
            Else
                dsLocalEmbarque = getDataSetCliente(dsCliente.Tables(0).Rows(0).Item("Cliente"), dsCliente.Tables(0).Rows(0).Item("EndCliente"))
            End If

            If Not HomologAlfasig Then
                If dsLocalEmbarque.Tables(0).Rows(0).Item("Cliente").ToString.Length > 11 Then
                    sb.Append("   CNPJ  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                    sb.Append("   IE	= " & dsLocalEmbarque.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
                Else
                    sb.Append("   CPF   = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                    sb.Append("   IE	= ISENTO" & ControlChars.CrLf)
                End If
                sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
            Else
                sb.Append("   CNPJ   = 99999999000191" & ControlChars.CrLf)
                sb.Append("   IE     = ISENTO" & ControlChars.CrLf)
                sb.Append("   XNOME  = CT-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
            End If
            sb.Append(ControlChars.CrLf)

            sb.Append("[ENDERREME]" & ControlChars.CrLf)
            sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
            sb.Append("   NRO	  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
            sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
            sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
            sb.Append("   CMUN    = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Estadoibge") & Format(dsLocalEmbarque.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
            sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
            sb.Append("   CEP	  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            sb.Append("   UF	  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
            sb.Append("   CPAIS   = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Pais") & ControlChars.CrLf)
            sb.Append("   XPAIS   = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("NomePais")) & ControlChars.CrLf)
        End If

        sb.Append(ControlChars.CrLf)
        For i = 0 To dsConhecimentoXNota.Tables(0).Rows.Count - 1
            sb.Append("[INFNFE]" & ControlChars.CrLf)
            sb.Append("   CHAVE  = " & dsConhecimentoXNota.Tables(0).Rows(i).Item("ChaveNfe") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        Next

        If dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEntradaSaida_id") = "S" Then
            dsLocalEmbarque = getDataSetCliente(dsConhecimento.Tables(0).Rows(0).Item("LocalEmbarque"), dsConhecimento.Tables(0).Rows(0).Item("EndLocalEmbarque"))
        Else
            dsLocalEmbarque = getDataSetCliente(dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEmpresa_id"), dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEndEmpresa_id"))
        End If

        sb.Append("[EXPED]" & ControlChars.CrLf)
        If Not HomologAlfasig Then
            If dsLocalEmbarque.Tables(0).Rows(0).Item("Cliente").ToString.Length > 11 Then
                sb.Append("   CNPJ = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                sb.Append("   IE   = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            Else
                sb.Append("   CPF  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                sb.Append("   IE   = ISENTO" & ControlChars.CrLf)
            End If
            sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
        Else
            sb.Append("   CNPJ   = 99999999000191" & ControlChars.CrLf)
            sb.Append("   IE     = ISENTO" & ControlChars.CrLf)
            sb.Append("   XNOME  = CT-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
        End If
        sb.Append(ControlChars.CrLf)

        sb.Append("[ENDEREXPED]" & ControlChars.CrLf)
        sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
        sb.Append("   NRO	  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
        sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
        sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
        sb.Append("   CMUN    = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Estadoibge") & Format(dsLocalEmbarque.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
        sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
        sb.Append("   CEP	  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        sb.Append("   UF	  = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        sb.Append("   CPAIS   = " & dsLocalEmbarque.Tables(0).Rows(0).Item("Pais") & ControlChars.CrLf)
        sb.Append("   XPAIS   = " & Funcoes.EliminarCaracteresEspeciais(dsLocalEmbarque.Tables(0).Rows(0).Item("NomePais")) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        If dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEntradaSaida_id") = "S" Then
            dsDestino = getDataSetCliente(dsConhecimento.Tables(0).Rows(0).Item("Destino"), dsConhecimento.Tables(0).Rows(0).Item("EndDestino"))
        Else
            dsDestino = getDataSetCliente(dsConhecimento.Tables(0).Rows(0).Item("Deposito"), dsConhecimento.Tables(0).Rows(0).Item("EndDeposito"))
        End If

        sb.Append("[RECEB]" & ControlChars.CrLf)
        If Not HomologAlfasig Then
            If dsDestino.Tables(0).Rows(0).Item("Cliente").ToString.Length > 11 Then
                sb.Append("   CNPJ   = " & dsDestino.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
            Else
                sb.Append("    CPF   = " & dsDestino.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
            End If
            If dsDestino.Tables(0).Rows(0).Item("Inscricao").ToString.Length > 0 Then sb.Append("   IE	  = " & dsDestino.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(dsDestino.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
        Else
            sb.Append("   CNPJ   = 99999999000191" & ControlChars.CrLf)
            sb.Append("   IE     = ISENTO" & ControlChars.CrLf)
            sb.Append("   XNOME  = CT-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
        End If
        sb.Append(ControlChars.CrLf)

        sb.Append("[ENDERRECEB]" & ControlChars.CrLf)
        sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(dsDestino.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
        sb.Append("   NRO	  = " & dsDestino.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
        sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(dsDestino.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
        sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(dsDestino.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
        sb.Append("   CMUN    = " & dsDestino.Tables(0).Rows(0).Item("Estadoibge") & Format(dsDestino.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
        sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(dsDestino.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
        sb.Append("   CEP	  = " & dsDestino.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        sb.Append("   UF	  = " & dsDestino.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        sb.Append("   CPAIS   = " & dsDestino.Tables(0).Rows(0).Item("Pais") & ControlChars.CrLf)
        sb.Append("   XPAIS   = " & Funcoes.EliminarCaracteresEspeciais(dsDestino.Tables(0).Rows(0).Item("NomePais")) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[DEST]" & ControlChars.CrLf)
        If dsConhecimentoXNota.Tables(0).Rows(0)("OrigemEntradaSaida_id") = "S" Then
            If Not HomologAlfasig Then
                If dsCliente.Tables(0).Rows(0).Item("Cliente").ToString.Length > 11 Then
                    sb.Append("   CNPJ  = " & dsCliente.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                    sb.Append("   IE	= " & dsCliente.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
                Else
                    sb.Append("    CPF = " & dsCliente.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                End If
                sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(dsCliente.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
            Else
                sb.Append("   CNPJ   = 99999999000191" & ControlChars.CrLf)
                sb.Append("   IE     = ISENTO" & ControlChars.CrLf)
                sb.Append("   XNOME  = CT-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
            End If
            sb.Append(ControlChars.CrLf)

            sb.Append("[ENDERDEST]" & ControlChars.CrLf)
            sb.Append("   XLGR    = " & Funcoes.EliminarCaracteresEspeciais(dsCliente.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
            sb.Append("   NRO	  = " & dsCliente.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
            sb.Append("   XCPL    = " & Funcoes.EliminarCaracteresEspeciais(dsCliente.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
            sb.Append("   XBAIRRO = " & Funcoes.EliminarCaracteresEspeciais(dsCliente.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
            sb.Append("   CMUN    = " & dsCliente.Tables(0).Rows(0).Item("Estadoibge") & Format(dsCliente.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
            sb.Append("   XMUN    = " & Funcoes.EliminarCaracteresEspeciais(dsCliente.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
            sb.Append("   CEP	  = " & dsCliente.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            sb.Append("   UF	  = " & dsCliente.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
            sb.Append("   CPAIS   = " & dsCliente.Tables(0).Rows(0).Item("Pais") & ControlChars.CrLf)
            sb.Append("   XPAIS   = " & Funcoes.EliminarCaracteresEspeciais(dsCliente.Tables(0).Rows(0).Item("NomePais")) & ControlChars.CrLf)
        Else
            If Not HomologAlfasig Then
                sb.Append("   CNPJ   = " & dsRemetente.Tables(0).Rows(0).Item("Cliente") & ControlChars.CrLf)
                sb.Append("   IE     = " & dsRemetente.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
            Else
                sb.Append("   CNPJ   = 99999999000191" & ControlChars.CrLf)
                sb.Append("   IE     = ISENTO" & ControlChars.CrLf)
                sb.Append("   XNOME  = CT-E EMITIDO EM AMBIENTE DE HOMOLOGACAO - SEM VALOR FISCAL" & ControlChars.CrLf)
            End If
            sb.Append(ControlChars.CrLf)

            sb.Append("[ENDERDEST]" & ControlChars.CrLf)
            sb.Append("   XLGR   = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
            sb.Append("   NRO	  = " & dsRemetente.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
            sb.Append("   XCPL   = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
            sb.Append("   XBAIRRO= " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
            sb.Append("   CMUN   = " & dsRemetente.Tables(0).Rows(0).Item("Estadoibge") & Format(dsRemetente.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
            sb.Append("   XMUN   = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
            sb.Append("   CEP	  = " & dsRemetente.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            sb.Append("   UF	  = " & dsRemetente.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
            sb.Append("   CPAIS  = " & dsRemetente.Tables(0).Rows(0).Item("Pais") & ControlChars.CrLf)
            sb.Append("   XPAIS  = " & Funcoes.EliminarCaracteresEspeciais(dsRemetente.Tables(0).Rows(0).Item("NomePais")) & ControlChars.CrLf)
        End If

        sb.Append(ControlChars.CrLf)
        sb.Append("[VPREST]" & ControlChars.CrLf)
        valor = CDec(dsItensConhecimento.Tables(0).Rows(0).Item("Valor")).ToString("N2")
        sb.Append("  VTPREST = " & valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        valor = CDec(dsItensConhecimento.Tables(0).Rows(0).Item("Valor")).ToString("N2")
        sb.Append("     VREC = " & valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[COMP]" & ControlChars.CrLf)
        sb.Append("   XNOME  = FRETE VALOR" & ControlChars.CrLf)
        valor = CDec(dsItensConhecimento.Tables(0).Rows(0).Item("Valor")).ToString("N2")
        sb.Append("   VCOMP  = " & valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[IMP]" & ControlChars.CrLf)
        sb.Append("   INFADFISCO = " & Funcoes.EliminarCaracteresEspeciais(strObservacoesFiscais) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        For Each drEncargos As DataRow In dsEncargosConhecimento.Tables(0).Rows
            situacaoTributaria = drEncargos("SituacaoTributaria")
            If drEncargos("Encargo").ToString.Contains("ICMS") Then
                baseICMS = drEncargos("Base").ToString.Replace(",", ".")
                pBaseICMS = drEncargos("PercentualBase").ToString.Replace(",", ".")
                aliqICMS = drEncargos("Percentual").ToString.Replace(",", ".")
                valorICMS = drEncargos("Valor").ToString.Replace(",", ".")
            End If
        Next

        Select Case situacaoTributaria
            Case 0
                sb.Append("[ICMS00]" & ControlChars.CrLf)
                sb.Append("   CST    = 00" & ControlChars.CrLf)
                sb.Append("   VBC    = " & baseICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   PICMS  = " & aliqICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VICMS  = " & valorICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
            Case 20
                sb.Append("[ICMS20]" & ControlChars.CrLf)
                sb.Append("   CST    = 20" & ControlChars.CrLf)
                sb.Append("PREDBC    = " & pBaseICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VBC    = " & baseICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   PICMS  = " & aliqICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                sb.Append("   VICMS  = " & valorICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
            Case 40
                sb.Append("[ICMS45]" & ControlChars.CrLf)
                sb.Append("   CST    = 40" & ControlChars.CrLf)
            Case 41
                sb.Append("[ICMS45]" & ControlChars.CrLf)
                sb.Append("   CST    = 41" & ControlChars.CrLf)
            Case 51
                sb.Append("[ICMS45]" & ControlChars.CrLf)
                sb.Append("   CST    = 51" & ControlChars.CrLf)
            Case 90
                sb.Append("[ICMS90]" & ControlChars.CrLf)
                sb.Append("   CST    = 90" & ControlChars.CrLf)
                If baseICMS.ToString.Length = 0 Then
                    sb.Append("   VBC   = 0" & ControlChars.CrLf)
                    sb.Append("   PICMS = 0" & ControlChars.CrLf)
                    sb.Append("   VICMS = 0" & ControlChars.CrLf)
                Else
                    sb.Append("   VBC   = " & baseICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("   PICMS = " & aliqICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                    sb.Append("   VICMS = " & valorICMS.ToString.Replace(",", ".") & ControlChars.CrLf)
                End If
        End Select

        sb.Append(ControlChars.CrLf)
        sb.Append("[INFCTENORM]" & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)
        sb.Append("   [INFCARGA]" & ControlChars.CrLf)

        For i = 0 To dsConhecimentoXNota.Tables(0).Rows.Count - 1
            valorMercadoria = valorMercadoria + dsConhecimentoXNota.Tables(0).Rows(i).Item("ValorMercadoria")
        Next
        valor = valorMercadoria.ToString("N2")

        sb.Append("      VCARGA  = " & valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append("      PROPRED = " & Funcoes.EliminarCaracteresEspeciais(dsConhecimentoXNota.Tables(0).Rows(0).Item("NomeProduto")) & ControlChars.CrLf)
        sb.Append("      XOUTCAT = " & Funcoes.EliminarCaracteresEspeciais(dsConhecimentoXNota.Tables(0).Rows(0).Item("EmbalagemProduto")) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        For i = 0 To dsConhecimentoXNota.Tables(0).Rows.Count - 1
            sb.Append("   [INFQ]" & ControlChars.CrLf)
            If dsConhecimentoXNota.Tables(0).Rows(i).Item("UnidadeProduto") = "MT3" Then
                sb.Append("      CUNID = 00" & ControlChars.CrLf)
            ElseIf dsConhecimentoXNota.Tables(0).Rows(i).Item("UnidadeProduto") = "KGS" Then
                sb.Append("      CUNID = 01" & ControlChars.CrLf)
            ElseIf dsConhecimentoXNota.Tables(0).Rows(i).Item("UnidadeProduto") = "LTS" Then
                sb.Append("      CUNID = 04" & ControlChars.CrLf)
            ElseIf dsConhecimentoXNota.Tables(0).Rows(i).Item("UnidadeProduto") = "UND" Then
                sb.Append("      CUNID = 03" & ControlChars.CrLf)
            ElseIf dsConhecimentoXNota.Tables(0).Rows(i).Item("UnidadeProduto") = "TON" Then
                sb.Append("      CUNID = 02" & ControlChars.CrLf)
            ElseIf dsConhecimentoXNota.Tables(0).Rows(i).Item("UnidadeProduto") = "SCS" Then
                sb.Append("      CUNID = 03" & ControlChars.CrLf)
            Else
                sb.Append("      CUNID =" & ControlChars.CrLf)
            End If
            sb.Append("      TPMED = " & dsConhecimentoXNota.Tables(0).Rows(i).Item("UnidadeProduto") & ControlChars.CrLf)
            valor = CDec(dsConhecimentoXNota.Tables(0).Rows(i).Item("QuantidadeFiscal")).ToString("N4")
            sb.Append("      QCARGA = " & valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        Next

        'INFORMAÇÃO TEMPORÁRIA ATÉ COLOCARMOS NA TELA
        sb.Append("[SEG]" & ControlChars.CrLf)
        sb.Append("   RESPSEG = 4" & ControlChars.CrLf)
        sb.Append("   XSEG = AGRICOLA ALVORADA LTDA" & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[RODO]" & ControlChars.CrLf)
        sb.Append("    RNTRC = 11111111" & ControlChars.CrLf)
        sb.Append("    DPREV = " & CDate(strData).AddDays(5).ToString("yyyy-MM-dd") & ControlChars.CrLf)
        sb.Append("     LOTA = 1" & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        If dsTransportador.Tables(0).Rows(0).Item("Placa03") <> "" Then
            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 0" & ControlChars.CrLf)
            sb.Append("    TPROD = 03" & ControlChars.CrLf)
            sb.Append("    TPCAR = 00" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[PROP]" & ControlChars.CrLf)
            If dsTransportador.Tables(0).Rows(0).Item("Proprietario").ToString.Length > 11 Then
                sb.Append("   CNPJ = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
                If dsTransportador.Tables(0).Rows(0).Item("IETransportador").ToString.Length > 0 Then
                    sb.Append("   IE = " & dsTransportador.Tables(0).Rows(0).Item("IETransportador").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                    sb.Append("   UF = " & dsTransportador.Tables(0).Rows(0).Item("UFTransportador") & ControlChars.CrLf)
                End If
            Else
                sb.Append("   CPF = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
            End If

            If dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length = 8 Then
                sb.Append("   RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador") & ControlChars.CrLf)
            ElseIf dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length = 8 Then
                sb.Append("   RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca") & ControlChars.CrLf)
            Else
                sb.Append("   RNTRC = 11111111" & ControlChars.CrLf)
            End If

            sb.Append("    XNOME = " & Funcoes.EliminarCaracteresEspeciais(dsTransportador.Tables(0).Rows(0).Item("NomeTransportador")) & ControlChars.CrLf)
            sb.Append("   TPPROP = 1" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa01").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 1" & ControlChars.CrLf)
            sb.Append("    TPROD = 05" & ControlChars.CrLf)
            sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca01") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa02").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 1" & ControlChars.CrLf)
            sb.Append("    TPROD = 05" & ControlChars.CrLf)
            sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca02") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa03").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 1" & ControlChars.CrLf)
            sb.Append("    TPROD = 05" & ControlChars.CrLf)
            sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca03") & ControlChars.CrLf)
        ElseIf dsTransportador.Tables(0).Rows(0).Item("Placa02") <> "" Then
            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 0" & ControlChars.CrLf)
            sb.Append("    TPROD = 03" & ControlChars.CrLf)
            sb.Append("    TPCAR = 00" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[PROP]" & ControlChars.CrLf)
            If dsTransportador.Tables(0).Rows(0).Item("Proprietario").ToString.Length > 11 Then
                sb.Append("     CNPJ = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
                If dsTransportador.Tables(0).Rows(0).Item("IETransportador").ToString.Length > 0 Then
                    sb.Append("       IE = " & dsTransportador.Tables(0).Rows(0).Item("IETransportador").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                    sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("UFTransportador") & ControlChars.CrLf)
                End If
            Else
                sb.Append("      CPF = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
            End If

            If dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador") & ControlChars.CrLf)
            ElseIf dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca") & ControlChars.CrLf)
            Else
                sb.Append("    RNTRC = 11111111" & ControlChars.CrLf)
            End If

            sb.Append("    XNOME = " & Funcoes.EliminarCaracteresEspeciais(dsTransportador.Tables(0).Rows(0).Item("NomeTransportador")) & ControlChars.CrLf)
            sb.Append("   TPPROP = 1" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa01").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 1" & ControlChars.CrLf)
            sb.Append("    TPROD = 05" & ControlChars.CrLf)
            sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca01") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa02").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 1" & ControlChars.CrLf)
            sb.Append("    TPROD = 05" & ControlChars.CrLf)
            sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca02") & ControlChars.CrLf)
        ElseIf dsTransportador.Tables(0).Rows(0).Item("Placa01") <> "" Then
            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 0" & ControlChars.CrLf)
            sb.Append("    TPROD = 03" & ControlChars.CrLf)
            sb.Append("    TPCAR = 00" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[PROP]" & ControlChars.CrLf)
            If dsTransportador.Tables(0).Rows(0).Item("Proprietario").ToString.Length > 11 Then
                sb.Append("     CNPJ = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
                If dsTransportador.Tables(0).Rows(0).Item("IETransportador").ToString.Length > 0 Then
                    sb.Append("       IE = " & dsTransportador.Tables(0).Rows(0).Item("IETransportador").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                    sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("UFTransportador") & ControlChars.CrLf)
                End If
            Else
                sb.Append("      CPF = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
            End If

            If dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador") & ControlChars.CrLf)
            ElseIf dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca") & ControlChars.CrLf)
            Else
                sb.Append("    RNTRC = 11111111" & ControlChars.CrLf)
            End If

            sb.Append("    XNOME = " & Funcoes.EliminarCaracteresEspeciais(dsTransportador.Tables(0).Rows(0).Item("NomeTransportador")) & ControlChars.CrLf)
            sb.Append("   TPPROP = 1" & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa01").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 1" & ControlChars.CrLf)
            sb.Append("    TPROD = 05" & ControlChars.CrLf)
            sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca01") & ControlChars.CrLf)
        Else
            sb.Append("[VEIC]" & ControlChars.CrLf)
            sb.Append("  RENAVAM = 000000000" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa").Replace("-", "") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)
            sb.Append("   TPPROP = T" & ControlChars.CrLf)
            sb.Append("   TPVEIC = 0" & ControlChars.CrLf)
            sb.Append("    TPROD = 03" & ControlChars.CrLf)
            sb.Append("    TPCAR = 00" & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca") & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)

            sb.Append("[prop]" & ControlChars.CrLf)
            If dsTransportador.Tables(0).Rows(0).Item("Proprietario").ToString.Length > 11 Then
                sb.Append("     CNPJ = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
                If dsTransportador.Tables(0).Rows(0).Item("IETransportador").ToString.Length > 0 Then
                    sb.Append("       IE = " & dsTransportador.Tables(0).Rows(0).Item("IETransportador").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                    sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("UFTransportador") & ControlChars.CrLf)
                End If
            Else
                sb.Append("      CPF = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
            End If

            If dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador") & ControlChars.CrLf)
            ElseIf dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca") & ControlChars.CrLf)
            Else
                sb.Append("    RNTRC = 11111111" & ControlChars.CrLf)
            End If

            sb.Append("    XNOME = " & Funcoes.EliminarCaracteresEspeciais(dsTransportador.Tables(0).Rows(0).Item("NomeTransportador")) & ControlChars.CrLf)
            sb.Append("   TPPROP = 1" & ControlChars.CrLf)
        End If

        sb.Append(ControlChars.CrLf)
        sb.Append("[MOTO]" & ControlChars.CrLf)
        sb.Append("    XNOME = " & Funcoes.EliminarCaracteresEspeciais(dsTransportador.Tables(0).Rows(0).Item("NomeMotorista")) & ControlChars.CrLf)
        sb.Append("      CPF = " & dsTransportador.Tables(0).Rows(0).Item("Motorista"))
        Return sb.ToString()
    End Function

    Private Function getTextoDACTE(ByVal cte As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        If cte.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Anulacao Then
            sb.Append("CHAVENFE=" & cte.ChaveNFE & ControlChars.CrLf)
        Else
            sb.Append("CHAVECTE=" & cte.ChaveNFE & ControlChars.CrLf)
        End If
        sb.Append("IMPRESSORA=pdf_nfe" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function getTextoCancelar(ByVal cte As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVECTE=" & cte.ChaveNFE & ControlChars.CrLf)
        sb.Append("NPROTOCOLO =" & cte.ProtocoloNota & ControlChars.CrLf)
        sb.Append("JUSTIFICATIVA=" & Funcoes.EliminarCaracteresEspeciaisNF(cte.ObservacaoCancelamento) & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function getTextoEmail(ByVal cte As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVECTE=" & cte.ChaveNFE & ControlChars.CrLf)
        sb.Append("DESTINATARIO=" & cte.Cliente.EmailNFE & ControlChars.CrLf)
        sb.Append("ASSUNTO=" & "CT-e Nr. " & cte.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & "Envio CT-e Nr. " & cte.Codigo & ControlChars.CrLf)
        sb.Append("EMAILEMITENTE=" & cte.Empresa.EmailNFE & ControlChars.CrLf)
        sb.Append("NOMEEMITENTE=" & cte.Empresa.Nome & ControlChars.CrLf)
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & "NAO" & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function getTextoEmail(ByVal cte As [Lib].Negocio.NotaFiscal, ByVal lstClientes As [Lib].Negocio.ListCliente, ByVal strAssunto As String, ByVal strMensagem As String, ByVal compactado As Boolean) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVECTE=" & cte.ChaveNFE & ControlChars.CrLf)
        For Each objCliente As [Lib].Negocio.Cliente In lstClientes.Where(Function(s) Not String.IsNullOrWhiteSpace(s.EmailNFE)).ToList()
            sb.Append("DESTINATARIO=" & objCliente.EmailNFE & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & strAssunto & " - CT-e Nr. " & cte.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & strMensagem & " - Envio CT-e Nr. " & cte.Codigo & ControlChars.CrLf)
        sb.Append("EMAILEMITENTE=" & cte.Empresa.EmailNFE & ControlChars.CrLf)
        sb.Append("NOMEEMITENTE=" & cte.Empresa.Nome & ControlChars.CrLf)
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & IIf(compactado, "SIM", "NAO") & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Function IsEnabled(ByVal cte As [Lib].Negocio.NotaFiscal) As Boolean
        Dim sql As String = "SELECT * FROM NFEPendencias " & vbCrLf &
            "WHERE 1=1 " & vbCrLf &
            "AND Empresa_Id = '" & cte.CodigoEmpresa & "' " & vbCrLf &
            "AND EndEmpresa_Id = '" & cte.EnderecoEmpresa & "' " & vbCrLf &
            "AND Cliente_Id = '" & cte.CodigoCliente & "' " & vbCrLf &
            "AND EndCliente_Id = '" & cte.EnderecoCliente & "' " & vbCrLf &
            "AND EntradaSaida_Id = '" & IIf(cte.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
            "AND Serie_Id = '" & cte.Serie & "' " & vbCrLf &
            "AND Nota_Id = '" & cte.Codigo & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NFEPendencias")
        Return ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0
    End Function

    Private Function IsDuplicar(ByVal origem As [Lib].Negocio.NotaFiscal) As Boolean
        If origem.NotaTrocaOrigem Is Nothing Then Return False

        Dim sql As String = "SELECT * FROM NotasFiscais " & vbCrLf &
            "WHERE 1=1 " & vbCrLf &
            "AND Empresa_Id = '" & origem.NotaTrocaOrigem.CodigoEmpresa & "' " & vbCrLf &
            "AND EndEmpresa_Id = '" & origem.NotaTrocaOrigem.EnderecoEmpresa & "' " & vbCrLf &
            "AND Cliente_Id = '" & origem.CodigoEmpresa & "' " & vbCrLf &
            "AND EndCliente_Id = '" & origem.EnderecoEmpresa & "' " & vbCrLf &
            "AND EntradaSaida_Id = '" & "E" & "' " & vbCrLf &
            "AND Serie_Id = '" & origem.Serie & "' " & vbCrLf &
            "AND Nota_Id = '" & origem.Codigo & "'" & vbCrLf &
            "AND TipoDeDocumento = '" & CInt(eTipoDeDocumento.CTRC) & "'" & vbCrLf &
            "AND TipoDeDocumentoFrete = '" & CInt(eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro) & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NotasFiscais")
        Return ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0
    End Function

    Private Function Duplicar(ByVal cte As [Lib].Negocio.NotaFiscal) As Boolean
        If cte Is Nothing Then
            MsgBox(Me.Page, "Não foi possível incluir a nota de entrada para o conhecimento de circulação!")
            Return False
        End If

        Dim Sqls As New ArrayList
        Dim obj As New [Lib].Negocio.NotaFiscal()
        Dim op As New [Lib].Negocio.OperacoesDeFretes(cte.CodigoOperacao, cte.CodigoSubOperacao)
        If op.OpContrapartida Is Nothing OrElse op.SubOpContrapartida Is Nothing Then
            'MsgBox(Me.Page, "Não foi possível encontrar a operação e suboperação de contrapartida!")
            Return False
        End If

        obj.IUD = "I"
        obj.CarregandoNota = True
        obj.CodigoEmpresa = objConhecimento.NotaTrocaOrigem.CodigoEmpresa
        obj.EnderecoEmpresa = objConhecimento.NotaTrocaOrigem.EnderecoEmpresa
        obj.CodigoUnidadeDeNegocio = objConhecimento.NotaTrocaOrigem.CodigoUnidadeDeNegocio
        obj.CodigoSafra = cte.CodigoSafra
        obj.CodigoCliente = cte.CodigoEmpresa
        obj.EnderecoCliente = cte.EnderecoEmpresa
        obj.EntradaSaida = eEntradaSaida.Entrada
        obj.Serie = cte.Serie
        obj.Codigo = cte.Codigo
        obj.CodigoTipoDeDocumento = eTipoDeDocumento.CTRC
        obj.Formulario = cte.Formulario
        'Não deve gravar o pedido da Nota Fiscal - Furlan - 27/07/2015
        'obj.CodigoPedido = cte.CodigoPedido
        obj.CodigoProcuracao = obj.CodigoProcuracao
        obj.CodigoOperacao = op.OpContrapartida
        obj.CodigoSubOperacao = op.SubOpContrapartida
        obj.CodigoFinalidade = cte.CodigoFinalidade
        obj.Movimento = cte.Movimento
        obj.DataNota = cte.DataNota
        obj.NossaEmissao = False
        obj.SerieNotaProdutor = cte.SerieNotaProdutor
        obj.NotaProdutor = cte.NotaProdutor
        obj.CodigoDeposito = objConhecimento.NotaTrocaOrigem.CodigoEmpresa
        obj.EnderecoDeposito = objConhecimento.NotaTrocaOrigem.EnderecoEmpresa
        obj.CodigoDestino = cte.CodigoEmpresa
        obj.EnderecoDestino = cte.EnderecoEmpresa
        obj.CodigoTransbordo = cte.CodigoTransbordo
        obj.EnderecoTransbordo = cte.EnderecoTransbordo
        obj.CodigoAgenciador = cte.CodigoAgenciador
        obj.EnderecoAgenciador = cte.EnderecoAgenciador
        obj.CIFFOB = eTiposFrete.FOB
        obj.Observacoes = cte.Observacoes
        obj.ObservacoesDeEmbarque = cte.ObservacoesDeEmbarque
        obj.ObservacoesDoProduto = cte.ObservacoesDoProduto
        obj.ObservacoesControleInterno = cte.ObservacoesControleInterno
        obj.CodigoSituacao = CInt(eSituacao.Normal)
        obj.CodigoTransportador = cte.CodigoTransportador
        obj.EnderecoTransportador = cte.EnderecoTransportador
        obj.PlacaTransportador = cte.PlacaTransportador
        obj.CodigoMotorista = cte.CodigoMotorista
        obj.EnderecoMotorista = cte.EnderecoMotorista
        obj.CodigoRomaneio = cte.CodigoRomaneio
        obj.CodigoAutorizacao = obj.CodigoAutorizacao
        obj.UsuarioInclusao = UsuarioServidor.NomeUsuario
        obj.DataInclusao = DateTime.Now
        obj.CodigoLocalEmbarque = cte.CodigoLocalEmbarque
        obj.EndLocalEmbarque = cte.EndLocalEmbarque
        obj.DataTermino = cte.DataTermino
        obj.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro
        obj.CodigoFavorecido = cte.CodigoFavorecido
        obj.EnderecoFavorecido = cte.EnderecoFavorecido
        obj.NFG = cte.NFG
        obj.CarregandoNota = False

        Dim objNotaFiscalXItem As New [Lib].Negocio.NotaFiscalXItem(obj)

        With objNotaFiscalXItem
            .CodigoProduto = cte.Itens(0).Produto.Codigo
            'Não deve gravar o pedido da Nota Fiscal - Furlan - 27/07/2015
            '.CodigoPedido = cte.CodigoPedido
            .PesoQuantidade = cte.Itens(0).Produto.PesoQuantidade
            .QuantidadeFiscal = cte.Itens(0).QuantidadeFiscal
            .QuantidadeFisica = cte.Itens(0).QuantidadeFisica
            .Unitario = cte.Itens(0).Unitario
            .NotaFiscal.CarregandoItens = True
            .ValorTotal = cte.Itens(0).ValorTotal
        End With

        objNotaFiscalXItem.CodigoOperacao = obj.CodigoOperacao
        objNotaFiscalXItem.CodigoSubOperacao = obj.CodigoSubOperacao

        objNotaFiscalXItem.CarregandoEncargos = True
        objNotaFiscalXItem.Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objNotaFiscalXItem, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
        objNotaFiscalXItem.CarregandoEncargos = False

        If CType(CInt(hdfTipo.Value), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Circulacao Then
            obj.CodigoTransportador = cte.CodigoTransportador
            obj.EnderecoTransportador = cte.EnderecoTransportador
        End If

        If Not objNotaFiscalXItem.Encargos IsNot Nothing AndAlso objNotaFiscalXItem.Encargos.Count > 0 Then
            txtUnitario.Text = 0
            MsgBox(Me.Page, "Não foi possível encontrar encargos para a operação " & obj.CodigoOperacao & " e suboperação " & obj.CodigoSubOperacao & "!")
            Exit Function
        End If

        obj.CarregandoItens = True
        obj.Itens.Clear()
        obj.Itens.Add(objNotaFiscalXItem)
        obj.AtualizaTotais()
        obj.CarregandoItens = False

        obj.NotaTrocaOrigem = cte
        If obj.NotasTrocaOrigem Is Nothing Then
            obj.NotasTrocaOrigem = New List(Of [Lib].Negocio.NotaFiscal)
        End If
        obj.NotasTrocaOrigem.Add(cte)
        obj.SalvarSql(Sqls)

        ContabilizarNota(obj, Sqls, True)

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Return False
        End If
        Return True
    End Function

    Private Function getDataSet(ByVal cte As [Lib].Negocio.NotaFiscal) As DataSet
        Dim sql As String = "-- CIRCULAÇÃO                                                          " & vbCrLf &
                "       SELECT                                                                      " & vbCrLf &
                "       	r.Empresa_Id,                                                           " & vbCrLf &
                "       	r.Conta_Id + ' - ' + pc.Titulo as Conta_Id,                             " & vbCrLf &
                "       	r.Cliente_Id,                                                           " & vbCrLf &
                "       	c.Nome,                                                                 " & vbCrLf &
                "       	r.Movimento_Id,                                                         " & vbCrLf &
                "       	r.Lote_Id,                                                              " & vbCrLf &
                "       	r.Encargo_Nf,                                                           " & vbCrLf &
                "       	r.DebitoOficial,                                                        " & vbCrLf &
                "       	r.CreditoOficial                                                        " & vbCrLf &
                "       FROM Razao r                                                                " & vbCrLf &
                "       INNER JOIN NotasFiscais nf                                                  " & vbCrLf &
                "       	 ON	nf.Cliente_Id = r.Cliente_Nf                                        " & vbCrLf &
                "       	AND nf.EndCliente_Id = r.EndCliente_Nf                                  " & vbCrLf &
                "       	AND nf.EntradaSaida_Id = r.EntradaSaida_Nf                              " & vbCrLf &
                "       	AND nf.Serie_Id = r.Serie_Nf                                            " & vbCrLf &
                "       	AND nf.Nota_Id = r.Numero_Nf                                            " & vbCrLf &
                "       	AND nf.TipoDeDocumento = 2                                              " & vbCrLf &
                "       	AND nf.TipoDeDocumentoFrete = 1                                         " & vbCrLf &
                "       LEFT JOIN PlanoDeContas pc                                                  " & vbCrLf &
                "       	ON pc.Conta_Id = r.Conta_Id                                             " & vbCrLf &
                "       LEFT JOIN Clientes c                                                        " & vbCrLf &
                "       	ON (c.Cliente_Id = r.Cliente_Id and c.Endereco_Id = r.EndCliente_Id)    " & vbCrLf &
                "       WHERE 1=1                                                                   " & vbCrLf &
                "       AND r.Numero_Nf = '" & objConhecimento.Codigo & "'                          " & vbCrLf &
                "       AND r.EntradaSaida_Nf = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                "       AND r.Serie_Nf = '" & objConhecimento.Serie & "'                            " & vbCrLf &
                "       ORDER BY r.Lote_Id, r.Sequencia_Id                                          " & vbCrLf &
                "                                                                                   " & vbCrLf &
                "-- CONTRAPARTIDA                                                                   " & vbCrLf &
                "       SELECT                                                                      " & vbCrLf &
                "       	r.Empresa_Id,                                                           " & vbCrLf &
                "       	r.Conta_Id + ' - ' + pc.Titulo as Conta_Id,                             " & vbCrLf &
                "       	r.Cliente_Id,                                                           " & vbCrLf &
                "       	c.Nome,                                                                 " & vbCrLf &
                "       	r.Movimento_Id,                                                         " & vbCrLf &
                "       	r.Lote_Id,                                                              " & vbCrLf &
                "       	r.Encargo_Nf,                                                           " & vbCrLf &
                "       	r.DebitoOficial,                                                        " & vbCrLf &
                "       	r.CreditoOficial                                                        " & vbCrLf &
                "       FROM Razao r                                                                " & vbCrLf &
                "       INNER JOIN NotasFiscais nf                                                  " & vbCrLf &
                "       	 ON	nf.Cliente_Id = r.Cliente_Nf                                        " & vbCrLf &
                "       	AND nf.EndCliente_Id = r.EndCliente_Nf                                  " & vbCrLf &
                "       	AND nf.EntradaSaida_Id = r.EntradaSaida_Nf                              " & vbCrLf &
                "       	AND nf.Serie_Id = r.Serie_Nf                                            " & vbCrLf &
                "       	AND nf.Nota_Id = r.Numero_Nf                                            " & vbCrLf &
                "       	AND nf.TipoDeDocumento = 2                                              " & vbCrLf &
                "       	AND nf.TipoDeDocumentoFrete = 5                                         " & vbCrLf &
                "       LEFT JOIN PlanoDeContas pc                                                  " & vbCrLf &
                "       	ON pc.Conta_Id = r.Conta_Id                                             " & vbCrLf &
                "       LEFT JOIN Clientes c                                                        " & vbCrLf &
                "       	ON (c.Cliente_Id = r.Cliente_Id and c.Endereco_Id = r.EndCliente_Id)    " & vbCrLf &
                "       WHERE 1=1                                                                   " & vbCrLf &
                "       AND r.Numero_Nf = '" & objConhecimento.Codigo & "'                          " & vbCrLf &
                "       AND r.EntradaSaida_Nf = 'E'                                                 " & vbCrLf &
                "       AND r.Serie_Nf = '" & objConhecimento.Serie & "'                            " & vbCrLf &
                "       ORDER BY r.Lote_Id, r.Sequencia_Id                                          " & vbCrLf &
                "                                                                                   " & vbCrLf &
                "-- COMPROVAÇÃO                                                                     " & vbCrLf &
                "       SELECT                                                                      " & vbCrLf &
                "       	r.Empresa_Id,                                                           " & vbCrLf &
                "       	r.Conta_Id + ' - ' + pc.Titulo as Conta_Id,                             " & vbCrLf &
                "       	r.Cliente_Id,                                                           " & vbCrLf &
                "       	c.Nome,                                                                 " & vbCrLf &
                "       	r.Movimento_Id,                                                         " & vbCrLf &
                "       	r.Lote_Id,                                                              " & vbCrLf &
                "       	r.Encargo_Nf,                                                           " & vbCrLf &
                "       	r.DebitoOficial,                                                        " & vbCrLf &
                "       	r.CreditoOficial                                                        " & vbCrLf &
                "       FROM NotasFiscais nf                                                        " & vbCrLf &
                "       INNER JOIN NotasXNotas nxn                                                  " & vbCrLf &
                "       	 ON nf.Empresa_Id = nxn.OrigemEmpresa_Id                                " & vbCrLf &
                "       	AND nf.EndEmpresa_Id = nxn.OrigemEndEmpresa_Id                          " & vbCrLf &
                "       	AND nf.Cliente_Id = nxn.OrigemCliente_Id                                " & vbCrLf &
                "       	AND nf.EndCliente_Id = nxn.OrigemEndCliente_Id                          " & vbCrLf &
                "       	AND nf.EntradaSaida_Id = nxn.OrigemEntradaSaida_Id                      " & vbCrLf &
                "       	AND nf.Serie_Id = nxn.OrigemSerie_Id                                    " & vbCrLf &
                "       	AND nf.Nota_Id = nxn.OrigemNota_Id                                      " & vbCrLf &
                "       INNER JOIN NotasFiscais comp                                                " & vbCrLf &
                "       	 ON comp.Empresa_Id = nxn.Empresa_Id                                    " & vbCrLf &
                "       	AND comp.EndEmpresa_Id = nxn.EndEmpresa_Id                              " & vbCrLf &
                "       	AND comp.Cliente_Id = nxn.Cliente_Id                                    " & vbCrLf &
                "       	AND comp.EndCliente_Id = nxn.EndCliente_Id                              " & vbCrLf &
                "       	AND comp.EntradaSaida_Id = nxn.EntradaSaida_Id                          " & vbCrLf &
                "       	AND comp.Serie_Id = nxn.Serie_Id                                        " & vbCrLf &
                "       	AND comp.Nota_Id = nxn.Nota_Id	                                        " & vbCrLf &
                "       	AND comp.TipoDeDocumento = 2                                            " & vbCrLf &
                "       	AND comp.TipoDeDocumentoFrete = 2                                       " & vbCrLf &
                "       INNER JOIN Razao r                                                          " & vbCrLf &
                "       	 ON	comp.Cliente_Id = r.Cliente_Nf                                      " & vbCrLf &
                "       	AND comp.EndCliente_Id = r.EndCliente_Nf                                " & vbCrLf &
                "       	AND comp.EntradaSaida_Id = r.EntradaSaida_Nf                            " & vbCrLf &
                "       	AND comp.Serie_Id = r.Serie_Nf                                          " & vbCrLf &
                "       	AND comp.Nota_Id = r.Numero_Nf                                          " & vbCrLf &
                "       LEFT JOIN PlanoDeContas pc                                                  " & vbCrLf &
                "       	ON pc.Conta_Id = r.Conta_Id                                             " & vbCrLf &
                "       LEFT JOIN Clientes c                                                        " & vbCrLf &
                "       	 ON c.Cliente_Id = r.Cliente_Id                                         " & vbCrLf &
                "       	AND c.Endereco_Id = r.EndCliente_Id                                     " & vbCrLf &
                "       WHERE 1=1                                                                   " & vbCrLf &
                "       AND nf.Nota_Id = '" & objConhecimento.Codigo & "'                           " & vbCrLf &
                "       AND nf.EntradaSaida_Id = '" & IIf(objConhecimento.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
                "       AND nf.Serie_Id = '" & objConhecimento.Serie & "'                           " & vbCrLf &
                "       ORDER BY r.Lote_Id, r.Sequencia_Id" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "CIRCULAÇÃO")
    End Function

    Private Sub viewReport()
        Try
            RecuperarSessaoConhecimento()
            If objConhecimento IsNot Nothing Then
                Dim ds As DataSet = getDataSet(objConhecimento)
                If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                    Dim parameters As New Dictionary(Of String, eTipoCampo)
                    parameters.Add("Movimento_Id", eTipoCampo.Data)
                    parameters.Add("DebitoOficial", eTipoCampo.ValorSemTotalizador)
                    parameters.Add("CreditoOficial", eTipoCampo.ValorSemTotalizador)
                    If ds.Tables.Count > 0 Then ds.Tables(0).TableName = "CIRCULAÇÃO"
                    If ds.Tables.Count > 1 Then ds.Tables(1).TableName = "CONTRAPARTIDA"
                    If ds.Tables.Count > 2 Then ds.Tables(2).TableName = "COMPROVAÇÃO"
                    Funcoes.BindExcelOffice(Me.Page, ds, "RAZÃO " & objConhecimento.Codigo, parameters, True)
                Else
                    MsgBox(Me.Page, "NENHUM RESULTADO ENCONTRADO!")
                End If
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConhecimentoDeTransporte")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Sub CalcularTotaisSelecionados()
        Dim txtPeso As TextBox
        Dim txtValor As TextBox
        Dim chkConhecimento As CheckBox

        Dim PesoTotal As Decimal = 0
        Dim ValorTotalFrete As Decimal = 0

        For Each row As GridViewRow In grdNFe.Rows
            If row.RowType = DataControlRowType.DataRow Then
                chkConhecimento = CType(grdNFe.Rows(row.RowIndex).FindControl("chkConhecimento"), CheckBox)
                If chkConhecimento.Checked Then
                    txtPeso = grdNFe.Rows(row.RowIndex).FindControl("txtPesoFiscal")
                    txtValor = grdNFe.Rows(row.RowIndex).FindControl("txtValorFrete")
                    ValorTotalFrete += CDec(txtValor.Text)
                    PesoTotal += CInt(txtPeso.Text)
                End If
            End If
        Next
        txtPesoTotalSelecionado.Text = CInt(PesoTotal)
        txtValorFreteSelecionado.Text = CDec(ValorTotalFrete)
    End Sub

    Protected Sub chkConhecimento_CheckedChanged(sender As Object, e As EventArgs)
        CalcularTotaisSelecionados()
    End Sub

    Protected Sub btnClienteRPA_Click(sender As Object, e As EventArgs) Handles btnClienteRPA.Click
        Session("ssCampo") = "LivreClasse"
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteRPA" & HID.Value, "txtNomeClienteRPA")
    End Sub

    Protected Sub btnPlaca_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPlaca.Click
        Try
            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlacaCTRC" & HID.Value)
            Popup.CenterDialog(Me.Page, "divConsultaPlacas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdVencimentos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grdVencimentos.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim txtValor As Label
            Dim txtValor2 As Label
            Dim tString As String

            txtValor = CType(e.Row.FindControl("lblSituacaoFinanceira"), Label)
            txtValor.Text = WebHelpers.GetEnumDescription(CType(txtValor.Text, eProvisao))

            txtValor2 = CType(e.Row.FindControl("lblReceberPagar"), Label)
            tString = txtValor2.Text
            txtValor2.Text = tString
        End If
    End Sub
End Class