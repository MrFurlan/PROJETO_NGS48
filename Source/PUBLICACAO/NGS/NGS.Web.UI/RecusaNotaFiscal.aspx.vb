Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.IO
Public Class RecusaNotaFiscal
    Inherits BasePage

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RecusaNotaFiscal", "ACESSAR") Then
                BuscaEmpresa()
                Funcoes.VerificaEmpresa(ddlEmpresa)

                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub Limpar()
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        lnkAtualizar.Parent.Visible = False
        txtChaveNFe.Text = String.Empty
        rdNFe.Checked = False
        rdCTe.Checked = False
        txtjustificativa.Text = String.Empty
        txtjustificativa.Enabled = False
        lnkVerificarChaveNFE.Visible = False
    End Sub

    Protected Sub lnkVerificarChaveNFE_Click(sender As Object, e As EventArgs)
        If String.IsNullOrWhiteSpace(txtChaveNFe.Text.Replace(".", "")) Then
            MsgBox(Me.Page, "Chave da Nota Fiscal não foi informada.")
        ElseIf txtChaveNFe.Text.Replace(".", "").Length <> 44 Then
            txtChaveNFe.Text = ""
            MsgBox(Me.Page, "Chave da nota eletrônica deve ter 44 dígitos.")
        ElseIf Not rdNFe.Checked AndAlso Not rdCTe.Checked Then
            MsgBox(Me.Page, "Selecione NFe ou CTe para continuar.")
        Else
            Dim Sqls As New ArrayList
            Dim Sql As String = String.Empty
            Dim chanfe As String = txtChaveNFe.Text.Replace(".", "")
            Dim numnfe As String = Mid(txtChaveNFe.Text.Replace(".", ""), 26, 9)

            Dim ModeloNFe As String = Mid(chanfe, 21, 2)

            If rdNFe.Checked AndAlso Not ModeloNFe.Equals("55") Then
                MsgBox(Me.Page, "Modelo " & ModeloNFe & " da chave eletrônica não é de Nota Fiscal, e não é válido para esse tipo de Operação")
                Exit Sub
            End If

            If rdCTe.Checked AndAlso Not ModeloNFe.Equals("57") Then
                MsgBox(Me.Page, "Modelo " & ModeloNFe & " da chave eletrônica não é de Conhecimento de Transporte, e não é válido para esse tipo de Operação")
                Exit Sub
            End If

            Sql = "SELECT ISNULL(ChaveNFE,'') AS ChaveNFE, ISNULL(Protocolo, '') AS Protocolo, DadosAdicionais " & vbCrLf &
                    "FROM NFERealizadas " & vbCrLf &
                    "WHERE ChaveNFE = '" & chanfe & "'" & vbCrLf &
                    "  AND Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'"

            Sqls.Add(Sql)

            Dim dsRealizada As New DataSet
            dsRealizada = Banco.ConsultaDataSet(Sql, "Consulta")

            If Not dsRealizada Is Nothing AndAlso dsRealizada.Tables(0).Rows.Count > 0 Then
                If rdNFe.Checked Then
                    MsgBox(Me.Page, "Nota Fiscal não pode ser recusada pois foi encontrada uma NFe com a mesma chave lançada no Sistema, verifique.")
                ElseIf rdCTe.Checked Then
                    MsgBox(Me.Page, "Conhecimento de Transporte não pode ser recusado pois foi encontrado um CTe com a mesma chave lançado no Sistema, verifique.")
                End If

                Exit Sub
            End If

            Sqls.Clear()

            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"

            If rdNFe.Checked Then
                obj.NomeArquivo = String.Format("consulta{0:000000000}#{1}.txt", numnfe, ddlEmpresa.SelectedValue.Split("-")(0))
                obj.Texto = "CHAVENFE=" & chanfe & ControlChars.CrLf
            ElseIf rdCTe.Checked Then
                obj.NomeArquivo = String.Format("consultacte{0:000000000}#{1}.txt", numnfe, ddlEmpresa.SelectedValue.Split("-")(0))
                obj.Texto = "CHAVECTE=" & chanfe & ControlChars.CrLf
            Else
                MsgBox(Me.Page, "Erro com seleção de tipo de recusa, verifique.")
                Exit Sub
            End If

            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Empty

            If rdNFe.Checked Then
                fileName = String.Format("resp-consulta{0:000000000}#{1}.txt", numnfe, ddlEmpresa.SelectedValue.Split("-")(0))
            Else
                fileName = String.Format("resp-consultacte{0:000000000}#{1}.txt", numnfe, ddlEmpresa.SelectedValue.Split("-")(0))
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
                Dim chave As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CHAVE")).FirstOrDefault()
                Dim recibo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RECIBO")).FirstOrDefault()
                Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()
                Dim lote As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NUMEROLOTE")).FirstOrDefault()

                Dim strCodigo As String = String.Empty
                If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                    strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                End If

                Dim strMsg As String = String.Empty
                'If rdNFe.Checked Then
                '    If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                '        strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                '    End If
                'End If

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

                If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "100" Or strCodigo = "150") Then
                    lnkAtualizar.Parent.Visible = True

                    If rdNFe.Checked Then txtjustificativa.Text = "OPERAÇÃO NÃO REALIZADA"

                    If rdCTe.Checked Then txtjustificativa.Text = "PRESTAÇÃO DE SERVIÇO EM DESACORDO"

                    txtjustificativa.Enabled = True
                Else
                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                End If
            Else
                MsgBox(Me.Page, "Sefaz não retornou nenhuma resposta, consulte novamente.")
            End If
        End If
    End Sub

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("RecusaNotaFiscal", "GRAVAR") Then
                If ddlEmpresa.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Empresa não foi informada")
                ElseIf txtChaveNFe.Text.Replace(".", "").Length <> 44 Then
                    MsgBox(Me.Page, "A chave da nota eletrônica deve ter 44 dígitos")
                ElseIf txtjustificativa.Text.Length < 10 Then
                    MsgBox(Me.Page, "Justificativa deve ter no mínimo 10 caracteres")
                Else
                    objNotaFiscal = New [Lib].Negocio.NotaFiscal
                    objNotaFiscal.IUD = "I"
                    objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")
                    objNotaFiscal.UsuarioInclusao = Session("ssNomeUsuario")
                    objNotaFiscal.DataInclusao = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
                    objNotaFiscal.Observacoes = Funcoes.EliminarCaracteresEspeciais(txtjustificativa.Text.ToUpper())
                    objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")
                    objNotaFiscal.Codigo = CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9))

                    SessaoSalvaNotaFiscal()

                    Dim TipoDeOperacao As String = String.Empty

                    Dim mensagem As String = String.Empty

                    If rdNFe.Checked Then

                        TipoDeOperacao = eTipoManifesto.OperacaoNaoRealizada

                        'TipoDeOperacao = eTipoManifesto.CienciaDaOperacao

                        If DocumentoEletronico.ManifestoNFe(objNotaFiscal, TipoDeOperacao, mensagem) Then
                            MsgBox(Me.Page, mensagem, eTitulo.Info, False)
                            Limpar()
                        Else
                            MsgBox(Me.Page, mensagem, eTitulo.Erro, False)
                        End If

                    ElseIf rdCTe.Checked Then

                        TipoDeOperacao = eTipoManifesto.OperacaoCTeNaoRealizada

                        If DocumentoEletronico.ManifestoCTe(objNotaFiscal, TipoDeOperacao, mensagem) Then
                            MsgBox(Me.Page, mensagem, eTitulo.Info, False)
                            Limpar()
                        Else
                            MsgBox(Me.Page, mensagem, eTitulo.Erro, False)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
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

    Protected Sub rdNFe_CheckedChanged(sender As Object, e As EventArgs) Handles rdNFe.CheckedChanged
        txtjustificativa.Text = String.Empty
        txtjustificativa.Enabled = False
        lnkVerificarChaveNFE.Visible = True
    End Sub

    Protected Sub rdCTe_CheckedChanged(sender As Object, e As EventArgs) Handles rdCTe.CheckedChanged
        txtjustificativa.Text = String.Empty
        txtjustificativa.Enabled = True
        lnkVerificarChaveNFE.Visible = True
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        If ddlEmpresa.SelectedIndex > 0 Then
            Dim mensagem As String = String.Empty
            Dim retorno As String = String.Empty

            DocumentoEletronico.VerificarModoOperacaoNFe(1, ddlEmpresa.SelectedValue.Split("-")(0), mensagem, retorno)

            If retorno = "107" Then
                lnkVerificarChaveNFE.Visible = True
            Else
                lnkVerificarChaveNFE.Visible = False
                MsgBox(Me.Page, retorno & " - " & mensagem)
            End If
        Else
            Limpar()
        End If
    End Sub
End Class