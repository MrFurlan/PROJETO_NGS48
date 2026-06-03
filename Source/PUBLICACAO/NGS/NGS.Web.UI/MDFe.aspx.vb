Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class MDFe
    Inherits BasePage

#Region "Variáveis"

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal
    Private objMDFe As [Lib].Negocio.NotaFiscal

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
            If Funcoes.VerificaPermissao("MDFe", "ACESSAR") Then
                CargaEmpresas()
                LimparCampos()
                TabContainer1.ActiveTabIndex = 0

                With DdlEmpresa
                    .SelectedIndex = .Items.IndexOf(.Items.FindByValue(Session("ssEmpresa") & "-" & Session("ssEndEmpresa")))
                End With

                If DdlEmpresa.SelectedIndex > 0 Then ConsultarNotas()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar a página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCTRC" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnEstados_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            RecuperarSessaoMDFe()
            Session.Remove("Estados" & HID.Value)
            If (objMDFe IsNot Nothing AndAlso objMDFe.NotaFiscalXPercursos IsNot Nothing) Then
                Session("Estados" & HID.Value) = objMDFe.NotaFiscalXPercursos
                Dim btnSelecionar As Button = CType(ucMDFeXEstado.FindControlRecursive("btnSelecionar"), Button)
                Popup.ConsultaMDFeXEstado(Me.Page, "objMDFeXEstado" & HID.Value, btnSelecionar.ClientID)
                ucMDFeXEstado.BindGridView()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmpresa.Click
        Try
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaCTRC" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClienteMDFe_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "LivreClasse"
            Dim lst As New List(Of [Lib].Negocio.Cliente)
            For Each nf As [Lib].Negocio.NotaFiscal In lstNotaFiscal
                If nf IsNot Nothing AndAlso nf.Cliente IsNot Nothing Then
                    lst.Add(nf.Cliente)
                End If
            Next
            ucConsultaClientesDireto.Limpar()
            ucConsultaClientesDireto.BindGridView(lst)
            Popup.ConsultaDeClientesDireto(Me.Page, "objClienteMDFe" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultarC_Click(sender As Object, e As EventArgs) Handles lnkConsultarC.Click
        Try
            If Funcoes.VerificaPermissao("MDFe", "LEITURA") Then
                If DdlEmpresa.SelectedIndex > 0 Then
                    ConsultarNotas()
                Else
                    MsgBox(Me.Page, "Empresa não foi selecionada.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkSelecionarC_Click(sender As Object, e As EventArgs) Handles lnkSelecionarC.Click
        Try
            If lstNotaFiscal IsNot Nothing Then
                lstNotaFiscal.Clear()
            End If

            Dim sql As String = String.Empty

            sql = "select N.Nota_Id, N.Movimento, N.Situacao " & vbCrLf &
                    "from NFEPendencias NFE " & vbCrLf &
                    "	INNER JOIN NotasFiscais N " & vbCrLf &
                    "		   ON N.Empresa_Id    = NFE.Empresa_Id" & vbCrLf &
                    "		AND N.EndEmpresa_Id   = NFE.EndEmpresa_Id" & vbCrLf &
                    "		AND N.Cliente_Id      = NFE.Cliente_Id" & vbCrLf &
                    "		AND N.EndCliente_Id   = NFE.EndCliente_Id" & vbCrLf &
                    "		AND N.EntradaSaida_Id = NFE.EntradaSaida_Id" & vbCrLf &
                    "		AND N.Serie_Id        = NFE.Serie_Id" & vbCrLf &
                    "		AND N.Nota_Id         = NFE.Nota_Id" & vbCrLf &
                    "WHERE N.Movimento < CONVERT(DATE, GETDATE())" & vbCrLf &
                    "AND N.TipoDeDocumento = 12"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ConsultaPendencia")

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                MsgBox(Me.Page, "Existe MDFE com data inferior pendente de Homologação, verifique antes de proceder com essa emissão.")
            Else
                Selecionar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparC_Click(sender As Object, e As EventArgs) Handles lnkLimparC.Click
        Try
            txtCodigoCliente.Value = ""
            txtNomeCliente.Text = ""
            txtNumMDFe.Text = ""
            rdoSaida.Checked = True
            rdoEntrada.Checked = False
            txtPeriodoInicial.Text = String.Format("01/01/{0}", DateTime.Now.Year)
            txtPeriodoFinal.Text = DateTime.Now.ToShortDateString()
            grdNFe.DataSource = New List(Of Object)
            grdNFe.DataBind()
            LiberaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaC_Click(sender As Object, e As EventArgs) Handles lnkAjudaC.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnTransportador_Click(sender As Object, e As EventArgs) Handles btnTransportador.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objTransportadorCTRC" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Dim fm As New FilesManager()
        Try
            If fm.IsConnect() Then
                Salvar()
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("MDFe", "EXCLUIR") Then
                Dim fm As New FilesManager()
                If fm.IsConnect() Then
                    Excluir("C")
                Else
                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("MDFe", "LEITURA") Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal()
                objNotaFiscal.DataNota = txtDataDeEmissao.Text.Trim()
                objNotaFiscal.Movimento = txtDataDeSaida.Text.Trim()
                objNotaFiscal.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ManifestoEletronico)
                objNotaFiscal.CodigoSituacao = CInt(eSituacao.Normal)
                If Not String.IsNullOrWhiteSpace(txtSerie.Text) Then objNotaFiscal.Serie = txtSerie.Text.Trim
                If Not String.IsNullOrWhiteSpace(txtNumero.Text) Then objNotaFiscal.Codigo = Convert.ToInt32(txtNumero.Text.Trim)
                Session("objNotaFiscal" & HID.Value) = objNotaFiscal
                Session("ssCampo" & HID.Value) = "MDFe"
                Popup.ConsultaDeMDFeXNotas(Me.Page, "NfConsultaMDFe" & HID.Value)
                objNotaFiscal.CarregandoNota = True
                ucConsultaMDFeXNotas.BindGridView()
                objNotaFiscal.CarregandoNota = False
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEspelho_Click(sender As Object, e As EventArgs) Handles lnkEspelho.Click
        Try
            ImprimirEspelho()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEncerrar_Click(sender As Object, e As EventArgs) Handles lnkEncerrar.Click
        Dim fm As New FilesManager()

        Try
            If fm.IsConnect() Then
                If EncerrarMDFe() Then
                    MsgBox(Me.Page, "MDF-e encerrado com Sucesso.", eTitulo.Sucess)
                    LimparCampos()
                End If
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEnviarSEFAZ_Click(sender As Object, e As EventArgs) Handles lnkEnviarSEFAZ.Click
        Dim fm As New FilesManager()

        Try
            If fm.IsConnect() Then
                RecuperarSessaoMDFe()
                EnviarSEFAZ()
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEnviarEmail_Click(sender As Object, e As EventArgs) Handles lnkEnviarEmail.Click
        Dim fm As New FilesManager()

        Try
            If fm.IsConnect() Then
                ucEmailNFe.Limpar()
                Dim txtDestinatario As TextBox = CType(ucEmailNFe.FindControlRecursive("txtDestinatario"), TextBox)
                Popup.ConsultaDeEmailNFe(Me.Page, "objEmailMDFe" & HID.Value, txtDestinatario.ClientID, 100)
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkDAMDFE_Click(sender As Object, e As EventArgs) Handles lnkDAMDFE.Click
        Dim fm As New FilesManager()

        Try
            If fm.IsConnect() Then
                RecuperarSessaoMDFe()
                ImprimirMDFe(objMDFe)
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkVisualizar_Click(sender As Object, e As EventArgs) Handles lnkVisualizar.Click
        Try
            RecuperarSessaoMDFe()
            If objMDFe IsNot Nothing Then
                Dim fileName As String = Server.MapPath(String.Format("~/Files/mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa))
                If System.IO.File.Exists(fileName) Then
                    System.IO.File.Delete(fileName)
                End If

                Dim str As String = getTextoMDFe(objMDFe)
                Using sw As New StreamWriter(fileName)
                    sw.WriteLine(str)
                    sw.Close()
                End Using

                Response.Clear()
                Response.ClearHeaders()
                Response.AddHeader("content-length", str.Length.ToString())
                Response.ContentType = "text/plain"
                Response.AppendHeader("content-disposition", "attachment;filename=" & String.Format("mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa))
                Response.Write(str)
                Response.End()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Métodos"

    Private Sub SalvarSessaoMDFe()
        Session("ssMDFe" & HID.Value) = objMDFe
    End Sub

    Private Sub RecuperarSessaoMDFe()
        objMDFe = CType(Session("ssMDFe" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub Salvar()
        Try
            If Funcoes.VerificaPermissao("MDFe", "GRAVAR") Then
                Dim Sqls As New ArrayList

                RecuperarSessaoMDFe()

                If String.IsNullOrWhiteSpace(txtDataDeEmissao.Text) OrElse String.IsNullOrWhiteSpace(txtDataDeSaida.Text) Then
                    MsgBox(Me.Page, "É necessário informar os campos de data de emissão e data de saída.")
                    Exit Sub
                End If
                If String.IsNullOrWhiteSpace(txtPesoBruto.Text) OrElse Not CDec(txtPesoBruto.Text) > 0 Then
                    MsgBox(Me.Page, "É necessário informar o peso bruto do MDF-e e o valor deve ser maior que zero.")
                    Exit Sub
                End If

                'DESTINO
                If String.IsNullOrWhiteSpace(objMDFe.CodigoDestino) Then
                    If objMDFe.Cliente.CodigoEstado = "EX" Then
                        MsgBox(Me.Page, "Destino Internacional não é permitido para emissão de MDF-e.")
                        Exit Sub
                    End If
                Else
                    If objMDFe.Destino.CodigoEstado = "EX" Then
                        MsgBox(Me.Page, "Destino Internacional não é permitido para emissão de MDF-e.")
                        Exit Sub
                    End If
                End If

                objMDFe.CarregandoNota = True
                objMDFe.NossaEmissao = True
                objMDFe.Eletronica = True
                objMDFe.CodigoSituacao = getSituacao()
                objMDFe.CIFFOB = eTiposFrete.CIF
                objMDFe.DataNota = CDate(txtDataDeEmissao.Text)
                objMDFe.Movimento = CDate(txtDataDeSaida.Text)
                objMDFe.PesoBruto = CDec(txtPesoBruto.Text)
                objMDFe.PesoLiquido = CDec(txtPesoBruto.Text)
                objMDFe.ObservacoesDeEmbarque = txtObservacao.Text.Trim()
                objMDFe.CarregandoNota = False

                'Furlan - 08/07/2021
                If objMDFe.IUD = "I" Then 'Colocado por Cáceres ser uma hora menos e estar no servidor da Matriz
                    If objMDFe.CodigoEmpresa = "03189063000800" Then
                        objMDFe.DataInclusao = Format(Date.Now.AddHours(-1), "yyyy-MM-dd HH:mm:ss")
                    Else
                        objMDFe.DataInclusao = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")
                    End If
                End If

                If objMDFe IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objMDFe.IUD) AndAlso objMDFe.IUD = "U" Then
                    Dim sql = " UPDATE NOTASFISCAIS SET " & vbCrLf & _
                              "       DataDaNota           = '" & objMDFe.DataNota.ToSqlDate() & "'" & vbCrLf & _
                              "     , Movimento            = '" & objMDFe.Movimento.ToSqlDate() & "'" & vbCrLf & _
                              "  WHERE Empresa_Id      = '" & objMDFe.CodigoEmpresa & "'" & vbCrLf & _
                              "    AND EndEmpresa_Id   = " & objMDFe.EnderecoEmpresa & vbCrLf & _
                              "	  AND Cliente_Id      = '" & objMDFe.CodigoCliente & "'" & vbCrLf & _
                              "    AND EndCliente_Id   = " & objMDFe.EnderecoCliente & vbCrLf & _
                              "    AND EntradaSaida_Id = '" & objMDFe.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                              "    AND Nota_Id         = " & objMDFe.Codigo & vbCrLf & _
                              "    AND Serie_Id        = '" & objMDFe.Serie & "';" & vbCrLf
                    Sqls.Add(sql)

                    sql = "UPDATE RAZAO SET DataMoeda = '" & objMDFe.Movimento.ToSqlDate() & "',  " & vbCrLf & _
                          " Movimento_Id = '" & objMDFe.Movimento.ToSqlDate() & "'  " & vbCrLf & _
                          " WHERE 1=1  " & vbCrLf & _
                          " AND Cliente_Nf = '" & objMDFe.CodigoCliente & "'  " & vbCrLf & _
                          " AND EndCliente_Nf = '" & objMDFe.EnderecoCliente & "'  " & vbCrLf & _
                          " AND EntradaSaida_Nf = '" & objMDFe.EntradaSaida.ToString.Substring(0, 1) & "'  " & vbCrLf & _
                          " AND Serie_Nf = '" & objMDFe.Serie & "'  " & vbCrLf & _
                          " AND Numero_Nf = '" & objMDFe.Codigo & "';  " & vbCrLf
                    Sqls.Add(sql)

                    If Not Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If

                    MsgBox(Me.Page, "MDF-e " & objMDFe.Codigo & "-" & objMDFe.Serie & " atualizado com Sucesso.", eTitulo.Sucess)
                    LimparCampos()
                    TabContainer1.ActiveTabIndex = 0
                    Exit Sub
                End If

                If objMDFe.Salvar(Sqls) Then
                    If Not Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        Exit Sub
                    End If

                    lnkNovo.Parent.Visible = False
                    Dim fm As New FilesManager()
                    If fm.IsConnect() AndAlso StatusMDFe() Then
                        EnviarSEFAZ()
                    Else
                        MsgBox(Me.Page, "Serviço de emissão da Alfasig(Guardian) não está em funcionamento, verifique seu servidor.")
                    End If
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function GravarNFeXpress(ByVal mdfe As [Lib].Negocio.NotaFiscal) As Boolean
        Dim aux As Boolean = True

        Try
            Dim Sqls As New ArrayList
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("mdfe{0:000000000}#{1}.txt", mdfe.Codigo, mdfe.CodigoEmpresa)
            obj.Texto = getTextoMDFe(mdfe)
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
        If GravarNFeXpress(objMDFe) Then
            objMDFe = New [Lib].Negocio.NotaFiscal(objMDFe)

            Dim obj As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)

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
                Dim chave As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CHAVEMDFE")).FirstOrDefault()
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

                Dim Sqls As New ArrayList
                Dim sql As String = String.Empty

                If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "100" Then
                    sql = "UPDATE NFEPendencias SET NfExpress = '" & String.Format("mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "', Operacao = 'INCLUIR', Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' "
                    If strRecibo.Length > 0 Then sql &= ", recibo = '" & strRecibo & "' "
                    If strChave.Length > 0 Then sql &= ", ChaveNfe = '" & strChave & "' "
                    If strLote.Length > 0 Then sql &= ", Lote = '" & strLote & "' "
                    If strProtocolo.Length > 0 Then sql &= ", Protocolo = '" & strProtocolo & "' "
                    sql &= "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa
                    sql &= "  AND Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente
                    sql &= "  AND EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'"
                    sql &= "  AND Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo
                    Sqls.Add(sql)

                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-status{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "'"
                    Sqls.Add(sql)
                    sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "'"
                    Sqls.Add(sql)

                    Banco.GravaBanco(Sqls)

                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                    LimparCampos()
                    Exit Sub
                End If

                'ATUALIZAR NFE PENDENCIAS COM OS DADOS DO RETORNO
                objMDFe.ChaveNFE = strChave
                objMDFe.ProtocoloNota = strProtocolo
                objMDFe.ReciboNota = strRecibo

                sql = "INSERT INTO NFERealizadas " & vbCrLf & _
                      "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, " & vbCrLf & _
                      "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) " & vbCrLf & _
                      "VALUES ('" & objMDFe.CodigoEmpresa & "'," & objMDFe.EnderecoEmpresa & ",'" & objMDFe.CodigoCliente & "'," & objMDFe.EnderecoCliente & vbCrLf & _
                      ",'" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "','" & objMDFe.Serie & "'," & objMDFe.Codigo & vbCrLf & _
                      ",'" & objMDFe.DataInclusao.ToSqlDate() & "','" & Format(objMDFe.DataInclusao, "HH:mm:ss") & "','" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                      ",'" & String.Format("mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "','INCLUIR'" & vbCrLf & _
                      ",'" & strCodigo & "','" & strMsg & "','" & strRecibo & "','" & strChave & "'," & strLote & ",1,'" & objMDFe.Observacoes & "', '','" & strProtocolo & "','')"
                Sqls.Add(sql)

                sql = "DELETE FROM NFEPendencias " & vbCrLf & _
                      "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' And EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & vbCrLf & _
                      "  AND Cliente_Id = '" & objMDFe.CodigoCliente & "' And EndCliente_Id = " & objMDFe.EnderecoCliente & vbCrLf & _
                      "  AND EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' And Serie_Id = '" & objMDFe.Serie & "'" & vbCrLf & _
                      "  AND Nota_Id = " & objMDFe.Codigo
                Sqls.Add(sql)

                sql = "UPDATE NotasFiscais Set Situacao = 1 " & vbCrLf & _
                      "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' And EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & vbCrLf & _
                      "  AND Cliente_Id = '" & objMDFe.CodigoCliente & "' And EndCliente_Id = " & objMDFe.EnderecoCliente & vbCrLf & _
                      "  AND EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' And Serie_Id = '" & objMDFe.Serie & "'" & vbCrLf & _
                      "  AND Nota_Id = " & objMDFe.Codigo
                Sqls.Add(sql)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If
                'FIM DA ROTINA

                SalvarSessaoMDFe()

                If SendMailMDFe(objMDFe) Then
                    ImprimirMDFe(objMDFe, True)
                End If

                Sqls.Clear()
                sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-statusmdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "';"
                Sqls.Add(sql)
                sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-mdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "';"
                Sqls.Add(sql)
                sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailmdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "';"
                Sqls.Add(sql)
                sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-damdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "';"
                Sqls.Add(sql)
                sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-eventomdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa) & "';"
                Sqls.Add(sql)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If

                MsgBox(Me.Page, "MDF-e " & objMDFe.Codigo & "-" & objMDFe.Serie & " incluído com Sucesso.", eTitulo.Sucess)
                ConsultarNotas()
                LimparCampos()
                TabContainer1.ActiveTabIndex = 0
            Else
                MsgBox(Me.Page, "Falha no retorno da Sefaz ref. a Homologação da MDFe, é necessário realizar o reenvio à SEFAZ!")
            End If
        End If
    End Sub

    Public Sub ImprimirMDFe(ByVal objMDFe As [Lib].Negocio.NotaFiscal, Optional ByVal IsFirst As Boolean = False)
        Dim Sqls As New ArrayList
        Dim pdf As New [Lib].Negocio.Fil()

        pdf.IUD = "I"
        pdf.NomeArquivo = String.Format("damdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)
        pdf.Texto = getTextoDAMDFe(objMDFe)
        pdf.SalvarSql(Sqls)

        If Not Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Exit Sub
        End If

        Dim obj As [Lib].Negocio.Resp = Nothing
        While obj Is Nothing
            obj = GetResp(String.Format("resp-{0}", pdf.NomeArquivo))
        End While

        If obj IsNot Nothing Then
            Dim lstMsg As String() = obj.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
            Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "7012" Then
                MsgBox(Me.Page, strMsg)
                Exit Sub
            End If

            Try
                Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objMDFe.ChaveNFE), eTipoDeDocumento.ManifestoEletronico)
                If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) Then
                    Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objMDFe.ChaveNFE))
                    System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                    Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objMDFe.ChaveNFE))
                End If
            Catch ex As Exception
                If IsFirst Then
                    MsgBox(Me.Page, "MDF-e " & objMDFe.Codigo & "-" & objMDFe.Serie & " incluído com Sucesso.", eTitulo.Sucess)
                    ConsultarNotas()
                    LimparCampos()
                    TabContainer1.ActiveTabIndex = 0
                Else
                    Throw New Exception(ex.Message)
                End If
            End Try
        End If
    End Sub

    Private Function StatusMDFe() As Boolean
        RecuperarSessaoMDFe()
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True

        Try
            Dim obj As New [Lib].Negocio.Fil()
            obj.IUD = "I"
            obj.NomeArquivo = String.Format("statusmdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)
            obj.Texto = String.Empty
            obj.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Return False
            End If

            'AGUARDAR RESPOSTA SEFAZ
            Dim resp As [Lib].Negocio.Resp = Nothing
            Dim fileName As String = String.Format("resp-statusmdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)

            While resp Is Nothing
                resp = GetResp(fileName)
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

    Private Function EncerrarMDFe() As Boolean
        RecuperarSessaoMDFe()
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True

        Try
            objMDFe = New [Lib].Negocio.NotaFiscal(objMDFe)
            If Not String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.ProtocoloNota) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("eventomdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)
                obj.Texto = getTextoEncerrar(objMDFe)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Return False
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Format("resp-eventomdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)

                While resp Is Nothing
                    resp = GetResp(fileName)
                End While

                If resp IsNot Nothing Then
                    Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                    Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                    Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                    Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "135" Then
                        MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                        Return False
                    End If

                    'ATUALIZAR NFE REALIZADAS COM O RETORNO
                    Sqls.Clear()
                    Sqls.Add(String.Format("UPDATE NFERealizadas SET Retorno = '{0}', MsgRetorno = '{1}' WHERE Empresa_Id = '{2}' AND EndEmpresa_Id = '{3}' AND Cliente_Id = '{4}' " & vbCrLf & _
                             "AND EndCliente_Id = '{5}' AND EntradaSaida_Id = '{6}' AND Serie_Id = '{7}' AND Nota_Id = '{8}'", strCodigo, strMsg, objMDFe.CodigoEmpresa, objMDFe.EnderecoEmpresa,
                             objMDFe.CodigoCliente, objMDFe.EnderecoCliente, IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S"), objMDFe.Serie, objMDFe.Codigo))

                    If Not Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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

    Private Function CancelarMDFe() As Boolean
        RecuperarSessaoMDFe()
        Dim Sqls As New ArrayList
        Dim aux As Boolean = True

        Try
            If Not String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.ProtocoloNota) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("eventomdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)
                obj.Texto = getTextoCancelar(objMDFe)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Return False
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Format("resp-eventomdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)

                While resp Is Nothing
                    resp = GetResp(fileName)
                End While

                If resp IsNot Nothing Then
                    Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                    Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                    Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                    Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    Dim strMsg As String = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "135" Then
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

    Public Function SendMailMDFe(ByVal objMDFe As [Lib].Negocio.NotaFiscal) As Boolean
        Dim aux As Boolean = True
        Dim Sqls As New ArrayList

        Try
            If Not String.IsNullOrWhiteSpace(objMDFe.Empresa.EmailNFE) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.Cliente.EmailNFE) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("emailmdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)
                obj.Texto = getTextoEmail(objMDFe)
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

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)

        If Not obj.Codigo > 0 Then
            Return Nothing
        End If

        Return obj
    End Function

    Private Function getDataSetEmpresa(ByVal mdfe As [Lib].Negocio.NotaFiscal)
        Dim sql As String = String.Empty

        sql &= "SELECT     Clientes.Nome, Clientes.Inscricao, Clientes.Endereco, ISNULL(Clientes.Numero, N'') AS Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, " & vbCrLf & _
               "                      ISNULL(Clientes.Bairro, '') AS Bairro, Municipios.Codigo_id AS CodMunicipio, Clientes.Cidade, Clientes.Estado, Clientes.Cep, Clientes.Pais, " & vbCrLf & _
               "                      Pais.Descricao AS NomePais, Clientes.Telefone, Municipios.Estadoibge, Clientes.EmailNFE, ClientesXEmpresas.Crt, Clientes.Email " & vbCrLf & _
               "FROM         Clientes INNER JOIN " & vbCrLf & _
               "                      ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id INNER JOIN " & vbCrLf & _
               "                      Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.Cidade = Municipios.Municipio_id INNER JOIN " & vbCrLf & _
               "                      Pais ON Clientes.Pais = Pais.Pais_Id INNER JOIN " & vbCrLf & _
               "                      ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id " & vbCrLf & _
               "WHERE     (Clientes.Cliente_Id = '" & mdfe.CodigoEmpresa & "') AND (Clientes.Endereco_Id = " & mdfe.EnderecoEmpresa & ") AND (ClientesXTipos.Tipo_Id = 1)" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "ConsultaEmpresa")
    End Function

    Private Function getDataSetNotaFiscal(ByVal mdfe As [Lib].Negocio.NotaFiscal)
        Dim Sql As String = String.Empty

        Sql &= "SELECT    NotasFiscais.Observacoes, NotasFiscais.ObservacoesDeEmbarque, NotasFiscais.Movimento, NotasFiscais.DataDaNota, isnull(NotasFiscais.LocalEmbarque,'') as LocalEmbarque, " & vbCrLf & _
               "          isnull(NotasFiscais.EndLocalEmbarque,0) as EndLocalEmbarque, isnull(Embarque.Estado,'') as EstadoEmbarque, isnull(Embarque.Cidade,'') AS CidadeEmbarque, " & _
               "          isnull(MunicipioEmbarque.Codigo_id,0) AS CodMunicipioEmbarque,  isnull(MunicipioEmbarque.Estadoibge,'') AS EstadoibgeEmbarque, " & vbCrLf & _
               "          isnull(NotasFiscais.Destino,'') as Destino, isnull(NotasFiscais.EndDestino,0) as EndDestino, isnull(DestinoMercadoria.Cidade,'') AS CidadeDestino, isnull(DestinoMercadoria.Estado,'') as EstadoDestino, " & vbCrLf & _
               "          isnull(MunicipioDestino.Codigo_id,0) AS CodMunicipioDestino,  isnull(MunicipioDestino.Estadoibge,'') AS EstadoibgeDestino, convert(decimal(18, 4), round(NotasXEmbalagens.PesoBruto, 4)) AS PesoBruto, isnull(DestinoMercadoria.CEP,'') As CepDestino, " & vbCrLf & _
               "          NotasFiscais.Movimento, NotasFiscais.UsuarioInclusaoData" & vbCrLf & _
               "FROM      NotasFiscais LEFT JOIN Clientes AS Embarque " & vbCrLf & _
               "                                 ON  NotasFiscais.LocalEmbarque = Embarque.Cliente_Id " & vbCrLf & _
               "                                 AND NotasFiscais.EndLocalEmbarque = Embarque.Endereco_Id " & vbCrLf & _
               "                          LEFT JOIN Municipios AS MunicipioEmbarque " & vbCrLf & _
               "                                 ON Embarque.Estado = MunicipioEmbarque.Estado_id " & _
               "                                 AND Embarque.Cidade = MunicipioEmbarque.Municipio_id " & vbCrLf & _
               "                          LEFT JOIN Clientes AS DestinoMercadoria " & vbCrLf & _
               "                                 ON  NotasFiscais.Destino = DestinoMercadoria.Cliente_Id " & vbCrLf & _
               "                                 AND NotasFiscais.EndDestino = DestinoMercadoria.Endereco_Id " & vbCrLf & _
               "                          LEFT JOIN Municipios AS MunicipioDestino " & vbCrLf & _
               "                                 ON DestinoMercadoria.Estado = MunicipioDestino.Estado_id " & vbCrLf & _
               "                                 AND DestinoMercadoria.Cidade = MunicipioDestino.Municipio_id " & vbCrLf & _
               "                        INNER JOIN NotasXEmbalagens " & vbCrLf & _
               "                                ON NotasXEmbalagens.Empresa_Id       = NotasFiscais.Empresa_Id " & vbCrLf & _
               "                                AND NotasXEmbalagens.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id " & vbCrLf & _
               "                                AND NotasXEmbalagens.Cliente_Id      = NotasFiscais.Cliente_Id " & vbCrLf & _
               "                                AND NotasXEmbalagens.EndCliente_Id   = NotasFiscais.EndCliente_Id " & vbCrLf & _
               "                                AND NotasXEmbalagens.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id " & vbCrLf & _
               "                                AND NotasXEmbalagens.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf & _
               "                                AND NotasXEmbalagens.Nota_Id         = NotasFiscais.Nota_Id " & vbCrLf & _
               "WHERE     (NotasFiscais.Empresa_Id = '" & mdfe.CodigoEmpresa & "') AND (NotasFiscais.EndEmpresa_Id = " & mdfe.EnderecoEmpresa & ") AND (NotasFiscais.Cliente_Id = '" & mdfe.CodigoCliente & "') AND " & vbCrLf & _
               "          (NotasFiscais.EndCliente_Id = " & mdfe.EnderecoCliente & ") AND (NotasFiscais.EntradaSaida_Id = '" & IIf(mdfe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') AND " & vbCrLf & _
               "          (NotasFiscais.Serie_Id = '" & mdfe.Serie & "') AND (NotasFiscais.Nota_Id = " & mdfe.Codigo & ")" & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Nota")
    End Function

    Private Function getDataSetTransportador(ByVal mdfe As [Lib].Negocio.NotaFiscal)
        Dim sql As String = String.Empty

        sql &= "SELECT     NotasFiscaisXTransportadores.Proprietario, NotasFiscaisXTransportadores.EndProprietario, DadosDoTransportador.Nome AS NomeTransportador, ISNULL(DadosDoTransportador.Inscricao, '') As IETransportador, DadosDoTransportador.Estado AS UFTransportador, NotasFiscaisXTransportadores.Motorista, " & vbCrLf & _
               "                      NotasFiscaisXTransportadores.EndMotorista, DadosDoMotorista.Nome AS NomeMotorista, Placas.Placa_Id AS Placa, isnull(DadosDoTransportador.RNTRCTransportador,'') AS RNTRCTransportador, Placas.CidadePlaca, Placas.EstadoPlaca, Placas.Placa01, Placas.CidadePlaca01, " & vbCrLf & _
               "                      Placas.EstadoPlaca01, Placas.Placa02, Placas.CidadePlaca02, Placas.EstadoPlaca02, Placas.Placa03, Placas.CidadePlaca03, Placas.EstadoPlaca03, " & vbCrLf & _
               "                      ISNULL(Placas.RNTRCPlaca,'') AS RNTRCPlaca " & vbCrLf & _
               "FROM         NotasFiscaisXTransportadores INNER JOIN " & vbCrLf & _
               "                      Placas ON NotasFiscaisXTransportadores.Placa = Placas.Placa_Id INNER JOIN " & vbCrLf & _
               "                      Clientes AS DadosDoMotorista ON NotasFiscaisXTransportadores.Motorista = DadosDoMotorista.Cliente_Id AND " & vbCrLf & _
               "                      NotasFiscaisXTransportadores.EndMotorista = DadosDoMotorista.Endereco_Id INNER JOIN " & vbCrLf & _
               "                      Clientes AS DadosDoTransportador ON NotasFiscaisXTransportadores.Proprietario = DadosDoTransportador.Cliente_Id AND " & vbCrLf & _
               "                      NotasFiscaisXTransportadores.EndProprietario = DadosDoTransportador.Endereco_Id " & vbCrLf & _
               "WHERE     (NotasFiscaisXTransportadores.Empresa_Id = '" & mdfe.CodigoEmpresa & "') AND (NotasFiscaisXTransportadores.EndEmpresa_Id = " & mdfe.EnderecoEmpresa & ") AND " & vbCrLf & _
               "                      (NotasFiscaisXTransportadores.Cliente_Id = '" & mdfe.CodigoCliente & "') AND (NotasFiscaisXTransportadores.EndCliente_Id = " & mdfe.EnderecoCliente & ") AND " & vbCrLf & _
               "                      (NotasFiscaisXTransportadores.EntradaSaida_Id = '" & IIf(mdfe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') AND (NotasFiscaisXTransportadores.Serie_Id = '" & mdfe.Serie & "') AND (NotasFiscaisXTransportadores.Nota_Id = " & mdfe.Codigo & ")" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "ConsultaTransportador")
    End Function

    Private Function getDataSetCliente(ByVal mdfe As [Lib].Negocio.NotaFiscal)
        Dim sql As String = String.Empty

        sql &= "SELECT     TOP (1) Clientes.Cliente_Id AS Cliente, Clientes.Endereco_Id AS EndCliente, Clientes.Nome, Clientes.Inscricao, Clientes.Endereco, ISNULL(Clientes.Numero, N'') AS Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, " & vbCrLf & _
               "                      ISNULL(Clientes.Bairro, '') AS Bairro, Municipios.Codigo_id AS CodMunicipio, Clientes.Cidade, Clientes.Estado, Clientes.Cep, Clientes.Pais, " & vbCrLf & _
               "                      Pais.Descricao AS NomePais, Clientes.Telefone, Clientes.Email, Municipios.Estadoibge, Clientes.EmailNFE, Clientes.Email " & vbCrLf & _
               "FROM         Clientes INNER JOIN " & vbCrLf & _
               "                      Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.Cidade = Municipios.Municipio_id INNER JOIN " & vbCrLf & _
               "                      Pais ON Clientes.Pais = Pais.Pais_Id " & vbCrLf & _
               "WHERE     (Clientes.Cliente_Id = '" & mdfe.CodigoCliente & "') AND (Clientes.Endereco_Id = " & mdfe.EnderecoCliente & ")" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "ConsultaCliente")
    End Function

    Private Function getDataSetMDFeXNota(ByVal mdfe As [Lib].Negocio.NotaFiscal) As DataSet
        Dim sql As String = String.Empty

        sql &= "SELECT Cli.Cidade, CliMun.Codigo_id AS CidadeIbge, CliMun.EstadoIbge " & vbCrLf &
               "FROM NotasXNotas  " & vbCrLf &
               "    INNER JOIN NotasFiscais NF " & vbCrLf &
               "            ON NF.Empresa_Id      = NotasXNotas.OrigemEmpresa_Id " & vbCrLf &
               "           AND NF.EndEmpresa_Id   = NotasXNotas.OrigemEndEmpresa_Id " & vbCrLf &
               "           AND NF.Cliente_Id      = NotasXNotas.OrigemCliente_Id " & vbCrLf &
               "           AND NF.EndCliente_Id   = NotasXNotas.OrigemEndCliente_Id " & vbCrLf &
               "           AND NF.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id " & vbCrLf &
               "           AND NF.Serie_Id        = NotasXNotas.OrigemSerie_Id " & vbCrLf &
               "           AND NF.Nota_Id         = NotasXNotas.OrigemNota_Id " & vbCrLf &
               "	INNER JOIN Clientes AS Cli " & vbCrLf &
               "		ON Cli.Cliente_Id   = NF.Destino " & vbCrLf &
               "		AND CLI.Endereco_Id = NF.EndDestino " & vbCrLf &
               "	INNER JOIN Municipios AS CliMun " & vbCrLf &
               "		ON CliMun.Estado_id     = Cli.Estado " & vbCrLf &
               "		AND CliMun.Municipio_id = Cli.Cidade " & vbCrLf &
               "WHERE     (NotasXNotas.Empresa_Id = '" & mdfe.CodigoEmpresa & "') AND (NotasXNotas.EndEmpresa_Id = " & mdfe.EnderecoEmpresa & ") AND (NotasXNotas.Cliente_Id = '" & mdfe.CodigoCliente & "') AND " & vbCrLf &
               "          (NotasXNotas.EndCliente_Id = " & mdfe.EnderecoCliente & ") AND (NotasXNotas.EntradaSaida_Id = '" & IIf(mdfe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') AND (NotasXNotas.Serie_Id = '" & mdfe.Serie & "') AND (NotasXNotas.Nota_Id = " & mdfe.Codigo & ")" & vbCrLf &
               "group by Cli.Cidade, CliMun.Codigo_id, Estadoibge " & vbCrLf &
               "order by CliMun.Codigo_id" & vbCrLf &
               "SELECT Cli.Cidade, CliMun.Codigo_id AS CidadeIbge, CliMun.EstadoIbge, NFERealizadas.ChaveNfe, ISNULL(NFERealizadas.SegCodBarra,'') AS SegCodBarra, sum(NotasFiscaisXItens.Valor) AS ValorMercadoria " & vbCrLf &
               "FROM  NotasXNotas " & vbCrLf &
               "    INNER JOIN NFERealizadas " & vbCrLf &
               "            ON NFERealizadas.Empresa_Id      = NotasXNotas.OrigemEmpresa_Id " & vbCrLf &
               "           AND NFERealizadas.EndEmpresa_Id   = NotasXNotas.OrigemEndEmpresa_Id " & vbCrLf &
               "           AND NFERealizadas.Cliente_Id      = NotasXNotas.OrigemCliente_Id " & vbCrLf &
               "           AND NFERealizadas.EndCliente_Id   = NotasXNotas.OrigemEndCliente_Id " & vbCrLf &
               "           AND NFERealizadas.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id " & vbCrLf &
               "           AND NFERealizadas.Serie_Id        = NotasXNotas.OrigemSerie_Id " & vbCrLf &
               "           AND NFERealizadas.Nota_Id         = NotasXNotas.OrigemNota_Id " & vbCrLf &
               "    INNER JOIN NotasFiscais NF " & vbCrLf &
               "            ON NF.Empresa_Id      = NotasXNotas.OrigemEmpresa_Id " & vbCrLf &
               "           AND NF.EndEmpresa_Id   = NotasXNotas.OrigemEndEmpresa_Id " & vbCrLf &
               "           AND NF.Cliente_Id      = NotasXNotas.OrigemCliente_Id " & vbCrLf &
               "           AND NF.EndCliente_Id   = NotasXNotas.OrigemEndCliente_Id " & vbCrLf &
               "           AND NF.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id " & vbCrLf &
               "           AND NF.Serie_Id        = NotasXNotas.OrigemSerie_Id " & vbCrLf &
               "           AND NF.Nota_Id         = NotasXNotas.OrigemNota_Id " & vbCrLf &
               "    INNER JOIN NotasFiscaisXItens " & vbCrLf &
               "            ON NotasFiscaisXItens.Empresa_Id       = NotasXNotas.OrigemEmpresa_Id " & vbCrLf &
               "           AND NotasFiscaisXItens.EndEmpresa_Id   = NotasXNotas.OrigemEndEmpresa_Id " & vbCrLf &
               "           AND NotasFiscaisXItens.Cliente_Id      = NotasXNotas.OrigemCliente_Id " & vbCrLf &
               "           AND NotasFiscaisXItens.EndCliente_Id   = NotasXNotas.OrigemEndCliente_Id " & vbCrLf &
               "           AND NotasFiscaisXItens.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id " & vbCrLf &
               "           AND NotasFiscaisXItens.Serie_Id        = NotasXNotas.OrigemSerie_Id  " & vbCrLf &
               "           AND NotasFiscaisXItens.Nota_Id         = NotasXNotas.OrigemNota_Id " & vbCrLf &
               "	INNER JOIN Clientes AS Cli " & vbCrLf &
               "	 	    ON Cli.Cliente_Id  = NF.Destino " & vbCrLf &
               "		   AND CLI.Endereco_Id = NF.EndDestino " & vbCrLf &
               "	INNER JOIN Municipios AS CliMun " & vbCrLf &
               "		    ON CliMun.Estado_id     = Cli.Estado " & vbCrLf &
               "		   AND CliMun.Municipio_id = Cli.Cidade " & vbCrLf &
               "WHERE     (NotasXNotas.Empresa_Id = '" & mdfe.CodigoEmpresa & "') AND (NotasXNotas.EndEmpresa_Id = " & mdfe.EnderecoEmpresa & ") AND (NotasXNotas.Cliente_Id = '" & mdfe.CodigoCliente & "') AND " & vbCrLf &
               "                      (NotasXNotas.EndCliente_Id = " & mdfe.EnderecoCliente & ") AND (NotasXNotas.EntradaSaida_Id = '" & IIf(mdfe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') AND (NotasXNotas.Serie_Id = '" & mdfe.Serie & "') AND (NotasXNotas.Nota_Id = " & mdfe.Codigo & ")" & vbCrLf &
               "group by Cli.Cidade, CliMun.Codigo_id, Estadoibge, NFERealizadas.ChaveNfe, ISNULL(NFERealizadas.SegCodBarra,'') " & vbCrLf &
               "order by CliMun.Codigo_id" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "mdfEXNota")
    End Function

    Private Function getDataSetMDFePendencias(ByVal mdfe As [Lib].Negocio.NotaFiscal) As DataSet
        Dim sql As String = String.Empty

        sql = "SELECT     Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, Serie_Id, Nota_Id, EntradaSaida_Id, " & vbCrLf & _
              " CONVERT(varchar, Data, 103) As Data, CONVERT(varchar(08),Hora,14) As Hora, Usuario, NfExpress, Operacao, ISNULL(Retorno, '') AS Retorno, MsgRetorno, Recibo, " & vbCrLf & _
              " isnull(ChaveNfe,'') AS ChaveNfe, Lote, TpEmis, ISNULL(ObservacoesFiscais, '') AS ObservacoesFiscais, ISNULL(DadosAdicionais, '') AS DadosAdicionais, ISNULL(Protocolo, '') AS Protocolo, RegDPEC " & vbCrLf & _
              " FROM NFEPendencias " & vbCrLf & _
              " WHERE 1=1 " & vbCrLf & _
              " AND (NFEPendencias.Empresa_Id = '" & mdfe.CodigoEmpresa & "') " & vbCrLf & vbCrLf & _
              " AND (NFEPendencias.EndEmpresa_Id = " & mdfe.EnderecoEmpresa & ") " & vbCrLf & vbCrLf & _
              " AND (NFEPendencias.Cliente_Id = '" & mdfe.CodigoCliente & "') " & vbCrLf & vbCrLf & _
              " AND (NFEPendencias.EndCliente_Id = " & mdfe.EnderecoCliente & ") " & vbCrLf & vbCrLf & _
              " AND (NFEPendencias.EntradaSaida_Id = '" & IIf(mdfe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "') " & vbCrLf & vbCrLf & _
              " AND (NFEPendencias.Serie_Id = '" & mdfe.Serie & "') " & vbCrLf & vbCrLf & _
              " AND (NFEPendencias.Nota_Id = " & mdfe.Codigo & ")" & vbCrLf & vbCrLf & _
              " ORDER BY Data DESC, Hora DESC" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "NFEPendencias")
    End Function

    Private Function getTextoMDFe(ByVal mdfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        Dim dsEmpresa As DataSet = getDataSetEmpresa(mdfe)
        Dim dsCliente As DataSet = getDataSetCliente(mdfe)
        Dim dsNotaFiscal As DataSet = getDataSetNotaFiscal(mdfe)
        Dim dsTransportador As DataSet = getDataSetTransportador(mdfe)
        Dim dsMDFeXNota As DataSet = getDataSetMDFeXNota(mdfe)
        Dim dsPendencias As DataSet = getDataSetMDFePendencias(mdfe)
        Dim valorMercadoria As Decimal = 0
        Dim observacoes As String = ""
        Dim valor As String = ""

        sb.Append("[IDE]" & ControlChars.CrLf)
        sb.Append("   CUF	 = " & dsEmpresa.Tables(0).Rows(0).Item("Estadoibge") & ControlChars.CrLf)
        sb.Append("   TPEMIT = 2" & ControlChars.CrLf)

        Dim tipoTac As Boolean = False

        'TPTRANSP - Tipo do Transportador - 1 - ETC   2 - TAC   3 - CTC
        If Left(mdfe.CodigoEmpresa, 8) = Left(dsTransportador.Tables(0).Rows(0).Item("Proprietario"), 8) Then
            'NÃO FAZ NADA
        Else
            If Not mdfe.CodigoEmpresa = dsTransportador.Tables(0).Rows(0).Item("Proprietario") And _
                       Not dsTransportador.Tables(0).Rows(0).Item("Proprietario") = "73901679987" Then
                'If Not dsEmpresa.Tables(0).Rows(0).Item("Estadoibge") = "41" Then
                sb.Append(" TPTRANSP = 2" & ControlChars.CrLf)
                tipoTac = True
                'End If
            End If
        End If

        sb.Append("   SERIE  = 0" & ControlChars.CrLf)
        sb.Append("   NMDF	 = " & mdfe.Codigo & ControlChars.CrLf)
        sb.Append("   MODAL  = 1" & ControlChars.CrLf)
        'sb.Append("   DHEMI  = " & String.Format("{0:yyyy-MM-dd}", CDate(dsPendencias.Tables(0).Rows(0).Item("Data"))) & "T" & String.Format("{0:HH:mm:ss}", dsPendencias.Tables(0).Rows(0).Item("Hora")) & ControlChars.CrLf)
        sb.Append("   DHEMI  = " & String.Format("{0:yyyy-MM-dd}", CDate(dsNotaFiscal.Tables(0).Rows(0).Item("Movimento"))) & "T" & String.Format("{0:HH:mm:ss}", dsNotaFiscal.Tables(0).Rows(0).Item("UsuarioInclusaoData")) & ControlChars.CrLf)
        'sb.Append("   DHEMI  = " & dsNotaFiscal.Tables(0).Rows(0).Item("Movimento").ToString("yyyy-MM-dd") & "T" & dsNotaFiscal.Tables(0).Rows(0).Item("UsuarioInclusaoData").ToString("HH:mm:ss") & ControlChars.CrLf)

        'ORIGEM
        If dsNotaFiscal.Tables(0).Rows(0)("LocalEmbarque") = "" Then
            sb.Append("   UFINI  = " & dsEmpresa.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        Else
            sb.Append("   UFINI  = " & dsNotaFiscal.Tables(0).Rows(0).Item("EstadoEmbarque") & ControlChars.CrLf)
        End If

        'DESTINO
        If dsNotaFiscal.Tables(0).Rows(0)("Destino") = "" Then
            sb.Append("   UFFIM  = " & dsCliente.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        Else
            sb.Append("   UFFIM  = " & dsNotaFiscal.Tables(0).Rows(0).Item("EstadoDestino") & ControlChars.CrLf)
        End If
        'DHINIVIAGEM - Data e hora previstas de início da Viagem - Formato AAAA-MM-DDTHH:MM:DD
        'sb.Append("DHINIVIAGEM= " & String.Format("{0:yyyy-MM-dd}", CDate(dsPendencias.Tables(0).Rows(0).Item("Data"))) & "T" & String.Format("{0:HH:mm:ss}", dsPendencias.Tables(0).Rows(0).Item("Hora")) & ControlChars.CrLf)
        sb.Append("DHINIVIAGEM= " & String.Format("{0:yyyy-MM-dd}", CDate(dsNotaFiscal.Tables(0).Rows(0).Item("Movimento"))) & "T" & String.Format("{0:HH:mm:ss}", dsNotaFiscal.Tables(0).Rows(0).Item("UsuarioInclusaoData")) & ControlChars.CrLf)
        'VERPROC - Versão deste MDF-e - Informar “ mdfe3g ”
        sb.Append("   VERPROC=mdfe3g" & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[INFMUNCARREGA]" & ControlChars.CrLf)
        If dsNotaFiscal.Tables(0).Rows(0)("LocalEmbarque") = "" Then
            sb.Append("   CMUNCARREGA = " & dsEmpresa.Tables(0).Rows(0).Item("EstadoIbge") & Format(dsEmpresa.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
            sb.Append("   XMUNCARREGA = " & dsEmpresa.Tables(0).Rows(0).Item("Cidade") & ControlChars.CrLf)
        Else
            sb.Append("   CMUNCARREGA = " & dsNotaFiscal.Tables(0).Rows(0).Item("EstadoibgeEmbarque") & Format(dsNotaFiscal.Tables(0).Rows(0).Item("CodMunicipioEmbarque"), "00000") & ControlChars.CrLf)
            sb.Append("   XMUNCARREGA = " & dsNotaFiscal.Tables(0).Rows(0).Item("CidadeEmbarque") & ControlChars.CrLf)
        End If
        sb.Append(ControlChars.CrLf)

        'PERCURSOS
        If mdfe.NotaFiscalXPercursos IsNot Nothing AndAlso mdfe.NotaFiscalXPercursos.Count > 0 Then
            For Each nfp As [Lib].Negocio.NotaFiscalXPercurso In mdfe.NotaFiscalXPercursos.OrderBy(Function(s) s.Ordem).ToList()
                If Not String.IsNullOrWhiteSpace(nfp.Estado_Id) Then
                    sb.Append("[INFPERCURSO]" & ControlChars.CrLf)
                    sb.Append("   UFPER  = " & nfp.Estado_Id & ControlChars.CrLf)
                    sb.Append(ControlChars.CrLf)
                End If
            Next
        End If

        'DADOS DO EMITENTE 
        sb.Append("[EMIT]" & ControlChars.CrLf)
        sb.Append("   CNPJ   = " & mdfe.CodigoEmpresa & ControlChars.CrLf)
        sb.Append("   IE     = " & dsEmpresa.Tables(0).Rows(0).Item("Inscricao").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
        sb.Append("   XNOME  = " & Funcoes.EliminarCaracteresEspeciaisNF(dsEmpresa.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
        sb.Append("   XFANT  = " & Funcoes.EliminarCaracteresEspeciaisNF(dsEmpresa.Tables(0).Rows(0).Item("Nome")) & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[ENDEREMIT]" & ControlChars.CrLf)
        sb.Append("   XLGR   = " & Funcoes.EliminarCaracteresEspeciaisNF(dsEmpresa.Tables(0).Rows(0).Item("Endereco")) & ControlChars.CrLf)
        sb.Append("   NRO	  = " & dsEmpresa.Tables(0).Rows(0).Item("Numero") & ControlChars.CrLf)
        sb.Append("   XCPL   = " & Funcoes.EliminarCaracteresEspeciaisNF(dsEmpresa.Tables(0).Rows(0).Item("Complemento")) & ControlChars.CrLf)
        sb.Append("   XBAIRRO= " & Funcoes.EliminarCaracteresEspeciaisNF(dsEmpresa.Tables(0).Rows(0).Item("Bairro")) & ControlChars.CrLf)
        sb.Append("   CMUN   = " & dsEmpresa.Tables(0).Rows(0).Item("Estadoibge") & Format(dsEmpresa.Tables(0).Rows(0).Item("CodMunicipio"), "00000") & ControlChars.CrLf)
        sb.Append("   XMUN   = " & Funcoes.EliminarCaracteresEspeciaisNF(dsEmpresa.Tables(0).Rows(0).Item("Cidade")) & ControlChars.CrLf)
        sb.Append("   UF	  = " & dsEmpresa.Tables(0).Rows(0).Item("Estado") & ControlChars.CrLf)
        sb.Append("   CEP	  = " & dsEmpresa.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        sb.Append("   FONE   = " & dsEmpresa.Tables(0).Rows(0).Item("Telefone") & ControlChars.CrLf)
        sb.Append("   EMAIL  = " & dsEmpresa.Tables(0).Rows(0).Item("Email") & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        sb.Append("[RODO]" & ControlChars.CrLf)
        'sb.Append("    RNTRC = 11111111" & ControlChars.CrLf)
        'sb.Append("    CIOT  = " & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        If mdfe.CodigoEmpresa = "03189063000800" OrElse tipoTac Then
            sb.Append("[INFANTT]" & ControlChars.CrLf)
            sb.Append("[INFCONTRATANTE]" & ControlChars.CrLf)
            sb.Append("     CNPJ = " & mdfe.CodigoEmpresa & ControlChars.CrLf)
            sb.Append(ControlChars.CrLf)
        End If

        sb.Append("[VEICTRACAO]" & ControlChars.CrLf)
        sb.Append("     CINT = " & ControlChars.CrLf)
        sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa").Replace("-", "") & ControlChars.CrLf)
        sb.Append("     TARA = 0" & ControlChars.CrLf)
        sb.Append("    CAPKG = 0" & ControlChars.CrLf)
        sb.Append("    CAPM3 = 0" & ControlChars.CrLf)

        If dsTransportador.Tables(0).Rows(0).Item("Placa01").ToString.Length > 0 Then
            sb.Append("    TPROD = 03" & ControlChars.CrLf)
        Else
            sb.Append("    TPROD = 01" & ControlChars.CrLf)

        End If

        'sb.Append("    TPCAR = 00" & ControlChars.CrLf)
        If mdfe.CodigoEmpresa = "03189063000800" Then
            sb.Append("    TPCAR = 03" & ControlChars.CrLf)
        Else
            sb.Append("    TPCAR = 01" & ControlChars.CrLf)
        End If

        sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca") & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)


        If dsTransportador.Tables(0).Rows(0).Item("Placa01").ToString.Length > 0 Then
            sb.Append("[VEICREBOQUE]" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa01").Replace("-", "") & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca01") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)

            If mdfe.CodigoEmpresa = "03189063000800" Then
                sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            Else
                sb.Append("    TPCAR = 01" & ControlChars.CrLf)
            End If
            sb.Append(ControlChars.CrLf)
        End If

        If dsTransportador.Tables(0).Rows(0).Item("Placa02").ToString.Length > 0 Then
            sb.Append("[VEICREBOQUE]" & ControlChars.CrLf)
            sb.Append("    PLACA = " & dsTransportador.Tables(0).Rows(0).Item("Placa02").Replace("-", "") & ControlChars.CrLf)
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("EstadoPlaca02") & ControlChars.CrLf)
            sb.Append("     TARA = 0" & ControlChars.CrLf)
            sb.Append("    CAPKG = 0" & ControlChars.CrLf)
            sb.Append("    CAPM3 = 0" & ControlChars.CrLf)

            If mdfe.CodigoEmpresa = "03189063000800" Then
                sb.Append("    TPCAR = 03" & ControlChars.CrLf)
            Else
                sb.Append("    TPCAR = 01" & ControlChars.CrLf)
            End If
            sb.Append(ControlChars.CrLf)
        End If


        'Só preenchido quando o veículo não pertencer à empresa emitente do MDF-e
        'If Not mdfe.CodigoEmpresa = dsTransportador.Tables(0).Rows(0).Item("Proprietario") And _
        '    Not dsTransportador.Tables(0).Rows(0).Item("Proprietario") = "73901679987" Then
        If Not Left(mdfe.CodigoEmpresa, 8) = Left(dsTransportador.Tables(0).Rows(0).Item("Proprietario"), 8) Then
            sb.Append("[PROP]" & ControlChars.CrLf)
            sb.Append("    XNOME = " & Funcoes.EliminarCaracteresEspeciaisNF(dsTransportador.Tables(0).Rows(0).Item("NomeTransportador")) & ControlChars.CrLf)
            If dsTransportador.Tables(0).Rows(0).Item("Proprietario").ToString.Length > 11 Then
                sb.Append("     CNPJ = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
                If dsTransportador.Tables(0).Rows(0).Item("IETransportador").ToString.Length > 0 Then
                    sb.Append("       IE = " & dsTransportador.Tables(0).Rows(0).Item("IETransportador").Replace(".", "").Replace("-", "").Replace("/", "") & ControlChars.CrLf)
                End If
            Else
                sb.Append("      CPF = " & dsTransportador.Tables(0).Rows(0).Item("Proprietario") & ControlChars.CrLf)
                sb.Append("       IE = ISENTO" & ControlChars.CrLf)
            End If
            sb.Append("       UF = " & dsTransportador.Tables(0).Rows(0).Item("UFTransportador") & ControlChars.CrLf)
            sb.Append("   TPPROP = 1" & ControlChars.CrLf)

            If dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCTransportador") & ControlChars.CrLf)
            ElseIf dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length > 0 AndAlso dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca").ToString.Length = 8 Then
                sb.Append("    RNTRC = " & dsTransportador.Tables(0).Rows(0).Item("RNTRCPlaca") & ControlChars.CrLf)
            Else
                sb.Append("    RNTRC = 11111111" & ControlChars.CrLf)
            End If
        End If

        sb.Append(ControlChars.CrLf)
        sb.Append("[CONDUTOR]" & ControlChars.CrLf)
        sb.Append("    XNOME = " & Funcoes.EliminarCaracteresEspeciaisNF(dsTransportador.Tables(0).Rows(0).Item("NomeMotorista")) & ControlChars.CrLf)
        sb.Append("      CPF = " & dsTransportador.Tables(0).Rows(0).Item("Motorista") & ControlChars.CrLf)

        Dim rowIndex As Integer = 1
        sb.Append(ControlChars.CrLf)
        sb.Append("[INFDOC]" & ControlChars.CrLf)

        Dim CMUNDESCARGA As String = ""
        Dim CIDADEIBGE As String = ""

        For Each drMunDes In dsMDFeXNota.Tables(0).Rows
            If rowIndex <> 1 Then
                sb.Append(ControlChars.CrLf)
            End If


            If mdfe.CodigoEmpresa = mdfe.CodigoDestino Then
                If Not (mdfe.Empresa.Municipio.EstadoIbge & Format(mdfe.Empresa.Municipio.CodigoIbge, "00000")) = CMUNDESCARGA Then
                    CMUNDESCARGA = (mdfe.Empresa.Municipio.EstadoIbge & Format(mdfe.Empresa.Municipio.CodigoIbge, "00000"))
                    sb.Append("     [INFMUNDESCARGA]" & ControlChars.CrLf)
                    sb.Append("         CMUNDESCARGA = " & mdfe.Empresa.Municipio.EstadoIbge & Format(mdfe.Empresa.Municipio.CodigoIbge, "00000") & ControlChars.CrLf)
                    sb.Append("         XMUNDESCARGA = " & mdfe.Empresa.Municipio.CodigoMunicipio & ControlChars.CrLf)
                    sb.Append(ControlChars.CrLf)
                End If
            Else
                If Not (drMunDes("Estadoibge") & Format(drMunDes("CidadeIbge"), "00000")) = CMUNDESCARGA Then
                    'CMUNDESCARGA = (mdfe.Empresa.Municipio.EstadoIbge & Format(mdfe.Empresa.Municipio.CodigoIbge, "00000"))
                    CMUNDESCARGA = (drMunDes("Estadoibge") & Format(drMunDes("CidadeIbge"), "00000"))

                    sb.Append("     [INFMUNDESCARGA]" & ControlChars.CrLf)

                    If drMunDes("Estadoibge") & Format(drMunDes("CidadeIbge"), "00000") = "999999999" Then
                        sb.Append("         CMUNDESCARGA = " & dsNotaFiscal.Tables(0).Rows(0).Item("EstadoibgeDestino") & Format(dsNotaFiscal.Tables(0).Rows(0).Item("CodMunicipioDestino"), "00000") & ControlChars.CrLf)
                        sb.Append("         XMUNDESCARGA = " & drMunDes("Cidade") & ControlChars.CrLf)
                    Else
                        sb.Append("         CMUNDESCARGA = " & drMunDes("Estadoibge") & Format(drMunDes("CidadeIbge"), "00000") & ControlChars.CrLf)
                        sb.Append("         XMUNDESCARGA = " & drMunDes("Cidade") & ControlChars.CrLf)
                        'sb.Append("         CMUNDESCARGA = " & (mdfe.Destino.Municipio.EstadoIbge & Format(mdfe.Destino.Municipio.CodigoIbge, "00000")) & ControlChars.CrLf)
                        'sb.Append("         XMUNDESCARGA = " & mdfe.Destino.Municipio.CodigoMunicipio & ControlChars.CrLf)
                    End If

                    sb.Append(ControlChars.CrLf)
                End If
            End If

            For Each row In dsMDFeXNota.Tables(1).Rows
                If row("CidadeIbge") = drMunDes("CidadeIbge") AndAlso Not CIDADEIBGE = "9999999" Then
                    sb.Append("         [INFNFE]" & ControlChars.CrLf)
                    sb.Append("             CHNFE  = " & row("ChaveNfe") & ControlChars.CrLf)

                    If row("SegCodBarra").ToString.Length > 0 Then
                        sb.Append("       SEGCODBARRA  = " & row("SegCodBarra") & ControlChars.CrLf)
                    End If

                    valorMercadoria += row("ValorMercadoria")
                End If
            Next

            CIDADEIBGE = drMunDes("CidadeIbge")

            rowIndex += 1
        Next

        sb.Append(ControlChars.CrLf)
        sb.Append("[TOT]" & ControlChars.CrLf)
        sb.Append("     QCTE = " & ControlChars.CrLf)
        sb.Append("      QCT = " & ControlChars.CrLf)
        sb.Append("     QNFE = " & dsMDFeXNota.Tables(1).Rows.Count & ControlChars.CrLf)
        sb.Append("      QNF = " & ControlChars.CrLf)
        sb.Append("    QMDFE = " & ControlChars.CrLf)

        valor = valorMercadoria.ToString("N2")
        sb.Append("   VCARGA = " & valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append("    CUNID = 01" & ControlChars.CrLf)
        valor = CDec(dsNotaFiscal.Tables(0).Rows(0).Item("PesoBruto")).ToString("N4")

        sb.Append("   QCARGA = " & valor.Replace(".", "").Replace(",", ".") & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)
        sb.Append("[AUTXML]" & ControlChars.CrLf)
        sb.Append("    CNPJ = 07858433000121" & ControlChars.CrLf)
        sb.Append(ControlChars.CrLf)

        If dsNotaFiscal.Tables(0).Rows(0).Item("ObservacoesDeEmbarque").ToString.Length > 0 Then
            observacoes = Funcoes.EliminarCaracteresEspeciaisNF(dsNotaFiscal.Tables(0).Rows(0).Item("Observacoes")) & "|" & Funcoes.EliminarCaracteresEspeciaisNF(dsNotaFiscal.Tables(0).Rows(0).Item("ObservacoesDeEmbarque"))
        Else
            observacoes = dsNotaFiscal.Tables(0).Rows(0).Item("Observacoes")
        End If

        If Not String.IsNullOrWhiteSpace(observacoes) Then
            sb.Append("[INFADIC]" & ControlChars.CrLf)
            sb.Append("    INFADFISCO   = " & ControlChars.CrLf)
            sb.Append("    INFCPL       = " & Funcoes.EliminarCaracteresEspeciaisNF(observacoes) & ControlChars.CrLf)
        End If

        If mdfe.CodigoEmpresa = "03189063000800" OrElse tipoTac Then
            sb.Append(ControlChars.CrLf)
            sb.Append("[PRODPRED]" & ControlChars.CrLf)
            sb.Append(" TPCARGA = 05" & ControlChars.CrLf)
            If mdfe.CodigoEmpresa = "03189063000800" Then
                sb.Append("   XPROD = COURO" & ControlChars.CrLf)
            Else
                sb.Append("   XPROD = CAL" & ControlChars.CrLf)
            End If
            sb.Append("[INFLOCALCARREGA]" & ControlChars.CrLf)
            sb.Append("   CEP   = " & dsEmpresa.Tables(0).Rows(0).Item("Cep").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
            sb.Append("[INFLOCALDESCARREGA]" & ControlChars.CrLf)
            sb.Append("   CEP   = " & dsNotaFiscal.Tables(0).Rows(0).Item("CepDestino").Replace(".", "").Replace("-", "") & ControlChars.CrLf)
        End If

        Return sb.ToString()
    End Function

    Private Function getTextoDAMDFe(ByVal mdfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()

        sb.Append("chaveMDFe=" & mdfe.ChaveNFE & ControlChars.CrLf)
        sb.Append("impressora=pdf_mdfe" & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Private Function getTextoEncerrar(ByVal mdfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()

        sb.Append("TPEVENTO = 110112" & ControlChars.CrLf)
        sb.Append("CHAVEMDFE=" & mdfe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NPROT =" & mdfe.ProtocoloNota & ControlChars.CrLf)
        sb.Append("DTENC = " & String.Format("{0:yyyy-MM-dd}", DateTime.Now) & ControlChars.CrLf)
        sb.Append("CUF = " & mdfe.Cliente.Municipio.EstadoIbge & ControlChars.CrLf)
        sb.Append("CMUN = " & String.Format("{0}{1}", mdfe.Cliente.Municipio.EstadoIbge, Format(mdfe.Cliente.CodigoMunicipio, "00000")) & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Private Function getTextoCancelar(ByVal mdfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()

        sb.Append("TPEVENTO = 110111" & ControlChars.CrLf)
        sb.Append("CHAVEMDFE=" & mdfe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NPROT =" & mdfe.ProtocoloNota & ControlChars.CrLf)
        sb.Append("XJUST=" & Funcoes.EliminarCaracteresEspeciaisNF(mdfe.ObservacaoCancelamento) & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Private Function getTextoEmail(ByVal mdfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVEMDFE=" & mdfe.ChaveNFE & ControlChars.CrLf)
        Dim strCliente As String() = mdfe.Cliente.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
        For Each cliMail As String In strCliente
            sb.Append("DESTINATARIO=" & cliMail & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & "MDF-e Nr. " & mdfe.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & "Envio MDF-e Nr. " & mdfe.Codigo & ControlChars.CrLf)
        If Not String.IsNullOrWhiteSpace(mdfe.Empresa.EmailNFE) Then
            Dim strEmpresa As String = mdfe.Empresa.EmailNFE.Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)(0).Trim()
            sb.Append("EMAILEMITENTE=" & strEmpresa & ControlChars.CrLf)
            sb.Append("NOMEEMITENTE=" & mdfe.Empresa.Nome & ControlChars.CrLf)
        End If
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & "NAO" & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Private Function getTextoEmail(ByVal mdfe As [Lib].Negocio.NotaFiscal, ByVal lstClientes As [Lib].Negocio.ListCliente, ByVal strAssunto As String, ByVal strMensagem As String, ByVal compactado As Boolean) As String
        Dim sb As New StringBuilder()

        sb.Append("CHAVEMDFE=" & mdfe.ChaveNFE & ControlChars.CrLf)
        For Each objCliente As [Lib].Negocio.Cliente In lstClientes.Where(Function(s) Not String.IsNullOrWhiteSpace(s.EmailNFE)).ToList()
            sb.Append("DESTINATARIO=" & objCliente.EmailNFE & ControlChars.CrLf)
        Next
        sb.Append("ASSUNTO=" & strAssunto & " - MDF-e Nr. " & mdfe.Codigo & ControlChars.CrLf)
        sb.Append("MENSAGEM=" & strMensagem & " - Envio MDF-e Nr. " & mdfe.Codigo & ControlChars.CrLf)
        sb.Append("EMAILEMITENTE=" & mdfe.Empresa.EmailNFE & ControlChars.CrLf)
        sb.Append("NOMEEMITENTE=" & mdfe.Empresa.Nome & ControlChars.CrLf)
        sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)
        sb.Append("COMPACTADO=" & IIf(compactado, "SIM", "NAO") & ControlChars.CrLf)

        Return sb.ToString()
    End Function

    Private Sub Excluir(Optional ByVal IUD As String = "D")
        Dim Sqls As New ArrayList

        RecuperarSessaoMDFe()
        objMDFe.IUD = IUD

        If IUD = "D" Then
            For Each item As [Lib].Negocio.NotaFiscalXItem In objMDFe.Itens
                item.IUD = IUD
                For Each enc As [Lib].Negocio.NotaFiscalXItemXEncargo In item.Encargos
                    enc.IUD = IUD
                Next
            Next
        End If
        Dim objClienteXEmpresa As New [Lib].Negocio.ClienteXEmpresa(objMDFe.CodigoEmpresa, objMDFe.EnderecoEmpresa)
        If (objMDFe.DataHoraNFE.HasValue AndAlso DateTime.Now > CDate(objMDFe.DataHoraNFE).AddDays(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2)) Then
            MsgBox(Me.Page, "Este MDF-e não pode ser cancelado, pois o prazo para cancelamento terminou em: " & CDate(objMDFe.DataHoraNFE).AddDays(objClienteXEmpresa.PrazoCancelamentoNFE).AddMinutes(-2).ToString("dd/MM/yyyy HH:mm:ss"))
            Exit Sub
        End If

        objMDFe.ObservacaoCancelamento = "Erro na digitação do mdf-e."

        SalvarSessaoMDFe()

        If CancelarMDFe() Then
            'objMDFe.SalvarSql(Sqls)

            Dim sql = "UPDATE NotasFiscais SET Situacao = 2 "
            sql &= "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & " AND " & vbCrLf & _
                   "      Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente & " AND " & vbCrLf & _
                   "      EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf & _
                   "      Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo & "; " & vbCrLf
            Sqls.Add(sql)

            sql = "DELETE FROM NFEPendencias " & vbCrLf & _
                  "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & " AND " & vbCrLf & _
                  "      Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente & " AND " & vbCrLf & _
                  "      EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf & _
                  "      Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo & "; " & vbCrLf
            Sqls.Add(sql)

            sql = "UPDATE NFERealizadas SET Retorno = '101', MsgRetorno = 'Cancelamento de MDF-e homologado' " & vbCrLf & _
                  "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & " AND " & vbCrLf & _
                  "      Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente & " AND " & vbCrLf & _
                  "      EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf & _
                  "      Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo & "; " & vbCrLf
            Sqls.Add(sql)

            sql = "DELETE FROM NotasXNotas " & vbCrLf & _
                  "WHERE Empresa_Id = '" & objMDFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objMDFe.EnderecoEmpresa & " AND " & vbCrLf & _
                  "      Cliente_Id = '" & objMDFe.CodigoCliente & "' AND EndCliente_Id = " & objMDFe.EnderecoCliente & " AND " & vbCrLf & _
                  "      EntradaSaida_Id = '" & IIf(objMDFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND " & vbCrLf & _
                  "      Serie_Id = '" & objMDFe.Serie & "' AND Nota_Id = " & objMDFe.Codigo & "; " & vbCrLf
            Sqls.Add(sql)

            If Not Banco.GravaBanco(Sqls) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Exit Sub
            End If

            Dim msg As String = String.Empty
            If IUD = "C" Then
                msg = "Solicitação de cancelamento do MDF-e " & objMDFe.Codigo & "-" & objMDFe.Serie & " enviado com sucesso para SEFAZ."
            Else
                msg = "MDF-e " & objMDFe.Codigo & "-" & objMDFe.Serie & " Excluído com sucesso."
            End If

            MsgBox(Me.Page, msg)
            LimparCampos()
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal str As String)
        Try
            If Session("objCancelamento" & HID.Value) IsNot Nothing Then
                Cancelar(str)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Cancelar(ByVal observacao As String)
        RecuperarSessaoMDFe()

        If objMDFe IsNot Nothing Then
            objMDFe.IUD = "C"
            objMDFe.ObservacaoCancelamento = observacao.Trim()
            Excluir(objMDFe.IUD)
            SalvarSessaoMDFe()
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
            RecuperarSessaoMDFe()
            objMDFe.CodigoEmpresa = objEmpresa.Codigo
            objMDFe.EnderecoEmpresa = objEmpresa.CodigoEndereco
            objMDFe.Empresa = objEmpresa
            SalvarSessaoMDFe()
            Session.Remove("objEmpresaCTRC" & HID.Value)
        ElseIf Session("objClienteCTRC" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtCodigoCliente.Value = Funcoes.FormatarListItemCliente(obj).Value
            txtNomeCliente.Text = Funcoes.FormatarListItemCliente(obj).Text
            RecuperarSessaoMDFe()
            objMDFe.CodigoCliente = objCliente.Codigo
            objMDFe.EnderecoCliente = objCliente.CodigoEndereco
            objMDFe.Empresa = objCliente
            SalvarSessaoMDFe()
            Session.Remove("objClienteCTRC" & HID.Value)
        ElseIf Session("objClienteMDFe" & HID.Value) IsNot Nothing Then
            Dim objCliente As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtNomeDoCliente.Text = objCliente.Nome
            txtCnpjDoCliente.Text = objCliente.CodigoFormatado
            txtEnderecoDoCliente.Text = objCliente.Endereco & IIf(objCliente.Numero.ToString.Length > 0, ", " & objCliente.Numero.ToString, "")
            txtComplementoDoCliente.Text = objCliente.Complemento
            txtBairroDoCliente.Text = objCliente.Bairro
            txtCepDoCliente.Text = objCliente.CEP
            txtCidadeDoCliente.Text = objCliente.Cidade
            txtTelefoneDoCliente.Text = objCliente.Telefone
            txtEstadoDoCliente.Text = objCliente.CodigoEstado
            txtInscricaoDoCliente.Text = objCliente.InscricaoEstadual
            Session.Remove("objClienteMDFe" & HID.Value)
        ElseIf Session("objTransportadorCTRC" & HID.Value) IsNot Nothing Then
            Dim objTransportador As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
            txtNomeDoTransportador.Text = objTransportador.Nome
            txtRNTRCTransportador.Text = objTransportador.RNTRCTransportador
            txtEnderecoDoTransportador.Text = objTransportador.Endereco & IIf(objTransportador.Numero.ToString.Length > 0, ", " & objTransportador.Numero.ToString, "")
            txtCidadeDoTransportador.Text = objTransportador.Cidade
            txtEstadoDoTransportador.Text = objTransportador.CodigoEstado
            txtCnpjDoTransportador.Text = objTransportador.CodigoFormatado
            txtInscricaoDoTransportador.Text = objTransportador.InscricaoEstadual

            RecuperarSessaoMDFe()
            objMDFe.CarregandoNota = True
            objMDFe.CodigoCliente = objTransportador.Codigo
            objMDFe.EnderecoCliente = objTransportador.CodigoEndereco
            objMDFe.CodigoTransportador = objTransportador.Codigo
            objMDFe.EnderecoTransportador = objTransportador.CodigoEndereco
            objMDFe.Transportador = objTransportador
            objMDFe.NotaTrocaOrigem.CodigoTransportador = objTransportador.Codigo
            objMDFe.NotaTrocaOrigem.EnderecoTransportador = objTransportador.CodigoEndereco
            objMDFe.NotaTrocaOrigem.Transportador = objTransportador

            'cruzamento da operação e suboperação atráves do tipo de pessoa do cliente da nota, da classe da suboperação e se a suboperação é entrada ou saída
            Dim classe As String = "FRETES"
            Dim entradaSaida As String = "S"
            Dim tipo As Nullable(Of eTipoOperacao) = Nothing
            Dim fisicaJuridica As String = String.Empty
            Dim financeiro As String = "N"

            If objMDFe.NotaTrocaOrigem IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objMDFe.NotaTrocaOrigem.CodigoTransportador) Then
                fisicaJuridica = IIf(objMDFe.NotaTrocaOrigem.CodigoTransportador.IsCnpj(), "J", "F")
            End If

            Dim objOperacaoDeFrete As New [Lib].Negocio.OperacoesDeFretes(classe, entradaSaida, fisicaJuridica, financeiro, tipo)
            If (objOperacaoDeFrete IsNot Nothing AndAlso objOperacaoDeFrete.OperacaoId > 0 AndAlso objOperacaoDeFrete.SubOperacaoId > 0) Then
                objMDFe.CodigoOperacao = objOperacaoDeFrete.OperacaoId
                objMDFe.CodigoSubOperacao = objOperacaoDeFrete.SubOperacaoId
            Else
                Dim msg As String = String.Format("Não foi possível encontrar uma operação de frete para a nota fiscal selecionada, com os seguintes parâmetros: CLASSE: {0} - CPF/CNPJ: {1} - E/S: {2} - FINANCEIRO: {3}!", classe, fisicaJuridica, entradaSaida, financeiro)
                MsgBox(Me.Page, msg)
                LimparCampos()
                Exit Sub
            End If

            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlacaCTRC" & HID.Value)
            objMDFe.CarregandoNota = False
            Session.Remove("objTransportadorCTRC" & HID.Value)
            SalvarSessaoMDFe()
        ElseIf Session("objPlacaCTRC" & HID.Value) IsNot Nothing Then
            RecuperarSessaoMDFe()
            objMDFe.PlacaDetalhes = CType(Session("objPlacaCTRC" & HID.Value), [Lib].Negocio.Placa)
            If objMDFe.PlacaDetalhes.Motorista IsNot Nothing Then
                txtNomeDoMotorista.Text = objMDFe.PlacaDetalhes.Motorista.Nome
                txtEnderecoDoMotorista.Text = objMDFe.PlacaDetalhes.Motorista.Endereco & IIf(objMDFe.PlacaDetalhes.Motorista.Numero.ToString.Length > 0, ", " & objMDFe.PlacaDetalhes.Motorista.Numero.ToString, "")
                txtCidadeDoMotorista.Text = objMDFe.PlacaDetalhes.Motorista.Cidade
                txtEstadoDoMotorista.Text = objMDFe.PlacaDetalhes.Motorista.CodigoEstado
                txtCPFDoMotorista.Text = objMDFe.PlacaDetalhes.Motorista.CodigoFormatado
                txtHabilitacaoDoMotorista.Text = objMDFe.PlacaDetalhes.Motorista.Habilitacao
            End If
            If objMDFe.PlacaDetalhes IsNot Nothing Then
                txtCPlaca1.Text = objMDFe.PlacaDetalhes.Placa01
                txtCidadePlaca1.Text = objMDFe.PlacaDetalhes.CidadePlaca01
                txtEstadoPlaca1.Text = objMDFe.PlacaDetalhes.EstadoPlaca01
                txtRNTRCPlaca1.Text = objMDFe.PlacaDetalhes.RNTRCPlaca01
                txtProprietarioPlaca1.Text = objMDFe.PlacaDetalhes.CodigoProprietario01
                If Not String.IsNullOrWhiteSpace(objMDFe.PlacaDetalhes.Placa02) Then
                    txtCPlaca2.Text = objMDFe.PlacaDetalhes.Placa02
                    txtCidadePlaca2.Text = objMDFe.PlacaDetalhes.CidadePlaca02
                    txtEstadoPlaca2.Text = objMDFe.PlacaDetalhes.EstadoPlaca02
                    txtRNTRCPlaca2.Text = objMDFe.PlacaDetalhes.RNTRCPlaca02
                    txtProprietarioPlaca2.Text = objMDFe.PlacaDetalhes.CodigoProprietario02
                End If
                If Not String.IsNullOrWhiteSpace(objMDFe.PlacaDetalhes.Placa03) Then
                    txtCPlaca3.Text = objMDFe.PlacaDetalhes.Placa03
                    txtCidadePlaca3.Text = objMDFe.PlacaDetalhes.CidadePlaca03
                    txtEstadoPlaca3.Text = objMDFe.PlacaDetalhes.EstadoPlaca03
                    txtRNTRCPlaca3.Text = objMDFe.PlacaDetalhes.RNTRCPlaca03
                    txtProprietarioPlaca3.Text = objMDFe.PlacaDetalhes.CodigoProprietario03
                End If
            End If
            SalvarSessaoMDFe()
        ElseIf Session("objEmailMDFe" & HID.Value) IsNot Nothing Then
            Dim fm As New FilesManager()
            If fm.IsConnect() Then
                Dim Sqls As New ArrayList
                Try
                    RecuperarSessaoMDFe()
                    objMDFe = New [Lib].Negocio.NotaFiscal(objMDFe)
                    Dim lst As [Lib].Negocio.ListCliente = CType(Session("objEmailMDFe" & HID.Value), [Lib].Negocio.ListCliente)
                    If objMDFe.Empresa IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objMDFe.Empresa.EmailNFE) Then
                        Dim objFil As New [Lib].Negocio.Fil()
                        objFil.IUD = "I"
                        objFil.NomeArquivo = String.Format("emailmdfe{0:000000000}#{1}.txt", objMDFe.Codigo, objMDFe.CodigoEmpresa)
                        objFil.Texto = getTextoEmail(objMDFe, lst, Session("strAssunto" & HID.Value), Session("strMensagem" & HID.Value), CBool(Session("strCompactado" & HID.Value)))
                        objFil.SalvarSql(Sqls)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If
                    End If
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
            Else
                MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
            End If
        ElseIf Session("objMDFeXEstado" & HID.Value) IsNot Nothing Then
            RecuperarSessaoMDFe()
            Dim lstEstados As [Lib].Negocio.Estados = CType(Session("objMDFeXEstado" & HID.Value), [Lib].Negocio.Estados)
            If lstEstados IsNot Nothing AndAlso lstEstados.Count > 0 Then
                Dim strEstados = String.Empty
                objMDFe.NotaFiscalXPercursos.Clear()
                For Each e As [Lib].Negocio.Estado In lstEstados.OrderBy(Function(s) s.Sequencia).ToList()
                    If Not String.IsNullOrWhiteSpace(strEstados) Then
                        strEstados &= ", "
                    End If
                    strEstados &= e.Codigo

                    Dim nfp As New [Lib].Negocio.NotaFiscalXPercurso()
                    nfp.IUD = "I"
                    nfp.Empresa_Id = objMDFe.CodigoEmpresa
                    nfp.EndEmpresa_Id = objMDFe.EnderecoEmpresa
                    nfp.Cliente_Id = objMDFe.CodigoCliente
                    nfp.EndCliente_Id = objMDFe.EnderecoCliente
                    nfp.EntradaSaida_Id = objMDFe.EntradaSaida
                    nfp.Serie_Id = objMDFe.Serie
                    nfp.Nota_Id = objMDFe.Codigo
                    nfp.Estado_Id = e.Codigo
                    nfp.Ordem = e.Sequencia
                    objMDFe.NotaFiscalXPercursos.Add(nfp)
                Next
                lblEstados.Text = strEstados
            End If
            SalvarSessaoMDFe()
        End If
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Public Sub CarregarNotaFiscal()
        If Session("NfConsultaMDFe" & HID.Value) IsNot Nothing Then
            objMDFe = New [Lib].Negocio.NotaFiscal(CType(Session("NfConsultaMDFe" & HID.Value), [Lib].Negocio.NotaFiscal))
            objMDFe.IUD = "U"
            txtNumero.Text = objMDFe.Codigo
            txtSerie.Text = objMDFe.Serie
            txtChaveNFe.Text = objMDFe.ChaveNFE
            SalvarSessaoMDFe()
            CarregarMDFe(False, True)
            txtChaveNFe.Enabled = False
            txtNumero.Enabled = False
            txtSerie.Enabled = False
            If objMDFe.NossaEmissao Then
                txtDataDeEmissao.Enabled = False
                txtDataDeSaida.Enabled = False
                HideCalendar(Me.Page, txtDataDeEmissao)
                HideCalendar(Me.Page, txtDataDeSaida)
            End If
            Session.Remove("NfConsultaMDFe" & HID.Value)
            SalvarSessaoMDFe()
        End If
    End Sub

    Private Function getSql(ByVal strEmpresa() As String, ByVal strCliente() As String) As String
        Dim Sql As String = "SELECT NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, Clientes.Nome, NotasFiscais.EntradaSaida_Id, " & vbCrLf & _
                            "       NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.Operacao, NotasFiscais.SubOperacao, " & vbCrLf & _
                            "       NFERealizadas.ChaveNfe, 'False' AS chkMDFe, NFERealizadas.Data, " & vbCrLf & _
                            "       ISNULL((SELECT SUM(NFIxP.QuantidadeFiscal) AS QuantidadeFiscal " & vbCrLf & _
                            "				    FROM NotasFiscaisXItens NFIxP " & vbCrLf & _
                            "                WHERE (NFIxP.Empresa_Id      = NotasFiscais.Empresa_Id) " & vbCrLf & _
                            "                  AND (NFIxP.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id) " & vbCrLf & _
                            "                  AND (NFIxP.Cliente_Id      = NotasFiscais.Cliente_Id) " & vbCrLf & _
                            "                  AND (NFIxP.EndCliente_Id   = NotasFiscais.EndCliente_Id) " & vbCrLf & _
                            "                  AND (NFIxP.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id) " & vbCrLf & _
                            "                  AND (NFIxP.Serie_Id        = NotasFiscais.Serie_Id) " & vbCrLf & _
                            "                  AND (NFIxP.Nota_Id         = NotasFiscais.Nota_Id)), 0) AS PesoFiscal, " & vbCrLf & _
                            "       ISNULL((SELECT SUM(NFIxV.Valor) AS Valor " & vbCrLf & _
                            "				    FROM NotasFiscaisXItens AS NFIxV " & vbCrLf & _
                            "                WHERE (NFIxV.Empresa_Id      = NotasFiscais.Empresa_Id) " & vbCrLf & _
                            "                  AND (NFIxV.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id) " & vbCrLf & _
                            "                  AND (NFIxV.Cliente_Id      = NotasFiscais.Cliente_Id) " & vbCrLf & _
                            "                  AND (NFIxV.EndCliente_Id   = NotasFiscais.EndCliente_Id) " & vbCrLf & _
                            "                  AND (NFIxV.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id) " & vbCrLf & _
                            "                  AND (NFIxV.Serie_Id        = NotasFiscais.Serie_Id) " & vbCrLf & _
                            "                  AND (NFIxV.Nota_Id         = NotasFiscais.Nota_Id)), 0) AS ValorOficial, " & vbCrLf & _
                            "     MDFe.Empresa_Id as MDFe_Empresa_Id, " & vbCrLf & _
                            "     MDFe.EndEmpresa_Id as MDFe_EndEmpresa_Id, " & vbCrLf & _
                            "     MDFe.Cliente_Id as MDFe_Cliente_Id, " & vbCrLf & _
                            "     MDFe.EndCliente_Id as MDFe_EndCliente_Id, " & vbCrLf & _
                            "     MDFe.EntradaSaida_Id as MDFe_EntradaSaida_Id, " & vbCrLf & _
                            "     MDFe.Serie_Id as MDFe_Serie_Id, " & vbCrLf & _
                            "     MDFe.Nota_Id as MDFe_Nota_Id, " & vbCrLf & _
                            "     MDFe.TipoDeDocumento as MDFe_TipoDeDocumento " & vbCrLf & _
                            " FROM NotasFiscais " & vbCrLf & _
                            "	INNER JOIN  NFERealizadas " & vbCrLf & _
                            "				ON  NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                            "				AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                            "				AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                            "				AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                            "				AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                            "				AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                            "				AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id " & vbCrLf & _
                            "	INNER JOIN  Clientes " & vbCrLf & _
                            "				ON NotasFiscais.Cliente_Id     = Clientes.Cliente_Id " & vbCrLf & _
                            "				AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf & _
                            " LEFT JOIN NotasXNotas                                                          " & vbCrLf & _
                            "  		ON  NotasFiscais.Empresa_Id      = NotasXNotas.OrigemEmpresa_Id          " & vbCrLf & _
                            "		AND NotasFiscais.EndEmpresa_Id   = NotasXNotas.OrigemEndEmpresa_Id       " & vbCrLf & _
                            "		AND NotasFiscais.Cliente_Id      = NotasXNotas.OrigemCliente_Id          " & vbCrLf & _
                            "		AND NotasFiscais.EndCliente_Id   = NotasXNotas.OrigemEndCliente_Id       " & vbCrLf & _
                            "		AND NotasFiscais.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id     " & vbCrLf & _
                            "		AND NotasFiscais.Serie_Id        = NotasXNotas.OrigemSerie_Id            " & vbCrLf & _
                            "		AND NotasFiscais.Nota_Id         = NotasXNotas.OrigemNota_Id             " & vbCrLf & _
                            "LEFT JOIN NotasFiscais MDFe " & vbCrLf & _
                            "		ON  MDFe.Empresa_Id      =  NotasXNotas.Empresa_Id " & vbCrLf & _
                            "		AND MDFe.EndEmpresa_Id   =  NotasXNotas.EndEmpresa_Id  " & vbCrLf & _
                            "		AND MDFe.Cliente_Id      =  NotasXNotas.Cliente_Id  " & vbCrLf & _
                            "		AND MDFe.EndCliente_Id   =  NotasXNotas.EndCliente_Id  " & vbCrLf & _
                            "		AND MDFe.EntradaSaida_Id =  NotasXNotas.EntradaSaida_Id  " & vbCrLf & _
                            "		AND MDFe.Serie_Id        =  NotasXNotas.Serie_Id  " & vbCrLf & _
                            "		AND MDFe.Nota_Id         =  NotasXNotas.Nota_Id  " & vbCrLf & _
                            "		AND MDFe.TipoDeDocumento = " & CInt(eTipoDeDocumento.ManifestoEletronico) & " " & vbCrLf & _
                            "LEFT JOIN NFERealizadas MDFeR " & vbCrLf & _
                            "		ON  MDFeR.Empresa_Id      =  MDFe.Empresa_Id " & vbCrLf & _
                            "		AND MDFeR.EndEmpresa_Id   =  MDFe.EndEmpresa_Id  " & vbCrLf & _
                            "		AND MDFeR.Cliente_Id      =  MDFe.Cliente_Id  " & vbCrLf & _
                            "		AND MDFeR.EndCliente_Id   =  MDFe.EndCliente_Id  " & vbCrLf & _
                            "		AND MDFeR.EntradaSaida_Id =  MDFe.EntradaSaida_Id  " & vbCrLf & _
                            "		AND MDFeR.Serie_Id        =  MDFe.Serie_Id  " & vbCrLf & _
                            "		AND MDFeR.Nota_Id         =  MDFe.Nota_Id  " & vbCrLf & _
                            "WHERE 1=1 " & vbCrLf & _
                            "  AND (NotasFiscais.Empresa_Id = '" & strEmpresa(0) & "') " & vbCrLf & _
                            "  AND (NotasFiscais.EndEmpresa_Id = " & strEmpresa(1) & ")" & vbCrLf & _
                            "  AND (NotasFiscais.EntradaSaida_Id = '" & IIf(rdoSaida.Checked, "S", "E") & "') " & vbCrLf & _
                            "  AND (NotasFiscais.Situacao = 1) " & vbCrLf & _
                            "  AND LEN(isnull(NFERealizadas.Chavenfe,'')) > 0 " & vbCrLf

        If strCliente IsNot Nothing AndAlso strCliente.Length > 0 Then
            Sql &= "  AND (NotasFiscais.Cliente_Id = '" & strCliente(0) & "') " & vbCrLf & _
                   "  AND (NotasFiscais.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtNumMDFe.Text) Then
            Dim strNota = txtNumMDFe.Text.Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)
            Sql &= "  AND (NotasFiscais.Nota_Id in (" & String.Join(",", strNota) & ")) " & vbCrLf
        End If

        Sql &= "  AND (NotasFiscais.TipoDeDocumento = 1) " & vbCrLf & _
               "  AND ((isnull(MDFe.Nota_Id,0) = 0) OR (isnull(MDFeR.Retorno,'') = '135'))" & vbCrLf

        Sql &= "  AND (NotasFiscais.Movimento between '" & txtPeriodoInicial.Text.ToSqlDate() & "' and '" & txtPeriodoFinal.Text.ToSqlDate() & "') " & vbCrLf & _
                   "ORDER BY NotasFiscais.Nota_Id, NotasFiscais.Serie_Id, NFERealizadas.Data DESC " & vbCrLf

        Return Sql
    End Function

    Private Sub ConsultarNotas()
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "É necessário selecionar a empresa.")
            Exit Sub
        End If

        Dim strEmpresa() As String = DdlEmpresa.SelectedValue.Split("-")
        Dim strCliente() As String = Nothing

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            strCliente = txtCodigoCliente.Value.Split("-")
        End If

        Dim sql As String = getSql(strEmpresa, strCliente)
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NotasFiscais")

        grdNFe.DataSource = ds
        grdNFe.DataBind()

        EnviarNFePendentes()

    End Sub

    Private Sub LimparCampos()
        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaClientesDireto.SetarHID(HID.Value)
        ucConsultaMDFeXNotas.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucMDFeXEstado.SetarHID(HID.Value)
        ucEmailNFe.SetarHID(HID.Value)

        txtCodigoCliente.Value = ""
        txtNomeCliente.Text = ""
        txtNumMDFe.Text = ""
        rdoSaida.Checked = True
        rdoEntrada.Checked = True
        'txtPeriodoInicial.Text = String.Format("01/01/{0}", DateTime.Now.Year)
        txtPeriodoInicial.Text = DateTime.Now.ToShortDateString()
        txtPeriodoFinal.Text = DateTime.Now.ToShortDateString()

        txtNomeDaEmpresa.Text = ""
        txtEnderecoDaEmpresa.Text = ""
        txtComplementoDaEmpresa.Text = ""
        txtBairroDaEmpresa.Text = ""
        txtCepDaEmpresa.Text = ""
        txtInscricaoDaEmpresa.Text = ""
        txtCnpjDaEmpresa.Text = ""

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

        txtEnderecoDoCliente.Text = ""
        txtComplementoDoCliente.Text = ""
        txtBairroDoCliente.Text = ""
        txtCepDoCliente.Text = ""
        txtDataDeSaida.Text = DateTime.Now.ToShortDateString()

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
        txtPesoBruto.Text = ""
        txtValor.Text = ""
        txtOperacao.Text = ""
        txtCfop.Text = ""

        txtNumero.Enabled = True
        txtSerie.Enabled = True

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
        txtObservacao.Height = Unit.Pixel(150)
        txtObservacao.Text = ""

        gridNotasFiscais.DataSource = Nothing
        gridNotasFiscais.DataBind()

        txtEmbarque.Text = ""
        txtMunicipioEmbarque.Text = ""
        txtEstadoEmbarque.Text = ""
        txtNomeOrigemDestino.Text = ""
        txtMunicipioOrigemDestino.Text = ""
        txtEstadoOrigemDestino.Text = ""
        txtChaveNFe.Text = ""

        ShowCalendar(Me.Page, txtDataDeEmissao)
        ShowCalendar(Me.Page, txtDataDeSaida)

        txtDataDeEmissao.Enabled = True
        txtDataDeSaida.Enabled = True

        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkLimpar.Parent.Visible = True
        lnkEspelho.Parent.Visible = True
        lnkEncerrar.Parent.Visible = False
        lnkEnviarSEFAZ.Parent.Visible = False
        lnkEnviarEmail.Parent.Visible = False
        lnkDAMDFE.Parent.Visible = False
        lnkVisualizar.Parent.Visible = False

        objMDFe = New [Lib].Negocio.NotaFiscal()
        objMDFe.IUD = "I"

        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            Dim strEmpresa() As String = DdlEmpresa.SelectedValue.ToString.Split("-")
            Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
            If objEmpresa IsNot Nothing Then
                objMDFe.CodigoEmpresa = objEmpresa.Codigo
                objMDFe.EnderecoEmpresa = objEmpresa.CodigoEndereco
            End If
        End If

        objMDFe.CodigoTipoDeDocumento = getTipoDeDocumento()
        objMDFe.EntradaSaida = eEntradaSaida.Saida
        objMDFe.CIFFOB = eTiposFrete.CIF
        objMDFe.Eletronica = True
        objMDFe.CodigoSituacao = getSituacao()
        SalvarSessaoMDFe()
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub Selecionar()
        Try
            Dim aux As Integer = 0
            For Each row As GridViewRow In grdNFe.Rows
                If CType(row.FindControl("chkMDFe"), CheckBox).Checked Then
                    aux += 1
                End If
            Next

            If Not aux > 0 Then
                For Each row As GridViewRow In grdNFe.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkMDFe"), CheckBox)
                    If (chk IsNot Nothing) Then
                        chk.Checked = False
                    End If
                Next
                MsgBox(Me.Page, "Selecione pelo menos um registro para emissão do MDF-e.")
            Else
                Dim auxMsg As Boolean = True
                Dim msg As String = ""

                For Each row As GridViewRow In grdNFe.Rows
                    Dim chk As CheckBox = CType(row.FindControl("chkMDFe"), CheckBox)
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

                        If (lstNotaFiscal Is Nothing) Then
                            lstNotaFiscal = New List(Of [Lib].Negocio.NotaFiscal)
                        End If
                        lstNotaFiscal.Add(objNotaFiscal)

                        If (objNotaFiscal Is Nothing) Then
                            auxMsg = False
                            Continue For
                        End If
                    End If
                Next

                Dim uf As String = lstNotaFiscal(0).Destino.CodigoEstado.ToUpper().Trim()
                If lstNotaFiscal.Any(Function(s) s.Destino.CodigoEstado.ToUpper().Trim() <> uf) AndAlso rdoSaida.Checked Then
                    msg &= "Não é possível emitir um documento MDF-e com clientes de UF diferentes."
                End If

                If Not String.IsNullOrWhiteSpace(msg) Then
                    MsgBox(Me.Page, msg)
                    Exit Sub
                Else
                    GerarMDFe()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function EnviarNFePendentes() As Boolean
        Dim pendenciaOK As Boolean = True

        Try
            Dim Sql As String = "SELECT nf.Empresa_Id, cast(nf.Nota_Id as nvarchar) + '-' + cast(nf.Serie_Id as nvarchar) as nf, nf.UsuarioInclusao, nf.UsuarioAlteracao " & vbCrLf &
                                  " FROM NotasFiscais nf  " & vbCrLf &
                                  "WHERE nf.Eletronica = 'S' " & vbCrLf &
                                  "  AND nf.NossaEmissao = 'S' " & vbCrLf &
                                  "  AND nf.TipoDeDocumento = 12 " & vbCrLf &
                                  "  AND nf.Situacao in(4,7) " & vbCrLf &
                                  "  AND datepart(hh, GETDATE() - nf.UsuarioInclusaoData) > 0 " & vbCrLf &
                                  "  AND datepart(mi, GETDATE() - nf.UsuarioInclusaoData) > 5 " & vbCrLf &
                                  "ORDER BY nf.Movimento DESC"

            Dim lstMail As New List(Of String)
            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotasFiscais")
            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                Dim Filial As String = String.Empty
                Dim bodyHTML = "Sr(s), <br/><br/> o(s) Mdfe(s) a seguir estão pendentes de envio para a SEFAZ: <br/>" & vbCrLf
                bodyHTML &= "<ul>"
                For Each row As DataRow In ds.Tables(0).Rows
                    If String.IsNullOrEmpty(Filial) Then Filial = row("Empresa_Id")

                    bodyHTML &= "<li>"
                    bodyHTML &= "<label>" & row("Empresa_Id") & " - " & row("nf") & " (Usuário Inclusão: " & row("UsuarioInclusao") & ")" & "</label>"
                    bodyHTML &= "</li>"

                    If Not IsDBNull(row("UsuarioInclusao")) AndAlso Not String.IsNullOrWhiteSpace(row("UsuarioInclusao")) Then
                        Dim objUsuario As New [Lib].Negocio.Usuario(row("UsuarioInclusao"))
                        If objUsuario IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objUsuario.Email) AndAlso objUsuario.Email.Trim() <> "&nbsp;" Then
                            lstMail.Add(objUsuario.Email)
                        End If
                    End If

                    If Not IsDBNull(row("UsuarioAlteracao")) AndAlso Not String.IsNullOrWhiteSpace(row("UsuarioAlteracao")) Then
                        Dim objUsuario As New [Lib].Negocio.Usuario(row("UsuarioAlteracao"))
                        If objUsuario IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objUsuario.Email) AndAlso objUsuario.Email.Trim() <> "&nbsp;" Then
                            lstMail.Add(objUsuario.Email)
                        End If
                    End If
                Next
                bodyHTML &= "</ul>"

                Dim Empresa = New [Lib].Negocio.Cliente(Filial, 0)
                Dim errorMsg As String = String.Empty
                Dim subject As String = String.Format("Empresa " & Empresa.CodigoFormatado & " (" & Empresa.Cidade & "/" & Empresa.CodigoEstado & ") - NF-e pendentes de envio para SEFAZ - {0:dd/MM/yyyy HH:mm}", DateTime.Now)
                Dim smtp = Funcoes.GetSmtpSettings()
                Dim fromMail = Funcoes.GetFromMail()

                Sql = "SELECT u.Email " & vbCrLf &
                      "  FROM ConfiguracaoXUsuario cxu " & vbCrLf &
                      " INNER JOIN Usuarios u " & vbCrLf &
                      "    ON (u.Usuario_Id = cxu.Usuario_Id) " & vbCrLf &
                      " WHERE cxu.Etapa_Id = " & eEtapa.NFePendencias

                Dim dsMail As DataSet = Banco.ConsultaDataSet(Sql, "ConfiguracaoXUsuario")
                If dsMail IsNot Nothing AndAlso dsMail.Tables IsNot Nothing AndAlso dsMail.Tables.Count > 0 AndAlso dsMail.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In dsMail.Tables(0).Rows
                        If Not String.IsNullOrWhiteSpace(row("Email")) Then
                            lstMail.Add(row("Email"))
                        End If
                    Next
                End If

                If lstMail IsNot Nothing AndAlso lstMail.Count > 0 Then
                    If Not Funcoes.SendMail(fromMail, "NGS SOLUÇÕES", lstMail, subject, bodyHTML, smtp, errorMsg) Then
                        MsgBox(Me.Page, "Erro ao tententar enviar E-Mail das Notas Pendentes. " & errorMsg, eTitulo.Erro)
                        pendenciaOK = False
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
            pendenciaOK = False
        End Try

        Return pendenciaOK
    End Function

    Private Function getSituacao() As Integer
        Return CInt(eSituacao.Bloqueado)
    End Function

    Private Function getTipoDeDocumento() As Integer
        Return Convert.ToInt32(eTipoDeDocumento.ManifestoEletronico)
    End Function

    Private Sub GerarMDFe()
        Dim index As Integer = 0

        For Each row As GridViewRow In grdNFe.Rows
            Dim chk As CheckBox = CType(row.FindControl("chkMDFe"), CheckBox)
            If (chk IsNot Nothing AndAlso chk.Checked) Then
                RecuperarSessaoMDFe()
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

                objMDFe = If(objMDFe, New [Lib].Negocio.NotaFiscal())

                If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                    Dim strEmpresa() As String = DdlEmpresa.SelectedValue.ToString.Split("-")
                    Dim objEmpresa As [Lib].Negocio.Cliente = New [Lib].Negocio.Cliente(strEmpresa(0), strEmpresa(1))
                    If objEmpresa IsNot Nothing Then
                        objMDFe.CodigoEmpresa = objEmpresa.Codigo
                        objMDFe.EnderecoEmpresa = objEmpresa.CodigoEndereco
                        objMDFe.Empresa = objEmpresa
                    End If
                End If

                objMDFe.CodigoCliente = objNotaFiscal.CodigoTransportador
                objMDFe.EnderecoCliente = objNotaFiscal.EnderecoTransportador
                objMDFe.CodigoDeposito = objNotaFiscal.CodigoDeposito
                objMDFe.EnderecoDeposito = objNotaFiscal.EnderecoDeposito
                objMDFe.EntradaSaida = eEntradaSaida.Saida
                objMDFe.NossaEmissao = True
                objMDFe.Movimento = objNotaFiscal.Movimento
                objMDFe.DataTermino = objNotaFiscal.DataTermino

                If row.Cells(1).Text() = "S" Then
                    objMDFe.CodigoDestino = objNotaFiscal.CodigoDestino
                    objMDFe.EnderecoDestino = objNotaFiscal.EnderecoDestino
                Else
                    objMDFe.CodigoDestino = objNotaFiscal.CodigoDeposito
                    objMDFe.EnderecoDestino = objNotaFiscal.EnderecoDeposito
                End If

                If objNotaFiscal.CodigoLocalEmbarque.Length > 0 Then
                    objMDFe.CodigoLocalEmbarque = objNotaFiscal.CodigoLocalEmbarque
                    objMDFe.EndLocalEmbarque = objNotaFiscal.EndLocalEmbarque
                Else
                    If row.Cells(1).Text() = "S" Then
                        objMDFe.CodigoLocalEmbarque = objNotaFiscal.CodigoDeposito
                        objMDFe.EndLocalEmbarque = objNotaFiscal.EnderecoDeposito
                    Else
                        objMDFe.CodigoLocalEmbarque = objNotaFiscal.CodigoCliente
                        objMDFe.EndLocalEmbarque = objNotaFiscal.EnderecoCliente
                    End If
                End If

                objMDFe.Empresa.Empresa = New ClienteXEmpresa(objMDFe.CodigoEmpresa, objMDFe.EnderecoEmpresa)
                objMDFe.CodigoMotorista = objNotaFiscal.CodigoMotorista
                objMDFe.EnderecoMotorista = objNotaFiscal.EnderecoMotorista
                objMDFe.CodigoTransportador = objNotaFiscal.CodigoTransportador
                objMDFe.EnderecoTransportador = objNotaFiscal.EnderecoTransportador
                objMDFe.PlacaDetalhes = objNotaFiscal.PlacaDetalhes
                objMDFe.ObservacoesDeEmbarque = String.Format("{0} {1}", objNotaFiscal.Observacoes, objNotaFiscal.ObservacoesDeEmbarque).Trim()
                objMDFe.NotaTrocaOrigem = objNotaFiscal
                objMDFe.NotasTrocaOrigem = lstNotaFiscal
                objMDFe.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                objMDFe.DataInclusao = Now()
                txtDataDeEmissao.Enabled = True
                txtDataDeSaida.Enabled = True
                SalvarSessaoMDFe()
            End If
        Next

        CarregarMDFe(True)
    End Sub

    Protected Sub CarregarMDFe(ByVal ctrNovo As Boolean, Optional ByVal IsLoading As Boolean = False)
        Try
            RecuperarSessaoMDFe()

            If ctrNovo Then
                objMDFe.CodigoTipoDeDocumento = getTipoDeDocumento()
                objMDFe.Movimento = DateTime.Now
                objMDFe.DataNota = DateTime.Now
                txtNumero.Enabled = False
                txtSerie.Enabled = False
            End If

            txtDataDeEmissao.Text = objMDFe.DataNota.ToShortDateString()
            txtDataDeSaida.Text = objMDFe.Movimento.ToShortDateString()
            txtNomeDaEmpresa.Text = objMDFe.Empresa.Nome
            txtEnderecoDaEmpresa.Text = objMDFe.Empresa.Endereco & IIf(objMDFe.Empresa.Numero.ToString.Length > 0, ", " & objMDFe.Empresa.Numero.ToString, "")
            txtComplementoDaEmpresa.Text = objMDFe.Empresa.Complemento
            txtBairroDaEmpresa.Text = objMDFe.Empresa.Bairro
            txtCepDaEmpresa.Text = objMDFe.Empresa.CEP
            txtInscricaoDaEmpresa.Text = objMDFe.Empresa.InscricaoEstadual
            txtCnpjDaEmpresa.Text = objMDFe.Empresa.CodigoFormatado

            If IsLoading Then
                txtObservacao.Text = String.Format("{0} {1}", objMDFe.Observacoes, objMDFe.ObservacoesDeEmbarque).Trim()
            Else
                Dim rowIndex = 1
                If lstNotaFiscal IsNot Nothing AndAlso lstNotaFiscal.Count > 0 Then
                    For Each nf As [Lib].Negocio.NotaFiscal In lstNotaFiscal
                        txtObservacao.Text &= String.Format("REF NF {0} {1}", nf.Codigo, nf.Cliente.Nome).Trim()
                        If Not String.IsNullOrWhiteSpace(txtObservacao.Text) AndAlso rowIndex < lstNotaFiscal.Count Then
                            txtObservacao.Text &= ", "
                        End If
                        rowIndex += 1
                    Next
                End If
            End If

            If objMDFe.NotaTrocaOrigem.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida Then
                txtNomeDoRemetente.Text = objMDFe.NotaTrocaOrigem.Empresa.Nome
                txtEnderecoDoRemetente.Text = objMDFe.NotaTrocaOrigem.Empresa.Endereco & IIf(objMDFe.NotaTrocaOrigem.Empresa.Numero.ToString.Length > 0, ", " & objMDFe.NotaTrocaOrigem.Empresa.Numero.ToString, "")
                txtComplementoDoRemetente.Text = objMDFe.NotaTrocaOrigem.Empresa.Complemento
                txtBairroDoRemetente.Text = objMDFe.NotaTrocaOrigem.Empresa.Bairro
                txtCepDoRemetente.Text = objMDFe.NotaTrocaOrigem.Empresa.CEP
                txtInscricaoDoRemetente.Text = objMDFe.NotaTrocaOrigem.Empresa.InscricaoEstadual
                txtCnpjDoRemetente.Text = objMDFe.NotaTrocaOrigem.Empresa.CodigoFormatado
            Else
                txtNomeDoRemetente.Text = objMDFe.NotaTrocaOrigem.Cliente.Nome
                txtEnderecoDoRemetente.Text = objMDFe.NotaTrocaOrigem.Cliente.Endereco & IIf(objMDFe.NotaTrocaOrigem.Cliente.Numero.ToString.Length > 0, ", " & objMDFe.NotaTrocaOrigem.Cliente.Numero.ToString, "")
                txtComplementoDoRemetente.Text = objMDFe.NotaTrocaOrigem.Cliente.Complemento
                txtBairroDoRemetente.Text = objMDFe.NotaTrocaOrigem.Cliente.Bairro
                txtCepDoRemetente.Text = objMDFe.NotaTrocaOrigem.Cliente.CEP
                txtInscricaoDoRemetente.Text = objMDFe.NotaTrocaOrigem.Cliente.InscricaoEstadual
                txtCnpjDoRemetente.Text = objMDFe.NotaTrocaOrigem.Cliente.CodigoFormatado
            End If
            If objMDFe.NotaTrocaOrigem.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida Then
                txtNomeDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.Nome
                txtCnpjDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.CodigoFormatado
                txtEnderecoDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.Endereco & IIf(objMDFe.NotaTrocaOrigem.Cliente.Numero.ToString.Length > 0, ", " & objMDFe.NotaTrocaOrigem.Cliente.Numero.ToString, "")
                txtComplementoDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.Complemento
                txtBairroDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.Bairro
                txtCepDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.CEP
                txtCidadeDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.Cidade
                txtTelefoneDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.Telefone
                txtEstadoDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.CodigoEstado
                txtInscricaoDoCliente.Text = objMDFe.NotaTrocaOrigem.Cliente.InscricaoEstadual
            Else
                txtNomeDoCliente.Text = objMDFe.Empresa.Nome
                txtCnpjDoCliente.Text = objMDFe.Empresa.CodigoFormatado
                txtEnderecoDoCliente.Text = objMDFe.Empresa.Endereco & IIf(objMDFe.Empresa.Numero.ToString.Length > 0, ", " & objMDFe.Empresa.Numero.ToString, "")
                txtComplementoDoCliente.Text = objMDFe.Empresa.Complemento
                txtBairroDoCliente.Text = objMDFe.Empresa.Bairro
                txtCepDoCliente.Text = objMDFe.Empresa.CEP
                txtCidadeDoCliente.Text = objMDFe.Empresa.Cidade
                txtTelefoneDoCliente.Text = objMDFe.Empresa.Telefone
                txtEstadoDoCliente.Text = objMDFe.Empresa.CodigoEstado
                txtInscricaoDoCliente.Text = objMDFe.Empresa.InscricaoEstadual
            End If
            If objMDFe.LocalEmbarque IsNot Nothing Then
                txtEmbarque.Text = objMDFe.LocalEmbarque.CodigoFormatado & " - " & objMDFe.LocalEmbarque.Nome
                txtMunicipioEmbarque.Text = objMDFe.LocalEmbarque.Cidade
                txtEstadoEmbarque.Text = objMDFe.LocalEmbarque.CodigoEstado
            End If

            If objMDFe.Destino IsNot Nothing Then
                txtNomeOrigemDestino.Text = objMDFe.Destino.CodigoFormatado & " - " & objMDFe.Destino.Nome
                txtMunicipioOrigemDestino.Text = objMDFe.Destino.Cidade
                txtEstadoOrigemDestino.Text = objMDFe.Destino.CodigoEstado
            End If
            If objMDFe.Transportador IsNot Nothing Then
                txtNomeDoTransportador.Text = objMDFe.Transportador.Nome
                txtRNTRCTransportador.Text = objMDFe.Transportador.RNTRCTransportador
                txtEnderecoDoTransportador.Text = objMDFe.Transportador.Endereco & IIf(objMDFe.Transportador.Numero.ToString.Length > 0, ", " & objMDFe.Transportador.Numero.ToString, "")
                txtCidadeDoTransportador.Text = objMDFe.Transportador.Cidade
                txtEstadoDoTransportador.Text = objMDFe.Transportador.CodigoEstado
                txtCnpjDoTransportador.Text = objMDFe.Transportador.CodigoFormatado
                txtInscricaoDoTransportador.Text = objMDFe.Transportador.InscricaoEstadual
            End If

            If ctrNovo Then
                objMDFe.CarregandoNota = False
                objMDFe.NossaEmissao = True
            End If
            If ctrNovo Then
                lblUsuario.Text = objMDFe.UsuarioInclusao
            Else
                lblUsuario.Text = IIf(objMDFe.UsuarioAlteracao.Length = 0, objMDFe.UsuarioInclusao, objMDFe.UsuarioAlteracao)
            End If

            If objMDFe.Motorista IsNot Nothing Then
                txtNomeDoMotorista.Text = objMDFe.Motorista.Nome
                txtEnderecoDoMotorista.Text = objMDFe.Motorista.Endereco & IIf(objMDFe.Motorista.Numero.ToString.Length > 0, ", " & objMDFe.Motorista.Numero.ToString, "")
                txtCidadeDoMotorista.Text = objMDFe.Motorista.Cidade
                txtEstadoDoMotorista.Text = objMDFe.Motorista.CodigoEstado
                txtCPFDoMotorista.Text = objMDFe.Motorista.CodigoFormatado
                txtHabilitacaoDoMotorista.Text = objMDFe.Motorista.Habilitacao
            End If
            If objMDFe.PlacaDetalhes IsNot Nothing Then
                txtCPlaca1.Text = objMDFe.PlacaDetalhes.Placa01
                txtCidadePlaca1.Text = objMDFe.PlacaDetalhes.CidadePlaca01
                txtEstadoPlaca1.Text = objMDFe.PlacaDetalhes.EstadoPlaca01
                txtRNTRCPlaca1.Text = objMDFe.PlacaDetalhes.RNTRCPlaca01
                txtProprietarioPlaca1.Text = objMDFe.PlacaDetalhes.CodigoProprietario01
                If objMDFe.PlacaDetalhes.Placa02.Length > 0 Then
                    txtCPlaca2.Text = objMDFe.PlacaDetalhes.Placa02
                    txtCidadePlaca2.Text = objMDFe.PlacaDetalhes.CidadePlaca02
                    txtEstadoPlaca2.Text = objMDFe.PlacaDetalhes.EstadoPlaca02
                    txtRNTRCPlaca2.Text = objMDFe.PlacaDetalhes.RNTRCPlaca02
                    txtProprietarioPlaca2.Text = objMDFe.PlacaDetalhes.CodigoProprietario02
                End If
                If objMDFe.PlacaDetalhes.Placa03.Length > 0 Then
                    txtCPlaca3.Text = objMDFe.PlacaDetalhes.Placa03
                    txtCidadePlaca3.Text = objMDFe.PlacaDetalhes.CidadePlaca03
                    txtEstadoPlaca3.Text = objMDFe.PlacaDetalhes.EstadoPlaca03
                    txtRNTRCPlaca3.Text = objMDFe.PlacaDetalhes.RNTRCPlaca03
                    txtProprietarioPlaca3.Text = objMDFe.PlacaDetalhes.CodigoProprietario03
                End If
            End If

            Dim ds As New DataSet
            Dim dt As New DataTable("NotaFiscal")
            dt.Columns.Add("Nota", Type.GetType("System.String"))
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

            If IsLoading Then
                If lstNotaFiscal Is Nothing Then
                    lstNotaFiscal = New List(Of [Lib].Negocio.NotaFiscal)
                End If
                If objMDFe.NotasTrocaOrigem IsNot Nothing AndAlso objMDFe.NotasTrocaOrigem.Count > 0 Then
                    For Each origem As [Lib].Negocio.NotaFiscal In objMDFe.NotasTrocaOrigem
                        lstNotaFiscal.Add(New [Lib].Negocio.NotaFiscal(origem))
                    Next
                ElseIf objMDFe.NotaTrocaOrigem IsNot Nothing Then
                    lstNotaFiscal.Add(New [Lib].Negocio.NotaFiscal(objMDFe.NotaTrocaOrigem))
                End If
            End If

            'cruzamento da operação e suboperação atráves do tipo de pessoa do cliente da nota, da classe da suboperação e se a suboperação é entrada ou saída
            objMDFe.CarregandoNota = True
            If Not IsLoading Then
                Dim classe As String = "FRETES"
                Dim entradaSaida As String = "S"
                Dim tipo As Nullable(Of eTipoOperacao) = Nothing
                Dim fisicaJuridica As String = String.Empty
                Dim financeiro As String = "N"

                If objMDFe.NotaTrocaOrigem IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objMDFe.NotaTrocaOrigem.CodigoTransportador) Then
                    fisicaJuridica = IIf(objMDFe.NotaTrocaOrigem.CodigoTransportador.IsCnpj(), "J", "F")
                Else
                    MsgBox(Me.Page, "Transportador da Nota Fiscal não foi encontrado!")
                    Exit Sub
                End If

                Dim objOperacaoDeFrete As New [Lib].Negocio.OperacoesDeFretes(classe, entradaSaida, fisicaJuridica, financeiro, tipo)
                If (objOperacaoDeFrete IsNot Nothing AndAlso objOperacaoDeFrete.OperacaoId > 0 AndAlso objOperacaoDeFrete.SubOperacaoId > 0) Then
                    objMDFe.CodigoOperacao = objOperacaoDeFrete.OperacaoId
                    objMDFe.CodigoSubOperacao = objOperacaoDeFrete.SubOperacaoId

                    Dim op As New SubOperacao(objMDFe.CodigoOperacao, objMDFe.CodigoSubOperacao)
                    txtOperacao.Text = String.Format("{0}-{1} - {2}", objMDFe.CodigoOperacao, objMDFe.CodigoSubOperacao, IIf(op IsNot Nothing, op.Descricao, String.Empty))

                    If String.IsNullOrWhiteSpace(objMDFe.CodigoTransportador) Then
                        btnTransportador.Enabled = True
                    End If
                End If
            Else
                Dim op As New SubOperacao(objMDFe.CodigoOperacao, objMDFe.CodigoSubOperacao)
                txtOperacao.Text = String.Format("{0}-{1} - {2}", objMDFe.CodigoOperacao, objMDFe.CodigoSubOperacao, IIf(op IsNot Nothing, op.Descricao, String.Empty))
            End If

            'pesquisa a lista de notas selecionadas para exbir todos os produtos na gridview
            If lstNotaFiscal IsNot Nothing AndAlso lstNotaFiscal.Count > 0 Then
                For Each item As [Lib].Negocio.NotaFiscalXItem In lstNotaFiscal.SelectMany(Function(s) s.Itens).OrderBy(Function(s) s.NotaFiscal.Codigo).ToList
                    Dim dr As DataRow = ds.Tables(0).NewRow()
                    dr("Nota") = IIf(item.NotaFiscal.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida, "S", "E") & "-" & item.NotaFiscal.Codigo & "-" & item.NotaFiscal.Serie
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

                Dim objClienteXEmpresa As New [Lib].Negocio.ClienteXEmpresa(objMDFe.CodigoEmpresa, objMDFe.EnderecoEmpresa)
                If Not (objClienteXEmpresa IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objClienteXEmpresa.CodigoProdutoDeMDFe)) Then
                    MsgBox(Me.Page, "É necessário informar o código do produto de MDF-e no cadastro de empresas!")
                    Exit Sub
                End If

                Dim objProduto As New [Lib].Negocio.Produto(objClienteXEmpresa.CodigoProdutoDeMDFe)
                Dim objNotaFiscalXItem As New [Lib].Negocio.NotaFiscalXItem(objMDFe)

                With objNotaFiscalXItem
                    .CodigoProduto = objProduto.Codigo
                    '.CodigoPedido = objMDFe.CodigoPedido
                    .PesoQuantidade = objProduto.PesoQuantidade
                    .QuantidadeFiscal = lstNotaFiscal.SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFisica)
                    .QuantidadeFisica = lstNotaFiscal.SelectMany(Function(s) s.Itens).Sum(Function(s) s.QuantidadeFisica)
                    .Unitario = lstNotaFiscal.SelectMany(Function(s) s.Itens).Sum(Function(s) s.Unitario)
                    .NotaFiscal.CarregandoItens = True
                    .ValorTotal = lstNotaFiscal.SelectMany(Function(s) s.Itens).Sum(Function(s) s.ValorTotal)
                End With

                objNotaFiscalXItem.CodigoOperacao = objMDFe.CodigoOperacao
                objNotaFiscalXItem.CodigoSubOperacao = objMDFe.CodigoSubOperacao
                objNotaFiscalXItem.CarregandoEncargos = True
                objNotaFiscalXItem.Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objNotaFiscalXItem, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
                objNotaFiscalXItem.CarregandoEncargos = False
                objMDFe.CarregandoItens = True
                objMDFe.Itens.Clear()
                objMDFe.Itens.Add(objNotaFiscalXItem)
                objMDFe.AtualizaTotais()
                objMDFe.CarregandoItens = False
            End If

            gridNotasFiscais.DataSource = ds
            gridNotasFiscais.DataBind()

            If objMDFe.Itens IsNot Nothing AndAlso objMDFe.Itens.Count > 0 Then
                Dim cfop As New [Lib].Negocio.CFOP(objMDFe.Itens(0).CFOP)
                txtCfop.Text = String.Format("{0} - {1}", cfop.Codigo, cfop.Descricao)
                If (Not IsLoading) AndAlso lstNotaFiscal IsNot Nothing AndAlso lstNotaFiscal.Count > 0 Then
                    txtPesoBruto.Text = lstNotaFiscal.Sum(Function(s) s.PesoBruto).ToString("N2")
                Else
                    Dim objNotaFiscalXEmbalagem As New [Lib].Negocio.NotasFiscaisXEmbalagens(objMDFe)
                    txtPesoBruto.Text = objNotaFiscalXEmbalagem.PesoBruto.ToString("N2")
                End If
                txtValor.Text = objMDFe.Itens.Sum(Function(s) s.ValorTotal).ToString("N2")
            End If

            If Not IsLoading Then
                objMDFe.CodigoFinalidade = objMDFe.NotaTrocaOrigem.CodigoFinalidade
                'objMDFe.CodigoPedido = objMDFe.NotaTrocaOrigem.CodigoPedido
            End If

            If objMDFe.NotaFiscalXPercursos IsNot Nothing AndAlso objMDFe.NotaFiscalXPercursos.Count > 0 Then
                Dim strEstados = String.Empty
                For Each e As [Lib].Negocio.NotaFiscalXPercurso In objMDFe.NotaFiscalXPercursos.OrderBy(Function(s) s.Ordem).ToList()
                    If Not String.IsNullOrWhiteSpace(strEstados) Then
                        strEstados &= ", "
                    End If
                    strEstados &= e.Estado_Id
                Next
                lblEstados.Text = strEstados
            End If

            objMDFe.CarregandoNota = False
            SalvarSessaoMDFe()

            lnkNovo.Parent.Visible = ctrNovo
            lnkExcluir.Parent.Visible = objMDFe IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objMDFe.IUD) AndAlso objMDFe.IUD <> "I" AndAlso objMDFe.CodigoSituacao = CInt(eSituacao.Normal)
            lnkEspelho.Parent.Visible = True
            lnkEncerrar.Parent.Visible = Not String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.ProtocoloNota) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.Retorno) AndAlso objMDFe.Retorno <> "135"
            lnkEnviarSEFAZ.Parent.Visible = Not String.IsNullOrWhiteSpace(objMDFe.IUD) AndAlso objMDFe.IUD <> "I" _
                AndAlso String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) AndAlso String.IsNullOrWhiteSpace(objMDFe.ProtocoloNota) _
                AndAlso (String.IsNullOrWhiteSpace(objMDFe.Retorno) OrElse (Not String.IsNullOrWhiteSpace(objMDFe.Retorno) AndAlso objMDFe.Retorno <> "100" AndAlso objMDFe.Retorno <> "135"))
            lnkDAMDFE.Parent.Visible = Not String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.ProtocoloNota)
            lnkEnviarEmail.Parent.Visible = Not String.IsNullOrWhiteSpace(objMDFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objMDFe.ProtocoloNota)
            lnkVisualizar.Parent.Visible = IsEnabled(objMDFe)

            btnEmpresa.Enabled = Not IsLoading
            btnClienteMDFe.Enabled = Not IsLoading
            btnTransportador.Enabled = Not IsLoading
            txtPesoBruto.Enabled = Not IsLoading
            txtObservacao.Height = Unit.Pixel(200)
            TabContainer1.ActiveTabIndex = 1
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BindGridView()
        grdNFe.DataSource = New List(Of Object)
        grdNFe.DataBind()
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

    Private Sub ImprimirEspelho()
        Try
            RecuperarSessaoMDFe()
            Dim espelho As New [Lib].Negocio.NotaFiscalEspelho
            espelho.ExibirEspelho(Me.Page, objMDFe)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function IsEnabled(ByVal mdfe As [Lib].Negocio.NotaFiscal) As Boolean
        Dim sql As String = "SELECT * FROM NFEPendencias " & vbCrLf & _
                            "WHERE 1=1 " & vbCrLf & _
                            "AND Empresa_Id = '" & mdfe.CodigoEmpresa & "' " & vbCrLf & _
                            "AND EndEmpresa_Id = '" & mdfe.EnderecoEmpresa & "' " & vbCrLf & _
                            "AND Cliente_Id = '" & mdfe.CodigoCliente & "' " & vbCrLf & _
                            "AND EndCliente_Id = '" & mdfe.EnderecoCliente & "' " & vbCrLf & _
                            "AND EntradaSaida_Id = '" & IIf(mdfe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
                            "AND Serie_Id = '" & mdfe.Serie & "' " & vbCrLf & _
                            "AND Nota_Id = '" & mdfe.Codigo & "'"
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NFEPendencias")

        Return ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0
    End Function

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "MDFe")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Info)
        End Try
    End Sub
End Class