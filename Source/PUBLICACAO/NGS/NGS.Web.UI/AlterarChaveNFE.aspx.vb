Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Xml
Imports System.Threading
Imports System.IO
Public Class AlterarChaveNFE
    Inherits BasePage

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AlterarChaveNFE", "ACESSAR") Then
                    ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                    limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged() Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CarregarEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If validaConsulta() Then
                If objNotaFiscal Is Nothing Then
                    objNotaFiscal = New [Lib].Negocio.NotaFiscal
                End If

                objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                objNotaFiscal.Movimento = CDate(txtDataFinal.Text)
                objNotaFiscal.NossaEmissao = False
                objNotaFiscal.CodigoSituacao = 1
                objNotaFiscal.CodigoTipoDeDocumento = 0

                If Not String.IsNullOrWhiteSpace(txtCliente.Text) Then
                    objNotaFiscal.CodigoCliente = txtCodigoCliente.Value.Split("-")(0)
                    objNotaFiscal.EnderecoCliente = txtCodigoCliente.Value.Split("-")(1)
                End If

                If txtES.Text.Length > 0 Then objNotaFiscal.EntradaSaida = IIf(txtES.Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)

                If txtNota.Text.Length > 0 Then
                    objNotaFiscal.Codigo = txtNota.Text
                Else
                    objNotaFiscal.Codigo = 0
                End If

                If txtSerie.Text.Length > 0 Then
                    objNotaFiscal.Serie = txtSerie.Text
                Else
                    objNotaFiscal.Serie = ""
                End If
                SalvaNotaFiscal()

                Session("ssCampo" & HID.Value) = "NotaXItens"

                If ucConsultaPedidosXNotas.BindGridView() > 0 Then
                    Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsulta" & HID.Value.ToString)
                Else
                    MsgBox(Me.Page, "Nenhum resultado encontrado.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AlterarChaveNFE", "ALTERAR") Then

                RecuperaNotaFiscal()

                objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")

                If validaAlteracao() Then
                    If objNotaFiscal IsNot Nothing Then

                        ucFile.Salvar(objNotaFiscal.Arquivos)

                        Dim Sqls As New ArrayList
                        Dim strSQL As String

                        Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            If objNotaFiscal.CodigoTipoDeDocumento = 2 Or objNotaFiscal.CodigoTipoDeDocumento = 57 Or objNotaFiscal.CodigoTipoDeDocumento = 58 Then
                                obs = obs & ". Alteração da CHAVE CTE em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                            Else
                                obs = obs & ". Alteração da CHAVE NFE em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                            End If
                        Else
                            If objNotaFiscal.CodigoTipoDeDocumento = 2 Or objNotaFiscal.CodigoTipoDeDocumento = 57 Or objNotaFiscal.CodigoTipoDeDocumento = 58 Then
                                obs = "Alteração da CHAVE CTE em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                            Else
                                obs = "Alteração da CHAVE NFE em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
                            End If
                        End If

                        strSQL = " Update NotasFiscais set" & vbCrLf &
                                 "     Eletronica                  = 'S'" & vbCrLf &
                                 "	  ,UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                                 "	  ,UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                                 "    ,ObservacoesControleInterno  = '" & obs & "'" & vbCrLf &
                                 "	Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                 "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                 "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                 "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                                 "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                 "    and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                                 "    and Nota_Id         = " & objNotaFiscal.Codigo & ";"
                        Sqls.Add(strSQL)

                        strSQL = "Delete NFERealizadas" & vbCrLf &
                              " where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "   and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                              "   and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                              "   and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                              "   and Serie_Id        ='" & objNotaFiscal.Serie & "'" & vbCrLf &
                              "   and Nota_Id         = " & objNotaFiscal.Codigo
                        Sqls.Add(strSQL)

                        strSQL = " Insert into NFERealizadas(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, ChaveNfe, ObservacoesFiscais, DadosAdicionais)" & vbCrLf &
                              " Values('" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                            "," & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                            ",'" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                            "," & objNotaFiscal.EnderecoCliente & vbCrLf &
                                            ",'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                            ",'" & objNotaFiscal.Serie & "'" & vbCrLf &
                                            "," & objNotaFiscal.Codigo & vbCrLf &
                                            ",'" & objNotaFiscal.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                                            ",'" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                                            ",'" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                                            ",'" & objNotaFiscal.ChaveNFE & "'" & vbCrLf &
                                            ",'" & objNotaFiscal.Observacoes & " " & objNotaFiscal.ObservacoesDeEmbarque & "'" & vbCrLf &
                                            ",'" & objNotaFiscal.ObservacoesDoProduto & "');"
                        Sqls.Add(strSQL)

                        If objNotaFiscal.Arquivos IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
                            For Each arq As [Lib].Negocio.Arquivo In objNotaFiscal.Arquivos
                                arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                                arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                                arq.CodigoCliente = objNotaFiscal.CodigoCliente
                                arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                                arq.CodigoNota = objNotaFiscal.Codigo
                                arq.Serie = objNotaFiscal.Serie
                                arq.CodigoPedido = objNotaFiscal.CodigoPedido
                                arq.SalvarSql(Sqls)
                            Next
                        End If

                        If Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                            limpar()
                        Else
                            Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                            MsgBox(Me.Page, Observacao)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro(s).")
                Exit Sub
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AlterarChaveNFE")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkVerificarChaveNFE_Click(sender As Object, e As EventArgs) Handles lnkVerificarChaveNFE.Click
        If String.IsNullOrWhiteSpace(txtChaveNFe.Text.Replace(".", "")) Then
            MsgBox(Me.Page, "Chave da Nota Fiscal não foi informada.")
        ElseIf txtChaveNFe.Text.Replace(".", "").Length <> 44 Then
            txtChaveNFe.Text = ""
            MsgBox(Me.Page, "Chave da nota eletrônica deve ter 44 dígitos.")
        Else
            RecuperaNotaFiscal()

            objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")

            'Verificar se foi feito a recusa da Nota Fiscal
            If objNotaFiscal.TemRecusa Then
                MsgBox(Me.Page, "Nota Fiscal não pode ser ajustada pois a mesma foi lançada como recusada.")
                limpar()
                Exit Sub
            End If

            If NotaJaPreenchida() Then
                'Se alguns campos foram preenchidos então serão validados em relação a chavenfe
                If ValidarPreenchimentoDaNFE() Then
                    If RealizarManifestoNFE() Then
                        lnkAtualizar.Parent.Visible = True
                    Else
                        lnkAtualizar.Parent.Visible = False
                    End If
                End If
            Else
                'Caso não tenham sido preenchido acontece o manifesto e a importação do xml com posterior preenchimento dos dados da nfe de entrada
                If RealizarManifestoNFE() Then
                    Dim DsXml As New DataSet

                    Dim ModeloNFe As String = Mid(objNotaFiscal.ChaveNFE, 21, 2)

                    If ModeloNFe.Equals("55") Then
                        DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE)))

                        If objNotaFiscal.CodigoEmpresa = DsXml.Tables("dest").Rows(0)("CNPJ").ToString() Then
                            lnkAtualizar.Parent.Visible = True
                        Else
                            MsgBox(Me.Page, "Empresa do XML está diferente da informanda na Nota Fiscal.")
                        End If
                    Else
                        DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}-CTe.xml", objNotaFiscal.ChaveNFE)))

                        If objNotaFiscal.CodigoEmpresa = DsXml.Tables("rem").Rows(0)("CNPJ").ToString() Then
                            lnkAtualizar.Parent.Visible = True
                        Else
                            MsgBox(Me.Page, "Empresa do XML está diferente do informando no Conhecimento de Transporte.")
                        End If
                    End If
                Else
                    lnkAtualizar.Parent.Visible = False
                End If
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsulta" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsulta" & HID.Value), NotaFiscal))

                divNFE.Visible = True
                lnkVerificarChaveNFE.Visible = True
                divArquivo.Visible = True

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                txtUf.Text = objNotaFiscal.Cliente.CodigoEstado
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie
                txtChaveNFe.Text = objNotaFiscal.ChaveNFE

                txtChaveNFe.ReadOnly = False
                txtUf.Enabled = False
                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False
                txtDataInicial.Enabled = False
                txtDataFinal.Enabled = False
                ddlEmpresa.Enabled = False
                ddlUnidadeDeNegocio.Enabled = False
                btnCliente.Enabled = False

                SalvaNotaFiscal()

                Session.Remove("objNFConsulta" & HID.Value)

                If ucFile.Parent.Visible Then
                    ucFile.Bind(objNotaFiscal.Arquivos)
                End If


                If objNotaFiscal.NossaEmissao Then
                    MsgBox(Me.Page, "Nota Fiscal NOSSA EMISSÃO não pode ser alterada.", eTitulo.Info)
                    lnkVerificarChaveNFE.Visible = False
                Else
                    If objNotaFiscal.CodigoTipoDeDocumento = 2 Or objNotaFiscal.CodigoTipoDeDocumento = 57 Or objNotaFiscal.CodigoTipoDeDocumento = 58 Then
                        lnkAtualizar.Parent.Visible = True
                    End If
                End If

                lnkConsultar.Parent.Visible = False

                txtChaveNFe.Focus()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub limpar()
        Session.Remove("objNFConsulta" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("ssCampo" & HID.Value)
        Session.Remove("objCliente" & HID.Value)

        txtUf.Enabled = True
        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtDataInicial.Enabled = True
        txtDataFinal.Enabled = True
        ddlEmpresa.Enabled = True
        ddlUnidadeDeNegocio.Enabled = True

        btnCliente.Enabled = True

        txtChaveNFe.Text = String.Empty
        txtChaveNFe.ReadOnly = True
        lnkVerificarChaveNFE.Visible = False
        ucFile.Clear()
        divNFE.Visible = False

        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtUf.Text = String.Empty
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)

        LiberaEmpresa()

        Try
            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            SalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CarregarEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Function validaConsulta() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período para consulta.")
            Return False
        End If
        Return True
    End Function

    Private Function validaAlteracao() As Boolean
        If Not objNotaFiscal.ChaveNFE.Length = 44 Then
            MsgBox(Me.Page, "Chave da nota eletrônica deve ter 44 dígitos.")
            Return False
        ElseIf objNotaFiscal.TemRecusa Then
            MsgBox(Me.Page, "Nota Fiscal não pode ser ajustada pois a mesma foi lançada como recusada.")
            Return False
        ElseIf System.IO.File.Exists(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE))) Then

            Dim DsXml As New DataSet

            'O metodo abaixo da erro de Nome de coluna 'Id' foi definido para diferentes tipos de mapeamento.
            'Isso ocorre quando tem mais de um ID no XML iguais, o codigo LimpaErro, faz a leitura sem o erro
            'DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE)))

            If Funcoes.LimparErroLeituraXML(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", objNotaFiscal.ChaveNFE)), DsXml) Then
                Return True
            End If

            If objNotaFiscal.CodigoEmpresa = DsXml.Tables("dest").Rows(0)("CNPJ").ToString() Then
                    Return True
                Else
                    MsgBox(Me.Page, "Empresa do XML está diferente da informanda na Nota Fiscal.")
                    Return False
                End If
            Else
                Return True
        End If
    End Function

    Private Function NotaJaPreenchida() As Boolean
        If Not String.IsNullOrWhiteSpace(txtNota.Text) AndAlso CInt(txtNota.Text) > 0 _
            AndAlso Not String.IsNullOrWhiteSpace(txtSerie.Text) _
            AndAlso Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function ValidarPreenchimentoDaNFE() As Boolean

        Dim valida As Boolean = True

        If (Mid(objNotaFiscal.ChaveNFE, 7, 14) = "02935843000105" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "01409655000180" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "78393592000146" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "82951310000156" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "87958674000181" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "76416890000189" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "58290502000184" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "03507415000578") Then
            valida = False
        End If

        Dim ModeloNFe As String = Mid(objNotaFiscal.ChaveNFE, 21, 2)

        If String.IsNullOrWhiteSpace(txtNota.Text) Then
            MsgBox(Me.Page, "Número da Nota Fiscal deve ser Informado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtSerie.Text) Then
            MsgBox(Me.Page, "Série da Nota Fiscal deve ser Informada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            MsgBox(Me.Page, "Cliente da Nota Fiscal deve ser Informado.")
            Return False
        ElseIf ModeloNFe.Equals("55") AndAlso (Not objNotaFiscal.CodigoTipoDeDocumento = 1 And Not objNotaFiscal.CodigoTipoDeDocumento = 15) Then
            MsgBox(Me.Page, "Validação só pode ser feita para Documento do Tipo Nota fiscal.")
            Return False
        ElseIf ModeloNFe.Equals("57") AndAlso (Not objNotaFiscal.CodigoTipoDeDocumento = 57 AndAlso Not objNotaFiscal.CodigoTipoDeDocumento = 58) Then
            MsgBox(Me.Page, "Validação só pode ser feita para Documento do Tipo Conhecimento de Transporte.")
            Return False
        ElseIf valida AndAlso objNotaFiscal.CodigoCliente.Length = 11 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 10, 11) = objNotaFiscal.CodigoCliente Then
            MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf valida AndAlso objNotaFiscal.CodigoCliente.Length = 14 AndAlso Not Mid(objNotaFiscal.ChaveNFE, 7, 14) = objNotaFiscal.CodigoCliente Then
            MsgBox(Me.Page, "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf objNotaFiscal.Cliente.Municipio.EstadoIbge = 0 Then
            MsgBox(Me.Page, "o Código IBGE do cliente não foi informado no cadastro! Favor revisar o cadastro do cliente.")
            Return False
        ElseIf valida AndAlso Not CInt(Left(objNotaFiscal.ChaveNFE, 2)) = objNotaFiscal.Cliente.Municipio.EstadoIbge Then
            MsgBox(Me.Page, "Estado do Cliente na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 26, 9)) = CInt(txtNota.Text) Then
            MsgBox(Me.Page, "Número da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf Not CInt(Mid(objNotaFiscal.ChaveNFE, 23, 3)) = CInt(txtSerie.Text) Then
            MsgBox(Me.Page, "Série da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        ElseIf Not Mid(objNotaFiscal.ChaveNFE, 3, 2) = String.Format("{0:yy}", objNotaFiscal.DataNota) Then
            MsgBox(Me.Page, "Ano da Nota na Chave Eletrônica diferente do informado na Nota Fiscal.")
            Return False
        End If
        Return True
    End Function

    Private Function RealizarManifestoNFE() As Boolean
        RecuperaNotaFiscal()

        Dim valida As Boolean = True

        If (Mid(objNotaFiscal.ChaveNFE, 7, 14) = "02935843000105" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "01409655000180" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "78393592000146" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "82951310000156" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "87958674000181" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "58290502000184" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "76416890000189" Or
            Mid(objNotaFiscal.ChaveNFE, 7, 14) = "03507415000578") Then
            valida = False
        End If

        objNotaFiscal.ChaveNFE = txtChaveNFe.Text.Replace(".", "")

        Dim msgResultado As String = String.Empty

        Dim ModeloNFe As String = Mid(objNotaFiscal.ChaveNFE, 21, 2)
        'Realiza o manifesto da NFe
        If objNotaFiscal.Empresa.Empresa.NotaFiscalEletronica AndAlso ucFile.Parent.Visible AndAlso (ModeloNFe.Equals("55") Or ModeloNFe.Equals("57")) Then '(Modelo: 55 NFe, 57 CTe)

            Dim tpExt As String = String.Empty

            If ModeloNFe.Equals("55") Then tpExt = "-nfe.xml"
            If ModeloNFe.Equals("57") Then tpExt = "-CTe.xml"

            'Download do Arquivo.
            Dim bytes As Byte() = New FilesManager().getFileXml(String.Format("{0}{1}", objNotaFiscal.ChaveNFE, tpExt))
            If bytes Is Nothing Then
                If ModeloNFe.Equals("55") Then
                    MsgBox(Me.Page, "XML da NFe não foi encontrado, favor inserir manualmente.")
                Else
                    MsgBox(Me.Page, "XML do CTe não foi encontrado, favor inserir manualmente.")
                End If

                Return False
            Else
                Dim caminhoArquivoFile As String = Server.MapPath(String.Format("~/Files/{0}{1}", objNotaFiscal.ChaveNFE, tpExt))
                If Not File.Exists(caminhoArquivoFile) Then
                    System.IO.File.WriteAllBytes(caminhoArquivoFile, bytes)
                End If
            End If

            If valida AndAlso ModeloNFe.Equals("55") Then
                If Not DocumentoEletronico.ManifestoNFe(objNotaFiscal, eTipoManifesto.ConfirmacaoDaOperacao, msgResultado) Then
                    MsgBox(Me.Page, msgResultado)
                    Return False
                End If
            End If

            If bytes IsNot Nothing Then
                Try
                    If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.ChaveNFE) Then
                        Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}{1}", objNotaFiscal.ChaveNFE, tpExt))
                        If System.IO.File.Exists(caminhoArquivo) Then
                            System.IO.File.Delete(caminhoArquivo)
                        End If

                        System.IO.File.WriteAllBytes(caminhoArquivo, bytes)

                        Dim arquivo As String = objNotaFiscal.ChaveNFE & tpExt

                        If File.Exists(caminhoArquivo) Then
                            Dim sourceDir As String = Server.MapPath("~/Files/")
                            Dim backupDir As String = "C:/Alfasig/LeituraNFe/-download/"

                            If Not File.Exists(backupDir & Arquivo) Then
                                File.Copy(Path.Combine(sourceDir, arquivo), Path.Combine(backupDir, arquivo))
                            End If
                        End If
                    End If

                    Dim temArquivo As Boolean = False

                    If objNotaFiscal.Arquivos IsNot Nothing AndAlso objNotaFiscal.Arquivos.Count > 0 Then
                        For Each arq In objNotaFiscal.Arquivos
                            If arq.Descricao = String.Format("{0}{1}", objNotaFiscal.ChaveNFE, tpExt) Then
                                temArquivo = True
                            End If
                        Next
                    End If

                    If Not temArquivo Then

                        objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                             .IUD = "I",
                             .Codigo = String.Empty,
                             .Descricao = String.Format("{0}{1}", objNotaFiscal.ChaveNFE, tpExt),
                             .Arquivo = bytes})

                    End If


                    ucFile.Bind(objNotaFiscal.Arquivos)

                    SalvaNotaFiscal()

                    Return True

                Catch ex As Exception
                    Throw New Exception(ex.Message)
                    Return False
                End Try
            Else
                MsgBox(Me.Page, "XML não encontrado.")
                'Dim obs As String = Funcoes.EliminarCaracteresEspeciais(msgResultado.ToUpper)
                'If obs.Contains("MANIFESTO NFE JA REALIZADO") Then
                '    Return True
                'Else
                '    Return False
                'End If
            End If
        Else
            MsgBox(Me.Page, "Só é permitido alterar a Chave para Nota Fiscal ou Conhecimento de Transporte.")
            Return False
        End If
    End Function

    Private Sub PreencherNFeXML(ByVal objNotaFiscal As [Lib].Negocio.NotaFiscal, ByVal pNomeArquivo As String, ByVal pOrigem As Boolean)
        Dim DsXml As New DataSet
        DsXml.ReadXml(Server.MapPath(String.Format("~/Files/{0}-nfe.xml", pNomeArquivo)))
        SessaoDsXML = DsXml
        DocumentoEletronico.PreencherNFeComXML(objNotaFiscal, DsXml, False, False, False, "", False, False)
        SalvaNotaFiscal()
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
End Class