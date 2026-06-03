Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.IO

Public Class CCeLocalEmbarque
    Inherits BasePage
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)

        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CCeLocalEmbarque", "ACESSAR") Then
                CargaUnidadeDeNegocio()

                limpar()

                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objCliente" & HID.Value)
            ElseIf Not Session("objLocalDeEmbarqueNovo" & HID.Value) Is Nothing Then
                If txtCodigoLocEmbAtual.Value = CType(Session("objLocalDeEmbarqueNovo" & HID.Value), [Lib].Negocio.Cliente).Codigo Then
                    MsgBox(Me.Page, "Local de Embarque selecionado não pode ser o mesmo do atual", eTitulo.Info)
                Else
                    RecuperaNotaFiscal()

                    objNotaFiscal.CodigoLocalEmbarque = CType(Session("objLocalDeEmbarqueNovo" & HID.Value), [Lib].Negocio.Cliente).Codigo
                    objNotaFiscal.EndLocalEmbarque = CType(Session("objLocalDeEmbarqueNovo" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco

                    Dim itemLocalDeEmbarqueNovo As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.LocalEmbarque)
                    txtLocEmbNovo.Text = itemLocalDeEmbarqueNovo.Text

                    txtUFLocEmbNovo.Text = objNotaFiscal.LocalEmbarque.CodigoEstado

                    Dim strLocal As String = ""
                    Dim strTraco As String = ""

                    If objNotaFiscal.LocalEmbarque.Endereco.Length > 0 Then
                        strLocal &= strTraco & objNotaFiscal.LocalEmbarque.Endereco
                        strTraco = " - "
                    End If

                    If objNotaFiscal.LocalEmbarque.Cidade.Length > 0 Then
                        strLocal &= strTraco & objNotaFiscal.LocalEmbarque.Cidade.ToString
                    End If

                    txtEnderecoLocEmbNovo.Text = Mid(strLocal, 1, 60)

                    Session.Remove("objLocalDeEmbarqueNovo" & HID.Value)

                    SalvaNotaFiscal()

                    lnkAtualizar.Parent.Visible = True
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objLocalDeEmbarqueNovo" & HID.Value)
        Session.Remove("objNFConsultaCCe" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)

        txtUf.Enabled = True
        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtCliente.Text = String.Empty
        txtUf.Text = String.Empty
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkReimpressao.Parent.Visible = False

        txtCodigoLocEmbAtual.Value = String.Empty
        txtLocEmbAtual.Text = String.Empty
        txtUFLocEmbAtual.Text = String.Empty
        txtEnderecoLocEmbAtual.Text = String.Empty

        txtCodigoLocEmbNovo.Value = String.Empty
        txtLocEmbNovo.Text = String.Empty
        txtUFLocEmbNovo.Text = String.Empty
        txtEnderecoLocEmbNovo.Text = String.Empty
        btnLocEmbNovo.Enabled = False

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        Try
            objNotaFiscal = New [Lib].Negocio.NotaFiscal()
            SalvaNotaFiscal()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Public Sub CarregarNotaFiscal()
        Try
            If Not Session("objNFConsultaCCe" & HID.Value) Is Nothing Then
                objNotaFiscal = New [Lib].Negocio.NotaFiscal(CType(Session("objNFConsultaCCe" & HID.Value), NotaFiscal))

                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                txtCliente.Text = itemCliente.Text
                txtUf.Text = objNotaFiscal.Cliente.CodigoEstado
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie

                txtUf.Enabled = False
                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False

                If System.IO.File.Exists("C:\Alfasig\LeituraNFe\-danfes\" & objNotaFiscal.ChaveNFE & "cce.pdf") Then
                    lnkReimpressao.Parent.Visible = True
                Else
                    lnkReimpressao.Parent.Visible = False
                End If

                If objNotaFiscal.CodigoLocalEmbarque IsNot Nothing AndAlso objNotaFiscal.CodigoLocalEmbarque.Length > 0 Then
                    txtCodigoLocEmbAtual.Value = objNotaFiscal.CodigoLocalEmbarque

                    Dim itemEmbarque As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.LocalEmbarque)
                    txtLocEmbAtual.Text = itemEmbarque.Text
                    txtUFLocEmbAtual.Text = objNotaFiscal.LocalEmbarque.CodigoEstado

                    Dim strLocal As String = ""
                    Dim strTraco As String = ""

                    If objNotaFiscal.LocalEmbarque.Endereco.Length > 0 Then
                        strLocal &= strTraco & objNotaFiscal.LocalEmbarque.Endereco
                        strTraco = " - "
                    End If

                    If objNotaFiscal.LocalEmbarque.Cidade.Length > 0 Then
                        strLocal &= strTraco & objNotaFiscal.LocalEmbarque.Cidade.ToString
                    End If

                    txtEnderecoLocEmbAtual.Text = Mid(strLocal, 1, 60)
                End If

                SalvaNotaFiscal()

                Session.Remove("objNFConsultaCCe" & HID.Value)

                lnkConsultar.Parent.Visible = False

                btnLocEmbNovo.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        CargaEmpresa()
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("4,5")
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("CCeLocalEmbarque", "LEITURA") Then
                RecuperaNotaFiscal()
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    MsgBox(Me.Page, "É necessário selecionar a empresa!")
                    Exit Sub
                End If

                objNotaFiscal.CodigoEmpresa = Empresa(0)
                objNotaFiscal.EnderecoEmpresa = Empresa(1)
                objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                objNotaFiscal.Movimento = CDate(txtDataFinal.Text)
                objNotaFiscal.NossaEmissao = True
                objNotaFiscal.CodigoSituacao = 1
                objNotaFiscal.CodigoTipoDeDocumento = 1

                If txtCodigoCliente.Value.ToString.Length > 0 Then
                    Dim Cliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                    objNotaFiscal.CodigoCliente = Cliente(0)
                    objNotaFiscal.EnderecoCliente = Cliente(1)
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

                Session("ssCampo" & HID.Value) = "NotaXItens"

                ucConsultaPedidosXNotas.BindGridView()
                Popup.ConsultaDePedidosXNotas(Me.Page, "objNFConsultaCCe" & HID.Value.ToString)
            Else
                MsgBox(Me.Page, "Usuario sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkReimpressao_Click(sender As Object, e As EventArgs) Handles lnkReimpressao.Click
        Try
            RecuperaNotaFiscal()

            Dim Sqls As New ArrayList
            Dim Observacao As String = String.Empty
            Dim strChave As String = objNotaFiscal.ChaveNFE & "cce"
            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", strChave), eTipoDeDocumento.Nota)

            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strChave) Then
                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", strChave))
                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", strChave))

                Sqls.Clear()

                Dim fs As Stream = New FileStream(Server.MapPath("Files/") & Server.HtmlDecode(String.Format("{0}.pdf", strChave)), FileMode.Open, FileAccess.Read)
                Dim br As New BinaryReader(fs)
                Dim Arqbytes As Byte() = br.ReadBytes(fs.Length)

                objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With { _
                                         .IUD = "I", _
                                         .Codigo = String.Empty, _
                                         .Descricao = String.Format("{0}.pdf", strChave),
                                         .Arquivo = Arqbytes})

                For Each arq As [Lib].Negocio.Arquivo In objNotaFiscal.Arquivos
                    arq.IUD = "I"
                    arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                    arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                    arq.CodigoCliente = objNotaFiscal.CodigoCliente
                    arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                    arq.CodigoNota = objNotaFiscal.Codigo
                    arq.Serie = objNotaFiscal.Serie
                    arq.CodigoPedido = objNotaFiscal.CodigoPedido
                    arq.SalvarSql(Sqls)
                Next

                If Sqls.Count > 0 Then
                    If Banco.GravaBanco(Sqls) Then
                        MsgBox(Me.Page, "Alterado com sucesso e Carta de correção inserida.", eTitulo.Sucess)
                        limpar()
                    Else
                        Observacao = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                        MsgBox(Me.Page, Observacao)
                        Exit Sub
                    End If
                End If
            Else
                MsgBox(Me.Page, "Carta de correção não foi encontrada.", eTitulo.Sucess)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CCeLocalEmbarque")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub btnLocEmbNovo_Click(sender As Object, e As EventArgs) Handles btnLocEmbNovo.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("3,4,5")
            Popup.ConsultaDeClientes(Me.Page, "objLocalDeEmbarqueNovo" & HID.Value.ToString, "txtLocEmbNovo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CCeLocalEmbarque", "ALTERAR") Then

                RecuperaNotaFiscal()

                If String.IsNullOrWhiteSpace(objNotaFiscal.CodigoLocalEmbarque) OrElse objNotaFiscal.CodigoLocalEmbarque.Length = 0 Then
                    MsgBox(Me.Page, "Local de Embarque não foi informado.", eTitulo.Info)
                    Exit Sub
                End If

                'objNotaFiscal.Cliente.CodigoEstado.Equals("EX") AndAlso
                'REMOVI POIS SENDO NOSSA EMISSÃO TEM QUE PODER FAZER

                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
                    If objNotaFiscal.CodigoLocalEmbarque.Length > 0 AndAlso
                        objNotaFiscal.EntradaSaida = eEntradaSaida.Saida AndAlso
                        Not [Enum].Parse(GetType(eClassesOperacoes), objNotaFiscal.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.IMPORTACOES Then

                        Dim msgNFE As String = String.Empty
                        Dim Sql As String = String.Empty
                        Dim Sqls As New ArrayList

                        Dim strLocal As String = ""
                        Dim strTraco As String = ""

                        If objNotaFiscal.LocalEmbarque.Endereco.Length > 0 Then
                            strLocal &= strTraco & objNotaFiscal.LocalEmbarque.Endereco
                            strTraco = " - "
                        End If

                        If objNotaFiscal.LocalEmbarque.Cidade.Length > 0 Then
                            strLocal &= strTraco & objNotaFiscal.LocalEmbarque.Cidade.ToString
                            strTraco = " - "
                        End If

                        If [Lib].Negocio.DocumentoEletronico.CCeLocalEmbarque(objNotaFiscal, msgNFE, strLocal) Then

                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-impcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)
                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)

                            Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                            If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                                obs = obs & ". Carta de correção do Local de Embarque em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                            Else
                                obs = "Carta de correção do Local de Embarque em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                            End If

                            Sql = " Update NotasFiscais set " & vbCrLf &
                                 "      LocalEmbarque               = '" & objNotaFiscal.LocalEmbarque.Codigo & "'," & vbCrLf &
                                 "      EndLocalEmbarque            = " & objNotaFiscal.LocalEmbarque.CodigoEndereco & "," & vbCrLf &
                                 "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                                 "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
                                 "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf &
                                 "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                                 "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                                 "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                                 "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                                 "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                                 "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                                 "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                            Sqls.Add(Sql)

                            If Banco.GravaBanco(Sqls) Then
                                SalvaNotaFiscal()
                            Else
                                Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                                MsgBox(Me.Page, Observacao)
                                Exit Sub
                            End If

                            Thread.Sleep(5000)

                            Dim strChave As String = objNotaFiscal.ChaveNFE & "cce"
                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", strChave), eTipoDeDocumento.Nota)
                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(strChave) Then
                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", strChave))
                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", strChave))

                                Sqls.Clear()

                                Dim fs As Stream = New FileStream(Server.MapPath("Files/") & Server.HtmlDecode(String.Format("{0}.pdf", strChave)), FileMode.Open, FileAccess.Read)
                                Dim br As New BinaryReader(fs)
                                Dim Arqbytes As Byte() = br.ReadBytes(fs.Length)

                                objNotaFiscal.Arquivos.Add(New [Lib].Negocio.Arquivo() With {
                                                         .IUD = "I",
                                                         .Codigo = String.Empty,
                                                         .Descricao = String.Format("{0}.pdf", strChave),
                                                         .Arquivo = Arqbytes})

                                For Each arq As [Lib].Negocio.Arquivo In objNotaFiscal.Arquivos
                                    arq.IUD = "I"
                                    arq.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                                    arq.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                                    arq.CodigoCliente = objNotaFiscal.CodigoCliente
                                    arq.EnderecoCliente = objNotaFiscal.EnderecoCliente
                                    arq.CodigoNota = objNotaFiscal.Codigo
                                    arq.Serie = objNotaFiscal.Serie
                                    arq.CodigoPedido = objNotaFiscal.CodigoPedido
                                    arq.SalvarSql(Sqls)
                                Next

                                If Sqls.Count > 0 Then
                                    If Banco.GravaBanco(Sqls) Then
                                        MsgBox(Me.Page, "Alterado com sucesso e Carta de correção inserida.", eTitulo.Sucess)
                                    Else
                                        Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                                        MsgBox(Me.Page, Observacao)
                                        Exit Sub
                                    End If
                                End If
                            Else
                                MsgBox(Me.Page, "Alterado com sucesso mais não foi possível inserir a Carta de correção.", eTitulo.Sucess)
                            End If

                            limpar()
                        Else
                            MsgBox(Me.Page, "Erro na Carta de Correção: " & msgNFE & ". Tente novamente ou entre em contato com o Suporte.")
                        End If
                    Else
                        AlterarNotaFiscal()
                    End If
                Else
                    AlterarNotaFiscal()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AlterarNotaFiscal()
        Dim Sql As String = String.Empty
        Dim Sqls As New ArrayList
        Dim Observacao As String = String.Empty

        Dim obs As String = objNotaFiscal.ObservacoesControleInterno
        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
            obs = obs & ". Alterado Local de Embarque em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
        Else
            obs = "Alterado Local de Embarque em " & Format(Now, "dd/MM/yyyy HH:mm:ss")
        End If

        Sql = " Update NotasFiscais set " & vbCrLf &
        "      LocalEmbarque               = '" & objNotaFiscal.LocalEmbarque.Codigo & "'," & vbCrLf &
        "      EndLocalEmbarque            = " & objNotaFiscal.LocalEmbarque.CodigoEndereco & "," & vbCrLf &
        "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
        "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
        "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf &
        "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
        "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
        "	 and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
        "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
        "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
        "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
        "    and Serie_Id        ='" & objNotaFiscal.Serie & "';"

        Sqls.Add(Sql)

        If Banco.GravaBanco(Sqls) Then
            MsgBox(Me.Page, "Local de Embarque Alterado com sucesso.", eTitulo.Sucess)
            limpar()
        Else
            Observacao = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
            MsgBox(Me.Page, Observacao)
        End If
    End Sub

End Class