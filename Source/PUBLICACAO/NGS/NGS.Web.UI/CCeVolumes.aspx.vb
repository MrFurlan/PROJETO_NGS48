Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.IO

Public Class CCeVolumes
    Inherits BasePage
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CCeVolumes", "ACESSAR") Then
                CargaUnidadeDeNegocio()
                limpar()

                Dim fm As New FilesManager()
                If Not fm.IsConnect() Then
                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor.")
                End If
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
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objNFConsultaCCe" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)

        txtUf.Enabled = True
        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtCliente.Text = ""
        txtUf.Text = ""
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkReimpressao.Parent.Visible = False

        chkVolume.Checked = False
        chkEspecie.Checked = False
        chkMarca.Checked = False
        chkNumeracao.Checked = False
        chkBruto.Checked = False
        chkLiquido.Checked = False

        txtVolumes.Text = ""
        txtEspecie.Text = ""
        txtMarca.Text = ""
        txtNumeracao.Text = ""
        txtPesoBruto.Text = ""
        txtPesoLiquido.Text = ""

        txtVolumes.Enabled = False
        txtEspecie.Enabled = False
        txtMarca.Enabled = False
        txtNumeracao.Enabled = False
        txtPesoBruto.Enabled = False
        txtPesoLiquido.Enabled = False

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

                If objNotaFiscal.CodigoSituacao = 1 OrElse objNotaFiscal.CodigoSituacao = 4 Then
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

                    txtVolumes.Text = objNotaFiscal.Quantidade
                    txtEspecie.Text = objNotaFiscal.Especie
                    txtMarca.Text = objNotaFiscal.Marca
                    txtNumeracao.Text = objNotaFiscal.Numero
                    txtPesoBruto.Text = objNotaFiscal.PesoBruto
                    txtPesoLiquido.Text = objNotaFiscal.PesoLiquido

                    SalvaNotaFiscal()

                    Session.Remove("objNFConsultaCCe" & HID.Value)
                    lnkAtualizar.Parent.Visible = True
                    lnkConsultar.Parent.Visible = False
                Else
                    MsgBox(Me.Page, "Nota Fiscal não pode ser alterado por aqui.")
                    limpar()
                End If
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
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnObservacoesFiscais_Click(ByVal sender As Object, ByVal e As System.EventArgs)

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

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("CCeVolumes", "LEITURA") Then
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
                ''objNotaFiscal.CodigoSituacao = 1
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

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        If Not chkVolume.Checked AndAlso _
           Not chkEspecie.Checked AndAlso _
           Not chkMarca.Checked AndAlso _
           Not chkNumeracao.Checked AndAlso _
           Not chkBruto.Checked AndAlso _
           Not chkLiquido.Checked Then
            MsgBox(Me.Page, "Nenhum campo foi selecionado para alteração.")
        Else
            Atualizar()
        End If
    End Sub

    Private Sub Atualizar()
        RecuperaNotaFiscal()

        Try
            If Funcoes.VerificaPermissao("CCeVolumes", "ALTERAR") Then
                Dim msgNFE As String = String.Empty
                Dim Sql As String = String.Empty
                Dim Sqls As New ArrayList

                If chkVolume.Checked Then objNotaFiscal.Quantidade = txtVolumes.Text
                If chkEspecie.Checked Then objNotaFiscal.Especie = txtEspecie.Text
                If chkMarca.Checked Then objNotaFiscal.Marca = txtMarca.Text
                If chkNumeracao.Checked Then objNotaFiscal.Numero = txtNumeracao.Text
                If chkBruto.Checked Then objNotaFiscal.PesoBruto = txtPesoBruto.Text
                If chkLiquido.Checked Then objNotaFiscal.PesoLiquido = txtPesoLiquido.Text

                If objNotaFiscal.CodigoSituacao = 4 Then
                    If chkVolume.Checked OrElse chkEspecie.Checked OrElse chkMarca.Checked OrElse chkNumeracao.Checked OrElse chkBruto.Checked OrElse chkLiquido.Checked Then
                        Sql = " Update NotasXEmbalagens set " & vbCrLf & _
                                 "      Quantidade            = " & Str(objNotaFiscal.Quantidade) & "," & vbCrLf & _
                                 "      Especie               = '" & objNotaFiscal.Especie & "'," & vbCrLf & _
                                 "      Marca                 = '" & objNotaFiscal.Marca & "'," & vbCrLf & _
                                 "      Numero                = '" & objNotaFiscal.Numero & "'," & vbCrLf & _
                                 "      PesoBruto             = " & Str(objNotaFiscal.PesoBruto) & "," & vbCrLf & _
                                 "      PesoLiquido           = " & Str(objNotaFiscal.PesoLiquido) & vbCrLf & _
                                 "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                 "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                 "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                                 "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                                 "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                 "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                                 "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                        Sqls.Add(Sql)

                        Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            obs = obs & ". Alterado Volumes para Homologação em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        Else
                            obs = "Alterado Volumes para Homologação em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        End If

                        Sql = " Update NotasFiscais set " & vbCrLf & _
                             "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                             "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                             "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                             "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                             "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                             "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                             "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                             "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                             "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                             "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                        Sqls.Add(Sql)

                        If Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, "Alterado com sucesso.", eTitulo.Sucess)
                            limpar()
                        Else
                            Dim Observacao As String = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                            MsgBox(Me.Page, Observacao)
                            Exit Sub
                        End If
                    Else
                        MsgBox(Me.Page, "Não foi checado nada para alteração.", eTitulo.Info)
                    End If
                Else
                    If [Lib].Negocio.DocumentoEletronico.CCeVolumes(objNotaFiscal, msgNFE, chkVolume.Checked, chkEspecie.Checked, chkMarca.Checked, chkNumeracao.Checked, chkBruto.Checked, chkLiquido.Checked) Then

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-impcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)

                        Sql = " Update NotasXEmbalagens set " & vbCrLf & _
                             "      Quantidade            = " & Str(objNotaFiscal.Quantidade) & "," & vbCrLf & _
                             "      Especie               = '" & objNotaFiscal.Especie & "'," & vbCrLf & _
                             "      Marca                 = '" & objNotaFiscal.Marca & "'," & vbCrLf & _
                             "      Numero                = '" & objNotaFiscal.Numero & "'," & vbCrLf & _
                             "      PesoBruto             = " & Str(objNotaFiscal.PesoBruto) & "," & vbCrLf & _
                             "      PesoLiquido           = " & Str(objNotaFiscal.PesoLiquido) & vbCrLf & _
                             "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                             "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                             "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                             "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                             "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                             "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                             "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                        Sqls.Add(Sql)

                        Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            obs = obs & ". Carta de correção Volumes em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        Else
                            obs = "Carta de correção Volumes em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        End If

                        Sql = " Update NotasFiscais set " & vbCrLf & _
                             "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf & _
                             "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf & _
                             "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf & _
                             "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                             "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                             "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                             "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                             "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                             "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
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
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
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
            Funcoes.Ajuda(Me.Page, "CCeVolumes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub chkVolume_CheckedChanged(sender As Object, e As EventArgs) Handles chkVolume.CheckedChanged
        RecuperaNotaFiscal()

        If chkVolume.Checked Then
            txtVolumes.Enabled = True
        Else
            txtVolumes.Text = objNotaFiscal.Quantidade
            txtVolumes.Enabled = False
        End If
    End Sub

    Protected Sub chkEspecie_CheckedChanged(sender As Object, e As EventArgs) Handles chkEspecie.CheckedChanged
        RecuperaNotaFiscal()

        If chkEspecie.Checked Then
            txtEspecie.Enabled = True
        Else
            txtEspecie.Text = objNotaFiscal.Especie
            txtEspecie.Enabled = False
        End If
    End Sub

    Protected Sub chkMarca_CheckedChanged(sender As Object, e As EventArgs) Handles chkMarca.CheckedChanged
        RecuperaNotaFiscal()

        If chkMarca.Checked Then
            txtMarca.Enabled = True
        Else
            txtMarca.Text = objNotaFiscal.Marca
            txtMarca.Enabled = False
        End If
    End Sub

    Protected Sub chkNumeracao_CheckedChanged(sender As Object, e As EventArgs) Handles chkNumeracao.CheckedChanged
        RecuperaNotaFiscal()

        If chkNumeracao.Checked Then
            txtNumeracao.Enabled = True
        Else
            txtNumeracao.Text = objNotaFiscal.Numero
            txtNumeracao.Enabled = False
        End If
    End Sub

    Protected Sub chkBruto_CheckedChanged(sender As Object, e As EventArgs) Handles chkBruto.CheckedChanged
        RecuperaNotaFiscal()

        If chkBruto.Checked Then
            txtPesoBruto.Enabled = True
        Else
            txtPesoBruto.Text = objNotaFiscal.PesoBruto
            txtPesoBruto.Enabled = False
        End If
    End Sub

    Protected Sub chkLiquido_CheckedChanged(sender As Object, e As EventArgs) Handles chkLiquido.CheckedChanged
        RecuperaNotaFiscal()

        If chkLiquido.Checked Then
            txtPesoLiquido.Enabled = True
        Else
            txtPesoLiquido.Text = objNotaFiscal.PesoLiquido
            txtPesoLiquido.Enabled = False
        End If
    End Sub
End Class