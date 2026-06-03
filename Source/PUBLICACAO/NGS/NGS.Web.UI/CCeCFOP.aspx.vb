Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.IO

Public Class CCeCFOP
    Inherits BasePage
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CCeCFOP", "ACESSAR") Then
                CargaUnidadeDeNegocio()
                CargaGrupoCFOP()
                limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaGrupoCFOP()
        ddl.Carregar(ddlGrupoCFOP, CarregarDDL.Tabela.CFOPGrupo, "", True)
    End Sub

    Private Sub CargaCFOP()
        ddl.Carregar(ddlCFOP, CarregarDDL.Tabela.CFOP, "GrupoCfop_Id = " & ddlGrupoCFOP.SelectedValue, True)
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objNFConsultaCCe" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)

        txtCliente.Text = ""
        txtUf.Text = ""
        txtES.Text = ""
        txtNota.Text = ""
        txtSerie.Text = ""
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        ddlGrupoCFOP.SelectedItem.Text = ""

        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True
        txtUf.Enabled = True

        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkReimpressao.Parent.Visible = False

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

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False

        ElseIf String.IsNullOrWhiteSpace(txtNota.Text) AndAlso String.IsNullOrWhiteSpace(txtCliente.Text) Then
            MsgBox(Me.Page, "Informe o cliente ou a nota fiscal.")
            Return False
        End If

        Return True
    End Function

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

                If System.IO.File.Exists("C:\Alfasig\LeituraNFe\-danfes\" & objNotaFiscal.ChaveNFE & "cce.pdf") Then
                    lnkReimpressao.Parent.Visible = True
                Else
                    lnkReimpressao.Parent.Visible = False
                End If

                ddlGrupoCFOP.SelectedValue = objNotaFiscal.CFOP.CodigoGrupo
                CargaCFOP()
                ddlCFOP.SelectedValue = objNotaFiscal.CFOP.Codigo

                txtUf.Enabled = False
                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False

                SalvaNotaFiscal()
                Session.Remove("objNFConsultaCCe" & HID.Value)
                lnkAtualizar.Parent.Visible = True
                lnkConsultar.Parent.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CargaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoCFOP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoCFOP.SelectedIndexChanged
        Try
            CargaCFOP()
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
            If Funcoes.VerificaPermissao("CCeCFOP", "LEITURA") Then
                If ValidaCampos() Then
                    RecuperaNotaFiscal()

                    If ddlEmpresa.SelectedValue.ToString.Length > 0 Then
                        objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.ToString.Split("-")(0)
                        objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.ToString.Split("-")(1)
                    Else
                        MsgBox(Me.Page, "É necessário selecionar a empresa!")
                    End If

                    objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                    objNotaFiscal.Movimento = CDate(txtDataFinal.Text)

                    If txtCodigoCliente.Value.ToString.Length > 0 Then
                        objNotaFiscal.CodigoCliente = txtCodigoCliente.Value.ToString.Split("-")(0)
                        objNotaFiscal.EnderecoCliente = txtCodigoCliente.Value.ToString.Split("-")(1)
                    End If

                    If txtES.Text.Length > 0 Then
                        objNotaFiscal.EntradaSaida = IIf(txtES.Text = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                    End If


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
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Not ddlCFOP.SelectedIndex > 0 Then
                MsgBox(Me.Page, "CFOP não foi selecionado.")
            ElseIf Funcoes.VerificaPermissao("CCeCFOP", "ALTERAR") Then
                RecuperaNotaFiscal()

                Dim msgNFE As String = String.Empty
                Dim Sql As String = String.Empty

                For Each item In objNotaFiscal.Itens
                    item.IUD = "U"
                    item.CFOP = ddlCFOP.SelectedValue
                Next

                If [Lib].Negocio.DocumentoEletronico.CCeCFOP(objNotaFiscal, msgNFE) Then
                    Dim Sqls As New ArrayList

                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-impcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)
                    Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                    Sqls.Add(Sql)

                    objNotaFiscal.Itens.SalvarSql(Sqls)

                    Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                    If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                        obs = obs & ". Carta de correção CFOP em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Carta de correção CFOP em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
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
                        msgNFE = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                        MsgBox(Me.Page, msgNFE)
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
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
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
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CCeCFOP")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class