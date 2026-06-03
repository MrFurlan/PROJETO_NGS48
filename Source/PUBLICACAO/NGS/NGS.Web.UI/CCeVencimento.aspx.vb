Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Threading
Imports System.IO

Public Class CCeVencimento
    Inherits BasePage
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CCeVencimento", "ACESSAR") Then
                ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                LimparCampos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub LimparCampos()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objNFConsultaCCe" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)

        txtCliente.Text = String.Empty
        txtUf.Text = String.Empty
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtValorNota.Text = String.Empty
        txtDataInicial.Text = Today.ToString("dd/MM/yyyy")
        txtDataFinal.Text = Today.ToString("dd/MM/yyyy")

        gridCCeVencimento.DataBind()

        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True
        txtUf.Enabled = True
        txtValorNota.Enabled = True

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

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtNota.Text) AndAlso String.IsNullOrWhiteSpace(txtCliente.Text) Then
            MsgBox(Me.Page, "Informe o cliente ou a nota fiscal.")
            Return False
        End If

        If gridCCeVencimento.Rows.Count > 0 Then
            RecuperaNotaFiscal()
            Dim lst As New [Lib].Negocio.ListTitulo(objNotaFiscal)
            Dim valor As Decimal
            For Each row As GridViewRow In gridCCeVencimento.Rows
                If CDate(CType(row.FindControl("txtProrrogacao"), TextBox).Text) < objNotaFiscal.Movimento Then
                    MsgBox(Me.Page, "Data de prorrogação não pode ser menor do que a data atual do vencimento da nota.")
                    Return False
                End If
                valor += CDec(CType(row.FindControl("txtValorDoDocumento"), TextBox).Text)
            Next

            If valor <> CDec(txtValorNota.Text) Then
                MsgBox(Me.Page, "O(s) valor(es) não corresponde(m) com o valor da nota.")
                Return False
            End If
        End If
        Return True
    End Function

    Private Function contemAtualizaveis() As Boolean
        For Each row As GridViewRow In gridCCeVencimento.Rows
            If CType(row.FindControl("txtProrrogacao"), TextBox).Enabled OrElse CType(row.FindControl("txtValorDoDocumento"), TextBox).Enabled Then
                Return True
            End If
        Next
        Return False
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
                txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                txtNota.Text = objNotaFiscal.Codigo
                txtSerie.Text = objNotaFiscal.Serie
                txtValorNota.Text = objNotaFiscal.TotalNota

                txtES.Enabled = False
                txtNota.Enabled = False
                txtSerie.Enabled = False
                txtValorNota.Enabled = False

                If System.IO.File.Exists("C:\Alfasig\LeituraNFe\-danfes\" & objNotaFiscal.ChaveNFE & "cce.pdf") Then
                    lnkReimpressao.Parent.Visible = True
                Else
                    lnkReimpressao.Parent.Visible = False
                End If

                
                gridCCeVencimento.DataSource = objNotaFiscal.VencimentosNota
                gridCCeVencimento.DataBind()

                If gridCCeVencimento.Rows.Count > 1 Then
                    For Each row As GridViewRow In gridCCeVencimento.Rows
                        CType(row.FindControl("txtValorDoDocumento"), TextBox).Enabled = True
                    Next
                End If

                SalvaNotaFiscal()
                Session.Remove("objNFConsultaCCe" & HID.Value)
                lnkAtualizar.Parent.Visible = contemAtualizaveis()
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

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CargaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CCeVencimento", "ALTERAR") Then
                If ValidaCampos() Then
                    RecuperaNotaFiscal()

                    Dim msgNFE As String = String.Empty
                    Dim Sql As String = String.Empty
                    Dim Sqls As New ArrayList

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
                        obs = obs & ". Carta de correção Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                    Else
                        obs = "Carta de correção Produto em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
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

                    LimparCampos()
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
                        LimparCampos()
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
            If Funcoes.VerificaPermissao("CCeVencimento", "LEITURA") Then
                If ValidaCampos() Then
                    RecuperaNotaFiscal()

                    objNotaFiscal.CodigoEmpresa = ddlEmpresa.SelectedValue.ToString.Split("-")(0)
                    objNotaFiscal.EnderecoEmpresa = ddlEmpresa.SelectedValue.ToString.Split("-")(1)
                    objNotaFiscal.DataNota = CDate(txtDataInicial.Text)
                    objNotaFiscal.Movimento = CDate(txtDataFinal.Text)
                    'objNotaFiscal.NossaEmissao = True
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
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissão para consultar registro.")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridCCeVencimento_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridCCeVencimento.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If CType(e.Row.FindControl("lblProvisao"), Label).Text = "BAIXA" OrElse e.Row.Cells(5).Text.Split("-")(0) <> 1 Then
                CType(e.Row.FindControl("txtProrrogacao"), TextBox).Enabled = False
                CType(e.Row.FindControl("txtValorDoDocumento"), TextBox).Enabled = False
            End If
        End If
    End Sub
End Class