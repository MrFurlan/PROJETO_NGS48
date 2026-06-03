Imports NGS.Lib.Negocio
Imports System.Threading
Imports System.IO

Public Class CCeTransporte
    Inherits BasePage
    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

#Region "Property"

    Public Property index() As Integer
        Get
            Return Session("index")
        End Get
        Set(ByVal value As Integer)
            Session("index") = value
        End Set
    End Property

#End Region

#Region "Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CCeTransporte", "ACESSAR") Then
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

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("4,5")
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnTransportador_Click(sender As Object, e As EventArgs) Handles btnTransportador.Click
        Try
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTipoCliente("7")
            Popup.ConsultaDeClientes(Me.Page, "objTransportador" & HID.Value.ToString, "txtTransportador")
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

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("CCeTransporte", "LEITURA") Then
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
                'objNotaFiscal.NossaEmissao = True
                'objNotaFiscal.CodigoSituacao = 1
                'objNotaFiscal.CodigoTipoDeDocumento = 1

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

    Protected Sub imgPlaca_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgPlaca.Click
        If txtTransportador.Text.Length > 0 Then
            ucConsultaPlacas.Limpar()
            Popup.ConsultaDePlacas(Me.Page, "objPlaca" & HID.Value)
            Popup.CenterDialog(Me.Page, "divConsultaPlacas")
        Else
            MsgBox(Me.Page, "Transportador não foi selecionado")
        End If
    End Sub

    Protected Sub imgEstadoVeiculo_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgEstadoVeiculo.Click
        Try
            If String.IsNullOrWhiteSpace(txtVeiculo.Text) Then
                MsgBox(Me.Page, "Placa do veículo não foi informada")
            Else
                ucConsultaEstados.Limpar()
                Popup.ConsultaDeEstados(Me.Page, "objEstadoVeiculo" & HID.Value.ToString)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgEstadoReboque1_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgEstadoReboque1.Click
        Try
            If String.IsNullOrWhiteSpace(txtReboque1.Text) Then
                MsgBox(Me.Page, "Reboque 1 não foi informado")
            Else
                ucConsultaEstados.Limpar()
                Popup.ConsultaDeEstados(Me.Page, "objEstado1" & HID.Value.ToString)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgEstadoReboque2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgEstadoReboque2.Click
        Try
            If String.IsNullOrWhiteSpace(txtReboque2.Text) Then
                MsgBox(Me.Page, "Reboque 2 não foi informado")
            Else
                ucConsultaEstados.Limpar()
                Popup.ConsultaDeEstados(Me.Page, "objEstado2" & HID.Value.ToString)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgEstadoReboque3_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgEstadoReboque3.Click
        Try
            If String.IsNullOrWhiteSpace(txtReboque3.Text) Then
                MsgBox(Me.Page, "Reboque 3 não foi informado")
            Else
                ucConsultaEstados.Limpar()
                Popup.ConsultaDeEstados(Me.Page, "objEstado3" & HID.Value.ToString)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CCeTransporte", "ALTERAR") Then
                If ValidaCampos() Then
                    Dim msgNFE As String = String.Empty
                    Dim Sqls As New ArrayList
                    Dim Sql As String = String.Empty
                    Dim Observacao As String = String.Empty

                    RecuperaNotaFiscal()

                    If (objNotaFiscal.CodigoSituacao = 1 AndAlso objNotaFiscal.NossaEmissao = False) OrElse ((objNotaFiscal.CodigoTipoDeDocumento = 1 OrElse objNotaFiscal.CodigoTipoDeDocumento = 12) AndAlso objNotaFiscal.CodigoSituacao = 4) Then

                        Dim objPlaca = New [Lib].Negocio.Placa(Trim(txtVeiculo.Text.ToUpper))

                        Sql = " Merge NotasFiscaisXTransportadores as nXt" & vbCrLf &
                              " USING (Select '" & objNotaFiscal.CodigoEmpresa & "' as Empresa_Id," & objNotaFiscal.EnderecoEmpresa & " as EndEmpresa_Id,'" & objNotaFiscal.CodigoCliente & "' as Cliente_Id," & objNotaFiscal.EnderecoCliente & " as EndCliente_Id,'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "' as EntradaSaida_Id,'" & objNotaFiscal.Serie & "' as Serie_Id," & objNotaFiscal.Codigo & " as Nota_Id) AS nXtOri" & vbCrLf &
                              "    ON nXt.Empresa_Id      = nXtOri.Empresa_Id" & vbCrLf &
                              "   and nXt.EndEmpresa_Id   = nXtOri.EndEmpresa_Id" & vbCrLf &
                              "   and nXt.Cliente_Id      = nXtOri.Cliente_Id" & vbCrLf &
                              "   and nXt.EndCliente_Id   = nXtOri.EndCliente_Id" & vbCrLf &
                              "   and nXt.EntradaSaida_Id = nXtOri.EntradaSaida_Id" & vbCrLf &
                              "   and nXt.Serie_Id        = nXtOri.Serie_Id" & vbCrLf &
                              "   and nXt.Nota_Id         = nXtOri.Nota_Id" & vbCrLf &
                              "  WHEN NOT MATCHED" & vbCrLf &
                              "    THEN Insert (Empresa_Id,EndEmpresa_Id,Cliente_Id,EndCliente_Id,EntradaSaida_Id,Serie_Id,Nota_Id,Proprietario,EndProprietario,Motorista,EndMotorista,Placa)" & vbCrLf &
                              " VALUES ('" & objNotaFiscal.CodigoEmpresa & "', " & objNotaFiscal.EnderecoEmpresa & ",'" & objNotaFiscal.CodigoCliente & "'," & objNotaFiscal.EnderecoCliente & vbCrLf &
                              ",'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "','" & objNotaFiscal.Serie & "'," & objNotaFiscal.Codigo & vbCrLf &
                              ",'" & objNotaFiscal.CodigoTransportador & "'," & objNotaFiscal.EnderecoTransportador & ",'" & objPlaca.CpfMotorista & "'," & objPlaca.EndCpfMotorista & ",'" & Trim(txtVeiculo.Text.ToUpper) & "')" & vbCrLf

                        Sql &= " WHEN MATCHED " & vbCrLf &
                               "   THEN Update set " & vbCrLf &
                               "      Proprietario            = '" & objNotaFiscal.CodigoTransportador & "'," & vbCrLf &
                               "      EndProprietario         = " & objNotaFiscal.EnderecoTransportador & "," & vbCrLf &
                               "      Motorista               = '" & objPlaca.CpfMotorista & "'," & vbCrLf &
                               "      EndMotorista            = " & objPlaca.EndCpfMotorista & "," & vbCrLf &
                               "      Placa                   = '" & Trim(txtVeiculo.Text.ToUpper) & "';"

                        Sqls.Add(Sql)

                        Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                        If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                            obs = obs & ". Alterado Transporte para Homologação em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        Else
                            obs = "Alterado Transporte para Homologação em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                        End If

                        Sql = " Update NotasFiscais set " & vbCrLf &
                              "      UsuarioAlteracao            = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                              "      UsuarioAlteracaoData        = '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
                              "      ObservacoesControleInterno  = '" & obs & "'" & vbCrLf &
                              "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                              "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf &
                              "	   and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                              "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf &
                              "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                              "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf &
                              "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                        Sqls.Add(Sql)

                        If Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, "Alterado com sucesso.", eTitulo.Sucess)
                            limpar()
                        Else
                            Observacao = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                            MsgBox(Me.Page, Observacao)
                            Exit Sub
                        End If

                    ElseIf objNotaFiscal.CodigoTipoDeDocumento = 1 AndAlso objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.CodigoSituacao = 1 Then

                        Dim trasportador As String = objNotaFiscal.CodigoTransportador & "-" & objNotaFiscal.EnderecoTransportador

                        If [Lib].Negocio.DocumentoEletronico.CCeTransportador(objNotaFiscal, msgNFE, chkTransportador.Checked, trasportador, _
                                                                              chkVeiculo.Checked, txtVeiculo.Text, chkEstadoVeiculo.Checked, lblEstadoVeiculo.Text, _
                                                                              chkReboque1.Checked, txtReboque1.Text, chkEstadoReboque1.Checked, lblEstadoReboque1.Text, _
                                                                              chkReboque2.Checked, txtReboque2.Text, chkEstadoReboque2.Checked, lblEstadoReboque2.Text, _
                                                                              chkReboque3.Checked, txtReboque3.Text, chkEstadoReboque3.Checked, lblEstadoReboque3.Text) Then

                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-impcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)
                            Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcarta{0:000000000}#{1}.txt", objNotaFiscal.Codigo, objNotaFiscal.CodigoEmpresa) & "';"
                            Sqls.Add(Sql)

                            Dim objPlaca = New [Lib].Negocio.Placa(Trim(txtVeiculo.Text.ToUpper))

                            Dim obs As String = objNotaFiscal.ObservacoesControleInterno
                            If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                                obs = obs & ". Carta de correção Transporte em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                            Else
                                obs = "Carta de correção Transporte em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                            End If

                            If chkTransportador.Checked OrElse chkVeiculo.Checked Then
                                If String.IsNullOrEmpty(txtCodigoTransportador.Value) Then
                                    Sql = "Insert Into NotasFiscaisXTransportadores (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Proprietario, EndProprietario, Motorista, EndMotorista, Placa)" & vbCrLf & _
                                          "Values('" & objNotaFiscal.CodigoEmpresa & "'," & objNotaFiscal.EnderecoEmpresa & ",'" & objNotaFiscal.CodigoCliente & "'," & _
                                          objNotaFiscal.EnderecoCliente & ",'" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "','" & objNotaFiscal.Serie & "'," & objNotaFiscal.Codigo & _
                                          ",'" & objNotaFiscal.CodigoTransportador & "'," & objNotaFiscal.EnderecoTransportador & ",'" & objPlaca.CpfMotorista & "'," & objPlaca.EndCpfMotorista & ",'" & Trim(txtVeiculo.Text.ToUpper) & "')"
                                Else
                                    Sql = " Update NotasFiscaisXTransportadores set " & vbCrLf & _
                                          "      Proprietario            = '" & objNotaFiscal.CodigoTransportador & "'," & vbCrLf & _
                                          "      EndProprietario         = " & objNotaFiscal.EnderecoTransportador & "," & vbCrLf & _
                                          "      Motorista               = '" & objPlaca.CpfMotorista & "'," & vbCrLf & _
                                          "      EndMotorista            = " & objPlaca.EndCpfMotorista & "," & vbCrLf & _
                                          "      Placa                   = '" & Trim(txtVeiculo.Text.ToUpper) & "'" & vbCrLf & _
                                          "  Where Empresa_Id      ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                          "    and EndEmpresa_Id   = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                                          "	  and Cliente_Id      ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf & _
                                          "    and EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf & _
                                          "    and EntradaSaida_Id ='" & objNotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                                          "    and Nota_Id         = " & objNotaFiscal.Codigo & vbCrLf & _
                                          "    and Serie_Id        ='" & objNotaFiscal.Serie & "'"
                                End If

                                Sqls.Add(Sql)
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

                            If chkEstadoVeiculo.Checked Then
                                Dim Estado() As String = lblEstadoVeiculo.Text.Split("-")
                                Sql = " Update Placas set " & vbCrLf & _
                                      "      CidadePlaca = '" & Trim(lblCidadeVeiculo.Text) & "'," & vbCrLf & _
                                      "      EstadoPlaca = '" & Estado(0) & "'" & vbCrLf & _
                                      "  Where Placa_id      ='" & Trim(txtVeiculo.Text.ToUpper) & "'"
                                Sqls.Add(Sql)
                            End If

                            If chkEstadoReboque1.Checked Then
                                Dim Estado() As String = lblEstadoReboque1.Text.Split("-")
                                Sql = " Update Placas set " & vbCrLf & _
                                      "      CidadePlaca01 = '" & Trim(lblCidadeReboque1.Text) & "'," & vbCrLf & _
                                      "      EstadoPlaca01 = '" & Estado(0) & "'" & vbCrLf & _
                                      "  Where Placa_id      ='" & Trim(txtVeiculo.Text.ToUpper) & "'"
                                Sqls.Add(Sql)
                            End If

                            If chkEstadoReboque2.Checked Then
                                Dim Estado() As String = lblEstadoReboque2.Text.Split("-")
                                Sql = " Update Placas set " & vbCrLf & _
                                      "      CidadePlaca02 = '" & Trim(lblCidadeReboque2.Text) & "'," & vbCrLf & _
                                      "      EstadoPlaca02 = '" & Estado(0) & "'" & vbCrLf & _
                                      "  Where Placa_id      ='" & Trim(txtVeiculo.Text.ToUpper) & "'"
                                Sqls.Add(Sql)
                            End If

                            If chkEstadoReboque3.Checked Then
                                Dim Estado() As String = lblEstadoReboque3.Text.Split("-")
                                Sql = " Update Placas set " & vbCrLf & _
                                      "      CidadePlaca03 = '" & Trim(lblCidadeReboque3.Text) & "'," & vbCrLf & _
                                      "      EstadoPlaca03 = '" & Estado(0) & "'" & vbCrLf & _
                                      "  Where Placa_id      ='" & Trim(txtVeiculo.Text.ToUpper) & "'"
                                Sqls.Add(Sql)
                            End If

                            If Not Banco.GravaBanco(Sqls) Then
                                Observacao = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
                                MsgBox(Me.Page, Observacao)
                                Exit Sub
                            End If

                            Thread.Sleep(10000)

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
                                        Observacao = Funcoes.EliminarCaracteresEspeciais(HttpContext.Current.Session("ssMessage"))
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
                    MsgBox(Me.Page, "Não foi(ram) selecionado(s) item(s).")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CCeTransporte")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Methods"
    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            RecuperaNotaFiscal()
            If Not Session("objCliente" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objCliente" & HID.Value)
            ElseIf Session("objPlaca" & HID.Value) IsNot Nothing Then
                txtVeiculo.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa01
                lblEstadoVeiculo.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).EstadoPlaca01Detalhes.Codigo & "-" & CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).EstadoPlaca01Detalhes.Descricao
                lblCidadeVeiculo.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca01
                chkVeiculo.Visible = True
                chkVeiculo.Enabled = True
                If objNotaFiscal.PlacaDetalhes IsNot Nothing AndAlso Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa01 = objNotaFiscal.PlacaDetalhes.CidadePlaca01 Then chkVeiculo.Checked = True

                chkEstadoVeiculo.Visible = True
                imgEstadoVeiculo.Visible = True

                If String.IsNullOrEmpty(txtCodigoPlaca.Value) Then
                    chkVeiculo.Checked = True
                    chkEstadoVeiculo.Checked = True
                End If

                If objNotaFiscal.PlacaDetalhes IsNot Nothing AndAlso Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).EstadoPlaca01 = objNotaFiscal.PlacaDetalhes.EstadoPlaca01 Then chkEstadoVeiculo.Checked = True
                If objNotaFiscal.PlacaDetalhes IsNot Nothing AndAlso Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca01 = objNotaFiscal.PlacaDetalhes.CidadePlaca01 Then chkEstadoVeiculo.Checked = True

                If String.IsNullOrWhiteSpace(CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa02) Then
                    txtReboque1.Text = ""
                    lblEstadoReboque1.Text = ""
                    lblCidadeReboque1.Text = ""
                Else

                    txtReboque1.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa02
                    lblEstadoReboque1.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Estado02Detalhes.Codigo & "-" & CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Estado02Detalhes.Descricao
                    lblCidadeReboque1.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca02
                    chkReboque1.Visible = True
                    chkReboque1.Enabled = True
                    chkEstadoReboque1.Visible = True
                    imgEstadoReboque1.Visible = True

                    If String.IsNullOrEmpty(txtCodigoPlaca.Value) Then
                        chkReboque1.Checked = True
                        chkEstadoReboque1.Checked = True
                    End If

                    If objNotaFiscal.PlacaDetalhes IsNot Nothing Then
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa02 = objNotaFiscal.PlacaDetalhes.CidadePlaca02 Then chkReboque1.Checked = True
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).EstadoPlaca02 = objNotaFiscal.PlacaDetalhes.EstadoPlaca02 Then chkEstadoReboque1.Checked = True
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca02 = objNotaFiscal.PlacaDetalhes.CidadePlaca02 Then chkEstadoReboque1.Checked = True
                    End If
                End If

                If String.IsNullOrWhiteSpace(CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa03) Then
                    txtReboque2.Text = ""
                    lblEstadoReboque2.Text = ""
                    lblCidadeReboque2.Text = ""
                Else
                    txtReboque2.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa03
                    lblEstadoReboque2.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Estado03Detalhes.Codigo & "-" & CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Estado03Detalhes.Descricao
                    lblCidadeReboque2.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca03
                    chkReboque2.Visible = True
                    chkReboque2.Enabled = True
                    chkEstadoReboque2.Visible = True
                    imgEstadoReboque2.Visible = True

                    If String.IsNullOrEmpty(txtCodigoPlaca.Value) Then
                        chkReboque2.Checked = True
                        chkEstadoReboque2.Checked = True
                    End If

                    If objNotaFiscal.PlacaDetalhes IsNot Nothing Then
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa03 = objNotaFiscal.PlacaDetalhes.CidadePlaca03 Then chkReboque2.Checked = True
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).EstadoPlaca03 = objNotaFiscal.PlacaDetalhes.EstadoPlaca03 Then chkEstadoReboque2.Checked = True
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca03 = objNotaFiscal.PlacaDetalhes.CidadePlaca03 Then chkEstadoReboque2.Checked = True
                    End If

                End If

                If String.IsNullOrWhiteSpace(CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa04) Then
                    txtReboque3.Text = ""
                    lblEstadoReboque3.Text = ""
                    lblCidadeReboque3.Text = ""
                Else
                    txtReboque3.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa04
                    lblEstadoReboque3.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Estado04Detalhes.Codigo & "-" & CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Estado04Detalhes.Descricao
                    lblCidadeReboque3.Text = CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca04
                    chkReboque3.Visible = True
                    chkReboque3.Enabled = True
                    chkEstadoReboque3.Visible = True
                    imgEstadoReboque3.Visible = True

                    If String.IsNullOrEmpty(txtCodigoPlaca.Value) Then
                        chkReboque3.Checked = True
                        chkEstadoReboque3.Checked = True
                    End If

                    If objNotaFiscal.PlacaDetalhes IsNot Nothing Then
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).Placa04 = objNotaFiscal.PlacaDetalhes.CidadePlaca04 Then chkReboque3.Checked = True
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).EstadoPlaca04 = objNotaFiscal.PlacaDetalhes.EstadoPlaca04 Then chkEstadoReboque3.Checked = True
                        If Not CType(Session("objPlaca" & HID.Value), [Lib].Negocio.Placa).CidadePlaca04 = objNotaFiscal.PlacaDetalhes.CidadePlaca04 Then chkEstadoReboque3.Checked = True
                    End If

                End If
                Session.Remove("objPlaca" & HID.Value)
            ElseIf Not Session("objTransportador" & HID.Value) Is Nothing Then
                Dim itemTransportador As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objTransportador" & HID.Value), [Lib].Negocio.Cliente))
                txtTransportador.Text = itemTransportador.Text

                If String.IsNullOrEmpty(txtCodigoTransportador.Value) Then
                    chkTransportador.Checked = True
                Else
                    If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoTransportador) Then
                        If CType(Session("objTransportador" & HID.Value), [Lib].Negocio.Cliente).Codigo = objNotaFiscal.Transportador.Codigo AndAlso CType(Session("objTransportador" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco = objNotaFiscal.Transportador.CodigoEndereco Then
                            chkTransportador.Checked = False
                        Else
                            chkTransportador.Checked = True
                        End If
                    Else
                        chkTransportador.Checked = True
                    End If
                End If

                objNotaFiscal.CodigoTransportador = CType(Session("objTransportador" & HID.Value), [Lib].Negocio.Cliente).Codigo
                objNotaFiscal.EnderecoTransportador = CType(Session("objTransportador" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco

                Session.Remove("objTransportador" & HID.Value)

                SalvaNotaFiscal()
            ElseIf Session("objEstadoVeiculo" & HID.Value) IsNot Nothing Then
                Dim objEstado As [Lib].Negocio.Estado = obj
                lblEstadoVeiculo.Text = objEstado.Codigo & "-" & objEstado.Descricao
                Session.Remove("objEstadoVeiculo" & HID.Value)

                If objEstado.Codigo = objNotaFiscal.PlacaDetalhes.EstadoPlaca01Detalhes.Codigo Then
                    chkEstadoVeiculo.Checked = False
                Else
                    chkEstadoVeiculo.Checked = True
                End If

                Session("ssUF" & HID.Value) = objEstado.Codigo
                Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                If ucConsultaCodMunicipios IsNot Nothing Then
                    ucConsultaCodMunicipios.Limpar()
                End If
                Popup.ConsultaDeMunicipios(Me.Page, "objCidadeVeiculo" & HID.Value)
            ElseIf Session("objCidadeVeiculo" & HID.Value) IsNot Nothing Then
                Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidadeVeiculo" & HID.Value), [Lib].Negocio.Municipio)
                lblCidadeVeiculo.Text = objMunicipio.CodigoMunicipio

                If Trim(objMunicipio.CodigoMunicipio.ToUpper) = Trim(objNotaFiscal.PlacaDetalhes.CidadePlaca01.ToUpper) Then
                    chkEstadoVeiculo.Checked = False
                Else
                    chkEstadoVeiculo.Checked = True
                End If

                Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
                Session.Remove("objCidadeVeiculo" & HID.Value)
            ElseIf Session("objEstado1" & HID.Value) IsNot Nothing Then
                Dim objEstado As [Lib].Negocio.Estado = obj
                lblEstadoReboque1.Text = objEstado.Codigo & "-" & objEstado.Descricao
                Session.Remove("objEstado1" & HID.Value)

                If objEstado.Codigo = objNotaFiscal.PlacaDetalhes.Estado02Detalhes.Codigo Then
                    chkEstadoReboque1.Checked = False
                Else
                    chkEstadoReboque1.Checked = True
                End If

                Session("ssUF" & HID.Value) = objEstado.Codigo
                Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                If ucConsultaCodMunicipios IsNot Nothing Then
                    ucConsultaCodMunicipios.Limpar()
                End If
                Popup.ConsultaDeMunicipios(Me.Page, "objCidade1" & HID.Value)
            ElseIf Session("objCidade1" & HID.Value) IsNot Nothing Then
                Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidade1" & HID.Value), [Lib].Negocio.Municipio)
                lblCidadeReboque1.Text = objMunicipio.CodigoMunicipio

                If Trim(objMunicipio.CodigoMunicipio.ToUpper) = Trim(objNotaFiscal.PlacaDetalhes.CidadePlaca02.ToUpper) Then
                    chkEstadoReboque1.Checked = False
                Else
                    chkEstadoReboque1.Checked = True
                End If

                Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
                Session.Remove("objCidade1" & HID.Value)
            ElseIf Session("objEstado2" & HID.Value) IsNot Nothing Then
                Dim objEstado As [Lib].Negocio.Estado = obj
                lblEstadoReboque2.Text = objEstado.Codigo & "-" & objEstado.Descricao
                Session.Remove("objEstado2" & HID.Value)

                If objEstado.Codigo = objNotaFiscal.PlacaDetalhes.Estado03Detalhes.Codigo Then
                    chkEstadoReboque2.Checked = False
                Else
                    chkEstadoReboque2.Checked = True
                End If

                Session("ssUF" & HID.Value) = objEstado.Codigo
                Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                If ucConsultaCodMunicipios IsNot Nothing Then
                    ucConsultaCodMunicipios.Limpar()
                End If
                Popup.ConsultaDeMunicipios(Me.Page, "objCidade2" & HID.Value)
            ElseIf Session("objCidade2" & HID.Value) IsNot Nothing Then
                Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidade2" & HID.Value), [Lib].Negocio.Municipio)
                lblCidadeReboque2.Text = objMunicipio.CodigoMunicipio

                If Trim(objMunicipio.CodigoMunicipio.ToUpper) = Trim(objNotaFiscal.PlacaDetalhes.CidadePlaca03.ToUpper) Then
                    chkEstadoReboque2.Checked = False
                Else
                    chkEstadoReboque2.Checked = True
                End If

                Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
                Session.Remove("objCidade2" & HID.Value)
            ElseIf Session("objEstado3" & HID.Value) IsNot Nothing Then
                Dim objEstado As [Lib].Negocio.Estado = obj
                lblEstadoReboque3.Text = objEstado.Codigo & "-" & objEstado.Descricao
                Session.Remove("objEstado3" & HID.Value)

                If objEstado.Codigo = objNotaFiscal.PlacaDetalhes.Estado03Detalhes.Codigo Then
                    chkEstadoReboque3.Checked = False
                Else
                    chkEstadoReboque3.Checked = True
                End If

                Session("ssUF" & HID.Value) = objEstado.Codigo
                Dim ucConsultaCodMunicipios = CType(Me.Page.FindControlRecursive("ucConsultaCodMunicipios"), ucConsultaCodMunicipios)
                If ucConsultaCodMunicipios IsNot Nothing Then
                    ucConsultaCodMunicipios.Limpar()
                End If
                Popup.ConsultaDeMunicipios(Me.Page, "objCidade3" & HID.Value)
            ElseIf Session("objCidade3" & HID.Value) IsNot Nothing Then
                Dim objMunicipio As [Lib].Negocio.Municipio = CType(Session("objCidade3" & HID.Value), [Lib].Negocio.Municipio)
                lblCidadeReboque3.Text = objMunicipio.CodigoMunicipio

                If Trim(objMunicipio.CodigoMunicipio.ToUpper) = Trim(objNotaFiscal.PlacaDetalhes.CidadePlaca04.ToUpper) Then
                    chkEstadoReboque3.Checked = False
                Else
                    chkEstadoReboque3.Checked = True
                End If

                Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
                Session.Remove("objCidade3" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub limpar()
        Session.Remove("objCliente" & HID.Value)
        Session.Remove("objTransportador" & HID.Value)
        Session.Remove("objNFConsultaCCe" & HID.Value)
        Session.Remove("objNotaFiscal" & HID.Value)
        Session.Remove("objEstadoVeiculo" & HID.Value)
        Session.Remove("objEstado1" & HID.Value)
        Session.Remove("objEstado2" & HID.Value)
        Session.Remove("objEstado3" & HID.Value)
        Session.Remove("ssUF" & HID.Value)
        Session.Remove(HttpContext.Current.Session("ssTipoRetorno"))
        Session.Remove("objCidadeVeiculo" & HID.Value)
        Session.Remove("objCidade1" & HID.Value)
        Session.Remove("objCidade2" & HID.Value)
        Session.Remove("objCidade3" & HID.Value)

        txtES.Enabled = True
        txtNota.Enabled = True
        txtSerie.Enabled = True

        txtCliente.Text = String.Empty
        txtES.Text = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        lnkAtualizar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        lnkReimpressao.Parent.Visible = False

        txtTransportador.Text = String.Empty
        txtCodigoTransportador.Value = String.Empty
        chkTransportador.Checked = False
        chkTransportador.Visible = False
        btnTransportador.Enabled = False

        txtVeiculo.Text = String.Empty
        txtCodigoPlaca.Value = String.Empty
        txtVeiculo.Enabled = False
        imgPlaca.Visible = False
        lblEstadoVeiculo.Text = String.Empty
        lblCidadeVeiculo.Text = String.Empty
        chkVeiculo.Checked = False
        chkEstadoVeiculo.Checked = False
        chkVeiculo.Visible = False
        chkEstadoVeiculo.Visible = False
        imgEstadoVeiculo.Visible = False

        txtReboque1.Text = String.Empty
        txtReboque1.Enabled = False
        lblEstadoReboque1.Text = String.Empty
        lblCidadeReboque1.Text = String.Empty
        chkReboque1.Checked = False
        chkEstadoReboque1.Checked = False
        chkReboque1.Visible = False
        chkEstadoReboque1.Visible = False
        imgEstadoReboque1.Visible = False

        txtReboque2.Text = String.Empty
        txtReboque2.Enabled = False
        lblEstadoReboque2.Text = String.Empty
        lblCidadeReboque2.Text = String.Empty
        chkReboque2.Checked = False
        chkEstadoReboque2.Checked = False
        chkReboque2.Visible = False
        chkEstadoReboque2.Visible = False
        imgEstadoReboque2.Visible = False

        txtReboque3.Text = String.Empty
        txtReboque3.Enabled = False
        lblEstadoReboque3.Text = String.Empty
        lblCidadeReboque3.Text = String.Empty
        chkReboque3.Checked = False
        chkEstadoReboque3.Checked = False
        chkReboque3.Visible = False
        chkEstadoReboque3.Visible = False
        imgEstadoReboque3.Visible = False

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidosXNotas.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPlacas.SetarHID(HID.Value)
        ucConsultaEstados.SetarHID(HID.Value)
        ucConsultaCodMunicipios.SetarHID(HID.Value)

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

                If (objNotaFiscal.CodigoSituacao = 1) OrElse ((objNotaFiscal.CodigoTipoDeDocumento = 1 OrElse objNotaFiscal.CodigoTipoDeDocumento = 12) AndAlso objNotaFiscal.CodigoSituacao = 4) Then
                    Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Cliente)
                    txtCliente.Text = itemCliente.Text
                    txtES.Text = IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S")
                    txtNota.Text = objNotaFiscal.Codigo
                    txtSerie.Text = objNotaFiscal.Serie

                    txtES.Enabled = False
                    txtNota.Enabled = False
                    txtSerie.Enabled = False

                    SalvaNotaFiscal()
                    Session.Remove("objNFConsultaCCe" & HID.Value)
                    lnkAtualizar.Parent.Visible = True
                    lnkConsultar.Parent.Visible = False
                    btnTransportador.Enabled = True
                    chkTransportador.Visible = True
                    chkTransportador.Enabled = True
                    imgPlaca.Visible = True

                    If System.IO.File.Exists("C:\Alfasig\LeituraNFe\-danfes\" & objNotaFiscal.ChaveNFE & "cce.pdf") Then
                        lnkReimpressao.Parent.Visible = True
                    Else
                        lnkReimpressao.Parent.Visible = False
                    End If

                    If Not objNotaFiscal.Transportador Is Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.Transportador.Codigo) Then
                        Dim itemTransportador As ListItem = Funcoes.FormatarListItemCliente(objNotaFiscal.Transportador)
                        txtTransportador.Text = itemTransportador.Text
                        txtCodigoTransportador.Value = itemTransportador.Value
                    Else
                        txtCodigoTransportador.Value = String.Empty
                    End If

                    If Not objNotaFiscal.PlacaDetalhes Is Nothing AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.PlacaDetalhes.Placa01) Then
                        txtVeiculo.Text = objNotaFiscal.PlacaDetalhes.Placa01
                        lblEstadoVeiculo.Text = objNotaFiscal.PlacaDetalhes.EstadoPlaca01Detalhes.Codigo & "-" & objNotaFiscal.PlacaDetalhes.EstadoPlaca01Detalhes.Descricao
                        lblCidadeVeiculo.Text = objNotaFiscal.PlacaDetalhes.CidadePlaca01
                        chkEstadoVeiculo.Visible = True
                        imgEstadoVeiculo.Visible = True
                        chkVeiculo.Visible = True
                        chkVeiculo.Enabled = True

                        If Not String.IsNullOrWhiteSpace(objNotaFiscal.PlacaDetalhes.Placa02) Then
                            txtReboque1.Text = objNotaFiscal.PlacaDetalhes.Placa02
                            lblEstadoReboque1.Text = objNotaFiscal.PlacaDetalhes.Estado02Detalhes.Codigo & "-" & objNotaFiscal.PlacaDetalhes.Estado02Detalhes.Descricao
                            lblCidadeReboque1.Text = objNotaFiscal.PlacaDetalhes.CidadePlaca02
                            chkReboque1.Visible = True
                            chkReboque1.Enabled = True
                            chkEstadoReboque1.Visible = True
                            imgEstadoReboque1.Visible = True
                        End If

                        If Not String.IsNullOrWhiteSpace(objNotaFiscal.PlacaDetalhes.Placa03) Then
                            txtReboque2.Text = objNotaFiscal.PlacaDetalhes.Placa03
                            lblEstadoReboque2.Text = objNotaFiscal.PlacaDetalhes.Estado03Detalhes.Codigo & "-" & objNotaFiscal.PlacaDetalhes.Estado03Detalhes.Descricao
                            lblCidadeReboque2.Text = objNotaFiscal.PlacaDetalhes.CidadePlaca03
                            chkReboque2.Visible = True
                            chkReboque2.Enabled = True
                            chkEstadoReboque2.Visible = True
                            imgEstadoReboque2.Visible = True
                        End If

                        If Not String.IsNullOrWhiteSpace(objNotaFiscal.PlacaDetalhes.Placa04) Then
                            txtReboque3.Text = objNotaFiscal.PlacaDetalhes.Placa04
                            lblEstadoReboque3.Text = objNotaFiscal.PlacaDetalhes.Estado04Detalhes.Codigo & "-" & objNotaFiscal.PlacaDetalhes.Estado04Detalhes.Descricao
                            lblCidadeReboque3.Text = objNotaFiscal.PlacaDetalhes.CidadePlaca04
                            chkReboque3.Visible = True
                            chkReboque3.Enabled = True
                            chkEstadoReboque3.Visible = True
                            imgEstadoReboque3.Visible = True
                        End If
                    Else
                        txtCodigoPlaca.Value = String.Empty
                    End If
                Else
                    MsgBox(Me.Page, "Documento não pode ser alterado por aqui.")
                    limpar()
                End If


            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCampos() As Boolean
        If chkTransportador.Checked Then
            Return True
        ElseIf chkVeiculo.Checked AndAlso Not String.IsNullOrWhiteSpace(txtVeiculo.Text) Then
            Return True
        ElseIf chkEstadoVeiculo.Checked Then
            Return True
        ElseIf chkReboque1.Checked AndAlso Not String.IsNullOrWhiteSpace(txtReboque1.Text) Then
            Return True
        ElseIf chkEstadoReboque1.Checked Then
            Return True
        ElseIf chkReboque2.Checked AndAlso Not String.IsNullOrWhiteSpace(txtReboque2.Text) Then
            Return True
        ElseIf chkEstadoReboque2.Checked Then
            Return True
        ElseIf chkReboque3.Checked AndAlso Not String.IsNullOrWhiteSpace(txtReboque3.Text) Then
            Return True
        ElseIf chkEstadoReboque3.Checked Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region
End Class