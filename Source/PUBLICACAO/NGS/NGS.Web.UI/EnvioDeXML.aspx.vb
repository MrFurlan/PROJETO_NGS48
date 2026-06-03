Imports NGS.Lib.Negocio
Imports NGS.[Lib].Uteis

Public Class EnvioDeXML
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("EnvioDeXML", "ACESSAR") Then
                    ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                    Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                    limpar()
                    LiberaEmpresa()
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
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

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultar.Click
        Try
            consultar()
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
            Popup.ConsultaDeClientes(Me.Page, "objClienteEmail" & HID.Value.ToString, "txtCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedido.Click
        Try
            If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) OrElse String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido.")
                Exit Sub
            End If

            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
            HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"

            Dim parameters As New Dictionary(Of String, Object)
            parameters("unidade") = ""
            parameters("empresa") = strEmpresa(0)
            parameters("enderecoEmpresa") = strEmpresa(1)
            parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
            parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")

            Popup.ConsultaDePedidos(Me.Page, "objPedidoEmail" & HID.Value, "txtNome")
            ucConsultaPedidos.BindGridView(parameters)
        Catch ex As Exception


            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objClienteEmail" & HID.Value) IsNot Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteEmail" & HID.Value), [Lib].Negocio.Cliente))
                txtCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClienteEmail" & HID.Value)
            ElseIf Session("objPedidoEmail" & HID.Value) IsNot Nothing Then
                Dim p As [Lib].Negocio.Pedido = CType(Session("objPedidoEmail" & HID.Value), [Lib].Negocio.Pedido)
                txtPedido.Text = p.Codigo
                Session.Remove("objPedidoEmail" & HID.Value)
                consultar()
            ElseIf Session("objEnvioEmail" & HID.Value) IsNot Nothing Then
                Dim fm As New FilesManager()
                If fm.IsConnect() Then
                    Try

                        Dim objEmpresa As New [Lib].Negocio.Cliente(ddlEmpresa.SelectedValue.ToString.Split("-")(0), ddlEmpresa.SelectedValue.ToString.Split("-")(1))
                        Dim Sqls As New ArrayList

                        Dim strMensagem As String = IIf(rdNFe.Checked, "ANEXO NFe(S): ", "ANEXO CTe(S): ")

                        Dim randomNumber As String = FuncoesStrings.GerarNumeroSeguro(9)

                        Dim objFil As New [Lib].Negocio.Fil()
                        objFil.IUD = "I"

                        If rdNFe.Checked Then
                            objFil.NomeArquivo = String.Format("email{0:000000000}#{1}.txt", randomNumber, ddlEmpresa.SelectedValue.ToString.Split("-")(0))
                        Else
                            objFil.NomeArquivo = String.Format("emailcte{0:000000000}#{1}.txt", randomNumber, ddlEmpresa.SelectedValue.ToString.Split("-")(0))
                        End If

                        Dim sb As New StringBuilder()

                        Dim primeira As Boolean = True

                        For Each row As GridViewRow In gridDocumentos.Rows
                            Dim chk As CheckBox = CType(row.FindControl("chkEnviar"), CheckBox)
                            If chk IsNot Nothing AndAlso chk.Checked Then

                                If Mid(row.Cells(12).Text, 21, 2) = "55" Then
                                    sb.Append("CHAVENFE=" & row.Cells(12).Text & ControlChars.CrLf)
                                ElseIf Mid(row.Cells(12).Text, 21, 2) = "57" Then
                                    sb.Append("CHAVECTE=" & row.Cells(12).Text & ControlChars.CrLf)
                                End If

                                If primeira Then
                                    strMensagem += row.Cells(9).Text
                                    primeira = False
                                Else
                                    strMensagem += ", " & row.Cells(9).Text
                                End If
                            End If
                        Next

                        Dim lst As [Lib].Negocio.ListCliente = CType(Session("objEnvioEmail" & HID.Value), [Lib].Negocio.ListCliente)

                        For Each objCliente As [Lib].Negocio.Cliente In lst
                            sb.Append("DESTINATARIO=" & objCliente.EmailNFE & ControlChars.CrLf)
                        Next

                        sb.Append("ASSUNTO=" & Session("strAssunto" & HID.Value) & ControlChars.CrLf)
                        sb.Append("MENSAGEM=" & Session("strMensagem" & HID.Value) & ControlChars.CrLf)
                        sb.Append("EMAILEMITENTE=" & objEmpresa.EmailNFE & ControlChars.CrLf)
                        sb.Append("NOMEEMITENTE=" & objEmpresa.Nome & ControlChars.CrLf)

                        If rdPDFSim.Checked Then
                            sb.Append("ANEXOPDF=" & "SIM" & ControlChars.CrLf)
                        Else
                            sb.Append("ANEXOPDF=" & "NAO" & ControlChars.CrLf)
                        End If

                        sb.Append("ANEXOXML=" & "SIM" & ControlChars.CrLf)
                        sb.Append("ANEXOPROTOCOLO=" & "SIM" & ControlChars.CrLf)

                        If CBool(Session("strCompactado" & HID.Value)) Then
                            sb.Append("COMPACTADO=" & "SIM" & ControlChars.CrLf)
                        Else
                            sb.Append("COMPACTADO=" & "NAO" & ControlChars.CrLf)
                        End If

                        objFil.Texto = sb.ToString()
                        objFil.SalvarSql(Sqls)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                        End If

                        Dim strMsg As String = String.Empty
                        Dim resp As [Lib].Negocio.Resp = Nothing
                        Dim fileName As String = String.Empty

                        If rdNFe.Checked Then
                            fileName = String.Format("resp-email{0:000000000}#{1}.txt", randomNumber, ddlEmpresa.SelectedValue.ToString.Split("-")(0))
                        Else
                            fileName = String.Format("resp-emailcte{0:000000000}#{1}.txt", randomNumber, ddlEmpresa.SelectedValue.ToString.Split("-")(0))
                        End If

                        While resp Is Nothing
                            resp = GetResp(fileName)
                        End While

                        If resp IsNot Nothing Then
                            Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

                            Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                            Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()

                            Dim strCodigo As String = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                            strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()

                            If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso strCodigo <> "4014" Then
                                MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                                Exit Sub
                            Else
                                MsgBox(Me.Page, strMsg)
                            End If

                            Sqls.Clear()
                            Dim Sql As String = String.Empty

                            If rdNFe.Checked Then
                                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", randomNumber, ddlEmpresa.SelectedValue.ToString.Split("-")(0)) & "';"
                            Else
                                Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-emailcte{0:000000000}#{1}.txt", randomNumber, ddlEmpresa.SelectedValue.ToString.Split("-")(0)) & "';"
                            End If

                            Sqls.Add(Sql)

                            If Not Banco.GravaBanco(Sqls) Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            End If

                            limpar()

                        End If

                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                    End Try
                Else
                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CarregarEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub limpar()
        Session.Remove("ssCampo" & HID.Value)
        Session.Remove("objClienteEmail" & HID.Value)
        Session.Remove("objPedidoEmail" & HID.Value)
        Session.Remove("objEnvioEmail" & HID.Value)
        Session.Remove("strAssunto" & HID.Value)
        Session.Remove("strMensagem" & HID.Value)
        Session.Remove("strCompactado" & HID.Value)

        rdNFe.Checked = True
        rdConferencia.Enabled = True
        rdConferencia.Checked = True
        rdBanco.Enabled = False
        tipoTitulo.Visible = False

        lnkEnviarEmail.Parent.Visible = False
        lnkConsultar.Parent.Visible = True

        txtDataInicial.Enabled = True
        txtDataFinal.Enabled = True
        ddlEmpresa.Enabled = True
        ddlUnidadeDeNegocio.Enabled = True
        btnCliente.Enabled = True
        rdPDFNao.Checked = True

        txtCliente.Text = ""
        txtPedido.Text = ""

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPedidos.SetarHID(HID.Value)
        ucConsultaClientes.SetarHID(HID.Value)
        ucEmailNFe.SetarHID(HID.Value)

        gridDocumentos.DataSource = Nothing
        gridDocumentos.DataBind()
    End Sub

    Private Sub consultar()

        Dim Sql As String = String.Empty

        If rdConferencia.Checked Then
            Sql = "SELECT DISTINCT" & vbCrLf &
                    "        NF.Movimento," & vbCrLf &
                    "        NF.Empresa_Id," & vbCrLf &
                    "        NF.EndEmpresa_Id," & vbCrLf &
                    "        NF.Cliente_Id As Cliente," & vbCrLf &
                    "        NF.EndCliente_Id As EndCliente," & vbCrLf &
                    "        C.Nome As NomeCliente," & vbCrLf &
                    "        NF.EntradaSaida_Id As ES," & vbCrLf &
                    "        NF.Serie_Id As Serie," & vbCrLf &
                    "        NF.Nota_Id As Nota," & vbCrLf &
                    "        NF.Operacao," & vbCrLf &
                    "        NF.SubOperacao," & vbCrLf &
                    "        NF.NossaEmissao," & vbCrLf &
                    "        SUM(NFxI.QuantidadeFiscal) As Quantidade," & vbCrLf &
                    "        SUM(NFxE.Valor) As Valor," & vbCrLf &
                    "        NFr.ChaveNfe," & vbCrLf &
                    "        1 AS Situacao," & vbCrLf &
                    "        '' AS Historico," & vbCrLf &
                    "        0 AS ValorLiberado" & vbCrLf &
                    "    FROM NotasFiscais NF" & vbCrLf &
                    "      INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
                    "         ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                    "        AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                    "        AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                    "        AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                    "        AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                    "        AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                    "        AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                    "      INNER JOIN NotasFiscaisXEncargos NFxE" & vbCrLf &
                    "         ON NFxE.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                    "        AND NFxE.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                    "        AND NFxE.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                    "        AND NFxE.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                    "        AND NFxE.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                    "        AND NFxE.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                    "        AND NFxE.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                    "        AND NFxE.Produto_Id      = NFxI.Produto_Id" & vbCrLf &
                    "        AND NFxE.CFOP_Id         = NFxI.CFOP_Id" & vbCrLf &
                    "        AND NFxE.Sequencia_Id    = NFxI.Sequencia_Id" & vbCrLf &
                    "        AND NFxE.Encargo_Id      = 'LIQUIDO'" & vbCrLf &
                    "      INNER JOIN Clientes C" & vbCrLf &
                    "         ON C.Cliente_Id   = NF.Cliente_Id" & vbCrLf &
                    "        AND C.Endereco_Id = NF.EndCliente_Id" & vbCrLf &
                    "      INNER JOIN NFERealizadas NFr" & vbCrLf &
                    "	     ON NFr.Empresa_Id       = NF.Empresa_Id" & vbCrLf &
                    "	     AND NFr.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                    "	     AND NFr.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                    "	     AND NFr.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                    "	     AND NFr.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                    "	     AND NFr.Serie_Id        = NF.Serie_Id" & vbCrLf &
                    "	     AND NFr.Nota_Id         = NF.Nota_Id" & vbCrLf &
                    "    WHERE NF.Eletronica      = 'S'" & vbCrLf &
                    "      AND NF.Situacao        = 1" & vbCrLf &
                    "      AND NF.Empresa_Id      = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "'" & vbCrLf &
                    "      AND NF.EndEmpresa_Id   = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & vbCrLf

            If rdNFe.Checked Then
                Sql &= "      AND NF.TipoDeDocumento IN(1)" & vbCrLf
            End If

            If rdCTe.Checked Then
                Sql &= "      AND NF.TipoDeDocumento IN(57,58,67)" & vbCrLf
            End If

            If rdEntradas.Checked Then
                Sql &= "      AND NF.EntradaSaida_Id = 'E'" & vbCrLf
            ElseIf rdSaidas.Checked Then
                Sql &= "      AND NF.EntradaSaida_Id = 'S'" & vbCrLf
            End If

            If txtCodigoCliente.Value.ToString.Length > 0 Then
                Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                Sql &= "    AND NF.Cliente_Id = '" & strCliente(0) & "' And NF.EndCliente_Id = " & strCliente(1) & vbCrLf
            End If

            If txtPedido.Text.Length > 0 Then
                Sql &= "    AND NF.Pedido = " & txtPedido.Text & vbCrLf
            Else
                Sql &= "    AND (NF.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "')" & vbCrLf
            End If

            Sql &= "GROUP BY NF.Movimento," & vbCrLf &
                    "          NF.Empresa_Id," & vbCrLf &
                    "          NF.EndEmpresa_Id," & vbCrLf &
                    "          NF.Cliente_Id," & vbCrLf &
                    "          NF.EndCliente_Id," & vbCrLf &
                    "          C.Nome," & vbCrLf &
                    "          NF.EntradaSaida_Id," & vbCrLf &
                    "          NF.Serie_Id," & vbCrLf &
                    "          NF.Nota_Id," & vbCrLf &
                    "          NF.Operacao," & vbCrLf &
                    "          NF.SubOperacao," & vbCrLf &
                    "          NF.NossaEmissao," & vbCrLf &
                    "		  NFr.ChaveNfe"
        End If

        If rdBanco.Checked Then
            Sql = "/*" & vbCrLf &
                "" & vbCrLf &
                "--Vermelho = 0" & vbCrLf &
                "--Amarelo  = 1" & vbCrLf &
                "--Verde    = 2" & vbCrLf &
                "" & vbCrLf &
                "*/" & vbCrLf &
                "" & vbCrLf &
                ";WITH NF_CTE AS (" & vbCrLf &
                "    /* SCRIPT 1 */" & vbCrLf &
                "    SELECT DISTINCT" & vbCrLf &
                "        NF.Movimento," & vbCrLf &
                "        NF.Empresa_Id," & vbCrLf &
                "        NF.EndEmpresa_Id," & vbCrLf &
                "        NF.Cliente_Id As Cliente," & vbCrLf &
                "        NF.EndCliente_Id As EndCliente," & vbCrLf &
                "        C.Nome As NomeCliente," & vbCrLf &
                "        NF.EntradaSaida_Id As ES," & vbCrLf &
                "        NF.Serie_Id As Serie," & vbCrLf &
                "        NF.Nota_Id As Nota," & vbCrLf &
                "        NF.Operacao," & vbCrLf &
                "        NF.SubOperacao," & vbCrLf &
                "        NF.NossaEmissao," & vbCrLf &
                "        SUM(NFxI.QuantidadeFiscal) As Quantidade," & vbCrLf &
                "        SUM(NFxE.Valor) As Valor," & vbCrLf &
                "        NFr.ChaveNfe" & vbCrLf &
                "    FROM NotasFiscais NF" & vbCrLf &
                "      INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
                "         ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                "        AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                "        AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                "        AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                "        AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                "        AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                "        AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                "      INNER JOIN NotasFiscaisXEncargos NFxE" & vbCrLf &
                "         ON NFxE.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                "        AND NFxE.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                "        AND NFxE.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                "        AND NFxE.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                "        AND NFxE.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                "        AND NFxE.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                "        AND NFxE.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                "        AND NFxE.Produto_Id      = NFxI.Produto_Id" & vbCrLf &
                "        AND NFxE.CFOP_Id         = NFxI.CFOP_Id" & vbCrLf &
                "        AND NFxE.Sequencia_Id    = NFxI.Sequencia_Id" & vbCrLf &
                "        AND NFxE.Encargo_Id      = 'LIQUIDO'" & vbCrLf &
                "      INNER JOIN Clientes C" & vbCrLf &
                "         ON C.Cliente_Id   = NF.Cliente_Id" & vbCrLf &
                "        AND C.Endereco_Id = NF.EndCliente_Id" & vbCrLf &
                "      INNER JOIN NFERealizadas NFr" & vbCrLf &
                "	     ON NFr.Empresa_Id       = NF.Empresa_Id" & vbCrLf &
                "	     AND NFr.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                "	     AND NFr.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                "	     AND NFr.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                "	     AND NFr.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                "	     AND NFr.Serie_Id        = NF.Serie_Id" & vbCrLf &
                "	     AND NFr.Nota_Id         = NF.Nota_Id" & vbCrLf &
                "      INNER JOIN NotaFiscalXTitulo NFxT" & vbCrLf &
                "	     ON NFxT.Empresa_Id       = NF.Empresa_Id" & vbCrLf &
                "	     AND NFxT.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                "	     AND NFxT.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
                "	     AND NFxT.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
                "	     AND NFxT.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
                "	     AND NFxT.Serie_Id        = NF.Serie_Id" & vbCrLf &
                "	     AND NFxT.Nota_Id         = NF.Nota_Id" & vbCrLf &
                "      INNER JOIN ContasAReceber Tit" & vbCrLf &
                "	     ON Tit.Registro_Id = NFxT.Titulo_Id" & vbCrLf &
                "    WHERE NF.TipoDeDocumento IN(1)" & vbCrLf &
                "      AND NF.Situacao        = 1" & vbCrLf &
                "      AND NF.NossaEmissao    = 'S'" & vbCrLf &
                "      AND NF.EntradaSaida_Id = 'S'" & vbCrLf &
                "      AND NF.Empresa_Id      = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "'" & vbCrLf &
                "      AND NF.EndEmpresa_Id   = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & vbCrLf

            If txtCodigoCliente.Value.ToString.Length > 0 Then
                Dim strCliente() As String = txtCodigoCliente.Value.ToString.Split("-")
                Sql &= "    AND NF.Cliente_Id = '" & strCliente(0) & "' And NF.EndCliente_Id = " & strCliente(1) & vbCrLf
            End If

            If txtPedido.Text.Length > 0 Then
                Sql &= "    AND NF.Pedido = " & txtPedido.Text & vbCrLf
            Else
                Sql &= "    AND (NF.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "')" & vbCrLf
            End If

            Sql &= "      AND Tit.Provisao = 2" & vbCrLf &
                "    GROUP BY NF.Movimento," & vbCrLf &
                "          NF.Empresa_Id," & vbCrLf &
                "          NF.EndEmpresa_Id," & vbCrLf &
                "          NF.Cliente_Id," & vbCrLf &
                "          NF.EndCliente_Id," & vbCrLf &
                "          C.Nome," & vbCrLf &
                "          NF.EntradaSaida_Id," & vbCrLf &
                "          NF.Serie_Id," & vbCrLf &
                "          NF.Nota_Id," & vbCrLf &
                "          NF.Operacao," & vbCrLf &
                "          NF.SubOperacao," & vbCrLf &
                "          NF.NossaEmissao," & vbCrLf &
                "		  NFr.ChaveNfe" & vbCrLf &
                ")," & vbCrLf &
                "TIT_CTE AS (" & vbCrLf &
                "    /* SCRIPT 2 (genérico para todos os títulos da NF) */" & vbCrLf &
                "    SELECT " & vbCrLf &
                "        NFxT.Empresa_Id," & vbCrLf &
                "        NFxT.EndEmpresa_Id," & vbCrLf &
                "        NFxT.Cliente_Id," & vbCrLf &
                "        NFxT.EndCliente_Id," & vbCrLf &
                "        NFxT.EntradaSaida_Id," & vbCrLf &
                "        NFxT.Serie_Id," & vbCrLf &
                "        NFxT.Nota_Id," & vbCrLf &
                "        Tit.Registro_Id as Titulo," & vbCrLf &
                "        Tit.Prorrogacao as Vencimento," & vbCrLf &
                "        Tit.ValorLiquido as Valor," & vbCrLf &
                "        CASE " & vbCrLf &
                "            WHEN Tit.Prorrogacao < CAST(GETDATE() AS DATE) THEN 1 ELSE 0" & vbCrLf &
                "        END AS Vencido" & vbCrLf &
                "    FROM NotaFiscalXTitulo NFxT" & vbCrLf &
                "      INNER JOIN ContasAReceber Tit " & vbCrLf &
                "        ON NFxT.Titulo_Id = Tit.Registro_Id" & vbCrLf &
                ")" & vbCrLf &
                "SELECT " & vbCrLf &
                "    NF.Movimento," & vbCrLf &
                "    NF.Empresa_Id," & vbCrLf &
                "    NF.EndEmpresa_Id," & vbCrLf &
                "    NF.Cliente," & vbCrLf &
                "    NF.EndCliente," & vbCrLf &
                "    NF.NomeCliente," & vbCrLf &
                "    NF.ES," & vbCrLf &
                "    NF.Serie," & vbCrLf &
                "    NF.Nota," & vbCrLf &
                "    NF.Operacao," & vbCrLf &
                "    NF.SubOperacao," & vbCrLf &
                "    NF.NossaEmissao," & vbCrLf &
                "    NF.Quantidade," & vbCrLf &
                "    NF.Valor," & vbCrLf &
                "    NF.ChaveNfe," & vbCrLf &
                "" & vbCrLf &
                "    -- Situação" & vbCrLf &
                "    CASE " & vbCrLf &
                "        WHEN COUNT(T.Titulo) = 1 " & vbCrLf &
                "             AND MAX(T.Vencido) = 1 THEN 0 -- Vermelho" & vbCrLf &
                "        WHEN COUNT(T.Titulo) = 1 " & vbCrLf &
                "             AND MAX(T.Vencido) = 0 THEN 2 -- Verde" & vbCrLf &
                "        WHEN COUNT(T.Titulo) > 1 " & vbCrLf &
                "             AND SUM(T.Vencido) = COUNT(T.Titulo) THEN 0 -- todos vencidos" & vbCrLf &
                "        WHEN COUNT(T.Titulo) > 1 " & vbCrLf &
                "             AND SUM(T.Vencido) = 0 THEN 2 -- todos a vencer" & vbCrLf &
                "        WHEN COUNT(T.Titulo) > 1 " & vbCrLf &
                "             AND SUM(T.Vencido) BETWEEN 1 AND COUNT(T.Titulo)-1 THEN 1 -- misto" & vbCrLf &
                "    END AS Situacao," & vbCrLf &
                "" & vbCrLf &
                "    -- Histórico (lista dos títulos)" & vbCrLf &
                "    STUFF((" & vbCrLf &
                "    SELECT CHAR(13) + CHAR(10) +" & vbCrLf &
                "           CAST(T2.Titulo AS VARCHAR(20)) + CHAR(9) +" & vbCrLf &
                "           CONVERT(VARCHAR(10), T2.Vencimento, 103) + CHAR(9) +" & vbCrLf &   ' dd/MM/yyyy
                "           CAST(T2.Valor AS VARCHAR(20)) + CHAR(9) +" & vbCrLf &
                "           CASE WHEN T2.Vencido = 1 THEN 'VENCIDO' ELSE 'LIBERADO' END" & vbCrLf &
                "    FROM TIT_CTE T2" & vbCrLf &
                "    WHERE T2.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                "      AND T2.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                "      AND T2.Cliente_Id      = NF.Cliente" & vbCrLf &
                "      AND T2.EndCliente_Id   = NF.EndCliente" & vbCrLf &
                "      AND T2.EntradaSaida_Id = NF.ES" & vbCrLf &
                "      AND T2.Serie_Id        = NF.Serie" & vbCrLf &
                "      AND T2.Nota_Id         = NF.Nota" & vbCrLf &
                "    FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')" & vbCrLf &
                ", 1, 2, '') AS Historico," & vbCrLf &
                "" & vbCrLf &
                "    -- ValorLiberado (somente títulos a vencer)" & vbCrLf &
                "    SUM(CASE WHEN T.Vencido = 0 THEN T.Valor ELSE 0 END) AS ValorLiberado" & vbCrLf &
                "" & vbCrLf &
                "FROM NF_CTE NF" & vbCrLf &
                "LEFT JOIN TIT_CTE T" & vbCrLf &
                "  ON T.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
                " AND T.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
                " AND T.Cliente_Id      = NF.Cliente" & vbCrLf &
                " AND T.EndCliente_Id   = NF.EndCliente" & vbCrLf &
                " AND T.EntradaSaida_Id = NF.ES" & vbCrLf &
                " AND T.Serie_Id        = NF.Serie" & vbCrLf &
                " AND T.Nota_Id         = NF.Nota" & vbCrLf &
                "GROUP BY NF.Movimento," & vbCrLf &
                "    NF.Empresa_Id," & vbCrLf &
                "    NF.EndEmpresa_Id," & vbCrLf &
                "    NF.Cliente," & vbCrLf &
                "    NF.EndCliente," & vbCrLf &
                "    NF.NomeCliente," & vbCrLf &
                "    NF.ES," & vbCrLf &
                "    NF.Serie," & vbCrLf &
                "    NF.Nota," & vbCrLf &
                "    NF.Operacao," & vbCrLf &
                "    NF.SubOperacao," & vbCrLf &
                "    NF.NossaEmissao," & vbCrLf &
                "    NF.Quantidade," & vbCrLf &
                "    NF.Valor," & vbCrLf &
                "    NF.ChaveNfe" & vbCrLf &
                "ORDER BY NF.Movimento DESC, NomeCliente,  Nota DESC, Serie DESC;"
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "EnvioDeXML")

        gridDocumentos.DataSource = ds
        gridDocumentos.DataBind()

        If rdConferencia.Checked Then
            gridDocumentos.Columns(13).Visible = False
            gridDocumentos.Columns(14).Visible = False
        Else
            gridDocumentos.Columns(13).Visible = True
            gridDocumentos.Columns(14).Visible = True
        End If

        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "init-tooltips", "initTooltips && initTooltips();", True)

    End Sub

    Protected Function TooltipHistorico(obj As Object) As String
        Dim raw As String = ""
        If obj IsNot Nothing AndAlso Not IsDBNull(obj) Then
            raw = Convert.ToString(obj)
        End If
        raw = System.Web.HttpUtility.HtmlEncode(raw) _
                .Replace(vbCrLf, "<br/>") _
                .Replace(vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
        Return raw
    End Function

    Protected Sub chkEnviarAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        Try
            If gridDocumentos.Rows.Count > 0 Then
                Dim chkEnviarAll As CheckBox = CType(sender, CheckBox)

                For Each rowgrid As GridViewRow In gridDocumentos.Rows

                    Dim chkEnviar As CheckBox = CType(rowgrid.FindControl("chkEnviar"), CheckBox)

                    RemoveHandler chkEnviar.CheckedChanged, AddressOf chkEnviar_CheckedChanged

                    If chkEnviar.Enabled = True Then
                        chkEnviar.Checked = chkEnviarAll.Checked

                        If chkEnviar.Checked Then
                            lnkEnviarEmail.Parent.Visible = True
                            lnkConsultar.Parent.Visible = False
                        Else
                            lnkEnviarEmail.Parent.Visible = False
                            lnkConsultar.Parent.Visible = True
                        End If
                    End If

                    AddHandler chkEnviar.CheckedChanged, AddressOf chkEnviar_CheckedChanged

                Next

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "init-tooltips", "initTooltips && initTooltips();", True)

    End Sub

    Protected Sub chkEnviar_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If gridDocumentos.Rows.Count > 0 Then

                Dim temCHK As Boolean = False

                For Each rowgrid As GridViewRow In gridDocumentos.Rows
                    Dim chkEnviar As CheckBox = CType(rowgrid.FindControl("chkEnviar"), CheckBox)

                    If chkEnviar.Checked Then
                        temCHK = True
                        Exit For
                    End If
                Next

                If temCHK Then
                    lnkEnviarEmail.Parent.Visible = True
                    lnkConsultar.Parent.Visible = False
                Else
                    lnkEnviarEmail.Parent.Visible = False
                    lnkConsultar.Parent.Visible = True
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEnviarEmail_Click(sender As Object, e As EventArgs) Handles lnkEnviarEmail.Click

        Dim objEmpresa As New [Lib].Negocio.Cliente(ddlEmpresa.SelectedValue.ToString.Split("-")(0), ddlEmpresa.SelectedValue.ToString.Split("-")(1))

        If String.IsNullOrWhiteSpace(objEmpresa.EmailNFE) Then
            MsgBox(Me.Page, "Email da Empresa " & objEmpresa.CodigoFormatado & "-" & objEmpresa.Nome & " não está cadastrado, favor verificar.")
        Else
            Dim itemSelecionado As Boolean = False

            Dim i As Integer = 0
            While i < gridDocumentos.Rows.Count
                If CType(gridDocumentos.Rows(i).FindControl("chkEnviar"), CheckBox).Checked Then itemSelecionado = True
                i += 1
            End While

            If itemSelecionado Then
                Dim fm As New FilesManager()
                If fm.IsConnect() Then

                    Dim primeira As Boolean = True

                    Dim strMensagem As String = IIf(rdNFe.Checked, "ANEXO NFe(S): ", "ANEXO CTe(S): ")

                    For Each row As GridViewRow In gridDocumentos.Rows
                        Dim chk As CheckBox = CType(row.FindControl("chkEnviar"), CheckBox)
                        If chk IsNot Nothing AndAlso chk.Checked Then
                            If primeira Then
                                strMensagem += row.Cells(9).Text
                                primeira = False
                            Else
                                strMensagem += ", " & row.Cells(9).Text
                            End If
                        End If
                    Next

                    Dim parameters As New Dictionary(Of String, Object)

                    If rdNFe.Checked Then
                        parameters("Assunto") = "XML NFe(S)"
                    Else
                        parameters("Assunto") = "XML CTe(S)"
                    End If

                    parameters("Mensagem") = strMensagem
                    If rdZIPSim.Checked Then
                        parameters("Compactado") = "SIM"
                    Else
                        parameters("Compactado") = "NAO"
                    End If

                    Dim ucEmailNFe As ucEmailNFe = CType(Me.Page.FindControlRecursive("ucEmailNFe"), ucEmailNFe)
                    ucEmailNFe.Limpar()
                    ucEmailNFe.EnvioDeEmail(parameters)

                    Dim txtDestinatario As TextBox = CType(ucEmailNFe.FindControlRecursive("txtDestinatario"), TextBox)
                    Popup.ConsultaDeEmailNFe(Me.Page, "objEnvioEmail" & HID.Value, txtDestinatario.ClientID, 100)
                Else
                    MsgBox(Me.Page, "Serviço de emissão da Alfasig (Guardian) não está em funcionamento, verifique seu servidor!")
                End If
            Else
                MsgBox(Me.Page, "Não foi selecionado nenhum item no Grid para Envio!")
            End If
        End If
    End Sub

    Protected Sub gridDocumentos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles gridDocumentos.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then

            ' Localiza a CheckBox
            Dim chkEnviar As CheckBox = TryCast(e.Row.FindControl("chkEnviar"), CheckBox)

            If chkEnviar IsNot Nothing Then
                ' Verifica a coluna "Situacao" (sem acento!)
                Dim status As Integer = 0
                Dim valObj = DataBinder.Eval(e.Row.DataItem, "Situacao")
                If valObj IsNot Nothing AndAlso valObj IsNot DBNull.Value Then
                    Integer.TryParse(valObj.ToString(), status)
                End If

                If status = 0 Then
                    chkEnviar.Enabled = False  ' desabilita
                End If
            End If

            Dim img As Image = TryCast(e.Row.FindControl("imgTitulos"), Image)
            If img IsNot Nothing Then
                ' ativa tooltip do Bootstrap via data-attrs
                img.Attributes("data-bs-toggle") = "tooltip"
                img.Attributes("data-bs-html") = "true"
                img.Attributes("data-bs-custom-class") = "tt-titulos"
                ' o conteúdo HTML já vem do ToolTip (TooltipHistorico) -> vira title
            End If

        End If
    End Sub

    Protected Sub rdNFe_CheckedChanged(sender As Object, e As EventArgs) Handles rdNFe.CheckedChanged
        Try
            rdConferencia.Checked = True
            rdConferencia.Enabled = True
            If rdSaidas.Checked Then rdBanco.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdCTe_CheckedChanged(sender As Object, e As EventArgs) Handles rdCTe.CheckedChanged
        Try
            rdConferencia.Checked = True
            rdBanco.Checked = False
            rdBanco.Enabled = False
            tipoTitulo.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdEntradas_CheckedChanged(sender As Object, e As EventArgs) Handles rdEntradas.CheckedChanged
        Try
            rdConferencia.Checked = True
            rdBanco.Checked = False
            rdBanco.Enabled = False
            tipoTitulo.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdSaidas_CheckedChanged(sender As Object, e As EventArgs) Handles rdSaidas.CheckedChanged
        Try
            If rdNFe.Checked Then
                rdBanco.Enabled = True
            Else
                rdBanco.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdTodas_CheckedChanged(sender As Object, e As EventArgs) Handles rdTodas.CheckedChanged
        Try
            rdConferencia.Checked = True
            rdBanco.Checked = False
            rdBanco.Enabled = False
            tipoTitulo.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdConferencia_CheckedChanged(sender As Object, e As EventArgs) Handles rdConferencia.CheckedChanged
        Try
            tipoTitulo.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdBanco_CheckedChanged(sender As Object, e As EventArgs) Handles rdBanco.CheckedChanged
        Try
            tipoTitulo.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class