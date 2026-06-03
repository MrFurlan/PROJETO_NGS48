Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioDeFretes
    Inherits BasePage

    Private objCliente As [Lib].Negocio.Cliente
    Private objPedido As [Lib].Negocio.Pedido

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeFretes", "ACESSAR") Then
                    BuscaEmpresa()
                    BuscarGruposProdutos()
                    txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaClientes.SetarHID(HID.Value)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteFRT" & HID.Value) Is Nothing Then
            objCliente = CType(Session("objClienteFRT" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteFRT" & HID.Value)
        ElseIf Not Session("objClienteFRTTrans" & HID.Value) Is Nothing Then
            objCliente = CType(Session("objClienteFRTTrans" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
            txtTransportador.Text = itemCliente.Text
            txtCodigoTransportador.Value = itemCliente.Value
            Session.Remove("objClienteFRTTrans" & HID.Value)
        ElseIf Not Session("objPedidoRDF" & HID.Value) Is Nothing Then
            objPedido = CType(Session("objPedidoRDF" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = objPedido.Codigo
            ddlGrupoProduto.Enabled = False
            ddlProduto.Enabled = False
            Session.Remove("objPedidoRDF" & HID.Value)
        End If
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, "", True)
        Funcoes.VerificaEmpresa(cmbEmpresa)
    End Sub

    Private Sub BuscarGruposProdutos()
        ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub BuscarProduto()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & ddlGrupoProduto.SelectedValue & "'", True)
    End Sub

    Private Sub Limpar()
        cmbEmpresa.SelectedIndex = 0
        txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtClientes.Text = ""
        txtCodigoCliente.Value = ""
        txtTransportador.Text = ""
        txtCodigoTransportador.Value = ""
        txtPedido.Text = ""
        ddlGrupoProduto.SelectedIndex = 0
        ddlProduto.Items.Clear()
        ddlGrupoProduto.Enabled = True
        ddlProduto.Enabled = True
        txtPedido.Enabled = True
        btnPedido.Enabled = True
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbEmpresa.Enabled = False
        End If
    End Sub

    Public Function ValidarCampos() As Boolean
        If cmbEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a empresa.")
            cmbEmpresa.Focus()
            Return False
        End If
        If txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida.")
            txtDataInicial.Focus()
            Return False
        End If
        If txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida.")
            txtDataFinal.Focus()
            Return False
        End If
        If CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final.")
            txtDataInicial.Focus()
            Return False
        End If

        Return True
    End Function

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteFRT" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPedido.Click
        Try
            Dim strJavaScript As String = ""

            If cmbEmpresa.SelectedIndex < 1 Or txtCodigoCliente.Value.Length = 0 Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido")
            Else
                Dim strEmpresa As String() = cmbEmpresa.SelectedValue.Split("-")
                Dim strCliente As String() = txtCodigoCliente.Value.Split("-")

                HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"

                Dim parameters As New Dictionary(Of String, Object)
                parameters("unidade") = ""
                parameters("empresa") = strEmpresa(0)
                parameters("enderecoEmpresa") = strEmpresa(1)
                parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
                parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
                Popup.ConsultaDePedidos(Me.Page, "objPedidoRDF" & HID.Value, "txtNome")
                ucConsultaPedidos.BindGridView(parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function BuscarRegistros() As DataSet
        Dim ds As New DataSet
        Dim Cliente As String = ""
        Dim strEmpresa() As String = cmbEmpresa.SelectedValue.Split("-")
        Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
        Dim strTransportador() As String = txtCodigoTransportador.Value.Split("-")
        Dim strProduto() As String = ddlProduto.SelectedValue.ToString.Split("_")
        Dim strSQL As String = ""

        If rdRecibo.Checked = True Then
            strSQL = "SELECT DISTINCT " & vbCrLf & _
                     "       NotasFiscais.Empresa_Id AS Empresa," & vbCrLf & _
                     "       NotasFiscais.EndEmpresa_Id AS EndEmpresa," & vbCrLf & _
                     "       NotasFiscais.Cliente_Id AS Cliente," & vbCrLf & _
                     "       NotasFiscais.EndCliente_Id AS EndCliente," & vbCrLf & _
                     "       Clientes.Nome AS NomeCliente," & vbCrLf & _
                     "       Recibo.EntradaSaida_Id AS EntradaSaida," & vbCrLf & _
                     "       Recibo.Serie_Id AS Serie," & vbCrLf & _
                     "       Recibo.Nota_Id AS Recibo," & vbCrLf & _
                     "       Recibo.CFOP_Id AS CFOP, " & vbCrLf & _
                     "       Recibo.QuantidadeFiscal AS Peso, " & vbCrLf & _
                     "       Recibo.Unitario, " & vbCrLf & _
                     "       Recibo.Valor," & vbCrLf & _
                     "       isnull(ReciboXEnc.Valor,0) AS Adiantamento," & vbCrLf & _
                     "       (Recibo.Valor - isnull(ReciboXEnc.Valor,0)) as Saldo," & vbCrLf & _
                     "       CONVERT(varchar, NotasFiscais.Movimento, 103) AS Movimento, " & vbCrLf & _
                     "       NotasFiscais.Pedido, " & vbCrLf & _
                     "       NotasFiscais.EntradaSaida_Id + '-' + CONVERT(varchar, NotasFiscais.Nota_Id) + '-' + NotasFiscais.Serie_Id AS NotaFiscal " & vbCrLf & _
                     "  FROM NotasFiscais " & vbCrLf & _
                     " INNER JOIN NotasFiscaisXItens " & vbCrLf & _
                     "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                     "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                     "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                     "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                     "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                     "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                     "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                     " INNER JOIN NotasXNotas AS NotasXCte " & vbCrLf & _
                     "    ON NotasFiscais.Empresa_Id      = NotasXCte.OrigemEmpresa_Id " & vbCrLf & _
                     "   AND NotasFiscais.EndEmpresa_Id   = NotasXCte.OrigemEndEmpresa_Id " & vbCrLf & _
                     "   AND NotasFiscais.Cliente_Id      = NotasXCte.OrigemCliente_Id " & vbCrLf & _
                     "   AND NotasFiscais.EndCliente_Id   = NotasXCte.OrigemEndCliente_Id " & vbCrLf & _
                     "   AND NotasFiscais.EntradaSaida_Id = NotasXCte.OrigemEntradaSaida_Id " & vbCrLf & _
                     "   AND NotasFiscais.Serie_Id        = NotasXCte.OrigemSerie_Id " & vbCrLf & _
                     "   AND NotasFiscais.Nota_Id         = NotasXCte.OrigemNota_Id" & vbCrLf & _
                     " INNER JOIN NotasFiscaisXTransportadores NFxT " & vbCrLf & _
                     "    ON NFxT.Empresa_Id      = NotasFiscais.Empresa_Id " & vbCrLf & _
                     "   AND NFxT.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id " & vbCrLf & _
                     "   AND NFxT.Cliente_Id      = NotasFiscais.Cliente_Id " & vbCrLf & _
                     "   AND NFxT.EndCliente_Id   = NotasFiscais.EndCliente_Id " & vbCrLf & _
                     "   AND NFxT.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id " & vbCrLf & _
                     "   AND NFxT.Serie_Id        = NotasFiscais.Serie_Id " & vbCrLf & _
                     "   AND NFxT.Nota_Id         = NotasFiscais.Nota_Id" & vbCrLf & _
                     " INNER JOIN NotasXNotas AS CteXRec" & vbCrLf & _
                     "    ON NotasXCte.Empresa_Id      = CteXRec.OrigemEmpresa_Id " & vbCrLf & _
                     "   AND NotasXCte.EndEmpresa_Id   = CteXRec.OrigemEndEmpresa_Id " & vbCrLf & _
                     "   AND NotasXCte.Cliente_Id      = CteXRec.OrigemCliente_Id " & vbCrLf & _
                     "   AND NotasXCte.EndCliente_Id   = CteXRec.OrigemEndCliente_Id " & vbCrLf & _
                     "   AND NotasXCte.EntradaSaida_Id = CteXRec.OrigemEntradaSaida_Id " & vbCrLf & _
                     "   AND NotasXCte.Serie_Id        = CteXRec.OrigemSerie_Id " & vbCrLf & _
                     "   AND NotasXCte.Nota_Id         = CteXRec.OrigemNota_Id " & vbCrLf & _
                     " INNER JOIN NotasFiscaisXItens AS Recibo " & vbCrLf & _
                     "    ON CteXRec.Empresa_Id      = Recibo.Empresa_Id" & vbCrLf & _
                     "   AND CteXRec.EndEmpresa_Id   = Recibo.EndEmpresa_Id " & vbCrLf & _
                     "   AND CteXRec.Cliente_Id      = Recibo.Cliente_Id " & vbCrLf & _
                     "   AND CteXRec.EndCliente_Id   = Recibo.EndCliente_Id " & vbCrLf & _
                     "   AND CteXRec.EntradaSaida_Id = Recibo.EntradaSaida_Id " & vbCrLf & _
                     "   AND CteXRec.Serie_Id        = 'REC' " & vbCrLf & _
                     "   AND CteXRec.Nota_Id         = Recibo.Nota_Id " & vbCrLf & _
                     " LEFT JOIN NotasFiscaisXEncargos AS ReciboXEnc " & vbCrLf & _
                     "    ON ReciboXEnc.Empresa_Id      = Recibo.Empresa_Id " & vbCrLf & _
                     "   AND ReciboXEnc.EndEmpresa_Id   = Recibo.EndEmpresa_Id " & vbCrLf & _
                     "   AND ReciboXEnc.Cliente_Id      = Recibo.Cliente_Id " & vbCrLf & _
                     "   AND ReciboXEnc.EndCliente_Id   = Recibo.EndCliente_Id " & vbCrLf & _
                     "   AND ReciboXEnc.EntradaSaida_Id = Recibo.EntradaSaida_Id " & vbCrLf & _
                     "   AND ReciboXEnc.Serie_Id        = Recibo.Serie_Id " & vbCrLf & _
                     "   AND ReciboXEnc.Nota_Id         = Recibo.Nota_Id " & vbCrLf & _
                     "   AND ReciboXEnc.Produto_Id      = Recibo.Produto_Id " & vbCrLf & _
                     "   AND ReciboXEnc.CFOP_Id         = Recibo.CFOP_Id " & vbCrLf & _
                     "   AND ReciboXEnc.Encargo_Id      = 'ADTODEFRETE' " & vbCrLf & _
                     " INNER JOIN Clientes " & vbCrLf & _
                     "    ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf & _
                     "   AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id " & vbCrLf & _
                     " WHERE NotasFiscais.Empresa_Id    ='" & strEmpresa(0) & "'" & vbCrLf & _
                     "   AND NotasFiscais.EndEmpresa_Id ='" & strEmpresa(1) & "'" & vbCrLf & _
                     "   AND ((Recibo.CFOP_Id > 5350 AND Recibo.CFOP_Id < 5360) OR (Recibo.CFOP_Id > 6350 AND Recibo.CFOP_Id < 6360)) " & vbCrLf
            If strCliente(0).Length > 0 Then
                strSQL &= " AND (NotasFiscais.Cliente_Id = '" & strCliente(0) & "')" & vbCrLf & _
                          " AND (NotasFiscais.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
            End If

            If strTransportador(0).Length > 0 Then
                strSQL &= " AND (NFxT.Proprietario = '" & strTransportador(0) & "')" & vbCrLf & _
                          " AND (NFxT.EndProprietario = " & strTransportador(1) & ")" & vbCrLf
            End If

            If ddlGrupoProduto.SelectedIndex > 0 Then strSQL &= " AND (left(NotasFiscaisXItens.Produto_Id,5) = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "') " & vbCrLf
            If ddlProduto.SelectedIndex > 0 Then strSQL &= " AND (NotasFiscaisXItens.Produto_Id = '" & strProduto(0) & "') " & vbCrLf
            If txtPedido.Text.Length > 0 Then
                strSQL &= " AND (NotasFiscais.Pedido = " & txtPedido.Text & ") "
            Else
                strSQL &= " AND (NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf & _
                          "                 AND '" & txtDataFinal.Text.ToSqlDate() & "') "
            End If
            strSQL &= "UNION " & vbCrLf & _
                      "" & vbCrLf & _
                      "SELECT DISTINCT " & vbCrLf & _
                      "       NotasFiscais.Empresa_Id AS Empresa," & vbCrLf & _
                      "       NotasFiscais.EndEmpresa_Id AS EndEmpresa," & vbCrLf & _
                      "       NotasFiscais.Cliente_Id AS Cliente," & vbCrLf & _
                      "       NotasFiscais.EndCliente_Id AS EndCliente," & vbCrLf & _
                      "       Clientes.Nome AS NomeCliente," & vbCrLf & _
                      "       Reciboitens.EntradaSaida_Id AS EntradaSaida," & vbCrLf & _
                      "       Reciboitens.Serie_Id AS Serie," & vbCrLf & _
                      "       Reciboitens.Nota_Id AS Recibo," & vbCrLf & _
                      "       Reciboitens.CFOP_Id AS CFOP," & vbCrLf & _
                      "       Reciboitens.QuantidadeFiscal AS Peso," & vbCrLf & _
                      "       Reciboitens.Unitario," & vbCrLf & _
                      "       Reciboitens.Valor," & vbCrLf & _
                      "       isnull(ReciboXEnc.Valor,0) AS Adiantamento," & vbCrLf & _
                      "       (Reciboitens.Valor - isnull(ReciboXEnc.Valor,0)) as Saldo," & vbCrLf & _
                      "       CONVERT(varchar, Recibo.Movimento, 103) AS Movimento," & vbCrLf & _
                      "       case" & vbCrLf & _
                      "         when ((SELECT count(*)" & vbCrLf & _
                      "                  From NotasXNotas" & vbCrLf & _
                      "                 WHERE Empresa_Id      = Reciboitens.Empresa_Id" & vbCrLf & _
                      "                   AND EndEmpresa_Id   = Reciboitens.EndEmpresa_Id" & vbCrLf & _
                      "                   AND Cliente_Id      = Reciboitens.Cliente_Id" & vbCrLf & _
                      "                   AND EndCliente_Id   = Reciboitens.EndCliente_Id" & vbCrLf & _
                      "                   AND EntradaSaida_Id = Reciboitens.EntradaSaida_Id" & vbCrLf & _
                      "                   AND Serie_Id        = Reciboitens.Serie_Id" & vbCrLf & _
                      "                   AND Nota_Id         = Reciboitens.Nota_Id )" & vbCrLf & _
                      "                ) > 1" & vbCrLf & _
                      "             then 0" & vbCrLf & _
                      "             else NotasFiscais.Pedido" & vbCrLf & _
                      "       end as Pedido," & vbCrLf & _
                      "       case" & vbCrLf & _
                      "         when ((SELECT count(*)" & vbCrLf & _
                      "                  From NotasXNotas" & vbCrLf & _
                      "                 WHERE Empresa_Id      = Reciboitens.Empresa_Id" & vbCrLf & _
                      "                   AND EndEmpresa_Id   = Reciboitens.EndEmpresa_Id" & vbCrLf & _
                      "                   AND Cliente_Id      = Reciboitens.Cliente_Id" & vbCrLf & _
                      "                   AND EndCliente_Id   = Reciboitens.EndCliente_Id" & vbCrLf & _
                      "                   AND EntradaSaida_Id = Reciboitens.EntradaSaida_Id" & vbCrLf & _
                      "                   AND Serie_Id        = Reciboitens.Serie_Id" & vbCrLf & _
                      "                   AND Nota_Id         = Reciboitens.Nota_Id )" & vbCrLf & _
                      "               ) > 1" & vbCrLf & _
                      "            then 'AGRUPADO'" & vbCrLf & _
                      "            else NotasFiscais.EntradaSaida_Id + '-' + CONVERT(varchar, NotasFiscais.Nota_Id) + '-' + NotasFiscais.Serie_Id" & vbCrLf & _
                      "       end as NotaFiscal" & vbCrLf & _
                      "  FROM NotasFiscais" & vbCrLf & _
                      " INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                      "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                      "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                      "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                      "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                      "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                      "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                      "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                      " INNER JOIN NotasXNotas AS NotasXRec" & vbCrLf & _
                      "    ON NotasFiscais.Empresa_Id      = NotasXRec.OrigemEmpresa_Id" & vbCrLf & _
                      "   AND NotasFiscais.EndEmpresa_Id   = NotasXRec.OrigemEndEmpresa_Id" & vbCrLf & _
                      "   AND NotasFiscais.Cliente_Id      = NotasXRec.OrigemCliente_Id" & vbCrLf & _
                      "   AND NotasFiscais.EndCliente_Id   = NotasXRec.OrigemEndCliente_Id" & vbCrLf & _
                      "   AND NotasFiscais.EntradaSaida_Id = NotasXRec.OrigemEntradaSaida_Id" & vbCrLf & _
                      "   AND NotasFiscais.Serie_Id        = NotasXRec.OrigemSerie_Id" & vbCrLf & _
                      "   AND NotasFiscais.Nota_Id         = NotasXRec.OrigemNota_Id" & vbCrLf & _
                      " INNER JOIN NotasFiscaisXItens AS Reciboitens" & vbCrLf & _
                      "    ON NotasXRec.Empresa_Id      = Reciboitens.Empresa_Id" & vbCrLf & _
                      "   AND NotasXRec.EndEmpresa_Id   = Reciboitens.EndEmpresa_Id" & vbCrLf & _
                      "   AND NotasXRec.Cliente_Id      = Reciboitens.Cliente_Id" & vbCrLf & _
                      "   AND NotasXRec.EndCliente_Id   = Reciboitens.EndCliente_Id" & vbCrLf & _
                      "   AND NotasXRec.EntradaSaida_Id = Reciboitens.EntradaSaida_Id" & vbCrLf & _
                      "   AND NotasXRec.Serie_Id        = 'REC'" & vbCrLf & _
                      "   AND NotasXRec.Nota_Id         = Reciboitens.Nota_Id" & vbCrLf & _
                      " LEFT JOIN NotasFiscaisXEncargos AS ReciboXEnc " & vbCrLf & _
                      "    ON ReciboXEnc.Empresa_Id      = Reciboitens.Empresa_Id " & vbCrLf & _
                      "   AND ReciboXEnc.EndEmpresa_Id   = Reciboitens.EndEmpresa_Id " & vbCrLf & _
                      "   AND ReciboXEnc.Cliente_Id      = Reciboitens.Cliente_Id " & vbCrLf & _
                      "   AND ReciboXEnc.EndCliente_Id   = Reciboitens.EndCliente_Id " & vbCrLf & _
                      "   AND ReciboXEnc.EntradaSaida_Id = Reciboitens.EntradaSaida_Id " & vbCrLf & _
                      "   AND ReciboXEnc.Serie_Id        = Reciboitens.Serie_Id " & vbCrLf & _
                      "   AND ReciboXEnc.Nota_Id         = Reciboitens.Nota_Id " & vbCrLf & _
                      "   AND ReciboXEnc.Produto_Id      = Reciboitens.Produto_Id " & vbCrLf & _
                      "   AND ReciboXEnc.CFOP_Id         = Reciboitens.CFOP_Id " & vbCrLf & _
                      "   AND ReciboXEnc.Encargo_Id      = 'ADTODEFRETE' " & vbCrLf & _
                      " inner join Notasfiscais as Recibo" & vbCrLf & _
                      "    ON Recibo.Empresa_Id      = Reciboitens.Empresa_Id" & vbCrLf & _
                      "   AND Recibo.EndEmpresa_Id   = Reciboitens.EndEmpresa_Id" & vbCrLf & _
                      "   AND Recibo.Cliente_Id      = Reciboitens.Cliente_Id" & vbCrLf & _
                      "   AND Recibo.EndCliente_Id   = Reciboitens.EndCliente_Id" & vbCrLf & _
                      "   AND Recibo.EntradaSaida_Id = Reciboitens.EntradaSaida_Id" & vbCrLf & _
                      "   AND Recibo.Serie_Id        = Reciboitens.Serie_id" & vbCrLf & _
                      "   AND Recibo.Nota_Id         = Reciboitens.Nota_Id" & vbCrLf & _
                      " INNER JOIN NotasFiscais AS ReciboN" & vbCrLf & _
                      "    ON NotasXRec.Empresa_Id      = ReciboN.Empresa_Id " & vbCrLf & _
                      "   AND NotasXRec.EndEmpresa_Id   = ReciboN.EndEmpresa_Id" & vbCrLf & _
                      "   AND NotasXRec.Cliente_Id      = ReciboN.Cliente_Id" & vbCrLf & _
                      "   AND NotasXRec.EndCliente_Id   = ReciboN.EndCliente_Id" & vbCrLf & _
                      "   AND NotasXRec.EntradaSaida_Id = ReciboN.EntradaSaida_Id " & vbCrLf & _
                      "   AND NotasXRec.Serie_Id        = 'REC'" & vbCrLf & _
                      "   AND NotasXRec.Nota_Id         = ReciboN.Nota_Id" & vbCrLf & _
                      " INNER JOIN NotasFiscaisXTransportadores NFxT" & vbCrLf & _
                      "    ON  Reciboitens.Empresa_Id     = NFxT.Empresa_Id" & vbCrLf & _
                      "   AND Reciboitens.EndEmpresa_Id   = NFxT.EndEmpresa_Id" & vbCrLf & _
                      "   AND Reciboitens.Cliente_Id      = NFxT.Cliente_Id" & vbCrLf & _
                      "   AND Reciboitens.EndCliente_Id   = NFxT.EndCliente_Id" & vbCrLf & _
                      "   AND Reciboitens.EntradaSaida_Id = NFxT.EntradaSaida_Id" & vbCrLf & _
                      "   AND Reciboitens.Serie_Id        = NFxT.Serie_Id" & vbCrLf & _
                      "   AND Reciboitens.Nota_Id         = NFxT.Nota_Id" & vbCrLf & _
                      " INNER JOIN Clientes" & vbCrLf & _
                      "    ON NotasFiscais.Cliente_Id    = Clientes.Cliente_Id" & vbCrLf & _
                      "   AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                      " WHERE NotasFiscais.Empresa_Id = '" & strEmpresa(0) & "'" & vbCrLf & _
                      "   AND NotasFiscais.EndEmpresa_Id = '" & strEmpresa(1) & "'" & vbCrLf & _
                      "   AND Reciboitens.Serie_id            = 'REC'" & vbCrLf & _
                      "   AND ((Reciboitens.CFOP_Id > 5350 AND Reciboitens.CFOP_Id < 5360) OR (Reciboitens.CFOP_Id > 6350 AND Reciboitens.CFOP_Id < 6360) or (Reciboitens.CFOP_Id > 1350 AND Reciboitens.CFOP_Id < 1360) OR (Reciboitens.CFOP_Id > 2350 AND Reciboitens.CFOP_Id < 2360)) " & vbCrLf
            If strCliente(0).Length > 0 Then
                strSQL &= " AND (NotasFiscais.Cliente_Id = '" & strCliente(0) & "')" & vbCrLf & _
                          " AND (NotasFiscais.EndCliente_Id = " & strCliente(1) & ")" & vbCrLf
            End If

            If strTransportador(0).Length > 0 Then
                strSQL &= " AND (NFxT.Proprietario = '" & strTransportador(0) & "')" & vbCrLf & _
                          " AND (NFxT.EndProprietario = " & strTransportador(1) & ")" & vbCrLf
            End If

            If ddlGrupoProduto.SelectedIndex > 0 Then strSQL &= " AND (left(NotasFiscaisXItens.Produto_Id,5) = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "') " & vbCrLf
            If ddlProduto.SelectedIndex > 0 Then strSQL &= " AND (NotasFiscaisXItens.Produto_Id = '" & strProduto(0) & "') " & vbCrLf
            If txtPedido.Text.Length > 0 Then
                strSQL &= " AND (NotasFiscais.Pedido = " & txtPedido.Text & ") "
            Else
                strSQL &= " AND (NotasFiscais.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf & _
                          "                 AND '" & txtDataFinal.Text.ToSqlDate() & "') "
            End If
        Else
            strSQL &= "SELECT DISTINCT " & vbCrLf & _
                      "       Ctrc.Empresa_Id AS Empresa," & vbCrLf & _
                      "       Ctrc.EndEmpresa_Id AS EndEmpresa," & vbCrLf & _
                      "       Ctrc.Cliente_Id AS Cliente," & vbCrLf & _
                      "       Ctrc.EndCliente_Id AS EndCliente," & vbCrLf & _
                      "       Clientes.Nome AS NomeCliente," & vbCrLf & _
                      "       Recibo.EntradaSaida_Id AS EntradaSaida," & vbCrLf & _
                      "       Recibo.Serie_Id AS Serie," & vbCrLf & _
                      "       Recibo.Nota_Id AS Recibo," & vbCrLf & _
                      "       CtrcXI.CFOP_Id AS CFOP," & vbCrLf & _
                      "       CtrcXI.QuantidadeFiscal AS Peso," & vbCrLf & _
                      "       CtrcXI.Unitario," & vbCrLf & _
                      "       CtrcXI.Valor," & vbCrLf & _
                      "       0 AS Adiantamento," & vbCrLf & _
                      "       0 AS Saldo," & vbCrLf & _
                      "       CONVERT(varchar, Ctrc.Movimento, 103) AS Movimento," & vbCrLf & _
                      "       Ctrc.Pedido," & vbCrLf & _
                      "       Ctrc.EntradaSaida_Id + '-' + CONVERT(varchar, Ctrc.Nota_Id) + '-' + Ctrc.Serie_Id AS NotaFiscal" & vbCrLf & _
                      "  FROM NotasFiscais as Ctrc" & vbCrLf & _
                      " INNER JOIN NotasFiscaisXItens as CtrcXI" & vbCrLf & _
                      "    ON Ctrc.Empresa_Id      = CtrcXI.Empresa_Id" & vbCrLf & _
                      "   AND Ctrc.EndEmpresa_Id   = CtrcXI.EndEmpresa_Id" & vbCrLf & _
                      "   AND Ctrc.Cliente_Id      = CtrcXI.Cliente_Id" & vbCrLf & _
                      "   AND Ctrc.EndCliente_Id   = CtrcXI.EndCliente_Id" & vbCrLf & _
                      "   AND Ctrc.EntradaSaida_Id = CtrcXI.EntradaSaida_Id" & vbCrLf & _
                      "   AND Ctrc.Serie_Id        = CtrcXI.Serie_Id" & vbCrLf & _
                      "   AND Ctrc.Nota_Id         = CtrcXI.Nota_Id" & vbCrLf & _
                      "  LEFT JOIN NotasFiscaisXTransportadores NFxT" & vbCrLf & _
                      "    ON NFxT.Empresa_Id      = Ctrc.Empresa_Id " & vbCrLf & _
                      "   AND NFxT.EndEmpresa_Id   = Ctrc.EndEmpresa_Id" & vbCrLf & _
                      "   AND NFxT.Cliente_Id      = Ctrc.Cliente_Id" & vbCrLf & _
                      "   AND NFxT.EndCliente_Id   = Ctrc.EndCliente_Id" & vbCrLf & _
                      "   AND NFxT.EntradaSaida_Id = Ctrc.EntradaSaida_Id" & vbCrLf & _
                      "   AND NFxT.Serie_Id        = Ctrc.Serie_Id" & vbCrLf & _
                      "   AND NFxT.Nota_Id         = Ctrc.Nota_Id" & vbCrLf & _
                      "  LEFT JOIN NotasXNotas AS NotasXRec" & vbCrLf & _
                      "    ON Ctrc.Empresa_Id      = NotasXRec.OrigemEmpresa_Id " & vbCrLf & _
                      "   AND Ctrc.EndEmpresa_Id   = NotasXRec.OrigemEndEmpresa_Id" & vbCrLf & _
                      "   AND Ctrc.Cliente_Id      = NotasXRec.OrigemCliente_Id" & vbCrLf & _
                      "   AND Ctrc.EndCliente_Id   = NotasXRec.OrigemEndCliente_Id" & vbCrLf & _
                      "   AND Ctrc.EntradaSaida_Id = NotasXRec.OrigemEntradaSaida_Id" & vbCrLf & _
                      "   AND Ctrc.Serie_Id        = NotasXRec.OrigemSerie_Id" & vbCrLf & _
                      "   AND Ctrc.Nota_Id         = NotasXRec.OrigemNota_Id" & vbCrLf & _
                      "  LEFT JOIN NotasFiscaisXItens AS Recibo" & vbCrLf & _
                      "    ON NotasXRec.Empresa_Id      = Recibo.Empresa_Id" & vbCrLf & _
                      "   AND NotasXRec.EndEmpresa_Id   = Recibo.EndEmpresa_Id" & vbCrLf & _
                      "   AND NotasXRec.Cliente_Id      = Recibo.Cliente_Id" & vbCrLf & _
                      "   AND NotasXRec.EndCliente_Id   = Recibo.EndCliente_Id" & vbCrLf & _
                      "   AND NotasXRec.EntradaSaida_Id = Recibo.EntradaSaida_Id" & vbCrLf & _
                      "   AND NotasXRec.Serie_Id        = 'UN'" & vbCrLf & _
                      "   AND NotasXRec.Nota_Id         = Recibo.Nota_Id" & vbCrLf & _
                      " INNER JOIN Clientes" & vbCrLf & _
                      "    ON Ctrc.Cliente_Id = Clientes.Cliente_Id" & vbCrLf & _
                      "   AND Ctrc.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                      " WHERE(Ctrc.Empresa_Id    ='" & strEmpresa(0) & "')" & vbCrLf & _
                      "   AND(Ctrc.EndEmpresa_Id ='" & strEmpresa(1) & "')" & vbCrLf & _
                      "   AND((CtrcXI.CFOP_Id > 5350 AND CtrcXI.CFOP_Id < 5360) OR (CtrcXI.CFOP_Id > 6350 AND CtrcXI.CFOP_Id < 6360) or (CtrcXI.CFOP_Id > 1350 AND CtrcXI.CFOP_Id < 1360) OR (CtrcXI.CFOP_Id > 2350 AND CtrcXI.CFOP_Id < 2360))  " & vbCrLf

            If strCliente(0).Length > 0 Then
                strSQL &= " AND (Ctrc.Destino = '" & strCliente(0) & "')" & vbCrLf & _
                          " AND (Ctrc.EndDestino = " & strCliente(1) & ")" & vbCrLf
            End If

            If strTransportador(0).Length > 0 Then
                strSQL &= " AND (Ctrc.Cliente_Id = '" & strTransportador(0) & "')" & vbCrLf & _
                          " AND (Ctrc.EndCliente_Id = " & strTransportador(1) & ")" & vbCrLf
            End If

            If ddlGrupoProduto.SelectedIndex > 0 Then strSQL &= " AND (left(CtrcXI.Produto_Id,5) = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "') " & vbCrLf
            If ddlProduto.SelectedIndex > 0 Then strSQL &= " AND (CtrcXI.Produto_Id = '" & strProduto(0) & "') " & vbCrLf
            If txtPedido.Text.Length > 0 Then
                strSQL &= " AND (Ctrc.Pedido = " & txtPedido.Text & ") "
            Else
                strSQL &= " AND (Ctrc.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') "
            End If
        End If

        ds = Banco.ConsultaDataSet(strSQL, "RelatorioDeFretes")

        For Each dr As DataRow In ds.Tables(0).Rows
            Cliente = Funcoes.FormatarCpfCnpj(dr("Cliente"))
            dr("Cliente") = Cliente
        Next

        Return ds
    End Function

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddlProduto.Items.Clear()
            txtPedido.Enabled = False
            btnPedido.Enabled = False
            BuscarProduto()
            ddlProduto.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub BtnTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnTransportador.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteFRTTrans" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioDeFretes", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim sql As String = ""
                    Dim DescricaoProduto As String = ""
                    Dim dsProduto As New DataSet
                    Dim ds As New DataSet
                    If ddlProduto.SelectedIndex > 0 Then
                        Dim strProduto() As String = ddlProduto.SelectedValue.ToString.Split("_")
                        sql &= "SELECT * FROM PRODUTOS WHERE PRODUTO_ID = '" & strProduto(0) & "'"
                        dsProduto = Banco.ConsultaDataSet(sql, "Produto")
                        DescricaoProduto = " - PRODUTO: " & Trim(dsProduto.Tables(0).Rows(0).Item("Nome")).ToString
                    ElseIf ddlGrupoProduto.SelectedIndex > 0 Then
                        sql &= "SELECT * FROM GRUPOSDEESTOQUES WHERE GRUPO_ID = '" & Mid(ddlGrupoProduto.SelectedValue, 1, 5) & "'"
                        dsProduto = Banco.ConsultaDataSet(sql, "Produto")
                        DescricaoProduto = " - GRUPO: " & Trim(dsProduto.Tables(0).Rows(0).Item("Descricao")).ToString
                    End If


                    Dim strEmpresa() As String = cmbEmpresa.SelectedValue.Split(New Char() {"-"})

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Nome", HttpContext.Current.Session("ssNomeEmpresa"))
                    parameters.Add("Cidade", Trim(HttpContext.Current.Session("ssCidadeEmpresa")) & "/" & HttpContext.Current.Session("ssEstadoEmpresa"))
                    parameters.Add("Cliente", IIf(txtClientes.Text.Length > 0, "Cliente: " & txtClientes.Text, ""))
                    parameters.Add("Transportador", IIf(txtTransportador.Text.Length > 0, "Transportador: " & txtTransportador.Text, ""))
                    parameters.Add("Cnpj", "CNPJ: " & Funcoes.FormatarCpfCnpj(strEmpresa(0)))
                    Dim texto As String = ""
                    If rdRecibo.Checked = True Then
                        texto = "Relatório de Recibos de Fretes - Periodo: " & txtDataInicial.Text & " à " & txtDataFinal.Text
                    Else
                        texto = "Relatório de CTRC de Transporte - Periodo: " & txtDataInicial.Text & " à " & txtDataFinal.Text
                    End If
                    parameters.Add("Titulo", texto & DescricaoProduto)

                    ds = BuscarRegistros()

                    Funcoes.BindReport(Me.Page, ds, "Cr_RelatorioDeFretes", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para listar relatório.")
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeFretes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class