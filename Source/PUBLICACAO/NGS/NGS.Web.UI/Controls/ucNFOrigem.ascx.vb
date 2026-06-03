Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucNFOrigem
    Inherits BaseUserControl

    Private objNotaFiscal As [Lib].Negocio.NotaFiscal

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" & HID.Value) = objNotaFiscal
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

        End If
    End Sub

    Protected Sub cmdConsultaClienteFrete_Click(sender As Object, e As EventArgs) Handles cmdClienteFrete.Click
        Dim ucConsultaClientes As ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)
        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            ucConsultaClientes.MainUserControl = Me
            Popup.ConsultaDeClientes(Me.Page, "objClienteNFGFr" & HID.Value)
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Consultar()
    End Sub

    Protected Sub lnkGravar_Click(sender As Object, e As EventArgs) Handles lnkGravar.Click
        Dim temNota As Boolean = False

        For Each row As GridViewRow In grdNotaFreteConsulta.Rows
            Dim chk As CheckBox = CType(row.FindControl("chkNotaXNota"), CheckBox)
            If chk IsNot Nothing AndAlso chk.Checked Then
                SessaoRecuperaNotaFiscal()
                Dim nf As New [Lib].Negocio.NotaFiscal()
                nf.CodigoEmpresa = objNotaFiscal.CodigoEmpresa
                nf.EnderecoEmpresa = objNotaFiscal.EnderecoEmpresa
                nf.CodigoCliente = row.Cells(1).Text.Split("-")(0)
                nf.EnderecoCliente = row.Cells(1).Text.Split("-")(1)
                nf.EntradaSaida = IIf(row.Cells(5).Text.Split("-")(0) = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                nf.Codigo = row.Cells(5).Text.Split("-")(1)
                nf.Serie = row.Cells(5).Text.Split("-")(2)
                nf = New [Lib].Negocio.NotaFiscal(nf)
                If Not (objNotaFiscal.NotasTrocaOrigem IsNot Nothing AndAlso objNotaFiscal.NotasTrocaOrigem.Count > 0) Then
                    objNotaFiscal.NotasTrocaOrigem = New List(Of [Lib].Negocio.NotaFiscal)
                End If
                objNotaFiscal.NotasTrocaOrigem.Add(nf)
                SessaoSalvaNotaFiscal()
                temNota = True
            End If
        Next

        If temNota Then
            If TypeOf Me.Page Is NotasFiscaisGerais Then
                CType(Me.Page, NotasFiscaisGerais).CarregarNotasDeOrigem()
            End If
            Limpar()
            Popup.CloseDialog(Me.Page, "divNFOrigem")
        Else
            MsgBox(Me.Page, "Não foi selecionada nenhuma Nota Fiscal no Grid")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divNFOrigem")
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objClienteNFGFr" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(obj, [Lib].Negocio.Cliente))
            txtCodigoClienteFrete.Value = itemCliente.Value
            txtClienteFrete.Text = itemCliente.Text
            Session.Remove("objClienteNFGFr" & HID.Value)
        End If
    End Sub

    Public Overrides Sub Limpar()
        txtCodigoClienteFrete.Value = String.Empty
        txtClienteFrete.Text = String.Empty
        txtDataIniFrete.Text = String.Empty
        txtDataFimFrete.Text = String.Empty
        txtNumNFFrete.Text = String.Empty
        txtSerieNFFrete.Text = String.Empty
        grdNotaFreteConsulta.DataSource = New List(Of Object)
        grdNotaFreteConsulta.DataBind()
    End Sub

    Private Sub Consultar()
        SessaoRecuperaNotaFiscal()
        If ValidarCamposFrete() Then
            Dim strSQL As String = "SELECT (NF.Cliente_Id + '-' + CAST(NF.EndCliente_Id AS varchar) + '-' + C.Nome) AS Cliente_ID, " & vbCrLf & _
                                   "       case " & vbCrLf & _
                                   "          when NF.TipoDeDocumento = 2 " & vbCrLf & _
                                   "             then case " & vbCrLf & _
                                   "       			    when isnull(NxI.Deposito,'') = '' " & vbCrLf & _
                                   "       				  then NF.Deposito + '-' + CAST(NF.EndDeposito AS varchar)" & vbCrLf & _
                                   "       				  else NxI.Deposito + '-' + CAST(NxI.EndDeposito AS varchar)" & vbCrLf & _
                                   "       			  end " & vbCrLf & _
                                   "             else (NF.Deposito + '-' + CAST(NF.EndDeposito AS varchar)) " & vbCrLf & _
                                   "        end AS Deposito," & vbCrLf & _
                                   "        NF.Destino + '-' + CAST(NF.EndDestino AS varchar) AS Destino, " & vbCrLf & _
                                   "        NF.EntradaSaida_Id + '-' + CAST(NF.Nota_Id AS varchar) + '-' + NF.Serie_Id AS NumNota, " & vbCrLf & _
                                   "        CAST(NF.Operacao AS varchar) + '-' + CAST(NF.SubOperacao AS varchar) AS Operacao," & vbCrLf & _
                                   "        NF.Movimento, " & vbCrLf & _
                                   "        NF.Cliente_Id + ';' + CAST(NF.EndCliente_Id AS varchar) + ';' + NF.EntradaSaida_Id + ';' + NF.Serie_Id + ';' + CAST(NF.Nota_Id AS varchar) AS Chave, " & vbCrLf & _
                                   "        (convert(varchar,NF.TipoDeDocumento) + '-' + TD.Descricao) AS Tipo," & vbCrLf & _
                                   "        NxI.Quantidade," & vbCrLf & _
                                   "        NxI.Valor " & vbCrLf & _
                                   "   FROM NotasFiscais NF " & vbCrLf & _
                                   "  INNER JOIN (SELECT NFI.Empresa_Id, NFI.EndEmpresa_Id," & vbCrLf & _
                                   "                     NFI.Cliente_Id, NFI.EndCliente_Id, " & vbCrLf & _
                                   "				     NFI.EntradaSaida_Id, NFI.Serie_Id, NFI.Nota_Id, " & vbCrLf & _
                                   "                     NxN.OrigemEmpresa_Id AS Deposito, NxN.OrigemEndEmpresa_Id As EndDeposito, " & vbCrLf & _
                                   "				     sum(NFI.QuantidadeFiscal) as Quantidade," & vbCrLf & _
                                   "                     sum(NFI.Valor) as Valor " & vbCrLf & _
                                   "			    FROM NotasFiscaisXItens as NFI " & vbCrLf & _
                                   "				LEFT JOIN NotasXNotas NxN " & vbCrLf & _
                                   "				  ON NxN.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                                   "				 AND NxN.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                                   "				 AND NxN.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                                   "				 AND NxN.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                                   "				 AND NxN.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                                   "				 AND NxN.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
                                   "				 AND NxN.Nota_Id         = NFI.Nota_Id " & vbCrLf & _
                                   "				LEFT JOIN NotasXNotas NxNO " & vbCrLf & _
                                   "				  ON NxNO.OrigemEmpresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                                   "				 AND NxNO.OrigemEndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                                   "				 AND NxNO.OrigemCliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                                   "				 AND NxNO.OrigemEndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                                   "				 AND NxNO.OrigemEntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                                   "				 AND NxNO.OrigemSerie_Id        = NFI.Serie_Id " & vbCrLf & _
                                   "				 AND NxNO.OrigemNota_Id         = NFI.Nota_Id " & vbCrLf & _
                                   "                 AND NxNO.Serie_id              <> 'REC' " & vbCrLf & _
                                   "				LEFT JOIN NotasFiscais NxNOXNF " & vbCrLf & _
                                   "				  ON  NxNOXNF.Empresa_Id     = NxNO.Empresa_Id " & vbCrLf & _
                                   "				 AND NxNOXNF.EndEmpresa_Id   = NxNO.EndEmpresa_Id " & vbCrLf & _
                                   "			     AND NxNOXNF.Cliente_Id      = NxNO.Cliente_Id " & vbCrLf & _
                                   "				 AND NxNOXNF.EndCliente_Id   = NxNO.EndCliente_Id " & vbCrLf & _
                                   "				 AND NxNOXNF.EntradaSaida_Id = NxNO.EntradaSaida_Id " & vbCrLf & _
                                   "				 AND NxNOXNF.Serie_Id        = NxNO.Serie_Id " & vbCrLf & _
                                   "				 AND NxNOXNF.Nota_Id         = NxNO.Nota_Id " & vbCrLf & _
                                   "		       WHERE NFI.Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                                   "			     AND NFI.EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf

            If Not String.IsNullOrWhiteSpace(txtCodigoClienteFrete.Value) Then
                Dim strCliente As String() = txtCodigoClienteFrete.Value.Split("-")
                strSQL &= "			  AND NFI.Cliente_Id = '" & strCliente(0) & "' " & vbCrLf & _
                          "			  AND NFI.EndCliente_Id = " & strCliente(1) & " " & vbCrLf
            End If

            If txtNumNFFrete.Text.Length > 0 Then

                If objNotaFiscal IsNot Nothing AndAlso (objNotaFiscal.CodigoTipoDeDocumento = 2 Or objNotaFiscal.CodigoTipoDeDocumento = 57) Then
                    If txtNumNFFrete.Text.Length > 0 Then strSQL &= "			  AND NFI.Nota_Id in (" & txtNumNFFrete.Text & ")" & vbCrLf
                Else
                    If txtNumNFFrete.Text.Length > 0 Then strSQL &= "			  AND NFI.Nota_Id = " & txtNumNFFrete.Text & vbCrLf
                End If

            End If

            If txtSerieNFFrete.Text.Length > 0 Then strSQL &= "			  AND NFI.Serie_Id = '" & txtSerieNFFrete.Text.ToUpper() & "' " & vbCrLf

            If rdEntradas.Checked Then
                strSQL &= "			  AND NFI.EntradaSaida_Id = 'E' " & vbCrLf
            End If

            If rdSaidas.Checked Then
                strSQL &= "			  AND NFI.EntradaSaida_Id = 'S' " & vbCrLf
            End If

            If objNotaFiscal IsNot Nothing AndAlso (objNotaFiscal.CodigoTipoDeDocumento = 2 Or objNotaFiscal.CodigoTipoDeDocumento = 57) Then
                'strSQL &= "		       group by NFI.Empresa_Id, NFI.EndEmpresa_Id, NFI.Cliente_Id, NFI.EndCliente_Id, " & vbCrLf & _
                '          "				        NFI.EntradaSaida_Id, NFI.Serie_Id, NFI.Nota_Id, NxN.OrigemEmpresa_Id, NxN.OrigemEndEmpresa_Id" & vbCrLf & _
                '          "				having  sum(case" & vbCrLf & _
                '          "				               when isnull(NxNOXNF.TipoDeDocumento,0) in(2,3,4,9,10,11)" & vbCrLf & _
                '          "				                then 1" & vbCrLf & _
                '          "				                else 0" & vbCrLf & _
                '          "				               end) = 0) AS NxI " & vbCrLf
                strSQL &= "		       group by NFI.Empresa_Id, NFI.EndEmpresa_Id, NFI.Cliente_Id, NFI.EndCliente_Id, " & vbCrLf & _
                          "				        NFI.EntradaSaida_Id, NFI.Serie_Id, NFI.Nota_Id, NxN.OrigemEmpresa_Id, NxN.OrigemEndEmpresa_Id) AS NxI " & vbCrLf
            ElseIf objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.CodigoTipoDeDocumento = 9 Then
                strSQL &= "		       group by NFI.Empresa_Id, NFI.EndEmpresa_Id, NFI.Cliente_Id, NFI.EndCliente_Id, " & vbCrLf & _
                          "				        NFI.EntradaSaida_Id, NFI.Serie_Id, NFI.Nota_Id, NxN.OrigemEmpresa_Id, NxN.OrigemEndEmpresa_Id" & vbCrLf & _
                          "				having  sum(case" & vbCrLf & _
                          "				               when isnull(NxNOXNF.TipoDeDocumento,0) in(2,3,4,9,10,11)" & vbCrLf & _
                          "				                then 1" & vbCrLf & _
                          "				                else 0" & vbCrLf & _
                          "				               end) = 0) AS NxI " & vbCrLf
            ElseIf objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.CodigoTipoDeDocumento = 10 Then
                strSQL &= "		       group by NFI.Empresa_Id, NFI.EndEmpresa_Id, NFI.Cliente_Id, NFI.EndCliente_Id, " & vbCrLf & _
                          "				        NFI.EntradaSaida_Id, NFI.Serie_Id, NFI.Nota_Id, NxN.OrigemEmpresa_Id, NxN.OrigemEndEmpresa_Id" & vbCrLf & _
                          "				having  sum(case" & vbCrLf & _
                          "				               when isnull(NxNOXNF.TipoDeDocumento,0) in(9,10,11)" & vbCrLf & _
                          "				                then 1" & vbCrLf & _
                          "				                else 0" & vbCrLf & _
                          "				               end) = 0) AS NxI " & vbCrLf
            Else
                strSQL &= "			  AND isnull(NxNO.Nota_Id,0) = 0 " & vbCrLf
                strSQL &= "		       group by NFI.Empresa_Id, NFI.EndEmpresa_Id, NFI.Cliente_Id, NFI.EndCliente_Id, " & vbCrLf & _
                          "				        NFI.EntradaSaida_Id, NFI.Serie_Id, NFI.Nota_Id, NxN.OrigemEmpresa_Id, NxN.OrigemEndEmpresa_Id" & vbCrLf & _
                          "           ) AS NxI " & vbCrLf
            End If

            strSQL &= "		   ON NF.Empresa_Id      = NxI.Empresa_Id " & vbCrLf & _
                      "		  AND NF.EndEmpresa_Id   = NxI.EndEmpresa_Id " & vbCrLf & _
                      "	      AND NF.Cliente_Id      = NxI.Cliente_Id " & vbCrLf & _
                      "		  AND NF.EndCliente_Id   = NxI.EndCliente_Id " & vbCrLf & _
                      "		  AND NF.EntradaSaida_Id = NxI.EntradaSaida_Id " & vbCrLf & _
                      "		  AND NF.Serie_Id        = NxI.Serie_Id " & vbCrLf & _
                      "		  AND NF.Nota_Id         = NxI.Nota_Id " & vbCrLf & _
                      "     INNER JOIN TipoDeDocumento TD " & vbCrLf & _
                      "        ON TD.Codigo_Id = NF.TipoDeDocumento " & vbCrLf & _
                      "     INNER JOIN SubOperacoes SO " & vbCrLf & _
                      "        ON SO.Operacao_Id = NF.Operacao " & vbCrLf & _
                      "       AND SO.SubOperacoes_Id = NF.SubOperacao " & vbCrLf & _
                      "     INNER JOIN Clientes C " & vbCrLf & _
                      "        ON C.Cliente_Id = NF.Cliente_Id " & vbCrLf & _
                      "       AND C.Endereco_Id = NF.EndCliente_Id " & vbCrLf & _
                      "     WHERE NF.Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf & _
                      "       AND NF.EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf

            If objNotaFiscal IsNot Nothing AndAlso Not objNotaFiscal.CodigoTipoDeDocumento = 14 Then
                strSQL &= "  AND SO.Classe <> '" & eClassesOperacoes.FRETES.ToString & "' " & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(txtCodigoClienteFrete.Value) Then
                Dim strCliente As String() = txtCodigoClienteFrete.Value.Split("-")

                strSQL &= "  AND NF.Cliente_Id = '" & strCliente(0) & "' " & vbCrLf & _
                          "  AND NF.EndCliente_Id = " & strCliente(1) & " " & vbCrLf
            End If

            If txtDataIniFrete.Text.Replace("_", "").Replace("/", "").Length > 0 Then
                Dim strDataIni As String() = txtDataIniFrete.Text.Split("/")
                Dim dateDataIni As New DateTime(strDataIni(2), strDataIni(1), strDataIni(0))

                If txtDataFimFrete.Text.Replace("_", "").Replace("/", "").Length > 0 Then
                    Dim strDataFim As String() = txtDataFimFrete.Text.Split("/")
                    Dim dateDataFim As New DateTime(strDataFim(2), strDataFim(1), strDataFim(0))

                    strSQL &= "  AND NF.Movimento BETWEEN '" & dateDataIni.ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                              "  AND '" & dateDataFim.ToString("yyyy-MM-dd") & "' " & vbCrLf
                Else
                    strSQL &= "  AND NF.Movimento = '" & dateDataIni.ToString("yyyy-MM-dd") & vbCrLf & "' "
                End If
            End If


            If objNotaFiscal IsNot Nothing AndAlso (objNotaFiscal.CodigoTipoDeDocumento = 2 Or objNotaFiscal.CodigoTipoDeDocumento = 57) Then
                If txtNumNFFrete.Text.Length > 0 Then strSQL &= "  AND NF.Nota_Id in (" & txtNumNFFrete.Text & ")" & vbCrLf
            Else
                If txtNumNFFrete.Text.Length > 0 Then strSQL &= "  AND NF.Nota_Id = " & txtNumNFFrete.Text & vbCrLf
            End If

            If txtSerieNFFrete.Text.Length > 0 Then strSQL &= "  AND NF.Serie_Id = '" & txtSerieNFFrete.Text.ToUpper() & "' " & vbCrLf

            If rdEntradas.Checked Then
                strSQL &= "			  AND NF.EntradaSaida_Id = 'E' " & vbCrLf
            End If

            If rdSaidas.Checked Then
                strSQL &= "			  AND NF.EntradaSaida_Id = 'S' " & vbCrLf
            End If

            If objNotaFiscal IsNot Nothing AndAlso _
                (objNotaFiscal.CodigoTipoDeDocumento = 2 OrElse objNotaFiscal.CodigoTipoDeDocumento = 3 OrElse objNotaFiscal.CodigoTipoDeDocumento = 57) Then
                strSQL &= "  AND ((NF.EntradaSaida_Id = 'E' AND NF.CIFFOB = 'FOB') OR (NF.EntradaSaida_Id = 'S' AND NF.CIFFOB = 'CIF')) " & vbCrLf
            End If

            If objNotaFiscal IsNot Nothing AndAlso (objNotaFiscal.CodigoTipoDeDocumento = 2 Or objNotaFiscal.CodigoTipoDeDocumento = 57) Then
                strSQL &= "  AND NOT EXISTS (SELECT isnull(NxNOXNF.TipoDeDocumento,0) AS TipoDeDocumento" & vbCrLf & _
                            "			    FROM NotasFiscaisXItens as NFI " & vbCrLf & _
                            "				LEFT JOIN NotasXNotas NxN " & vbCrLf & _
                            "				  ON NxN.Empresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                            "				 AND NxN.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                            "				 AND NxN.Cliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                            "				 AND NxN.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                            "				 AND NxN.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                            "				 AND NxN.Serie_Id        = NFI.Serie_Id " & vbCrLf & _
                            "				 AND NxN.Nota_Id         = NFI.Nota_Id " & vbCrLf & _
                            "				LEFT JOIN NotasXNotas NxNO " & vbCrLf & _
                            "				  ON NxNO.OrigemEmpresa_Id      = NFI.Empresa_Id " & vbCrLf & _
                            "				 AND NxNO.OrigemEndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf & _
                            "				 AND NxNO.OrigemCliente_Id      = NFI.Cliente_Id " & vbCrLf & _
                            "				 AND NxNO.OrigemEndCliente_Id   = NFI.EndCliente_Id " & vbCrLf & _
                            "				 AND NxNO.OrigemEntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf & _
                            "				 AND NxNO.OrigemSerie_Id        = NFI.Serie_Id " & vbCrLf & _
                            "				 AND NxNO.OrigemNota_Id         = NFI.Nota_Id " & vbCrLf & _
                            "                 AND NxNO.Serie_id              <> 'REC' " & vbCrLf & _
                            "				LEFT JOIN NotasFiscais NxNOXNF " & vbCrLf & _
                            "				  ON  NxNOXNF.Empresa_Id     = NxNO.Empresa_Id " & vbCrLf & _
                            "				 AND NxNOXNF.EndEmpresa_Id   = NxNO.EndEmpresa_Id " & vbCrLf & _
                            "			     AND NxNOXNF.Cliente_Id      = NxNO.Cliente_Id " & vbCrLf & _
                            "				 AND NxNOXNF.EndCliente_Id   = NxNO.EndCliente_Id " & vbCrLf & _
                            "				 AND NxNOXNF.EntradaSaida_Id = NxNO.EntradaSaida_Id " & vbCrLf & _
                            "				 AND NxNOXNF.Serie_Id        = NxNO.Serie_Id " & vbCrLf & _
                            "				 AND NxNOXNF.Nota_Id         = NxNO.Nota_Id " & vbCrLf & _
                            "		       WHERE NFI.Empresa_Id    ='" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                            "			     AND NFI.EndEmpresa_Id = " & objNotaFiscal.EnderecoEmpresa & vbCrLf & _
                            "			     AND NFI.Nota_Id in (" & txtNumNFFrete.Text & ")" & vbCrLf

                If Not String.IsNullOrWhiteSpace(txtCodigoClienteFrete.Value) Then
                    Dim strCliente As String() = txtCodigoClienteFrete.Value.Split("-")

                    strSQL &= "           AND NFI.Cliente_Id      = '" & strCliente(0) & "' " & vbCrLf & _
                              "           AND NFI.EndCliente_Id   = " & strCliente(1) & " " & vbCrLf
                End If

                If rdEntradas.Checked Then
                    strSQL &= "			  AND NFI.EntradaSaida_Id = 'E' " & vbCrLf
                End If

                If rdSaidas.Checked Then
                    strSQL &= "			  AND NFI.EntradaSaida_Id = 'S' " & vbCrLf
                End If

                strSQL &= "			     AND ISNULL(NxNOXNF.TipoDeDocumento,0) IN(2,57))" & vbCrLf
            End If

            strSQL &= "ORDER BY NF.Serie_Id"

            Dim dsNotasFrete As DataSet = Banco.ConsultaDataSet(strSQL, "NotasFiscais")

            If dsNotasFrete Is Nothing OrElse dsNotasFrete.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nota(s) de origem não encontrada(s)!")
                Exit Sub
            End If

            grdNotaFreteConsulta.DataSource = dsNotasFrete
            grdNotaFreteConsulta.DataBind()
        End If
    End Sub

    Private Function ValidarCamposFrete() As Boolean
        If Not String.IsNullOrWhiteSpace(txtDataIniFrete.Text) AndAlso String.IsNullOrWhiteSpace(txtDataFimFrete.Text) Then
            MsgBox(Me.Page, "Informe a data de fim do frete!")
            Return False
        End If

        If Not String.IsNullOrWhiteSpace(txtDataFimFrete.Text) AndAlso String.IsNullOrWhiteSpace(txtDataIniFrete.Text) Then
            MsgBox(Me.Page, "Informe a data de início do frete!")
            Return False
        End If

        If Not String.IsNullOrWhiteSpace(txtDataIniFrete.Text) AndAlso Not String.IsNullOrWhiteSpace(txtDataFimFrete.Text) AndAlso CDate(txtDataIniFrete.Text) > CDate(txtDataFimFrete.Text) Then
            MsgBox(Me.Page, "A data inicial deve ser menor que a data final do período!")
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtDataIniFrete.Text) AndAlso String.IsNullOrWhiteSpace(txtDataFimFrete.Text) AndAlso String.IsNullOrWhiteSpace(txtNumNFFrete.Text) Then
            MsgBox(Me.Page, "Informe a data de início e fim ou o número da Nota Fiscal para pesquisa!")
            Return False
        End If

        Return True
    End Function

End Class