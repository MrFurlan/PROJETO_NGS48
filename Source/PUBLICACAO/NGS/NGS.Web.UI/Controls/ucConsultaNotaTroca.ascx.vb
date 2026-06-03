Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaNotaTroca
    Inherits BaseUserControl

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Page.IsPostBack Then
            Return
        End If
        txtNumAno.Text = DateTime.Now.Year.ToString
    End Sub

    Public Sub InicializarFormulario(Optional SomenteTroca As Boolean = False, Optional objNotaVinculo As [Lib].Negocio.NotaFiscal = Nothing)

        gridNotas.DataSource = Nothing
        gridNotas.DataBind()

        'Feito para tela de troca de nota
        If SomenteTroca Then
            lnkCompra.Parent.Visible = False
            lnkTransferencia.Parent.Visible = False
            lnkCessaoDeCredito.Parent.Visible = False
            lnkTroca.Parent.Visible = True
            Exit Sub
        End If


        If Not objNotaVinculo Is Nothing Then
            objNotaFiscal = objNotaVinculo
        Else
            SessaoRecuperaNotaFiscal()
        End If

        If objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.VincularNotas Then

            txtNumNota.Text = objNotaFiscal.Codigo
            txtNumAno.Text = objNotaVinculo.DataTermino.Year
            Consultar()
            lnkCompra.Parent.Visible = False
            lnkTransferencia.Parent.Visible = False
            lnkCessaoDeCredito.Parent.Visible = False
            lnkTroca.Parent.Visible = False

        Else

            If objNotaFiscal.SubOperacao.EntradaSaida = eEntradaSaida.Entrada And (objNotaFiscal.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Or objNotaFiscal.Operacao.CodigoClasse = eClassesOperacoes.DEPOSITOS.ToString) Then
                If objNotaFiscal.CodigoFinalidade <> 13 And objNotaFiscal.CodigoFinalidade <> 23 Then
                    lnkCompra.Parent.Visible = True
                    lnkTransferencia.Parent.Visible = True
                Else
                    lnkCompra.Parent.Visible = objNotaFiscal.CodigoFinalidade = 13
                    lnkTransferencia.Parent.Visible = objNotaFiscal.CodigoFinalidade = 23
                End If

                If objNotaFiscal.Operacao.CodigoClasse = eClassesOperacoes.DEPOSITOS.ToString Then
                    lnkCompra.Parent.Visible = False
                    lnkCessaoDeCredito.Parent.Visible = False
                Else
                    lnkCessaoDeCredito.Parent.Visible = True
                End If
                lnkTroca.Parent.Visible = False
            Else
                lnkCompra.Parent.Visible = False
                lnkTransferencia.Parent.Visible = False
                lnkCessaoDeCredito.Parent.Visible = False
                lnkTroca.Parent.Visible = True
            End If

        End If

    End Sub

    Protected Sub gridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim NotaCons As New [Lib].Negocio.NotaFiscal
        NotaCons.CodigoEmpresa = gridNotas.SelectedRow.Cells(2).Text
        NotaCons.EnderecoEmpresa = gridNotas.SelectedRow.Cells(3).Text
        NotaCons.CodigoCliente = gridNotas.SelectedRow.Cells(7).Text
        NotaCons.EnderecoCliente = gridNotas.SelectedRow.Cells(8).Text
        NotaCons.Codigo = gridNotas.SelectedRow.Cells(13).Text
        NotaCons.Serie = gridNotas.SelectedRow.Cells(14).Text
        NotaCons.EntradaSaida = IIf(gridNotas.SelectedRow.Cells(12).Text = "S", [Lib].Negocio.eEntradaSaida.Saida, [Lib].Negocio.eEntradaSaida.Entrada)

        SessaoRecuperaNotaFiscal()

        Dim Nota As New [Lib].Negocio.NotaFiscal(NotaCons)
        Nota.Itens(0).SaldoPedidoFiscal = gridNotas.SelectedRow.Cells(18).Text

        If TypeOf Me.Page Is NotaFiscalXItens Then
            Session("objTrocaDeNotaNXI" & HID.Value.ToString) = Nota
            CType(Me.Page, NotaFiscalXItens).CarregarNotaTroca()
        ElseIf TypeOf Me.Page Is TrocaDeNota Then
            Session("objTrocaDeNota" & HID.Value.ToString) = Nota
            CType(Me.Page, TrocaDeNota).CarregarNotaTroca()
        ElseIf TypeOf Me.Page Is VinculoDeNotaFiscal Then
            Session("objTrocaDeNotaVinculadas" & HID.Value.ToString) = Nota
        End If

        Popup.CloseDialog(Me.Page, "divConsultaNotaTroca")
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub SessaoSalvarNotaFiscal()
        Session("objNotaFiscal" & HID.Value.ToString) = objNotaFiscal
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Protected Sub Consultar()
        If (String.IsNullOrWhiteSpace(txtNumAno.Text)) Then
            MsgBox(Me.Page, "É necessário informar o campo ano!", eTitulo.Info)
            Exit Sub
        End If

        If (objNotaFiscal IsNot Nothing AndAlso objNotaFiscal.Itens IsNot Nothing AndAlso objNotaFiscal.Itens.Count > 0) Then
            'Dim Prod As String = objNotaFiscal.Itens(0).CodigoProduto
            Dim Emp As String = objNotaFiscal.CodigoEmpresa
            Dim EndEmp As String = objNotaFiscal.EnderecoEmpresa
            Dim ES As String = objNotaFiscal.EntradaSaida.ToString.Substring(0, 1)
            Dim ds As DataSet
            Dim sql As String
            'Dim prd As New [Lib].Negocio.Produto(Prod)

            sql = "Select CONVERT(varchar, NF.Movimento, 103) as Movimento, NF.Empresa_Id, NF.EndEmpresa_Id, " & vbCrLf & _
                  "	      Empresa.Nome As EmpresaNome, Empresa.Cidade as EmpresaCidade, Empresa.Estado as EmpresaEstado," & vbCrLf & _
                  "	      NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf & _
                  "	      Clientes.Nome, Clientes.Cidade, Clientes.Estado," & vbCrLf & _
                  "	      NF.EntradaSaida_Id," & vbCrLf & _
                  "	      NF.Serie_Id," & vbCrLf & _
                  "	      NF.Nota_Id," & vbCrLf & _
                  "	      NFxI.Produto_id," & vbCrLf & _
                  "	      P.Nome as NomeProduto," & vbCrLf & _
                  "	      P.Unidade," & vbCrLf & _
                  "       max(NFxI.QuantidadeFisica) AS QtdeNota," & vbCrLf & _
                  "       max(NFxI.QuantidadeFisica) - sum(isnull(NFxIT.QuantidadeFisica,0)) as SaldoTroca" & vbCrLf & _
                  "  from NotasFiscais NF" & vbCrLf & _
                  " Inner Join NotasFiscaisxItens as NFxI" & vbCrLf & _
                  "	   ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
                  "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
                  "   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
                  "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
                  "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
                  "   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
                  "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                  " Inner Join Suboperacoes SO" & vbCrLf & _
                  "    on So.Operacao_Id     = NFxI.Operacao" & vbCrLf & _
                  "   and So.Suboperacoes_Id = NFxI.Suboperacao" & vbCrLf & _
                  " Inner Join Produtos P" & vbCrLf & _
                  "	   on P.Produto_id = NFxI.Produto_id" & vbCrLf & _
                  " Inner Join Clientes Empresa" & vbCrLf & _
                  "	   on Empresa.Cliente_id  = NF.Empresa_id" & vbCrLf & _
                  "   and Empresa.Endereco_id = NF.EndEmpresa_Id" & vbCrLf & _
                  " Inner Join Clientes " & vbCrLf & _
                  "	   on Clientes.Cliente_id  = NF.Cliente_Id" & vbCrLf & _
                  "   and Clientes.Endereco_id = NF.EndCliente_Id" & vbCrLf & _
                  "  left Join NotasXNotas NN" & vbCrLf & _
                  "    on NN.OrigemEmpresa_id      = NF.Empresa_id" & vbCrLf & _
                  "   and NN.OrigemEndEmpresa_id   = NF.EndEmpresa_id" & vbCrLf & _
                  "   and NN.OrigemCliente_id      = NF.Cliente_id" & vbCrLf & _
                  "   and NN.OrigemEndCliente_id   = NF.EndCliente_id" & vbCrLf & _
                  "   and NN.OrigemEntradaSaida_id = NF.EntradaSaida_id" & vbCrLf & _
                  "   and NN.OrigemSerie_id        = NF.Serie_id" & vbCrLf & _
                  "   and NN.OrigemNota_id         = NF.Nota_id" & vbCrLf & _
                  "   and NN.Serie_id         not in ('UN','REC')" & vbCrLf & _
                  "  left join NotasFiscaisxItens NFxIT" & vbCrLf & _
                  "    on NN.Empresa_id      = NFxIT.Empresa_id" & vbCrLf & _
                  "   and NN.EndEmpresa_id   = NFxIT.EndEmpresa_id" & vbCrLf & _
                  "   and NN.Cliente_id      = NFxIT.Cliente_id" & vbCrLf & _
                  "   and NN.EndCliente_id   = NFxIT.EndCliente_id" & vbCrLf & _
                  "   and NN.EntradaSaida_id = NFxIT.EntradaSaida_id" & vbCrLf & _
                  "   and NN.Serie_id        = NFxIT.Serie_id" & vbCrLf & _
                  "   and NN.Nota_id         = NFxIT.Nota_id" & vbCrLf & _
                  "   and NFxI.Produto_id    = NFxIT.Produto_id  " & vbCrLf & _
                  " where 1 = 1" & vbCrLf

            If objNotaFiscal.CodigoFinalidade <> 13 And objNotaFiscal.CodigoFinalidade <> 23 Then
                sql &= " 	   and NF.Empresa_Id    ='" & Emp & "'" & vbCrLf & _
                       "       and NF.EndEmpresa_Id = " & EndEmp & vbCrLf
            End If

            If objNotaFiscal.VincularNotas Then
                'Não vai consultar finalidade
            ElseIf objNotaFiscal.CodigoFinalidade = 13 Then
                sql &= "        and nf.finalidade         = 20" & vbCrLf &
                       "        and left(NF.Cliente_Id,8) ='" & Left(objNotaFiscal.CodigoCliente, 8) & "'" & vbCrLf
                '"        and NF.EndCliente_Id    = " & objNotaFiscal.EnderecoCliente & vbCrLf
            ElseIf objNotaFiscal.CodigoFinalidade = 23 Then


                sql &= "        and nf.finalidade = 22" & vbCrLf
            ElseIf objNotaFiscal.CodigoFinalidade = 14 Then
                sql &= "        and nf.finalidade = 14" & vbCrLf
            Else
                sql &= "        and isnull(NF.troca,0) = 1" & vbCrLf & _
                       "        and nf.finalidade not in (13, 23, 20, 22, 14)" & vbCrLf
            End If

            If (TypeOf Me.Page Is TrocaDeNota Or objNotaFiscal.VincularNotas) AndAlso Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) AndAlso objNotaFiscal.CodigoFinalidade <> 13 Then
                sql &= "       and left(NF.Cliente_Id,8)     ='" & Left(objNotaFiscal.CodigoCliente, 8) & "'" & vbCrLf
                '"       and NF.EndCliente_Id   = " & objNotaFiscal.EnderecoCliente & vbCrLf
            End If

            If objNotaFiscal.VincularNotas Then

                If objNotaFiscal.EntradaSaida.ToString() = "Entrada" Then
                    sql &= " 	   and NF.EntradaSaida_Id ='E'" & vbCrLf
                Else
                    sql &= " 	   and NF.EntradaSaida_Id ='S'" & vbCrLf
                End If

                sql &= " 	   and p.produto_id in (" & vbCrLf &
                           "                            select produtoAgrupado_id" & vbCrLf &
                           "                              from ProdutosAgrupados" & vbCrLf &
                           "                             where Produto_id = (select Produto_id" & vbCrLf &
                           "                                                   from ProdutosAgrupados" & vbCrLf &
                           "                                                  where produtoAgrupado_id = NFxI.Produto_id)" & vbCrLf &
                           "                            union" & vbCrLf &
                           "                         select '" & objNotaFiscal.Itens(0).CodigoProduto & "'" & vbCrLf &
                           "                         )" & vbCrLf

            Else

                'Edson 27/10/2014 comentei essa clausula pq com ela nao viria as notas que foram feitas parcialmente
                '"       and NN.OrigemNota_Id IS NULL " & vbCrLf & _
                sql &= " 	   and NF.EntradaSaida_Id ='" & IIf(ES = "E", "S", "E") & "'" & vbCrLf &
                       "       and p.produto_id in (" & vbCrLf &
                       "                            select produtoAgrupado_id" & vbCrLf &
                       "                              from ProdutosAgrupados" & vbCrLf &
                       "                             where Produto_id = (select Produto_id" & vbCrLf &
                       "                                                   from ProdutosAgrupados" & vbCrLf &
                       "                                                  where produtoAgrupado_id = NFxI.Produto_id)" & vbCrLf &
                       "                            union" & vbCrLf &
                       "                         select '" & objNotaFiscal.Itens(0).CodigoProduto & "'" & vbCrLf &
                       "                         )" & vbCrLf

            End If


            If objNotaFiscal.VincularNotas Then

                sql &= IIf(String.IsNullOrWhiteSpace(txtNumNota.Text), "", "        AND NF.Nota_Id = " & txtNumNota.Text & "") & vbCrLf

                sql &= IIf(Not TypeOf Me.Page Is TrocaDeNota, "        AND YEAR(NF.Movimento) = " & txtNumAno.Text, " AND NF.Movimento between '" & objNotaFiscal.DataInclusao.ToSqlDate() & "' and '" & objNotaFiscal.DataTermino.ToSqlDate() & "'") & vbCrLf &
                   "     Group by CONVERT(varchar, NF.Movimento, 103), " & vbCrLf &
                   "              NF.Empresa_Id, NF.EndEmpresa_Id, " & vbCrLf &
                   " 	          Empresa.Nome, Empresa.Cidade, Empresa.Estado," & vbCrLf &
                   "              NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
                   "              Clientes.Nome, Clientes.Cidade, Clientes.Estado," & vbCrLf &
                   "              NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id," & vbCrLf &
                   "              NFxI.Produto_id, P.Nome, P.Unidade" & vbCrLf &
                   "    having max(NFxI.QuantidadeFisica) - sum(isnull(NFxIT.QuantidadeFisica,0)) > 0 " & vbCrLf &
                   "    or max(NFxI.QuantidadeFiscal) - sum(isnull(NFxIT.QuantidadeFiscal,0)) > 0; "

            Else

                sql &= IIf(String.IsNullOrWhiteSpace(txtNumNota.Text), "", "        AND NF.Nota_Id LIKE '%" & txtNumNota.Text & "%' ") & vbCrLf

                sql &= IIf(Not TypeOf Me.Page Is TrocaDeNota, "        AND YEAR(NF.Movimento) = " & txtNumAno.Text, " AND NF.Movimento between '" & objNotaFiscal.DataInclusao.ToSqlDate() & "' and '" & objNotaFiscal.DataTermino.ToSqlDate() & "'") & vbCrLf &
                   "     Group by CONVERT(varchar, NF.Movimento, 103), " & vbCrLf &
                   "              NF.Empresa_Id, NF.EndEmpresa_Id, " & vbCrLf &
                   " 	          Empresa.Nome, Empresa.Cidade, Empresa.Estado," & vbCrLf &
                   "              NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
                   "              Clientes.Nome, Clientes.Cidade, Clientes.Estado," & vbCrLf &
                   "              NF.EntradaSaida_Id, NF.Serie_Id, NF.Nota_Id," & vbCrLf &
                   "              NFxI.Produto_id, P.Nome, P.Unidade" & vbCrLf &
                   "    having max(NFxI.QuantidadeFisica) - sum(isnull(NFxIT.QuantidadeFisica,0)) > 0"

            End If

            ds = Banco.ConsultaDataSet(sql, "Notas")
            gridNotas.DataSource = ds
            gridNotas.DataBind()

            If ds.Tables(0).Rows.Count = 0 Then
                Session.Remove("objTrocaDeNotaNXI" & HID.Value.ToString)
                Session("SemTrocaDeNota" & HID.Value.ToString) = True
                If TypeOf Me.Page Is NotaFiscalXItens Then
                    CType(Me.Page, NotaFiscalXItens).CarregarNotaTroca()
                End If
                Popup.CloseDialog(Me.Page, "divConsultaNotaTroca")
            End If
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        txtNumNota.Text = String.Empty
        txtNumAno.Text = DateTime.Now.Year.ToString
    End Sub

    Protected Sub LnkSair_Click(sender As Object, e As EventArgs) Handles LnkSair.Click
        Popup.CloseDialog(Me.Page, "divConsultaNotaTroca")
        If TypeOf Me.Page Is NotaFiscalXItens Then
            CType(Me.Page, NotaFiscalXItens).CarregarNotaTroca()
        End If
    End Sub

    Protected Sub lnkCompra_Click(sender As Object, e As EventArgs) Handles lnkCompra.Click
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.CodigoFinalidade = 13
        SessaoSalvarNotaFiscal()
        Consultar()
    End Sub

    Protected Sub lnkTransferencia_Click(sender As Object, e As EventArgs) Handles lnkTransferencia.Click
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.CodigoFinalidade = 23
        SessaoSalvarNotaFiscal()
        Consultar()
    End Sub

    Protected Sub lnkTroca_Click(sender As Object, e As EventArgs) Handles lnkTroca.Click
        SessaoRecuperaNotaFiscal()
        Consultar()
    End Sub

    Protected Sub lnkCessaoDeCredito_Click(sender As Object, e As EventArgs) Handles lnkCessaoDeCredito.Click
        SessaoRecuperaNotaFiscal()
        objNotaFiscal.CodigoFinalidade = 14
        SessaoSalvarNotaFiscal()
        Consultar()
    End Sub
End Class