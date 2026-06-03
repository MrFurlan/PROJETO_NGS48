Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucComissoesXBaixas
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
            Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        grd.DataSource = New List(Of Object)
        grd.DataBind()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            SessaoRecuperaNotaFiscal()
            Dim aux As Boolean = True
            Dim lst As New ListComissoesXBaixas()
            If grd.Rows.Count > 0 Then
                For Each row As GridViewRow In grd.Rows
                    Dim txtValor As TextBox = CType(row.FindControl("txtValor"), TextBox)
                    Dim vlrSaldo As Decimal = CDec(row.Cells(7).Text.Trim())

                    If (CDec(txtValor.Text) > vlrSaldo) Then
                        aux &= False
                        MsgBox(Me.Page, "Valor informado do pedido '" & row.Cells(0).Text.Trim() & "' não pode ser maior que o saldo!")
                        Continue For
                    End If

                    If txtValor IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txtValor.Text) Then
                        Dim pId As String = row.Cells(0).Text.Trim()
                        Dim objComissoesXBaixas As ComissoesXBaixas = objNotaFiscal.ComissoesXBaixas.SingleOrDefault(Function(s) s.Pedido_Id = pId)
                        If objComissoesXBaixas IsNot Nothing Then
                            objComissoesXBaixas.Valor = CDec(txtValor.Text.Trim())
                        Else
                            If CDec(txtValor.Text) > 0 Then
                                Dim obj As New ComissoesXBaixas(objNotaFiscal)
                                Dim strEmpresa() As String = row.Cells(1).Text.Trim().Split("-")
                                Dim p As New [Lib].Negocio.Pedido(strEmpresa(0), strEmpresa(1), pId)
                                obj.Empresa_Id = p.CodigoEmpresa
                                obj.EndEmpresa_Id = p.EnderecoEmpresa
                                obj.Pedido_Id = p.Codigo
                                obj.EndRepresentante_Id = objNotaFiscal.EnderecoCliente
                                obj.Representante_Id = objNotaFiscal.CodigoCliente
                                obj.EmpresaNota_Id = objNotaFiscal.CodigoEmpresa
                                obj.EndEmpresaNota_Id = objNotaFiscal.EnderecoEmpresa
                                obj.EntradaSaida_Id = objNotaFiscal.EntradaSaida
                                obj.Serie_Id = objNotaFiscal.Serie
                                obj.Nota_Id = objNotaFiscal.Codigo
                                obj.Valor = CDec(txtValor.Text.Trim())
                                lst.Add(obj)
                            End If
                        End If
                    End If
                Next
            End If

            If Not aux Then
                Exit Sub
            End If

            If (lst.Sum(Function(s) s.Valor) > CDec(txtVlrTotal.Text)) Then
                MsgBox(Me.Page, "Valor rateado não pode ser maior que o valor do produto!")
                Exit Sub
            End If

            Session(Session("ssTipoRetorno")) = lst
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(lst)
                Else
                    CType(Me.Page, IBasePage).Carregar(lst)
                End If
                Popup.CloseDialog(Me.Page, "divComissoesXBaixas")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub BindGridView(ByVal parameters As Dictionary(Of String, Object))
        Dim sql As String = "SELECT " & vbCrLf & _
                            "	p.Pedido_Id as Pedido, " & vbCrLf & _
                            "	p.Empresa_Id + '-' + cast(p.EndEmpresa_Id as varchar) as Empresa, " & vbCrLf & _
                            "	p.Cliente + '-' + cast(p.EndCliente as varchar) + ' - ' + cli.Nome as Cliente, " & vbCrLf & _
                            "	cast(p.Operacao as varchar) + '-' + cast(p.SubOperacao as varchar) + ' - ' + sop.Descricao as Operacao, " & vbCrLf & _
                            "	dbo.ProdutosPedido(p.Empresa_Id, p.EndEmpresa_Id, p.Pedido_Id) as Produto, " & vbCrLf & _
                            "	c.ValorComissao as Comissao, " & vbCrLf & _
                            "	SUM(isnull(cxb.Valor,0)) as Baixado, " & vbCrLf & _
                            "	c.ValorComissao - SUM(isnull(cxb.Valor,0)) as Saldo " & vbCrLf & _
                            "FROM Comissoes c " & vbCrLf & _
                            "INNER JOIN Pedidos p " & vbCrLf & _
                            "	 ON p.Empresa_Id = c.Empresa_Id " & vbCrLf & _
                            "	AND p.EndEmpresa_Id = c.EndEmpresa_Id " & vbCrLf & _
                            "	AND p.Pedido_Id = c.Pedido_Id " & vbCrLf & _
                            "INNER JOIN Clientes cli " & vbCrLf & _
                            "	 ON cli.Cliente_Id = p.Cliente " & vbCrLf & _
                            "	AND cli.Endereco_Id = p.EndCliente " & vbCrLf & _
                            "INNER JOIN SubOperacoes sop " & vbCrLf & _
                            "	 ON sop.SubOperacoes_Id = p.SubOperacao " & vbCrLf & _
                            "	AND sop.Operacao_Id = p.Operacao " & vbCrLf & _
                            "LEFT JOIN ComissoesXBaixas cxb " & vbCrLf & _
                            "	 ON c.Empresa_Id = cxb.Empresa_Id " & vbCrLf & _
                            "	AND c.EndEmpresa_Id = cxb.EndEmpresa_Id " & vbCrLf & _
                            "	AND c.Pedido_Id = cxb.Pedido_Id " & vbCrLf & _
                            "	AND c.Representante_Id = cxb.Representante_Id " & vbCrLf & _
                            "	AND c.EndRepresentante_Id = cxb.EndRepresentante_Id " & vbCrLf & _
                            "WHERE 1=1 " & vbCrLf & _
                            "AND c.Representante_Id = '" & parameters("representanteId") & "' " & vbCrLf & _
                            "AND c.EndRepresentante_Id = '" & parameters("enderecoRepId") & "' " & vbCrLf & _
                            "AND (c.Liquidado is null OR c.Liquidado = 0) " & vbCrLf & _
                            "GROUP BY p.Empresa_Id, p.EndEmpresa_Id, p.Pedido_Id, p.Cliente, p.EndCliente, " & vbCrLf & _
                            "		 p.Operacao, p.SubOperacao, sop.Descricao, c.Pedido_Id, c.ValorComissao, cli.Nome " & vbCrLf & _
                            "HAVING (c.ValorComissao - SUM(isnull(cxb.Valor,0))) > 0 " & vbCrLf & _
                            "ORDER BY c.Pedido_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Comissoes")
        grd.DataSource = ds
        grd.DataBind()
        txtVlrTotal.Text = String.Format("{0:N2}", parameters("valorTotal"))
        txtVlrRateado.Text = String.Format("{0:N2}", Decimal.Zero)
    End Sub

    Protected Sub lnkNovoAut_Click(sender As Object, e As EventArgs) Handles lnkNovoAut.Click
        Selecionar()
    End Sub

    Protected Sub lnkLimparAut_Click(sender As Object, e As EventArgs) Handles lnkLimparAut.Click
        Limpar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divComissoesXBaixas")
    End Sub

    Protected Sub grd_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles grd.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.DataItem IsNot Nothing Then
                SessaoRecuperaNotaFiscal()
                Dim objComissoesXBaixas As ComissoesXBaixas = objNotaFiscal.ComissoesXBaixas.SingleOrDefault(Function(s) s.Pedido_Id = e.Row.Cells(0).Text.Trim())
                If objComissoesXBaixas IsNot Nothing Then
                    Dim txtValor As TextBox = CType(e.Row.FindControl("txtValor"), TextBox)
                    txtValor.Text = String.Format("{0:N2}", objComissoesXBaixas.Valor)
                End If
            End If
        End If
    End Sub

End Class