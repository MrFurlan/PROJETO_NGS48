Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaNotaFiscal
    Inherits BaseUserControl
    Dim notas As String = ""

#Region "Eventos"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then

        End If
    End Sub

    Protected Sub btnSelecionar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSelecionar.Click
        Selecionar()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaNotaFiscal")
    End Sub
#End Region

#Region "Métodos"
    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub
    Dim pram As New Dictionary(Of String, Object)

    Public Overrides Sub Limpar()
        grdNotasDoPedido.DataSource = New List(Of Object)
        grdNotasDoPedido.DataBind()
    End Sub

    Protected Overrides Sub Selecionar()
        Dim listNotas As New ListNotasFiscais()
        For Each row As GridViewRow In grdNotasDoPedido.Rows
            Dim chk As CheckBox = CType(row.FindControl("chkSelecionado"), CheckBox)
            If chk IsNot Nothing AndAlso chk.Checked Then

                Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal

                objNotaFiscal.CodigoEmpresa = CType(row.FindControl("lblEmpresa"), Label).Text.Trim().Split("-")(0)
                objNotaFiscal.EnderecoEmpresa = CType(row.FindControl("lblEmpresa"), Label).Text.Trim().Split("-")(1)
                objNotaFiscal.CodigoCliente = CType(row.FindControl("lblCliente"), Label).Text.Trim().Split("-")(0)
                objNotaFiscal.EnderecoCliente = CType(row.FindControl("lblCliente"), Label).Text.Trim().Split("-")(1)
                objNotaFiscal.Codigo = CType(row.FindControl("lblNota"), Label).Text.Trim().Split("-")(0)
                objNotaFiscal.Serie = CType(row.FindControl("lblNota"), Label).Text.Trim().Split("-")(1)
                objNotaFiscal.EntradaSaida = IIf(CType(row.FindControl("lblEntradaSaida"), Label).Text.Trim() = "S", [Lib].Negocio.eEntradaSaida.Saida, [Lib].Negocio.eEntradaSaida.Entrada)
                Dim nf As New [Lib].Negocio.NotaFiscal(objNotaFiscal)
                listNotas.Add(nf)
            End If
        Next

        Session(Session("ssTipoRetorno")) = listNotas
        If Session("ssTipoRetorno") IsNot Nothing Then
            If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                Dim ucName = MainUserControl.ClientID.Split("_")
                Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                CType(uc, IBaseUserControl).Carregar(listNotas)
            Else
                CType(Me.Page, IBasePage).Carregar(listNotas)
            End If
            Popup.CloseDialog(Me.Page, "divConsultaNotaFiscal")
        End If
    End Sub

    Public Sub BindGridView(ByVal parameters As Dictionary(Of String, Object))

        Dim sql As String = ""

        If parameters IsNot Nothing AndAlso parameters.ContainsKey("pedido") AndAlso Not String.IsNullOrWhiteSpace(parameters("pedido")) Then
            sql &= "   SELECT       NFI.Pedido, CAST(NFI.Nota_Id as varchar ) + '-' + NFI.Serie_Id as Nota_Id, NFI.Empresa_Id + '-' + cast(NFI.EndEmpresa_Id as varchar) as Empresa_Id, NFI.EntradaSaida_Id, " & vbCrLf & _
                       "            NFI.CLIENTE_ID + '-' + CAST(NFI.ENDCLIENTE_ID AS VARCHAR) AS Cliente_ID, NFI.Produto_Id, NFI.CFOP_Id," & vbCrLf & _
                       "   			SUM(isnull((CASE WHEN NFI.QuantidadeFisica <= 0 THEN NFI.QuantidadeFiscal ELSE NFI.QuantidadeFisica END),0)) as Quantidade,  " & vbCrLf & _
                       "   			SUM(isnull(NFI.Valor,0)) as Valor                                                                                                                " & vbCrLf & _
                       "   FROM         NotasFiscais AS NF                                                                                                       " & vbCrLf & _
                       "   INNER JOIN	 NotasFiscaisXItens AS NFI                                                                                               " & vbCrLf & _
                       "   	ON	NF.Nota_Id = NFI.Nota_Id                                                                                                     " & vbCrLf & _
                       "   	AND NF.Empresa_Id		= NFI.Empresa_Id                                                                                         " & vbCrLf & _
                       "   	AND NF.EndEmpresa_Id	= NFI.EndEmpresa_Id                                                                                      " & vbCrLf & _
                       "   	AND NF.Cliente_Id		= NFI.Cliente_Id                                                                                         " & vbCrLf & _
                       "      AND NF.EndCliente_Id	= NFI.EndCliente_Id                                                                                          " & vbCrLf & _
                       "      AND NF.EntradaSaida_Id	= NFI.EntradaSaida_Id                                                                                    " & vbCrLf & _
                       "      AND NF.Serie_Id			= NFI.Serie_Id                                                                                           " & vbCrLf & _
                       "   INNER JOIN Clientes CLI                                                                                                               " & vbCrLf & _
                       "   	ON	CLI.Cliente_Id		= NF.Cliente_Id                                                                                          " & vbCrLf & _
                       "   	AND CLI.Endereco_Id		= NF.EndCliente_Id                                                                                       " & vbCrLf & _
                       "                                                                                                                                        " & vbCrLf & _
                       "   WHERE	NF.SITUACAO = 1                                                                                                              " & vbCrLf & _
                       "   	AND NF.PEDIDO = " & parameters("pedido") & "                                                                                                            " & vbCrLf & _
                       "    AND NFI.PRODUTO_ID = '" & parameters("produto") & "'" & vbCrLf & _
                       "   GROUP BY  NFI.Pedido, NFI.Nota_Id, NFI.Serie_Id, NFI.Empresa_Id, NFI.EndEmpresa_Id, NFI.Cliente_Id, NFI.EndCliente_Id, NFI.EntradaSaida_Id, NFI.Produto_Id, NFI.CFOP_Id    " & vbCrLf
        End If

        Try
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NotasDoPedido")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                grdNotasDoPedido.DataSource = ds
                grdNotasDoPedido.DataBind()
            Else
                MsgBox(Me.Page, "Não foi encontrado nenhum resultado referente ao pedido")
                Popup.CloseDialog(Me.Page, "divConsultaNotaFiscal")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub chkSelecionarTodos_CheckedChanged(sender As Object, e As EventArgs)

    End Sub
End Class