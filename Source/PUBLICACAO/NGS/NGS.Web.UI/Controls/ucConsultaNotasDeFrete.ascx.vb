Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaNotasDeFrete
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Sub BindGridView(ByVal parameters As Dictionary(Of String, Object))
        Dim TipoDeDocumento As String = parameters("Tipo")
        Dim Emp As String = parameters("Emp")
        Dim EndEmp As String = parameters("EndEmp")
        'Dim Cli As String = parameters("Cli")
        'Dim EndCli As String = parameters("EndCli")
        Dim Nota As String = parameters("Nota")
        Dim PagarReceber As String = parameters("PagarReceber")
        Dim ds As DataSet
        Dim sql As String

        sql = " 	select CONVERT(varchar, NF.Movimento, 103) as Movimento, NF.Empresa_Id, NF.EndEmpresa_Id, " & vbCrLf &
              " 		   Empresa.Nome As EmpresaNome, Empresa.Cidade as EmpresaCidade, Empresa.Estado as EmpresaEstado," & vbCrLf &
              " 		   NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf &
              " 		   Clientes.Nome, Clientes.Cidade, Clientes.Estado," & vbCrLf &
              " 		   NF.EntradaSaida_Id," & vbCrLf &
              " 		   NF.Serie_Id," & vbCrLf &
              " 		   NF.Nota_Id," & vbCrLf &
              "            convert(nvarchar,NF.Operacao) + '-' + convert(nvarchar,NF.SubOperacao) as OpSb, " & vbCrLf &
              " 		   NFxI.Produto_id," & vbCrLf &
              " 		   P.Nome as NomeProduto," & vbCrLf &
              " 		   P.Unidade," & vbCrLf &
              "            NF.Operacao, NF.SubOperacao, SubOperacoes.Descricao as DescOperacao," & vbCrLf &
              "            sum(NFxIT.Valor) as ValorNotaOficial," & vbCrLf &
              "            sum(NFxIT.Unitario) as  Unitario," & vbCrLf &
              "            sum(NFxIT.PesoFiscal) as Peso " & vbCrLf &
              " 	  from NotasFiscais NF" & vbCrLf &
              " 	 Inner Join NotasFiscaisxItens as NFxI" & vbCrLf &
              " 		ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
              " 	   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
              " 	   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
              " 	   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
              " 	   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
              " 	   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
              " 	   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
             "      INNER JOIN FaturasDeFretesXItens FXI" & vbCrLf &
             "          ON NFxI.Empresa_Id          = FXI.Empresa_Id " & vbCrLf &
             "          AND NFxI.EndEmpresa_Id      = FXI.EndEmpresa_Id " & vbCrLf &
             "          AND NFxI.Cliente_Id         = FXI.Cliente_Id " & vbCrLf &
             "          AND NFxI.EndCliente_Id      = FXI.EndCliente_Id " & vbCrLf &
             "          AND NFxI.EntradaSaida_Id    = FXI.EntradaSaida_Id " & vbCrLf &
             "          AND NFxI.Serie_Id           = FXI.Serie_Id " & vbCrLf &
             "          AND NFxI.Nota_Id            = FXI.Nota_Id" & vbCrLf &
             "      INNER JOIN FaturaDeFreteXTitulo FFXT " & vbCrLf &
             "          ON FXI.Empresa_Id           = FFXT.Empresa_Id " & vbCrLf &
             "          AND FXI.EndEmpresa_Id       = FFXT.EndEmpresa_Id " & vbCrLf &
             "          AND FXI.Fatura_Id           = FFXT.Fatura_Id" & vbCrLf

        If Trim(PagarReceber) = "P" Then
            sql &= "    INNER JOIN ContasAPagar CP 
                            ON CP.Registro_Id = FFXT.Titulo_Id
	                        AND CP.Situacao = 1
	                        AND CP.Provisao = 3 " & vbCrLf
        Else
            sql &= "    INNER JOIN ContasAReceber CR 
                            ON CR.Registro_Id = FFXT.Titulo_Id
	                        AND CR.Situacao = 1
	                        AND CR.Provisao = 3 " & vbCrLf
        End If

        sql &= " 	 Inner Join Produtos P" & vbCrLf &
              " 		On P.Produto_id = NFxI.Produto_id" & vbCrLf &
              " 	 Inner Join Clientes Empresa" & vbCrLf &
              " 		On Empresa.Cliente_id  = NF.Empresa_id" & vbCrLf &
              " 	   And Empresa.Endereco_id = NF.EndEmpresa_Id" & vbCrLf &
              " 	 Inner Join Clientes " & vbCrLf &
              " 		On Clientes.Cliente_id  = NF.Cliente_Id" & vbCrLf &
              " 	   And Clientes.Endereco_id = NF.EndCliente_Id" & vbCrLf &
              "       left Join NotasXNotas NN" & vbCrLf &
              "         On NN.OrigemEmpresa_id      = NF.Empresa_id" & vbCrLf &
              " 	   And NN.OrigemEndEmpresa_id   = NF.EndEmpresa_id" & vbCrLf &
              " 	   And NN.OrigemCliente_id      = NF.Cliente_id" & vbCrLf &
              " 	   And NN.OrigemEndCliente_id   = NF.EndCliente_id" & vbCrLf &
              " 	   And NN.OrigemEntradaSaida_id = NF.EntradaSaida_id" & vbCrLf &
              " 	   And NN.OrigemSerie_id        = NF.Serie_id" & vbCrLf &
              " 	   And NN.OrigemNota_id         = NF.Nota_id" & vbCrLf &
              "       left join NotasFiscaisxItens NFxIT" & vbCrLf &
              "         On NN.Empresa_id      = NFxIT.Empresa_id" & vbCrLf &
              " 	   And NN.EndEmpresa_id   = NFxIT.EndEmpresa_id" & vbCrLf &
              " 	   And NN.Cliente_id      = NFxIT.Cliente_id" & vbCrLf &
              " 	   And NN.EndCliente_id   = NFxIT.EndCliente_id" & vbCrLf &
              " 	   And NN.EntradaSaida_id = NFxIT.EntradaSaida_id" & vbCrLf &
              " 	   And NN.Serie_id        = NFxIT.Serie_id" & vbCrLf &
              " 	   And NN.Nota_id         = NFxIT.Nota_id" & vbCrLf &
              "        And NFxI.Produto_id    = NFxIT.Produto_id  " & vbCrLf &
              "       INNER JOIN SubOperacoes " & vbCrLf &
              "        On SubOperacoes.Operacao_Id     = NF.Operacao" & vbCrLf &
              "        And SubOperacoes.SubOperacoes_Id = NF.SubOperacao" & vbCrLf &
              " 	 where 1=1 And (NF.TipoDeDocumento = 1 Or NF.TipoDeDocumento = 2 Or NF.TipoDeDocumento    ='" & TipoDeDocumento & "')" & vbCrLf


        If Trim(Emp) <> "" AndAlso Trim(EndEmp) <> "" Then
            sql &= "        AND NF.Empresa_Id     ='" & Emp & "'" & vbCrLf & _
                   "        AND NF.EndEmpresa_Id   = " & EndEmp & vbCrLf
        End If

        'If Trim(Cli) <> "" AndAlso Trim(EndCli) <> "" Then
        '    sql &= "        AND NF.Cliente_Id     ='" & Cli & "'" & vbCrLf & _
        '           "        AND NF.EndCliente_Id   = " & EndCli & vbCrLf
        'End If

        If Nota.Length > 0 Then
            sql &= "   AND NF.Nota_Id = " & Nota & vbCrLf
        End If

        sql &= "   Group By " & vbCrLf & _
        "  NF.Movimento, NF.Empresa_Id, NF.EndEmpresa_Id, " & vbCrLf & _
        "  Empresa.Nome, Empresa.Cidade, Empresa.Estado, " & vbCrLf & _
        "  NF.Cliente_Id, NF.EndCliente_Id, " & vbCrLf & _
        "  Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf & _
        "  NF.EntradaSaida_Id, " & vbCrLf & _
        "  NF.Serie_Id, " & vbCrLf & _
        "  NF.Nota_Id, " & vbCrLf & _
        "  NFxI.Produto_id, " & vbCrLf & _
        "  P.Nome, " & vbCrLf & _
        "  P.Unidade, " & vbCrLf & _
        "  NF.Operacao, NF.SubOperacao, " & vbCrLf & _
        "  SubOperacoes.Descricao" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Notas")
        gridNotasFrete.DataSource = ds
        gridNotasFrete.DataBind()

        If ds.Tables(0).Rows.Count = 0 Then
            Session("SemNotasDeFrete" & HID.Value) = True
            MsgBox(Me.Page, "Nenhum registro encontrado!")
            Popup.CloseDialog(Me.Page, "divConsultaNotasDeFrete")
        ElseIf ds.Tables(0).Rows.Count = 1 Then
            gridNotasFrete.SelectedIndex = 0
            gridNotas_SelectedIndexChanged(Nothing, Nothing)
        End If
    End Sub

    Protected Sub gridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal
        objNotaFiscal.CodigoEmpresa = gridNotasFrete.SelectedRow.Cells(2).Text.Trim()
        objNotaFiscal.EnderecoEmpresa = gridNotasFrete.SelectedRow.Cells(3).Text.Trim()
        objNotaFiscal.CodigoCliente = gridNotasFrete.SelectedRow.Cells(7).Text.Trim()
        objNotaFiscal.EnderecoCliente = gridNotasFrete.SelectedRow.Cells(8).Text.Trim()
        objNotaFiscal.Codigo = gridNotasFrete.SelectedRow.Cells(13).Text.Trim()
        objNotaFiscal.Serie = gridNotasFrete.SelectedRow.Cells(14).Text.Trim()
        objNotaFiscal.EntradaSaida = IIf(gridNotasFrete.SelectedRow.Cells(12).Text.Trim() = "S", [Lib].Negocio.eEntradaSaida.Saida, [Lib].Negocio.eEntradaSaida.Entrada)
        Dim nf As New [Lib].Negocio.NotaFiscal(objNotaFiscal)
        Session("ConsultaNotasDeFrete" & HID.Value) = nf
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("ConsultaNotasDeFrete")) Then
            If TypeOf Me.Page Is ControleDeFretes Then
                CType(Me.Page, ControleDeFretes).CarregarNotasdeFrete()
            End If
        End If
        Popup.CloseDialog(Me.Page, "divConsultaNotasDeFrete")
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaNotasDeFrete")
    End Sub

End Class