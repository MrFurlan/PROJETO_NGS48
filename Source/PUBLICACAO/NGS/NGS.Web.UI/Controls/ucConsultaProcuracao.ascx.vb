Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaProcuracao
    Inherits BaseUserControl

    Dim sql As String
    Dim strTipo As String = ""
    Dim strProcuracao As String = ""
    Dim strPedido As String = ""
    Dim strSituacao As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Protected Sub GridProcuracoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not String.IsNullOrWhiteSpace(HIDTipo.Value) Then
            strTipo = HIDTipo.Value.ToString
        End If
        Selecionar()
    End Sub

    Protected Sub GridProcuracao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not String.IsNullOrWhiteSpace(HIDTipo.Value) Then
            strTipo = HIDTipo.Value.ToString
        End If
        Selecionar()
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objProcuracaoNxI")) Then
            If TypeOf Me.Page Is NotaFiscalXItens Then
                CType(Me.Page, NotaFiscalXItens).CarregarProcuracao()
            End If
        End If
        Popup.CloseDialog(Me.Page, "divCessaoDeCredito")
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
        updConsultaCessaoDeCredito.Update()
    End Sub

    Public Function BindGridView(ByVal parameters As Dictionary(Of String, Object)) As Boolean
        If parameters("tipo") IsNot Nothing Then
            strTipo = parameters("tipo").ToString
            HIDTipo.Value = strTipo
        End If

        If strTipo = "NxI" Then
            CarregarProcuracoes(parameters)
            Return GridProcuracoes.Rows.Count > 0
        Else
            CargaProcuracao(parameters)
            Return GridProcuracao.Rows.Count > 0
        End If

    End Function

    Private Sub CarregarProcuracoes(ByVal parameters As Dictionary(Of String, Object))
        If parameters("pedc") IsNot Nothing Then
            strPedido = parameters("pedc")
        End If

        Dim Procuracoes As New [Lib].Negocio.ListProcuracao(strPedido, False)
        GridProcuracao.Visible = False
        GridProcuracoes.Visible = True
        GridProcuracoes.DataSource = Procuracoes.ToArray
        GridProcuracoes.DataBind()
        updConsultaCessaoDeCredito.Update()
        If Not Procuracoes IsNot Nothing AndAlso Procuracoes.Count > 0 Then
            Session("SemProcuracao" & HID.Value) = True
            Popup.CloseDialog(Me.Page, "divCessaoDeCredito")
        End If
    End Sub

    Private Sub CargaProcuracao(ByVal parameters As Dictionary(Of String, Object))
        If parameters("situacao") IsNot Nothing Then strSituacao = parameters("situacao")
        If parameters("pedc") IsNot Nothing Then strPedido = parameters("pedc")

        sql = " SELECT PR.Procuracao_ID AS Procuracao,                                                  " & vbCrLf & _
              "        PR.PedidoCedente AS Pedido,                                                      " & vbCrLf & _
              "        PR.Movimento,                                                                    " & vbCrLf & _
              "        C.Cliente_Id AS Cliente,                                                         " & vbCrLf & _
              "        C.Endereco_Id AS EndCliente,                                                     " & vbCrLf & _
              "        C.Nome,                                                                          " & vbCrLf & _
              "        PR.Quantidade,                                                                   " & vbCrLf & _
              "        ISNULL(sum(sb_Real.Quantidade), 0) AS Realizado,                                 " & vbCrLf & _
              "        PR.Quantidade - ISNULL(sum(sb_Real.Quantidade), 0) AS Saldo                      " & vbCrLf & _
              "   FROM Procuracoes PR                                                                   " & vbCrLf & _
              "  INNER JOIN Clientes C                                                                  " & vbCrLf & _
              "     ON PR.Cessionario    = C.Cliente_Id                                                 " & vbCrLf & _
              "    AND PR.EndCessionario = C.Endereco_Id                                                " & vbCrLf & _
              "   LEFT JOIN (SELECT NF.Empresa_Id,                                                      " & vbCrLf & _
              "                     NF.EndEmpresa_Id,                                                   " & vbCrLf & _
              "	                    NF.Pedido,                                                          " & vbCrLf & _
              "	                    NF.Procuracao,                                                      " & vbCrLf & _
              "	                    SUM(nfi.QuantidadeFisica) AS Quantidade                             " & vbCrLf & _
              "                FROM NotasFiscais NF                                                     " & vbCrLf & _
              "               INNER JOIN NotasFiscaisXItens nfi                                         " & vbCrLf & _
              "                  ON NF.Empresa_Id      = nfi.Empresa_Id                                 " & vbCrLf & _
              "                 AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id                              " & vbCrLf & _
              "                 AND NF.Cliente_Id      = nfi.Cliente_Id                                 " & vbCrLf & _
              "                 AND NF.EndCliente_Id   = nfi.EndCliente_Id                              " & vbCrLf & _
              "                 AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id                            " & vbCrLf & _
              "                 AND NF.Serie_Id        = nfi.Serie_Id                                   " & vbCrLf & _
              "                 AND NF.Nota_Id         = nfi.Nota_Id                                    " & vbCrLf & _
              "               Inner join SubOperacoes SO                                                " & vbCrLf & _
              "	                 on so.Operacao_Id     = NF.Operacao                                    " & vbCrLf & _
              "                 and so.SubOperacoes_Id = NF.SubOperacao                                 " & vbCrLf & _
              "               Where so.Classe <> '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'                                      " & vbCrLf & _
              "                 and NF.situacao in (1,4,7)                                              " & vbCrLf & _
              "               GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Procuracao        " & vbCrLf & _
              "	              Union All                                                                 " & vbCrLf & _
              "	             Select P.Empresa_Id,                                                       " & vbCrLf & _
              "                     P.EndEmpresa_Id,                                                    " & vbCrLf & _
              "	                    P.Pedido_id,                                                        " & vbCrLf & _
              "	                    PIF.Procuracao,                                                     " & vbCrLf & _
              "	              	    sum(pif.Quantidade)                                                 " & vbCrLf & _
              "	               from Pedidos P                                                           " & vbCrLf & _
              "               Inner join PedidosXItensXFixacoes PIF                                     " & vbCrLf & _
              "	                 on P.Empresa_Id    = PIF.Empresa_Id                                    " & vbCrLf & _
              "                 and P.EndEmpresa_Id = PIF.EndEmpresa_Id                                 " & vbCrLf & _
              "	                and p.Pedido_Id     = PIF.Pedido_Id                                     " & vbCrLf & _
              "	              Where P.Situacao      = 1                                                 " & vbCrLf & _
              "	              Group by P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_id, PIF.Procuracao       " & vbCrLf & _
              "        ) AS sb_Real                                                                     " & vbCrLf & _
              "     ON sb_Real.Empresa_Id    = PR.Empresa_Id                                            " & vbCrLf & _
              "    AND sb_Real.EndEmpresa_Id = PR.EndEmpresa_Id                                         " & vbCrLf & _
              "    AND sb_Real.Pedido        = PR.PedidoCedente                                         " & vbCrLf & _
              "    AND sb_Real.Procuracao    = PR.Procuracao_ID                                         " & vbCrLf & _
              "  WHERE PR.Empresa_id         = '" & parameters("emp") & "'                              " & vbCrLf & _
              "    AND PR.EndEmpresa_id      =  " & parameters("ende") & "                              " & vbCrLf

        If parameters("proc") IsNot Nothing AndAlso parameters("proc").Length > 0 Then
            sql &= "    AND PR.Procuracao_ID = " & parameters("proc") & vbCrLf
        Else

            If parameters("cliCedente") IsNot Nothing AndAlso parameters("cliCedente").Length > 0 Then sql &= "    AND PR.Cedente = '" & parameters("cliCedente") & "' AND PR.EndCedente = " & parameters("endCedente") & vbCrLf
            If parameters("cliCessionario") IsNot Nothing AndAlso parameters("cliCessionario").Length > 0 Then sql &= "    AND PR.Cessionario = '" & parameters("cliCessionario") & "' AND PR.EndCessionario = " & parameters("endCessionario") & vbCrLf
            If parameters("periodo") IsNot Nothing AndAlso parameters("periodo").Length > 0 Then sql &= "  AND PR.Movimento Between " & parameters("periodo") & vbCrLf
            If parameters("safra") IsNot Nothing AndAlso parameters("safra").Length > 0 Then sql &= "  AND PR.Safra = '" & parameters("safra") & "'" & vbCrLf

            If strPedido.Length > 0 Then sql &= "    AND PR.PedidoCedente = " & strPedido & vbCrLf
            If strSituacao.Length > 0 Then sql &= "    AND PR.Situacao = " & strSituacao & vbCrLf
        End If

        sql &= "  group by PR.Procuracao_ID, PR.PedidoCedente, PR.Movimento, Cliente_Id, C.Endereco_Id, C.Nome, PR.Quantidade" & vbCrLf

        If parameters("proc").Length = 0 Then
            If parameters("filtro") IsNot Nothing AndAlso parameters("filtro") = "L" Then sql &= "  having PR.Quantidade - ISNULL(sum(sb_Real.Quantidade), 0) = 0" & vbCrLf
            If parameters("filtro") IsNot Nothing AndAlso parameters("filtro") = "P" Then sql &= "  having PR.Quantidade - ISNULL(sum(sb_Real.Quantidade), 0) > 0" & vbCrLf
        End If

        sql &= "  order by PR.Movimento desc"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Procuracoes")

        GridProcuracoes.Visible = False
        GridProcuracao.Visible = True
        GridProcuracao.DataSource = ds
        GridProcuracao.DataBind()
        updConsultaCessaoDeCredito.Update()

    End Sub

    Private Overloads Sub Selecionar()
        Dim procuracao As String = ""
        If (GridProcuracoes.Visible) Then
            procuracao = GridProcuracoes.SelectedRow.Cells(1).Text()
        ElseIf (GridProcuracao.Visible) Then
            procuracao = GridProcuracao.SelectedRow.Cells(1).Text()
        End If

        Session(Session("ssTipoRetorno")) = procuracao

        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objProcuracaoNxI")) Then
            If TypeOf Me.Page Is NotaFiscalXItens Then
                Popup.CloseDialog(Me.Page, "divCessaoDeCredito")
                CType(Me.Page, NotaFiscalXItens).CarregarProcuracao()
            End If
        ElseIf (Session("ssTipoRetorno") IsNot Nothing AndAlso TypeOf Me.Page Is CessaoDeCredito) Then
            Popup.CloseDialog(Me.Page, "divCessaoDeCredito")
            CType(Me.Page, CessaoDeCredito).CarregarProcuracao()
        End If
    End Sub

End Class