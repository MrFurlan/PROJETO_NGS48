Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaPedidoDeTroca
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
        Dim Empresa As String = ""
        Dim EndEmpresa As String = ""
        Dim Cliente As String = ""
        Dim EndCliente As String = ""
        Dim Safra As String = ""
        Dim Moeda As String = ""
        Dim Where As String = ""

        If Not parameters("Empresa") Is Nothing Then Empresa = parameters("Empresa")
        If Not parameters("EndEmpresa") Is Nothing Then EndEmpresa = parameters("EndEmpresa")
        If Not parameters("Cliente") Is Nothing Then Cliente = parameters("Cliente")
        If Not parameters("EndCliente") Is Nothing Then EndCliente = parameters("EndCliente")
        If Not parameters("Safra") Is Nothing Then Safra = parameters("Safra")
        If Not parameters("Moeda") Is Nothing Then Moeda = parameters("Moeda")
        If Not parameters("Where") Is Nothing Then Where = parameters("Where")

        Dim Sql As String = String.Empty
        Sql &= "SELECT P.Empresa_Id, P.EndEmpresa_Id, Empresa.Nome, P.Safra, " & vbCrLf & _
               "       P.DataPedido, P.Pedido_Id, P.PedidoEfetivo, P.Cliente, P.EndCliente,  Clientes.Nome AS ClienteNome," & vbCrLf & _
               "       Clientes.Complemento, P.Moeda, M.Descricao as DescMoeda," & vbCrLf & _
               "       Sum(case " & vbCrLf & _
               "              when p.Moeda = 1" & vbCrLf & _
               "                then PxE.ValorOficial" & vbCrLf & _
               "                else PxE.ValorMoeda" & vbCrLf & _
               "           end) as Origem," & vbCrLf & _
               "       case " & vbCrLf & _
               "          when p.Moeda = 1" & vbCrLf & _
               "             then isnull(sb.OficialDestino,0)" & vbCrLf & _
               "             else isnull(sb.MoedaDestino,0)" & vbCrLf & _
               "       end as Destino" & vbCrLf & _
               "  Into #Temp" & vbCrLf & _
               "  FROM Pedidos AS P" & vbCrLf & _
               " Inner Join PedidosxEncargos PxE" & vbCrLf & _
               "    on P.Empresa_id    = PxE.Empresa_Id" & vbCrLf & _
               "   and P.EndEmpresa_Id = PxE.EndEmpresa_Id" & vbCrLf & _
               "   and P.Pedido_Id     = PxE.Pedido_Id" & vbCrLf & _
               "   and PxE.Encargo_Id  = 'LIQUIDO'" & vbCrLf & _
               " INNER JOIN Clientes AS Empresa " & vbCrLf & _
               "    ON P.Empresa_Id    = Empresa.Cliente_Id" & vbCrLf & _
               "   AND P.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
               " INNER JOIN Clientes" & vbCrLf & _
               "    ON P.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
               "   AND P.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
               " Inner Join Operacoes op " & vbCrLf & _
               "    on op.Operacao_Id     = P.Operacao" & vbCrLf & _
               " Inner Join Suboperacoes SO " & vbCrLf & _
               "    on SO.Operacao_Id     = P.Operacao" & vbCrLf & _
               "   and SO.SubOperacoes_Id = P.SubOperacao" & vbCrLf & _
               " Inner Join Moedas M" & vbCrLf & _
               "    on M.Moeda_id = P.Moeda" & vbCrLf & _
               "  Left Join ( " & vbCrLf & _
               "             Select PP.Empresa_Id," & vbCrLf & _
               "                    PP.EndEmpresa_Id," & vbCrLf & _
               "                    PP.Pedido_Id," & vbCrLf & _
               "                    Sum(PxE2.ValorOficial) as OficialDestino," & vbCrLf & _
               "                    Sum(PxE2.ValorMoeda) as MoedaDestino" & vbCrLf & _
               "               from Pedidos PP" & vbCrLf & _
               "              inner Join PedidosxEncargos PxE2" & vbCrLf & _
               "                 on PP.EmpresaTroca    = PxE2.Empresa_Id" & vbCrLf & _
               "                And PP.EndEmpresaTroca = PxE2.EndEmpresa_Id" & vbCrLf & _
               "                And PP.PedidoTroca     = PxE2.Pedido_Id" & vbCrLf & _
               "                and PxE2.Encargo_Id    = 'LIQUIDO'" & vbCrLf & _
               "              Group by PP.Empresa_Id, PP.EndEmpresa_Id, PP.Pedido_Id" & vbCrLf & _
               "          ) as sb" & vbCrLf & _
               "   on P.Empresa_id      = sb.Empresa_Id" & vbCrLf & _
               "  and P.EndEmpresa_Id   = sb.EndEmpresa_Id" & vbCrLf & _
               "  and P.Pedido_Id       = sb.Pedido_Id" & vbCrLf & _
               "  where isnull(P.Troca,0) = 1" & vbCrLf & _
               "    and op.classe         = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
               "    And P.situacao        = 1" & vbCrLf

        If Empresa.Length > 0 Then
            Sql &= " And P.Empresa_Id ='" & Empresa & "'" & vbCrLf
            Sql &= " And P.EndEmpresa_Id =" & EndEmpresa & vbCrLf
        End If

        If Cliente.Length > 0 Then
            Sql &= " And P.Cliente ='" & Cliente & "'" & vbCrLf
            Sql &= " And P.EndCliente =" & EndCliente & vbCrLf
        End If

        If Moeda.Length > 0 Then Sql &= " And P.Moeda =" & Moeda & vbCrLf
        If Safra.Length > 0 Then Sql &= " And P.Safra  ='" & Safra & "'" & vbCrLf
        If Where.Length > 0 Then Sql &= " And " & Where & vbCrLf

        Sql &= "  group by P.Empresa_Id, P.EndEmpresa_Id, Empresa.Nome, P.Safra, " & vbCrLf & _
               "         P.DataPedido, P.Pedido_Id, P.PedidoEfetivo, P.Cliente, P.EndCliente,  Clientes.Nome," & vbCrLf & _
               "         Clientes.Complemento, P.Moeda, M.Descricao, isnull(sb.OficialDestino,0), isnull(sb.MoedaDestino,0); " & vbCrLf & _
               "select Empresa_Id, EndEmpresa_Id, Nome," & vbCrLf & _
               "       Safra, DataPedido, Pedido_Id, PedidoEfetivo," & vbCrLf & _
               "       Cliente, EndCliente, ClienteNome," & vbCrLf & _
               "       Complemento, Moeda, DescMoeda," & vbCrLf & _
               "       Origem, Destino, Origem - Destino as Saldo" & vbCrLf & _
               "  from #Temp" & vbCrLf & _
               " where Origem <> Destino;" & vbCrLf & _
               " drop table #temp;" & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(Sql, "Pedido")
        gridPedido.DataSource = ds
        gridPedido.DataBind()
    End Sub
    Protected Overrides Sub Selecionar()
        Try
            If gridPedido.SelectedRow Is Nothing Then
                MsgBox(Me.Page, "Nenhum Pedido foi selecionado.")
            Else
                Dim objPedido As New [Lib].Negocio.Pedido(gridPedido.SelectedRow.Cells(1).Text, gridPedido.SelectedRow.Cells(2).Text, gridPedido.SelectedRow.Cells(6).Text)
                Session(Session("ssTipoRetorno")) = objPedido
                If Session("ssTipoRetorno") IsNot Nothing Then
                    If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                        Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, MainUserControl.ClientID.Split("_")(1)), UserControl)
                        CType(uc, IBaseUserControl).Carregar(objPedido)
                    Else
                        CType(Me.Page, IBasePage).Carregar(objPedido)
                    End If
                    Popup.CloseDialog(Me.Page, "divConsultaPedidoDeTroca")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Selecionar()
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaPedidoDeTroca")
    End Sub

    Protected Sub gridPedido_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridPedido.SelectedIndexChanged
        Dim Ped As New [Lib].Negocio.Pedido(gridPedido.SelectedRow.Cells(1).Text, gridPedido.SelectedRow.Cells(2).Text, gridPedido.SelectedRow.Cells(6).Text)
        GridItens.DataSource = Ped.Itens.ToArray
        GridItens.DataBind()
        divItens.Style.Clear()
    End Sub
End Class