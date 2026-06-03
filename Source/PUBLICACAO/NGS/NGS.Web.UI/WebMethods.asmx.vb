Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Web.Script.Serialization
Imports System.ComponentModel
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class WebMethods
    Inherits System.Web.Services.WebService

    <WebMethod(EnableSession:=True)> _
    Public Function getAutocomplete() As String
        Dim lst As New List(Of Object)
        Dim sql = " Select Processo_Id, Descricao From Processos "
        sql &= " where descricao <> '' and descricao is not null "
        sql &= " Order by Processo_Id "
        Dim ds As DataSet = New [Lib].Negocio.AcessaBanco().ConsultaDataSet(sql, "Processos")
        If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj = New With { _
                     Key .label = row("Descricao").Trim(), _
                     Key .url = row("Processo_Id").Trim() _
                }
                lst.Add(obj)
            Next
        End If
        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(lst)
        Return jsonResult
    End Function

#Region "Seleção de Produto"
    Dim ListGrupo As [Lib].Negocio.ListGrupoProduto
    Private Sub SessaoSalvarGrupo(ByVal HIDSelecaoProduto As String)
        Session("ListGrupo" + HIDSelecaoProduto) = ListGrupo
    End Sub

    Private Sub SessaoRecuperaGrupo(ByVal HIDSelecaoProduto As String)
        If Session("ListGrupo" + HIDSelecaoProduto) Is Nothing Then
            ListGrupo = New [Lib].Negocio.ListGrupoProduto()
        Else
            ListGrupo = Session("ListGrupo" + HIDSelecaoProduto)
        End If
    End Sub

    <WebMethod(EnableSession:=True)> _
    Public Function CarregarNivel(ByVal nivel As Integer, ByVal HWhereProduto As String, ByVal HIDSelecaoProduto As String) As String
        Dim lst As New List(Of Object)
        Dim sql As String = ""
        sql = " SELECT Grupo_id, Descricao " & _
              "   FROM GruposDeEstoques " & _
              "  Where Len(Grupo_Id) = " & nivel & _
              "    and exists (select 1 from produtos where left(grupo," & nivel & ") = grupo_id " & IIf(HWhereProduto.Length > 0, " and " & HWhereProduto, "") & ")" & vbCrLf & _
              "  ORDER BY Grupo_id"

        ListGrupo = New [Lib].Negocio.ListGrupoProduto()
        Dim ds As DataSet = New [Lib].Negocio.AcessaBanco().ConsultaDataSet(sql, "Consulta")
        If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj = New With { _
                    Key .Codigo = row("Grupo_Id").Trim(), _
                    Key .Descricao = row("Descricao").Trim(), _
                    Key .Selecionado = False _
               }
                lst.Add(obj)

                Dim gp As New [Lib].Negocio.GrupoProduto With {
                   .Codigo = row("Grupo_id"), _
                   .Descricao = row("Descricao"), _
                   .Selecionado = False
               }
                ListGrupo.Add(gp)
            Next
        End If

        SessaoSalvarGrupo(HIDSelecaoProduto)

        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(lst)
        Return jsonResult
    End Function


    <WebMethod(EnableSession:=True)> _
    Public Function CarregarProdutosGrupo(ByVal checked As Boolean, ByVal idchk As String, ByVal HWhereProduto As String, ByVal HIDSelecaoProduto As String) As String
        Dim index As Integer = idchk.Split("_")(1)
        Dim lst As New List(Of Object)
        SessaoRecuperaGrupo(HIDSelecaoProduto)
        If checked Then
            ListGrupo(index).Selecionado = True
            ListGrupo(index).Produtos = New ListProduto(ListGrupo(index).Codigo, "", "", "", "", HWhereProduto.ToString, True)
            For Each Produto As Produto In ListGrupo(index).Produtos
                Dim obj = New With { _
                    Key .Codigo = Produto.Codigo.Trim(), _
                    Key .Descricao = Produto.Descricao.Trim()
               }
                lst.Add(obj)
            Next
        Else
            ListGrupo(index).Selecionado = False
            ListGrupo(index).Produtos = Nothing
        End If
        SessaoSalvarGrupo(HIDSelecaoProduto)

        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(lst)
        Return jsonResult
    End Function

#End Region


    <WebMethod(EnableSession:=True)> _
    Public Function calcValor(ByVal unitario As String, ByVal quantidade As String) As String
        Dim vlrUnitario As Decimal = Decimal.Zero
        Dim vlrQuantidade As Decimal = Decimal.Zero


        If Not String.IsNullOrWhiteSpace(unitario) AndAlso IsNumeric(unitario) Then
            vlrUnitario = CDec(unitario)
        End If

        If Not String.IsNullOrWhiteSpace(quantidade) AndAlso IsNumeric(quantidade) Then
            vlrQuantidade = CDec(quantidade)
        End If


        Dim valor As String = String.Format("{0:N2}", (vlrUnitario * vlrQuantidade))
        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(valor)
        Return jsonResult
    End Function

    <WebMethod(EnableSession:=True)> _
    Public Function calcValorEdson(ByVal unitario As String, ByVal quantidade As String, ByVal unitarioFat As String, ByVal quantidadeFat As String, ByVal fatorConversao As String, ByVal Quem As String) As String
        Dim vlrUnitario As Decimal = Decimal.Zero
        Dim vlrQuantidade As Decimal = Decimal.Zero
        Dim vlrUnitarioFat As Decimal = Decimal.Zero
        Dim vlrQuantidadeFat As Decimal = Decimal.Zero
        Dim vlrFatorConversao As Decimal = Decimal.Zero
        Dim valor As String = ""

        If Not String.IsNullOrWhiteSpace(unitario) AndAlso IsNumeric(unitario) Then
            vlrUnitario = CDec(unitario)
        End If

        If Not String.IsNullOrWhiteSpace(quantidade) AndAlso IsNumeric(quantidade) Then
            vlrQuantidade = CDec(quantidade)
        End If

        If Not String.IsNullOrWhiteSpace(unitarioFat) AndAlso IsNumeric(unitarioFat) Then
            vlrUnitarioFat = CDec(unitarioFat)
        End If

        If Not String.IsNullOrWhiteSpace(quantidadeFat) AndAlso IsNumeric(quantidadeFat) Then
            vlrQuantidadeFat = CDec(quantidadeFat)
        End If

        If Not String.IsNullOrWhiteSpace(fatorConversao) AndAlso IsNumeric(fatorConversao) Then
            vlrFatorConversao = CDec(fatorConversao)
        End If

        If vlrFatorConversao = 1 Then
            valor = vlrUnitario.ToString("N10") & ";" & vlrQuantidade.ToString("N4") & ";0,0000000000;0,0000;" & (vlrUnitario * vlrQuantidade).ToString("N2")
            'ElseIf vlrFatorConversao = 1000 Then
            '    valor = vlrUnitario.ToString("N10") & ";" & vlrQuantidade.ToString("N4") & ";0,0000000000;0,0000;" & ((vlrUnitario * vlrQuantidade) * 1000).ToString("N2")
        Else
            Select Case CInt(Quem)
                Case 0
                    vlrUnitario = vlrUnitarioFat * fatorConversao
                    vlrQuantidade = vlrQuantidadeFat * fatorConversao
                Case 1
                    vlrUnitarioFat = vlrUnitario / fatorConversao
                Case 2
                    vlrQuantidadeFat = vlrQuantidade * fatorConversao
                Case 3
                    vlrUnitario = vlrUnitarioFat * fatorConversao
                Case 4
                    vlrQuantidade = vlrQuantidadeFat / fatorConversao
            End Select

            valor = vlrUnitario.ToString("N10") & ";" & vlrQuantidade.ToString("N4") & ";" & vlrUnitarioFat.ToString("N10") & ";" & vlrQuantidadeFat.ToString("N4") & ";" & (vlrUnitario * vlrQuantidade).ToString("N2")
        End If

        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(valor)
        Return jsonResult
    End Function

    <WebMethod(EnableSession:=True)> _
    Public Function calcTotal(ByVal valor As String, ByVal total As String, ByVal casasDecimais As Integer) As String
        Dim vlrValor As Decimal = Decimal.Zero
        Dim vlrTotal As Decimal = Decimal.Zero

        If Not String.IsNullOrWhiteSpace(valor) AndAlso IsNumeric(valor) Then
            vlrValor = CDec(valor)
        End If

        If Not String.IsNullOrWhiteSpace(total) AndAlso IsNumeric(total) Then
            vlrTotal = CDec(total)
        End If

        Dim resultado As String = String.Format("{0:N" & casasDecimais & "}", (vlrTotal + vlrValor))
        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(resultado)
        Return jsonResult
    End Function

    <WebMethod(EnableSession:=True)> _
    Public Function calcDiff(ByVal total As String, ByVal valor As String) As String
        Dim vlrTotal As Decimal = Decimal.Zero
        Dim vlrValor As Decimal = Decimal.Zero

        If Not String.IsNullOrWhiteSpace(total) AndAlso IsNumeric(total) Then
            vlrTotal = CDec(total)
        End If

        If Not String.IsNullOrWhiteSpace(valor) AndAlso IsNumeric(valor) Then
            vlrValor = CDec(valor)
        End If

        Dim resultado As String = String.Format("{0:N2}", (vlrTotal - vlrValor))
        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(resultado)
        Return jsonResult
    End Function

    <WebMethod(EnableSession:=True)> _
    Public Function calcValores(ByVal valores As String) As String
        Dim lstValores As New List(Of Decimal)
        Dim lstVlr As String() = valores.Trim().Split(New Char() {";"c}, StringSplitOptions.RemoveEmptyEntries)

        For Each item As String In lstVlr
            If Not String.IsNullOrWhiteSpace(item) Then
                lstValores.Add(CDec(item))
            End If
        Next

        Dim resultado As String = String.Format("{0:N2}", lstValores.Sum(Function(s) s))
        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(resultado)
        Return jsonResult
    End Function

    <WebMethod(EnableSession:=True)>
    Public Function getEmpresa() As String
        Dim lst As New List(Of Object)
        Dim sql = "SELECT unid.Cliente_Id, unid.Nome, emp.Cliente_Id, emp.Nome, emp.Cidade, emp.Estado, emp.Reduzido " & vbCrLf & _
                  "FROM gruposXempresas " & vbCrLf & _
                  "INNER JOIN clientes as emp " & vbCrLf & _
                  " ON GruposXEmpresas.Cliente_Id    = emp.Cliente_Id " & vbCrLf & _
                  " AND GruposXEmpresas.EndCliente_Id = emp.Endereco_Id " & vbCrLf & _
                  "INNER JOIN clientes as unid " & vbCrLf & _
                  " ON GruposXEmpresas.Empresa_Id    = unid.Cliente_Id " & vbCrLf & _
                  " AND GruposXEmpresas.EndEmpresa_Id = unid.Endereco_Id "
        Dim ds As DataSet = New [Lib].Negocio.AcessaBanco().ConsultaDataSet(sql, "Processos")
        If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj = New With { _
                     Key .label = row("Nome").Trim(), _
                     Key .url = row("Reduzido").Trim() _
                }
                lst.Add(obj)
            Next
        End If
        Dim oSerializer As New JavaScriptSerializer()
        Dim jsonResult As String = oSerializer.Serialize(lst)
        Return jsonResult
    End Function

End Class