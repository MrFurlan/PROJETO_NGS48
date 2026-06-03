Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'****************************************************************************************************************************************
'*******************************************  LISTA DE Romaneios X Descontos   **********************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class ListRomaneioXDesconto
    Inherits List(Of RomaneioXDesconto)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByRef Parent As Romaneio)
        Me.Parent = Parent

        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String

            sql = "Select CLA.Analise_Id," & vbCrLf &
                  "       Analises.Descricao," & vbCrLf &
                  "	      CASE " & vbCrLf &
                  "	        WHEN LEFT(RxD.Empresa_Id, 8) = '05366261' AND CLA.Analise_Id = 3 AND CLA.Produto_id   ='101020001' THEN" & vbCrLf &
                  "	            5" & vbCrLf &
                  "	        ELSE" & vbCrLf &
                  "	            ISNULL(RxD.Percentual, 0)" & vbCrLf &
                  "	      END As Percentual," & vbCrLf &
                  "	      isnull(RxD.indice,CLA.indice) As Indice," & vbCrLf &
                  "	      isnull(RxD.Desconto,0) As Desconto" & vbCrLf &
                  "  from classificacoes CLA" & vbCrLf &
                  "  Left Join RomaneiosXDescontos RxD" & vbCrLf &
                  "    On RxD.analise_id    = cla.Analise_id" & vbCrLf &
                  "   And RxD.Empresa_Id    ='" & Parent.CodigoEmpresa & "'" & vbCrLf &
                  "   AND RxD.EndEmpresa_Id = " & Parent.EnderecoEmpresa & vbCrLf &
                  "   AND RxD.Romaneio_Id   = " & Parent.Codigo & vbCrLf &
                  "  left join Analises" & vbCrLf &
                  "    on CLA.Analise_Id = Analises.Analise_Id" & vbCrLf &
                  " Where CLA.Tabela_id    = " & Parent.Pedido.Itens(0).CodigoClassificacao & vbCrLf &
                  "   and CLA.Produto_id   ='" & Parent.CodigoProduto & "'" & vbCrLf &
                  "   and CLA.Sequencia_Id = 1" & vbCrLf

            Dim dsDescontos As DataSet = objBanco.ConsultaDataSet(sql, "RomaneiosXDescontos")

            For Each drDesconto As DataRow In dsDescontos.Tables(0).Rows
                Dim objDesconto As New RomaneioXDesconto(Parent)
                objDesconto.CodigoAnalise = drDesconto("Analise_Id")
                objDesconto.Descricao = drDesconto("Descricao")
                objDesconto.Percentual = drDesconto("Percentual")
                objDesconto.Indice = drDesconto("Indice")
                objDesconto.Desconto = drDesconto("Desconto")
                Me.Add(objDesconto)
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            objBanco = Nothing
        End Try

    End Sub
#End Region

#Region "field"
    Private _Parent As Romaneio
#End Region

#Region "Property"
    Public Property Parent() As Romaneio
        Get
            Return _Parent
        End Get
        Set(ByVal value As Romaneio)
            _Parent = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Sqls.Add("DELETE RomaneiosXDescontos " & vbCrLf & _
                 " WHERE Empresa_Id    ='" & Parent.CodigoEmpresa & "'" & vbCrLf & _
                 "	 AND EndEmpresa_Id = " & Parent.EnderecoEmpresa & vbCrLf & _
                 "	 AND Romaneio_Id   = " & Parent.Codigo)

        For Each item As RomaneioXDesconto In Me
            If Not Parent.IUD.Contains("D") Then
                item.IUD = "I"
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function CalcularDescontos() As String
        'Mesmo Method em PesagemXAnalises.vb
        Dim Sql As String = String.Empty
        Dim ds As New DataSet
        Dim Banco As New AcessaBanco
        Dim tolerancia, indice, decIndice As Decimal
        Dim dra As DataRow
        Dim ErroAnalise As String = String.Empty
        Parent.Desconto = 0
        Parent.PesoBruto = Math.Abs(Parent.SegundaPesagem - Parent.PrimeiraPesagem)

        For Each Analise In Me
            If Analise.Analise.Opcao.Length = 0 Then
                If (Analise.CodigoAnalise.Equals(1)) Then
                    Sql = "SELECT FaixaInicial, FaixaFinal, Indice, Tolerancia, 0.00 As IndiceDesconto, 0 AS Desconto, FaixaValida " & vbCrLf & _
                          "  FROM Classificacoes " & vbCrLf & _
                          " Where Tabela_Id  = " & Parent.Pedido.Itens(0).CodigoClassificacao & vbCrLf & _
                          "   And Produto_Id ='" & Parent.CodigoProduto & "'" & vbCrLf & _
                          "   And Analise_Id = " & Analise.CodigoAnalise & vbCrLf & _
                          "   And " & Str(Analise.Percentual) & " BETWEEN FaixaInicial And FaixaFinal "
                Else
                    Sql = "SELECT FaixaInicial, FaixaFinal, Indice, Tolerancia, Indice * (" & Str(Analise.Percentual) & " - Tolerancia) As IndiceDesconto, " & vbCrLf & _
                          "       round(Indice * (" & Str(Analise.Percentual) & " - Tolerancia) * " & Parent.PesoBruto & " / 100,0) AS Desconto, FaixaValida " & vbCrLf & _
                          "  FROM Classificacoes " & vbCrLf & _
                          " WHERE Tabela_Id  = " & Parent.Pedido.Itens(0).CodigoClassificacao & vbCrLf & _
                          "   AND Produto_Id ='" & Parent.CodigoProduto & "'" & vbCrLf & _
                          "   AND Analise_Id = " & Analise.CodigoAnalise & vbCrLf & _
                          "   AND " & Str(Analise.Percentual) & " BETWEEN FaixaInicial AND FaixaFinal "
                End If

                ds = Banco.ConsultaDataSet(Sql, "Consulta")

                If (Parent.EntradaSaida.Equals("E") And ds.Tables(0).Rows.Count.Equals(0)) Then
                    ErroAnalise = "Classificação " & Analise.Descricao & " informada não foi encontrada, verifique se informou corretamente ou entre em contato com o Suporte"
                    Exit For
                End If
                If (Not ds.Tables(0).Rows.Count.Equals(0)) Then
                    If (Parent.EntradaSaida.Equals("E") And Not ds.Tables(0).Rows(0).Item("FaixaValida").Equals("S")) Then
                        ErroAnalise = "Faixa digitada da Análise " & Analise.Descricao & " é inválida"
                        Exit For
                    End If
                End If

                If (Not ds.Tables(0).Rows.Count.Equals(0)) Then
                    If Analise.CodigoAnalise = "1" Then
                        Dim j As Integer = 0
                        For Each dra In ds.Tables(0).Rows
                            tolerancia = FormatNumber(dra("FaixaInicial"), 2)
                            indice = dra("Indice")
                            If Analise.Percentual > dra("FaixaFinal") Then
                                decIndice = decIndice + (FormatNumber(dra("FaixaFinal"), 2) - FormatNumber(dra("FaixaInicial"), 2)) * indice
                            Else
                                decIndice = (Analise.Percentual - FormatNumber(dra("Tolerancia"), 2)) * indice
                            End If
                            j += 1
                            If j > 1 Then
                                dra.Delete()
                            End If
                        Next
                        If j > 1 Then
                            ds.AcceptChanges()
                        End If
                        ds.Tables(0).Rows(0).Item("IndiceDesconto") = FormatNumber(decIndice, 2)
                        ds.Tables(0).Rows(0).Item("Desconto") = CStr(CInt((CDec(Parent.PesoBruto) * (ds.Tables(0).Rows(0).Item("IndiceDesconto") / 100)) + 0.0000000000001))
                    End If

                    If Parent.EntradaSaida.Equals("E") Then
                        Analise.Indice = Convert.ToDecimal(ds.Tables(0).Rows(0).Item("IndiceDesconto")).ToString("N2")
                        Analise.Desconto = CInt(ds.Tables(0).Rows(0).Item("Desconto"))
                        Parent.Desconto += Analise.Desconto
                    Else
                        Analise.Indice = FormatNumber(0, 2)
                        Analise.Desconto = 0
                    End If
                End If
            Else
                Analise.Indice = FormatNumber(0, 2)
                Analise.Desconto = 0
            End If
        Next
        Parent.PesoLiquido = Parent.PesoBruto - Parent.Desconto
        Return ErroAnalise
    End Function
#End Region

End Class


'****************************************************************************************************************************************
'*******************************************  CLASSE BASE DE Romaneio X Desconto   ******************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class RomaneioXDesconto

#Region "Fields"
    Private _IUD As String
    Private _Romaneio As Romaneio

    Private _CodigoAnalise As Integer
    Private _Analise As Analise
    Private _Descricao As String = ""
    Private _Percentual As Decimal
    Private _Indice As Decimal
    Private _Desconto As Decimal
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub
    Public Sub New(ByRef pRomaneio As Romaneio)
        Me.Romaneio = pRomaneio
    End Sub
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Romaneio() As Romaneio
        Get
            Return _Romaneio
        End Get
        Set(ByVal value As Romaneio)
            _Romaneio = value
        End Set
    End Property

    Public Property CodigoAnalise() As Integer
        Get
            Return _CodigoAnalise
        End Get
        Set(ByVal value As Integer)
            _CodigoAnalise = value
            _Analise = Nothing
        End Set
    End Property

    Public ReadOnly Property Analise() As Analise
        Get
            If _Analise Is Nothing And _CodigoAnalise > 0 Then _Analise = New Analise(_CodigoAnalise)
            Return _Analise
        End Get
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property Percentual() As Double
        Get
            Return _Percentual
        End Get
        Set(ByVal value As Double)
            _Percentual = value
        End Set
    End Property

    Public Property Indice() As Double
        Get
            Return _Indice
        End Get
        Set(ByVal value As Double)
            _Indice = value
        End Set
    End Property

    Public Property Desconto() As Double
        Get
            Return _Desconto
        End Get
        Set(ByVal value As Double)
            _Desconto = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String = String.Empty
        Select Case Me.IUD
            Case "I"
                strSQL &= " Insert Into Romaneiosxdescontos (Empresa_Id, EndEmpresa_Id, Romaneio_Id, Analise_Id, Desconto, Percentual, Indice) " & vbCrLf & _
                          " Values ('" & Romaneio.CodigoEmpresa & "'," & Romaneio.EnderecoEmpresa & "," & Romaneio.Codigo & "," & Me.CodigoAnalise & "," & Str(Me.Desconto) & "," & Str(Me.Percentual) & "," & Str(Me.Indice) & ")" & vbCrLf
                Sqls.Add(strSQL)
            Case "U"
                strSQL &= "Update Romaneiosxdescontos set" & vbCrLf & _
                          "    Desconto   =" & Str(Me.Desconto) & vbCrLf & _
                          "   ,Percentual =" & Str(Me.Percentual) & vbCrLf & _
                          "   ,Indice     =" & Str(Me.Indice) & vbCrLf & _
                          " Where Empresa_Id    ='" & Romaneio.CodigoEmpresa & "'" & vbCrLf & _
                          "   and EndEmpresa_Id = " & Romaneio.EnderecoEmpresa & vbCrLf & _
                          "   and Romaneio_Id   = " & Romaneio.Codigo & vbCrLf & _
                          "   and Analise_Id    = " & Me.CodigoAnalise
            Case "D"
                strSQL = " Delete Romaneiosxdescontos" & vbCrLf & _
                          " Where Empresa_Id    ='" & Romaneio.CodigoEmpresa & "'" & vbCrLf & _
                          "   and EndEmpresa_Id = " & Romaneio.EnderecoEmpresa & vbCrLf & _
                          "   and Romaneio_Id   = " & Romaneio.Codigo & vbCrLf & _
                          "   and Analise_Id    = " & Me.CodigoAnalise
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class