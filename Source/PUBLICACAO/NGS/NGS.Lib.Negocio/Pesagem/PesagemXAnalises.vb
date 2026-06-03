Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'****************************************************************************************************************************************
'*********************************************  LISTA DE PESAGEM X ANALISES   ***********************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class ListPesagemXAnalises
    Inherits List(Of PesagemXAnalises)

#Region "Contrutor"
    Public Sub New(ByVal Parent As Pesagem, Optional ByVal CarregarClassificacaoPadrao As Boolean = False)
        Me.Parent = Parent

        Dim objBanco As New AcessaBanco()

        Dim sql As String
        sql = " Select CLA.Analise_Id," & vbCrLf & _
              "        Analises.Descricao," & vbCrLf & _
              "		   Case" & vbCrLf & _
              "          When 1 = " & IIf(CarregarClassificacaoPadrao, 1, 2) & vbCrLf & _
              "            then isnull(pxa.Percentual,cla.tolerancia)" & vbCrLf & _
              "            Else isnull(pxa.Percentual,0)" & vbCrLf & _
              "        end Percentual," & vbCrLf & _
              "		   isnull(pxa.indice,CLA.indice) as Indice," & vbCrLf & _
              "		   isnull(pxa.Desconto,0) as Desconto" & vbCrLf & _
              "   from classificacoes CLA" & vbCrLf & _
              "   Left Join PesagemXAnalises pxa" & vbCrLf & _
              "     ON pxa.analise_id    = cla.Analise_id" & vbCrLf & _
              "	   and PxA.Empresa_Id    ='" & Me.Parent.CodigoEmpresa & "'" & vbCrLf & _
              "    AND PxA.EndEmpresa_Id = " & Me.Parent.EnderecoEmpresa & vbCrLf & _
              "    AND PxA.Pesagem_Id    = " & Me.Parent.Codigo & vbCrLf & _
              "    AND PxA.Sequencia_Id  = " & Me.Parent.Sequencia & vbCrLf & _
              "   left join Analises " & vbCrLf & _
              "     on CLA.Analise_Id = Analises.Analise_Id " & vbCrLf & _
              "  Where CLA.Tabela_id    = " & Parent.CodigoTabelaDeClassificacao & vbCrLf & _
              "    and CLA.Produto_id   ='" & Parent.CodigoProduto & "'" & vbCrLf & _
              "    and CLA.Sequencia_Id = 1"

        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "PxA")

        'sql = "SELECT PxA.Empresa_Id, PxA.EndEmpresa_Id, PxA.Pesagem_Id, PxA.Sequencia_Id, PxA.Analise_Id, Analises.Descricao, PxA.Percentual, PxA.Indice, PxA.Desconto " & vbCrLf & _
        '      "  FROM PesagemXAnalises PxA " & vbCrLf & _
        '      "  LEFT JOIN Analises " & vbCrLf & _
        '      "    ON PxA.Analise_Id = Analises.Analise_Id " & vbCrLf & _
        '      " WHERE PxA.Empresa_Id    ='" & Me.Parent.CodigoEmpresa & "'" & vbCrLf & _
        '      "   AND PxA.EndEmpresa_Id = " & Me.Parent.EnderecoEmpresa & vbCrLf & _
        '      "   AND PxA.Pesagem_Id    = " & Me.Parent.Codigo & vbCrLf & _
        '      "   AND PxA.Sequencia_Id  = " & Me.Parent.Sequencia & vbCrLf
        'Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "PesagemXAnalises")

        'If ds.Tables(0).Rows.Count = 0 Then
        '    sql = " Select CLA.Analise_Id, Analises.Descricao, " & IIf(CarregarClassificacaoPadrao, "CLA.Tolerancia", "0") & " as Percentual , CLA.indice, 0 as Desconto" & vbCrLf & _
        '          "   from classificacoes CLA " & vbCrLf & _
        '          "      left join Analises " & vbCrLf & _
        '          "             on CLA.Analise_Id = Analises.Analise_Id " & vbCrLf & _
        '          "  Where CLA.Tabela_id    = " & Parent.CodigoTabelaDeClassificacao & vbCrLf & _
        '          "    and CLA.Produto_id   ='" & Parent.CodigoProduto & "'" & vbCrLf & _
        '          "    and CLA.Sequencia_Id = 1 "
        '    '"    and CLA.Sequencia_Id = 1 " & vbCrLf & _
        '    '"    and CLA.FaixaValida  = 'S'" & vbCrLf
        '    ds = objBanco.ConsultaDataSet(sql, "Classificacao")
        'End If

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Analise As New PesagemXAnalises(Me.Parent)
            Analise.CodigoAnalise = row("Analise_Id")
            Analise.Descricao = row("Descricao")
            Analise.Percentual = row("Percentual")
            Analise.Indice = row("Indice")
            Analise.Desconto = row("Desconto")
            Me.Add(Analise)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Parent As Pesagem
#End Region

#Region "Property"
    Public Property Parent() As Pesagem
        Get
            Return _Parent
        End Get
        Set(ByVal value As Pesagem)
            _Parent = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Sqls.Add("Delete PesagemXAnalises " & vbCrLf & _
                 " Where Empresa_Id    ='" & Parent.CodigoEmpresa & "'" & vbCrLf & _
                 "   and EndEmpresa_Id = " & Parent.EnderecoEmpresa & vbCrLf & _
                 "   and Pesagem_Id    = " & Parent.Codigo & "" & vbCrLf & _
                 "   and Sequencia_Id  = " & Parent.Sequencia)

        If Not Parent.IUD.Contains("D") Then
            For Each analise As PesagemXAnalises In Me
                analise.IUD = "I"
                analise.SalvarSql(Sqls)
            Next
        End If
    End Sub

    Public Function CalcularDescontos() As String
        'Mesmo Method em RomaneioxDesconto.vb
        Dim Sql As String = String.Empty
        Dim ds As New DataSet
        Dim Banco As New AcessaBanco
        Dim tolerancia, indice, decIndice As Decimal
        Dim dra As DataRow
        Dim ErroAnalise As String = String.Empty
        Parent.Desconto = 0
        Parent.BrutoBalanca = Math.Abs(Parent.SegundaPesagem - Parent.PrimeiraPesagem)

        For Each Analise In Me
            If Analise.Analise.Opcao.Length = 0 Then
                If (Analise.CodigoAnalise.Equals(1)) Then
                    Sql = "SELECT FaixaInicial, FaixaFinal, Indice, Tolerancia, 0.00 As IndiceDesconto, 0 AS Desconto, FaixaValida " & vbCrLf & _
                          "  FROM Classificacoes " & vbCrLf & _
                          " Where Tabela_Id  = " & Parent.CodigoTabelaDeClassificacao & vbCrLf & _
                          "   And Produto_Id ='" & Parent.CodigoProduto & "'" & vbCrLf & _
                          "   And Analise_Id = " & Analise.CodigoAnalise & vbCrLf & _
                          "   And " & Str(Analise.Percentual) & " BETWEEN FaixaInicial And FaixaFinal "
                Else
                    Sql = "SELECT FaixaInicial, FaixaFinal, Indice, Tolerancia, Indice * (" & Str(Analise.Percentual) & " - Tolerancia) As IndiceDesconto, " & vbCrLf & _
                          "       round(Indice * (" & Str(Analise.Percentual) & " - Tolerancia) * " & Parent.BrutoBalanca & " / 100,0) AS Desconto, FaixaValida " & vbCrLf & _
                          "  FROM Classificacoes " & vbCrLf & _
                          " WHERE Tabela_Id  = " & Parent.CodigoTabelaDeClassificacao & vbCrLf & _
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
                        ds.Tables(0).Rows(0).Item("Desconto") = CStr(CInt((CDec(Parent.BrutoBalanca) * (ds.Tables(0).Rows(0).Item("IndiceDesconto") / 100)) + 0.0000000000001))
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
        Parent.Liquido = Parent.BrutoBalanca - Parent.Desconto
        Return ErroAnalise
    End Function
#End Region

End Class


'****************************************************************************************************************************************
'*******************************************  CLASSE BASE DE PESAGEM X ANALISES   *******************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class PesagemXAnalises

#Region "Fields"
    Private _IUD As String
    Private _Pesagem As Pesagem

    Private _CodigoAnalise As Integer
    Private _Analise As Analise
    Private _Descricao As String = ""
    Private _Percentual As Decimal
    Private _Indice As Decimal
    Private _Desconto As Decimal
#End Region

#Region "Construtor"
    Public Sub New(ByVal Pesagem As Pesagem)
        Me.Pesagem = Pesagem
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

    Public Property Pesagem() As Pesagem
        Get
            Return _Pesagem
        End Get
        Set(ByVal value As Pesagem)
            _Pesagem = value
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

    Public ReadOnly Property Analise As Analise
        Get
            If _Analise Is Nothing And Me.CodigoAnalise > 0 Then _Analise = New Analise(Me.CodigoAnalise)
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

    Public Property Percentual() As Decimal
        Get
            Return _Percentual
        End Get
        Set(ByVal value As Decimal)
            _Percentual = value
        End Set
    End Property

    Public Property Indice() As Decimal
        Get
            Return _Indice
        End Get
        Set(ByVal value As Decimal)
            _Indice = value
        End Set
    End Property

    Public Property Desconto() As Decimal
        Get
            Return _Desconto
        End Get
        Set(ByVal value As Decimal)
            _Desconto = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function SalvarSql(ByRef Sqls As ArrayList)
        If Me.Percentual > 0 Or Len(Me.Analise.Opcao) > 0 Then
            Dim sql As String
            Select Case Me.IUD
                Case "I"
                    sql = "INSERT INTO PesagemXAnalises (Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id, Analise_Id, Percentual, Indice, Desconto)" & vbCrLf & _
                          "VALUES ('" & Pesagem.CodigoEmpresa & "', " & Pesagem.EnderecoEmpresa & ", " & Pesagem.Codigo & ", " & Pesagem.Sequencia & ", " & Me.CodigoAnalise & ", " & Str(Me.Percentual) & ", " & Str(Me.Indice) & ", " & Str(Me.Desconto) & ")"
                    Sqls.Add(sql)
                Case "D"
                    sql = " Delete PesagemXAnalises " & vbCrLf & _
                          " Where  Empresa_Id     ='" & Pesagem.CodigoEmpresa & "'" & vbCrLf & _
                          "   and EndEmpresa_Id   = " & Pesagem.EnderecoEmpresa & vbCrLf & _
                          "   and Pesagem_Id      = " & Pesagem.Codigo & "" & vbCrLf & _
                          "   and Sequencia_Id    = " & Pesagem.Sequencia & vbCrLf & _
                          "   and Analise_Id      = " & Me.CodigoAnalise
                    Sqls.Add(sql)
            End Select
        End If

        Return Sqls
    End Function
#End Region

End Class