Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListFaixaDeComissao
    Inherits List(Of FaixaDeComissao)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pTabela As TabelaDeComissao)
        Dim objBanco As New AcessaBanco()


        Dim strSQL As String = "SELECT FaixaInicial_Id, FaixaFinal, Indice, Observacoes " & _
                               "  FROM FaixaDeComissao " & _
                               " Where Tabela_id = " & pTabela.Codigo


        Dim dsFxC As DataSet = objBanco.ConsultaDataSet(strSQL, "FaixaDeComissao")

        For Each drFxC As DataRow In dsFxC.Tables(0).Rows
            Dim objFxC As New FaixaDeComissao

            objFxC.FaixaInicial = drFxC("FaixaInicial_Id")
            objFxC.FaixaFinal = drFxC("FaixaFinal")
            objFxC.Indice = drFxC("Indice")
            objFxC.Observacoes = drFxC("Observacoes")
            Me.Add(objFxC)
        Next
    End Sub
#End Region

#Region "Fields"

#End Region

#Region "Property"

#End Region

#Region "Methods"

#End Region

End Class

<Serializable()> _
Public Class FaixaDeComissao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pTabelaComissao As TabelaDeComissao)
        _Parent = pTabelaComissao
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _FaixaInicial As Decimal
    Private _FaixaFinal As Decimal
    Private _Indice As Decimal
    Private _Observacoes As String = ""
    Private _Parent As TabelaDeComissao
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

    Public Property FaixaInicial() As Decimal
        Get
            Return _FaixaInicial
        End Get
        Set(ByVal value As Decimal)
            _FaixaInicial = value
        End Set
    End Property

    Public Property FaixaFinal() As Decimal
        Get
            Return _FaixaFinal
        End Get
        Set(ByVal value As Decimal)
            _FaixaFinal = value
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

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public ReadOnly Property Parent() As TabelaDeComissao
        Get
            Return _Parent
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = "Insert Into FaixaDeComissao(Tabela_Id, FaixaInicial_Id, FaixaFinal, Indice, Observacoes)" & vbCrLf & _
                      " values(" & Parent.Codigo & "," & Str(_FaixaInicial) & ", " & Str(_FaixaFinal) & "," & Str(_Indice) & ", '" & _Observacoes & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update FaixaDeComissao Set  " & _
                      "  FaixaFinal = " & Str(_FaixaFinal) & ", Indice = " & Str(_Indice) & ", Observacoes = '" & _Observacoes & "' " & _
                      " Where Tabela_Id = " & Parent.Codigo & "  and FaixaInicial_Id = " & Str(_FaixaInicial) & " "
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete FaixaDeComissao " & vbCrLf & _
                      " Where Tabela_Id = " & Parent.Codigo & "  and FaixaInicial_Id = " & Str(_FaixaInicial) & ""
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function ConsultarFaixaInicial() As Boolean
        Dim objBanco As New AcessaBanco()

        Dim strSQL As String = "SELECT FaixaInicial_Id, FaixaFinal, Indice, Observacoes " & _
                               "  FROM FaixaDeComissao " & _
                               " Where Tabela_id       = " & Parent.Codigo & _
                               "   And FaixaInicial_Id = " & Str(_FaixaInicial)

        Dim dsFxC As DataSet = objBanco.ConsultaDataSet(strSQL, "FaixaDeComissao")

        If dsFxC Is Nothing OrElse dsFxC.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function
#End Region

End Class