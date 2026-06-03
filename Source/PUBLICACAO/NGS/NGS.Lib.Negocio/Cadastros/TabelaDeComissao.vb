Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTabelaDeComissao
    Inherits List(Of TabelaDeComissao)

#Region "Construtor"
    Public Sub New(Optional ByVal CarregarDados As Boolean = False, Optional ByVal pSafra As String = "", Optional ByVal Order As String = "Codigo, Descricao")
        If CarregarDados Then
            'Dim Primeiro As New TabelaDeComissao
            'Primeiro.Codigo = 0
            'Primeiro.Descricao = "SELECIONE"
            'Primeiro.CodigoSafra = ""
            'Me.Add(Primeiro)

            Dim objBanco As New AcessaBanco()

            Dim strSQL As String
            strSQL = "SELECT Codigo_Id As Codigo, Descricao, Safra " & _
                     "  FROM TabelaDeComissao "

            If pSafra.Length > 0 Then strSQL &= " Where Safra = '" & pSafra & "'"
            If Order.Length > 0 Then strSQL &= "Order By " & Order

            Dim dsTxC As DataSet = objBanco.ConsultaDataSet(strSQL, "TabelaDeComissao")

            For Each drTxC As DataRow In dsTxC.Tables(0).Rows
                Dim objTxC As New TabelaDeComissao
                objTxC.Codigo = drTxC("Codigo")
                objTxC.Descricao = drTxC("Descricao")
                objTxC.CodigoSafra = drTxC("Safra")
                Me.Add(objTxC)
            Next
        End If
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
Public Class TabelaDeComissao
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Dim objBanco As New AcessaBanco()

        Dim strSQL As String = "SELECT Codigo_Id As Codigo, Descricao, Safra " & _
                                   "FROM TabelaDeComissao " & _
                                   "WHERE Codigo_Id = " & Codigo.ToString()

        Dim dsTxC As DataSet = objBanco.ConsultaDataSet(strSQL, "TabelaDeComissao")

        For Each drTxC As DataRow In dsTxC.Tables(0).Rows
            _Codigo = drTxC("Codigo")
            _Descricao = drTxC("Descricao")
            _CodigoSafra = drTxC("Safra")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
    Private _CodigoSafra As String
    Private _Safra As Safra
    Private _Faixas As ListFaixaDeComissao
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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
            _Safra = Nothing
        End Set
    End Property

    Public ReadOnly Property Safra() As Safra
        Get
            If _Safra Is Nothing And _CodigoSafra.Length > 0 Then _Safra = New Safra(_CodigoSafra)
            Return _Safra
        End Get
    End Property

    Public Property Faixas() As ListFaixaDeComissao
        Get
            If _Faixas Is Nothing And _Codigo > 0 Then _Faixas = New ListFaixaDeComissao(Me)
            Return _Faixas
        End Get
        Set(ByVal value As ListFaixaDeComissao)
            _Faixas = value
        End Set
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
                Sql = "Insert Into TabelaDeComissao(Codigo_Id, Descricao, Safra)" & vbCrLf & _
                      " values(" & _Codigo & ",'" & _Descricao & "','" & _CodigoSafra & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update TabelaDeComissao Set  " & _
                      "  Descricao ='" & _Descricao & "'" & _
                      " ,Safra     ='" & _CodigoSafra & "'" & _
                      " Where Codigo_Id = " & _Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete TabelaDeComissao " & vbCrLf & _
                      " Where Codigo_Id = " & _Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class