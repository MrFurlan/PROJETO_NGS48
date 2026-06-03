Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXReceitasDespesas
    Inherits List(Of ClienteXReceitasDespesas)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "SELECT CxRD.Cliente_Id, CxRD.Endereco_Id, CxRD.Ano_Id, CxRD.Sequencia_Id, CxRD.ReceitaDespesa, " & vbCrLf & _
              "       CxRD.TipoReceitaDespesa, CxRD.Descricao, CxRD.ValorAno" & vbCrLf & _
              "  FROM ClientesXReceitasDespesas AS CxRD " & vbCrLf & _
              " Where CxRD.Cliente_Id    ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and CxRD.Endereco_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXReceitasDespesas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxRD As New ClienteXReceitasDespesas(pcliente)

            CxRD.Ano = row("Ano_Id")
            CxRD.Sequencia = row("Sequencia_Id")
            CxRD.ReceitaDespesa = row("ReceitaDespesa")
            CxRD.TipoReceitaDespesa = row("TipoReceitaDespesa")
            CxRD.Descricao = row("Descricao")
            CxRD.ValorAno = row("ValorAno")
            Me.Add(CxRD)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Cliente As Cliente
#End Region

#Region "Property"
    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each CxRD As ClienteXReceitasDespesas In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxRD.IUD = _Cliente.IUD
            If CxRD.IUD <> "" Then
                CxRD.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXReceitasDespesas
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _Ano As Integer
    Private _Sequencia As Integer
    Private _ReceitaDespesa As String
    Private _TipoReceitaDespesa As String
    Private _Descricao As String
    Private _ValorAno As String
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

    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property ReceitaDespesa() As String
        Get
            Return _ReceitaDespesa
        End Get
        Set(ByVal value As String)
            _ReceitaDespesa = value
        End Set
    End Property

    Public Property TipoReceitaDespesa() As String
        Get
            Return _TipoReceitaDespesa
        End Get
        Set(ByVal value As String)
            _TipoReceitaDespesa = value
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

    Public Property ValorAno() As String
        Get
            Return _ValorAno
        End Get
        Set(ByVal value As String)
            _ValorAno = value
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

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""

        Select Case _IUD
            Case "I"
                sql = "Insert into ClientesXReceitasDespesas(Cliente_Id, Endereco_Id, Ano_Id, Sequencia_Id, ReceitaDespesa, TipoReceitaDespesa, Descricao, ValorAno)" & vbCrLf & _
                      " values('" & Me.Cliente.Codigo & "'," & Me.Cliente.CodigoEndereco & "," & Me.Ano & "," & Me.Sequencia & ",'" & Me.ReceitaDespesa & "','" & Me.TipoReceitaDespesa & "','" & _Descricao & "'," & Str(Me.ValorAno) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClientesXReceitasDespesas set " & vbCrLf & _
                      "    ReceitaDespesa     ='" & Me.ReceitaDespesa & "'" & vbCrLf & _
                      "   ,TipoReceitaDespesa ='" & Me.TipoReceitaDespesa & "'" & vbCrLf & _
                      "   ,Descricao          ='" & Me.Descricao & "'" & vbCrLf & _
                      "   ,ValorAno           = " & Str(Me.ValorAno) & vbCrLf & _
                      " Where Cliente_Id    ='" & Me.Cliente.Codigo & "'" & vbCrLf & _
                      "   and Endereco_Id   = " & Me.Cliente.CodigoEndereco & vbCrLf & _
                      "   and Ano_Id        = " & Me.Ano & vbCrLf & _
                      "   and Sequencia_Id  = " & Me.Sequencia
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClientesXReceitasDespesas " & vbCrLf & _
                      " Where Cliente_Id   ='" & Me.Cliente.Codigo & "'" & vbCrLf & _
                      "   and Endereco_Id  = " & Me.Cliente.CodigoEndereco & vbCrLf & _
                      "   and Ano_Id       = " & Me.Ano & vbCrLf & _
                      "   and Sequencia_Id = " & Me.Sequencia
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class