Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListReceita
    Inherits List(Of Receita)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pNota As NotaFiscal)
        Dim sql As String
        sql = "SELECT Receita_Id, Cultura, ART, ARTReceita, DataReceita, Status, RespTecnico, EndRespTecnico" & vbCrLf & _
              "  FROM Receita" & vbCrLf & _
              " Where Receita_Id in (Select distinct isnull(nfi.receita,0)" & vbCrLf & _
              "                        From NotasFiscaisxItens nfi" & vbCrLf & _
              "                       Where nfi.Empresa_Id      ='" & pNota.CodigoEmpresa & "'" & vbCrLf & _
              "                         and nfi.EndEmpresa_Id   = " & pNota.EnderecoEmpresa & vbCrLf & _
              "                         and nfi.Cliente_Id      ='" & pNota.CodigoCliente & "'" & vbCrLf & _
              "                         and nfi.EndCliente_Id   = " & pNota.EnderecoCliente & vbCrLf & _
              "                         and nfi.EntradaSaida_Id ='" & pNota.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
              "                         and nfi.Nota_Id         = " & pNota.Codigo & vbCrLf & _
              "                         and nfi.Serie_Id        ='" & pNota.Serie & "')"



        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Receitas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Rec As New Receita
            Rec.CodigoReceita = row("Receita_Id")
            Rec.CodigoCultura = row("Cultura")
            Rec.CodigoArt = row("ART")
            Rec.NumReceita = row("ARTReceita")
            Rec.DataReceita = row("DataReceita")
            Rec.Status = row("Status")
            Rec.CodigoRespTecnico = row("RespTecnico")
            Rec.EndRespTecnico = row("EndRespTecnico")
        Next
    End Sub

End Class

<Serializable()> _
Public Class Receita

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoReceita As Integer)
        Dim sql As String
        sql = "SELECT Receita_Id, Cultura, ART, ARTReceita, DataReceita, Status, RespTecnico, EndRespTecnico" & vbCrLf & _
              "  FROM Receita" & vbCrLf & _
              " Where Receita_Id = " & pCodigoReceita

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Receita")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _CodigoReceita = row("Receita_Id")
            _CodigoCultura = row("Cultura")
            _CodigoArt = row("ART")
            _NumReceita = row("ARTReceita")
            _DataReceita = row("DataReceita")
            _Status = row("Status")
            _CodigoRespTecnico = row("RespTecnico")
            _EndRespTecnico = row("EndRespTecnico")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoReceita As Integer

    Private _Cultura As Cultura
    Private _CodigoCultura As Integer

    Private _DataReceita As DateTime

    Private _ART As ART
    Private _CodigoArt As String

    Private _NumReceita As Integer
    Private _Status As Boolean

    Private _RespTecnico As Cliente
    Private _CodigoRespTecnico As String
    Private _EndRespTecnico As Integer

    Private _Produtos As ListReceitaXProduto
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

    Public Property CodigoReceita() As Integer
        Get
            Return _CodigoReceita
        End Get
        Set(ByVal value As Integer)
            _CodigoReceita = value
        End Set
    End Property

    Public ReadOnly Property Cultura() As Cultura
        Get
            If _Cultura Is Nothing And _CodigoCultura > 0 Then _Cultura = New Cultura(_CodigoCultura)
            Return _Cultura
        End Get
    End Property

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
            _Cultura = Nothing
        End Set
    End Property

    Public Property DataReceita() As DateTime
        Get
            Return _DataReceita
        End Get
        Set(ByVal value As DateTime)
            _DataReceita = value
        End Set
    End Property

    Public ReadOnly Property ART() As ART
        Get
            If _ART Is Nothing And _CodigoArt > 0 Then _ART = New ART(_CodigoArt)
            Return _ART
        End Get
    End Property

    Public Property CodigoArt() As String
        Get
            Return _CodigoArt
        End Get
        Set(ByVal value As String)
            _CodigoArt = value
            _ART = Nothing
        End Set
    End Property

    Public Property NumReceita() As Integer
        Get
            Return _NumReceita
        End Get
        Set(ByVal value As Integer)
            _NumReceita = value
        End Set
    End Property

    Public Property Status() As Boolean
        Get
            Return _Status
        End Get
        Set(ByVal value As Boolean)
            _Status = value
        End Set
    End Property

    Public ReadOnly Property RespTecnico() As Cliente
        Get
            If _RespTecnico Is Nothing And _CodigoRespTecnico.Length > 0 Then _RespTecnico = New Cliente(_CodigoRespTecnico, _EndRespTecnico)
            Return _RespTecnico
        End Get
    End Property

    Public Property CodigoRespTecnico() As String
        Get
            Return _CodigoRespTecnico
        End Get
        Set(ByVal value As String)
            _CodigoRespTecnico = value
        End Set
    End Property

    Public Property EndRespTecnico() As Integer
        Get
            Return _EndRespTecnico
        End Get
        Set(ByVal value As Integer)
            _EndRespTecnico = value
        End Set
    End Property

    Public Property Produtos() As ListReceitaXProduto
        Get
            If _Produtos Is Nothing Then _Produtos = New ListReceitaXProduto(Me)
            Return _Produtos
        End Get
        Set(ByVal value As ListReceitaXProduto)
            _Produtos = value
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
                Sql = " INSERT INTO RECEITA(Receita_Id, Cultura, ART, ARTReceita, DataReceita, Status, RespTecnico, EndRespTecnico) " & vbCrLf & _
                      " VALUES (" & _CodigoReceita & "," & _CodigoCultura & "," & _CodigoArt & "," & _NumReceita & ",'" & _DataReceita.ToString("yyyy-MM-dd") & "','" & IIf(_Status, "S", "N") & "','" & _CodigoRespTecnico & "'," & _EndRespTecnico & ")"
                Sqls.Add(Sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                Sql = " UPDATE RECEITA SET" & vbCrLf & _
                      "    Cultura        = " & _CodigoCultura & vbCrLf & _
                      "   ,ART            = " & _CodigoArt & vbCrLf & _
                      "   ,ARTReceita     = " & _NumReceita & vbCrLf & _
                      "   ,DataReceita    ='" & _DataReceita.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Status         ='" & IIf(_Status, "S", "N") & "'" & vbCrLf & _
                      "   ,RespTecnico    ='" & _CodigoRespTecnico & "'" & vbCrLf & _
                      "   ,EndRespTecnico = " & _EndRespTecnico & vbCrLf & _
                      "  WHERE Receita_Id = " & _CodigoReceita
                Sqls.Add(Sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                Sql = " DELETE RECEITA" & vbCrLf & _
                      "  WHERE Receita_Id ='" & _CodigoReceita & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Not Produtos Is Nothing Then Produtos.SalvarSql(Sqls)
    End Sub
#End Region

End Class