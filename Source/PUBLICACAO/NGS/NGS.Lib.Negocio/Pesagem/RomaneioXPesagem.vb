Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListRomaneioXPesagem
    Inherits List(Of RomaneioXPesagem)

#Region "Field"
    Private _parent As Romaneio
#End Region

#Region "Property"
    Public Property Parent() As Romaneio
        Get
            Return _parent
        End Get
        Set(ByVal value As Romaneio)
            _parent = value
        End Set
    End Property
#End Region

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal pRomaneio As Romaneio)
        _parent = pRomaneio
        Dim sql As String
        sql = "Select Empresa_Id, EndEmpresa_Id, Romaneio_Id, Pesagem_Id, Sequencia_Id" & vbCrLf & _
              "  From RomaneiosXPesagens" & vbCrLf & _
              " Where Empresa_id    ='" & pRomaneio.CodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_id = " & pRomaneio.EnderecoEmpresa & vbCrLf & _
              "   and Romaneio_id   = " & pRomaneio.Codigo & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "TBL")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim RxP As New RomaneioXPesagem
            RxP.CodigoEmpresa = row("Empresa_Id")
            RxP.EndEmpresa = row("EndEmpresa_Id")
            RxP.CodigoRomaneio = row("Romaneio_Id")
            RxP.CodigoPesagem = row("Pesagem_Id")
            RxP.Sequencia = row("Sequencia_Id")
            Me.Add(RxP)
        Next
    End Sub

    Public Sub New(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pCodigoRomaneio As Integer, ByVal pCodigoPesagem As Integer)
        _parent = New Romaneio(pCodigoEmpresa, pEndEmpresa, pCodigoRomaneio)
        Dim sql As String
        sql = "Select Empresa_Id, EndEmpresa_Id, Romaneio_Id, Pesagem_Id, Sequencia_Id" & vbCrLf & _
              "  From RomaneiosXPesagens" & vbCrLf & _
              " Where Empresa_id    ='" & pCodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_id = " & pEndEmpresa & vbCrLf

        If pCodigoRomaneio > 0 Then
            sql &= " and Romaneio_id = " & pCodigoRomaneio & vbCrLf
        End If

        If pCodigoPesagem > 0 Then
            sql &= " and Pesagem_Id = " & pCodigoPesagem & vbCrLf
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "TBL")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim RxP As New RomaneioXPesagem
            RxP.CodigoEmpresa = row("Empresa_Id")
            RxP.EndEmpresa = row("EndEmpresa_Id")
            RxP.CodigoRomaneio = row("Romaneio_Id")
            RxP.CodigoPesagem = row("Pesagem_Id")
            RxP.Sequencia = row("Sequencia_Id")
            Me.Add(RxP)
        Next
    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As RomaneioXPesagem In Me
            If Parent.IUD.Contains("D") Or Parent.IUD = "I" Then item.IUD = Parent.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region
End Class

<Serializable()> _
Public Class RomaneioXPesagem

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByRef pRomaneio As Romaneio, ByRef pPesagem As Pesagem)
        Me.CodigoEmpresa = pRomaneio.CodigoEmpresa
        Me.EndEmpresa = pRomaneio.EnderecoEmpresa
        Me.CodigoRomaneio = pRomaneio.Codigo
        Me.Romaneio = pRomaneio
        Me.CodigoPesagem = pPesagem.Codigo
        Me.Sequencia = pPesagem.Sequencia
        Me.Pesagem = pPesagem
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente
    Private _CodigoRomaneio As Integer
    Private _Romaneio As Romaneio
    Private _CodigoPesagem As Integer
    Private _Sequencia As Integer
    Private _Pesagem As Pesagem
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

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property CodigoRomaneio() As Integer
        Get
            Return _CodigoRomaneio
        End Get
        Set(ByVal value As Integer)
            _CodigoRomaneio = value
            Me.Romaneio = Nothing
        End Set
    End Property

    Public Property Romaneio() As Romaneio
        Get
            If _Romaneio Is Nothing And Me.CodigoEmpresa.Length > 0 And Me.CodigoRomaneio > 0 Then _Romaneio = New Romaneio(Me.CodigoEmpresa, Me.EndEmpresa, Me.CodigoRomaneio)
            Return _Romaneio
        End Get
        Set(ByVal value As Romaneio)
            _Romaneio = value
        End Set
    End Property

    Public Property CodigoPesagem() As Integer
        Get
            Return _CodigoPesagem
        End Get
        Set(ByVal value As Integer)
            _CodigoPesagem = value
            Me.Pesagem = Nothing
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
            Me.Pesagem = Nothing
        End Set
    End Property

    Public Property Pesagem() As Pesagem
        Get
            If _Pesagem Is Nothing And Me.CodigoEmpresa.Length > 0 And Me.CodigoPesagem > 0 Then _Pesagem = New Pesagem(Me.CodigoEmpresa, Me.EndEmpresa, Me.CodigoPesagem)
            Return _Pesagem
        End Get
        Set(ByVal value As Pesagem)
            _Pesagem = value
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
                Sql = " INSERT INTO RomaneiosXPesagens(Empresa_Id, EndEmpresa_Id, Romaneio_Id, Pesagem_Id, Sequencia_Id) " & vbCrLf & _
                      " VALUES ('" & Me.CodigoEmpresa & "'," & Me.EndEmpresa & "," & Me.CodigoRomaneio & "," & Me.CodigoPesagem & "," & Me.Sequencia & ")"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE RomaneiosXPesagens" & vbCrLf & _
                      "  WHERE Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                      "    and EndEmpresa_Id = " & Me.EndEmpresa & vbCrLf & _
                      "    and Romaneio_Id   = " & Me.CodigoRomaneio & vbCrLf & _
                      "    and Pesagem_Id    = " & Me.CodigoPesagem & vbCrLf & _
                      "    and Sequencia_Id  = " & Me.Sequencia
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class