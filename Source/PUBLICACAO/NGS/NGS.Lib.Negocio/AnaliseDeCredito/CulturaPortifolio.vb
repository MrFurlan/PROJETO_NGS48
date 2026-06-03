Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCulturaPortifolio
    Inherits List(Of CulturaPortifolio)

#Region "Contrutor"
    Public Sub New(ByVal pParametrosAnaliseDeCreditoCultura As ParametrosAnaliseDeCreditoCultura)
        _Parametro = pParametrosAnaliseDeCreditoCultura

        Dim sql As String
        sql = "SELECT Cp.Grupo_Id, Cp.Quantidade, Cp.Valor" & vbCrLf & _
              "  from CulturaPortifolio Cp" & vbCrLf & _
              " where Cp.Ano_Id = " & pParametrosAnaliseDeCreditoCultura.Parametros.Ano & vbCrLf & _
              "   and Cp.DefinicaoAno_Id =" & pParametrosAnaliseDeCreditoCultura.Parametros.DefinicaoAno & vbCrLf & _
              "   and Cp.Safra_Id ='" & pParametrosAnaliseDeCreditoCultura.CodigoSafra & "'" & vbCrLf & _
              "   and Cp.Cultura_Id =" & pParametrosAnaliseDeCreditoCultura.CodigoCultura & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "CulturaPortifolio")

        If ds.Tables.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Cp As New CulturaPortifolio(pParametrosAnaliseDeCreditoCultura)
                Cp.IUD = "U"
                Cp.CodigoGrupoProduto = row("Grupo_Id")
                Cp.Quantidade = row("Quantidade")
                Cp.Valor = row("Valor")
                Me.Add(Cp)
            Next
        End If
    End Sub
#End Region

#Region "Fields"
    Private _Parametro As ParametrosAnaliseDeCreditoCultura
#End Region

#Region "Property"
    Public ReadOnly Property Parametro() As ParametrosAnaliseDeCreditoCultura
        Get
            Return _Parametro
        End Get
    End Property


    Public ReadOnly Property CustoTotalPortifolio
        Get
            Return (From c In Me Select c.Valor).Sum
        End Get
    End Property

#End Region

#Region "Methods"
    'Public Sub SalvarSql(ByVal Sqls As ArrayList)
    '    For Each cult As CulturaPortifolio In Me
    '        If Parametro.IUD = "D" Or Parametro.IUD = "I" Then cult.IUD = Parametro.IUD
    '        If cult.IUD <> "" Then
    '            cult.SalvarSql(Sqls)
    '        End If
    '    Next
    'End Sub

    Public Function Salvar() As Boolean
        Dim sqls As New ArrayList
        SalvarSql(sqls)

        Dim Banco As New AcessaBanco
        If Banco.GravaBanco(sqls) Then
            For Each row In Me
                row.IUD = ""
            Next
            Return True
        Else
            Return False
        End If
    End Function


    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As CulturaPortifolio In Me
            item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CulturaPortifolio
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pParametrosAnaliseDeCreditoCultura As ParametrosAnaliseDeCreditoCultura)
        _Parametro = pParametrosAnaliseDeCreditoCultura
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _Parametro As ParametrosAnaliseDeCreditoCultura
    'Private _CodigoSafra As String = ""
    'Private _Safra As Safra
    'Private _CodigoCultura As Integer
    'Private _Cultura As Cultura
    Private _CodigoGrupoProduto As String = ""
    Private _GrupoProduto As GrupoProduto
    Private _Quantidade As Decimal
    Private _Valor As Decimal

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

    Public ReadOnly Property ParametrosAnaliseDeCreditoCultura() As ParametrosAnaliseDeCreditoCultura
        Get
            Return _Parametro
        End Get
    End Property

    Public Property CodigoGrupoProduto As String
        Get
            Return _CodigoGrupoProduto
        End Get
        Set(ByVal value As String)
            _CodigoGrupoProduto = value
            _GrupoProduto = Nothing
        End Set
    End Property

    Public ReadOnly Property GrupoProduto As GrupoProduto
        Get
            If _GrupoProduto Is Nothing And CodigoGrupoProduto.Length > 0 Then _GrupoProduto = New GrupoProduto(CodigoGrupoProduto)
            Return _GrupoProduto
        End Get

    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
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

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case IUD
            Case "I"
                sql = " INSERT INTO CulturaPortifolio (Ano_Id, DefinicaoAno_Id, Safra_id, Cultura_Id, Grupo_Id, Quantidade, Valor)" & vbCrLf & _
                      " VALUES ('" & Me.ParametrosAnaliseDeCreditoCultura.Parametros.Ano & "'," & Me.ParametrosAnaliseDeCreditoCultura.Parametros.DefinicaoAno & ",'" & _
                                   Me.ParametrosAnaliseDeCreditoCultura.CodigoSafra & "','" & Me.ParametrosAnaliseDeCreditoCultura.CodigoCultura & "'," & _
                                   Me.CodigoGrupoProduto & "," & Me.Quantidade & ", " & Str(Me.Valor) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = " Update CulturaPortifolio " & vbCrLf & _
                      "    set Quantidade      = " & Str(Me.Quantidade) & "," & vbCrLf & _
                      "        Valor           = " & Str(Me.Valor) & vbCrLf & _
                      "  Where Ano_Id          = " & Me.ParametrosAnaliseDeCreditoCultura.Parametros.Ano & vbCrLf & _
                      "    and DefinicaoAno_Id = " & Me.ParametrosAnaliseDeCreditoCultura.Parametros.DefinicaoAno & vbCrLf & _
                      "    and Safra_Id        ='" & Me.ParametrosAnaliseDeCreditoCultura.CodigoSafra & "'" & vbCrLf & _
                      "    and Cultura_Id      = " & Me.ParametrosAnaliseDeCreditoCultura.CodigoCultura & vbCrLf & _
                      "    and Grupo_Id        ='" & Me.CodigoGrupoProduto & "'"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete CulturaPortifolio " & vbCrLf & _
                      "  Where Ano_Id          = " & Me.ParametrosAnaliseDeCreditoCultura.Parametros.Ano & vbCrLf & _
                      "    and DefinicaoAno_Id = " & Me.ParametrosAnaliseDeCreditoCultura.Parametros.DefinicaoAno & vbCrLf & _
                      "    and Safra_Id        ='" & Me.ParametrosAnaliseDeCreditoCultura.CodigoSafra & "'" & vbCrLf & _
                      "    and Cultura_Id      = " & Me.ParametrosAnaliseDeCreditoCultura.CodigoCultura & _
                      "    and Grupo_Id        ='" & Me.CodigoGrupoProduto & "'"
                Sqls.Add(sql)
        End Select

    End Sub
#End Region

End Class
