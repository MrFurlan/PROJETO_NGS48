Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotaFiscalXDestinoXDescontos
    Inherits List(Of NotaFiscalXDestinoXDescontos)

#Region "Fields"
    Private _Parent As NotaFiscalXItem
#End Region

#Region "Property"
    Public Property Parent() As NotaFiscalXItem
        Get
            Return _Parent
        End Get
        Set(ByVal value As NotaFiscalXItem)
            _Parent = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Parent As NotaFiscalXItem)
        Me.Parent = Parent

        Try
            Dim sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
                                "       Produto_Id, Analise_Id, Percentual, Indice, Desconto " & vbCrLf & _
                                "FROM  NotasXDestinosXDescontos " & vbCrLf & _
                                "Where Empresa_Id      ='" & Me.Parent.NotaFiscal.CodigoEmpresa & "' " & vbCrLf & _
                                "  and EndEmpresa_Id   = " & Me.Parent.NotaFiscal.EnderecoEmpresa & vbCrLf & _
                                "  and Cliente_Id      ='" & Me.Parent.NotaFiscal.CodigoCliente & "' " & vbCrLf & _
                                "  and EndCliente_Id   = " & Me.Parent.NotaFiscal.EnderecoCliente & vbCrLf & _
                                "  and EntradaSaida_Id ='" & Me.Parent.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "' " & vbCrLf & _
                                "  and Serie_Id        ='" & Me.Parent.NotaFiscal.Serie & "' " & vbCrLf & _
                                "  and Nota_id         = " & Me.Parent.NotaFiscal.Codigo & vbCrLf & _
                                "  and Produto_Id      ='" & Me.Parent.CodigoProduto & "' " & vbCrLf & _
                                "  and Nota_id         > 0"

            Dim ds As DataSet
            Dim Banco As New AcessaBanco
            ds = Banco.ConsultaDataSet(sql, "NotasXDestinosXDescontos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objNFXD As New NotaFiscalXDestinoXDescontos(Me.Parent)
                objNFXD.CodigoAnalise = row("Analise_Id")
                objNFXD.Percentual = row("Percentual")
                objNFXD.Indice = row("Indice")
                objNFXD.Desconto = row("Desconto")
                Me.Add(objNFXD)
            Next
        Catch ex As Exception
            HttpContext.Current.Session("ssMessage") = "Message: " & Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString)
        End Try
    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As NotaFiscalXDestinoXDescontos In Me
            If Parent.IUD = "D" Or Parent.IUD = "I" Then item.IUD = Parent.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class NotaFiscalXDestinoXDescontos

#Region "Fields"
    Private _NotaItem As NotaFiscalXItem
    Private _IUD As String
    Private _CodigoAnalise As Integer
    Private _Percentual As Decimal
    Private _Indice As Decimal
    Private _Desconto As Decimal
#End Region

#Region "Property"
    Public Property NotaItem() As NotaFiscalXItem
        Get
            Return _NotaItem
        End Get
        Set(ByVal value As NotaFiscalXItem)
            _NotaItem = value
        End Set
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoAnalise() As Integer
        Get
            Return _CodigoAnalise
        End Get
        Set(ByVal value As Integer)
            _CodigoAnalise = value
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

#Region "Construtor"
    Public Sub New(ByVal Item As NotaFiscalXItem)
        Me.NotaItem = Item
    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case IUD
            Case "I"
                sql = "INSERT INTO NotasXDestinosXDescontos (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, " & vbCrLf & _
                                                             "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
                                                             "Produto_Id,  CFOP_ID, Sequencia_Id, Analise_Id, " & vbCrLf & _
                                                             " Percentual, Indice, Desconto)" & vbCrLf & _
                                                  "  VALUES ('" & Me.NotaItem.NotaFiscal.CodigoEmpresa & "', " & Me.NotaItem.NotaFiscal.EnderecoEmpresa & ", " & vbCrLf & _
                                                  "'" & Me.NotaItem.NotaFiscal.CodigoCliente & "', " & Me.NotaItem.NotaFiscal.EnderecoCliente & ", " & vbCrLf & _
                                                  "'" & Me.NotaItem.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "', '" & Me.NotaItem.NotaFiscal.Serie & "', " & Me.NotaItem.NotaFiscal.Codigo & ", " & vbCrLf & _
                                                  "'" & Me.NotaItem.CodigoProduto & "', " & Me.NotaItem.CFOP & "," & Me.NotaItem.Sequencia & "," & Me.CodigoAnalise & "," & vbCrLf & _
                                                  "" & Str(Me.Percentual) & ", " & Str(Me.Indice) & "," & Str(Me.Desconto) & ")"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotasXDestinosXDescontos " & vbCrLf & _
                      " Where Empresa_Id      ='" & Me.NotaItem.NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & Me.NotaItem.NotaFiscal.EnderecoEmpresa & vbCrLf & _
                      "	  and Cliente_Id      ='" & Me.NotaItem.NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & Me.NotaItem.NotaFiscal.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & Me.NotaItem.NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "   and Nota_Id         = " & Me.NotaItem.NotaFiscal.Codigo & vbCrLf & _
                      "   and Serie_Id        ='" & Me.NotaItem.NotaFiscal.Serie & "'" & vbCrLf & _
                      "   and Produto_Id      ='" & Me.NotaItem.CodigoProduto & "'" & vbCrLf & _
                      "   and CFOP_Id         ='" & Me.NotaItem.CFOP & "'" & vbCrLf & _
                      "   and Sequencia_Id    ='" & Me.NotaItem.Sequencia & "'" & vbCrLf & _
                      "   and Analise_Id      = " & Me.CodigoAnalise & ";"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class