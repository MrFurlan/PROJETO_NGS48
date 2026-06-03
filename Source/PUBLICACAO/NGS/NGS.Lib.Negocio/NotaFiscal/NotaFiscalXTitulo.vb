Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis


<Serializable()> _
Public Class ListNotaFiscalXTitulo
    Inherits List(Of NotaFiscalXTitulo)

#Region "Contrutor"
    Public Sub New(pNotaFiscal As NotaFiscal)
        _NotaFiscal = pNotaFiscal
        Dim sql As String
        sql = "SELECT Titulo_Id" & vbCrLf & _
              "  FROM NotaFiscalXTitulo" & vbCrLf & _
              " Where Empresa_Id      ='" & pNotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_Id   = " & pNotaFiscal.EnderecoEmpresa & vbCrLf & _
              "   and Cliente_Id      ='" & pNotaFiscal.CodigoCliente & "'" & vbCrLf & _
              "   and EndCliente_Id   = " & pNotaFiscal.EnderecoCliente & vbCrLf & _
              "   and EntradaSaida_Id ='" & IIf(pNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "   and Serie_Id        ='" & pNotaFiscal.Serie & "'" & vbCrLf & _
              "   and Nota_Id         = " & pNotaFiscal.Codigo & _
              "   and Nota_id > 0"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NFxT")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim NxT As New NotaFiscalXTitulo(pNotaFiscal, New TituloV(row("Titulo_Id")))
            Me.Add(NxT)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _NotaFiscal As NotaFiscal
#End Region

#Region "Property"
    Public ReadOnly Property Notafiscal As NotaFiscal
        Get
            Return _NotaFiscal
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList

        sqls.Clear()
        Me.SalvarSQL(sqls)

        If sqls.Count = 0 OrElse Banco.GravaBanco(sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSQL(ByRef Sqls As ArrayList)
        For Each item As NotaFiscalXTitulo In Me
            If Notafiscal.IUD = "D" Or Notafiscal.IUD = "I" Then item.IUD = Notafiscal.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

'**************************************************************************************************
'*********************************   Classe Base NotaXTitulo  *************************************
'**************************************************************************************************
<Serializable()> _
Public Class NotaFiscalXTitulo

#Region "Construtor"
    'Usado somente para facilitar a insecao
    Public Sub New(pNota As NotaFiscal, pTitulo As TituloV)
        _Titulo = pTitulo
        _NotaFiscal = pNota
    End Sub

    'Retorna a nota ligada ao titulo se houver
    Public Sub New(pTitulo As TituloV)
        Dim sql As String
        sql = "SELECT Empresa_Id, EndEmpresa_Id," & vbCrLf & _
              "       Cliente_Id, EndCliente_Id," & vbCrLf & _
              "       EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id" & vbCrLf & _
              "  FROM NotaFiscalXTitulo" & vbCrLf & _
              " Where Titulo_Id = " & pTitulo.Codigo

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NFxT")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim row = ds.Tables(0).Rows(0)

        Dim nfConsulta As New NotaFiscal
        nfConsulta.CodigoEmpresa = row("Empresa_Id")
        nfConsulta.EnderecoEmpresa = row("EndEmpresa_Id")
        nfConsulta.CodigoCliente = row("Cliente_Id")
        nfConsulta.EnderecoCliente = row("EndCliente_Id")
        nfConsulta.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
        nfConsulta.Serie = row("Serie_Id")
        nfConsulta.Codigo = row("Nota_Id")

        Me.IUD = "U"
        Me.Titulo = pTitulo
        Me.NotaFiscal = New NotaFiscal(nfConsulta)
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _NotaFiscal As NotaFiscal
    Private _Titulo As TituloV
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public Property NotaFiscal As NotaFiscal
        Get
            Return _NotaFiscal
        End Get
        Set(value As NotaFiscal)
            _NotaFiscal = value
        End Set
    End Property

    Public Property Titulo As TituloV
        Get
            Return _Titulo
        End Get
        Set(value As TituloV)
            _Titulo = value
        End Set
    End Property

    Public ReadOnly Property Existe As Boolean
        Get
            Return Titulo IsNot Nothing
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        SalvarSql(sqls)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function SalvarSql(ByRef sqls As ArrayList) As ArrayList
        Dim strSQL As String = String.Empty
        Dim db As New AcessaBanco

        Select Case Me.IUD
            Case "I"
                strSQL = "INSERT INTO NotaFiscalXTitulo(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Titulo_Id) " & vbCrLf & _
                         " VALUES ('" & NotaFiscal.CodigoEmpresa & "'," & NotaFiscal.EnderecoEmpresa & ",'" & NotaFiscal.CodigoCliente & "'," & NotaFiscal.EnderecoCliente & ",'" & NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "','" & NotaFiscal.Serie & "'," & NotaFiscal.Codigo & "," & Titulo.Codigo & ")"
            Case "U"
                strSQL = "UPDATE NotaFiscalXTitulo " & vbCrLf & _
                         "SET Empresa_Id = '" & NotaFiscal.CodigoEmpresa & "'" & vbCrLf & _
                         " ,EndEmpresa_Id = " & NotaFiscal.EnderecoEmpresa & vbCrLf & _
                         " ,Cliente_Id = '" & NotaFiscal.CodigoCliente & "'" & vbCrLf & _
                         " ,EndCliente_Id = " & NotaFiscal.EnderecoCliente & vbCrLf & _
                         " ,EntradaSaida_Id = '" & NotaFiscal.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                         " ,Serie_Id = '" & NotaFiscal.Serie & "'" & vbCrLf & _
                         " ,Nota_Id = " & NotaFiscal.Codigo & vbCrLf & _
                         "WHERE Titulo_Id = " & Titulo.Codigo
            Case "D"
                strSQL = "Delete NotaFiscalXTitulo " & vbCrLf & _
                         " WHERE Titulo_Id       = " & Titulo.Codigo & vbCrLf
        End Select
        If Not String.IsNullOrWhiteSpace(strSQL) Then sqls.Add(strSQL)
        Return sqls
    End Function
#End Region

End Class

