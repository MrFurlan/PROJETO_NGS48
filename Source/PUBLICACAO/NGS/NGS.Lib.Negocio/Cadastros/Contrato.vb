Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio
Imports System.IO

<Serializable()> _
Public Class ListContrato
    Inherits List(Of Contrato)

    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        If CarregarDados Then
            Dim sql As String = "Select Contrato_Id, Descricao, Arquivo from Contrato"
            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Contrato")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Mc As New Contrato
                Mc.Codigo = row("Contrato_Id")
                Mc.Descricao = row("Descricao")
                Mc.Arquivo = DirectCast(row("Arquivo"), Byte())
                Me.Add(Mc)
            Next

        End If
    End Sub

End Class

<Serializable()> _
Public Class Contrato
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Guid)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Contrato_Id, Descricao, Arquivo FROM Contrato WHERE Contrato_Id = '" & pCodigo.ToString & "'"

        ds = Banco.ConsultaDataSet(sql, "Contrato")
        If ds.Tables(0).Rows.Count > 0 Then
            Codigo = ds.Tables(0).Rows(0)("Contrato_Id")
            Descricao = ds.Tables(0).Rows(0)("Descricao")
            Arquivo = DirectCast(ds.Tables(0).Rows(0)("Arquivo"), Byte())
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Guid
    Private _Descricao As String
    Private _Arquivo As Byte()

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

    Public Property Codigo() As Guid
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Guid)
            _Codigo = value
        End Set
    End Property

    'Public Property NomeDoArquivo() As String
    '    Get
    '        Return _NomeDoArquivo
    '    End Get
    '    Set(ByVal value As String)
    '        _NomeDoArquivo = value
    '    End Set
    'End Property

    Public Property Arquivo() As Byte()
        Get
            Return _Arquivo
        End Get
        Set(ByVal value As Byte())
            _Arquivo = value
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

    'Public Function LiberaExclusaoDaMarca() As Boolean
    '    Dim sql As String = String.Empty
    '    Dim Banco As New AcessaBanco

    '    sql = "SELECT 1 " & vbCrLf & _
    '        "    FROM Produtos " & vbCrLf & _
    '        "   WHERE Marca = " & Me.Codigo
    '    If Banco.ConsultaDataSet(sql, "MarcaProduto").Tables(0).Rows.Count > 0 Then
    '        Throw New Exception("Esta Marca já está sendo usada em pelo menos um produto.")
    '        Return False
    '    Else
    '        Return True
    '    End If
    'End Function
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

    Private Function BytesToString(ByVal Input As Byte()) As String
        Dim Result As New System.Text.StringBuilder(Input.Length * 2)
        Dim Part As String
        For Each b As Byte In Input
            Part = Conversion.Hex(b)
            If Part.Length = 1 Then Part = "0" & Part
            Result.Append(Part)
        Next
        Return "0x" & Result.ToString()
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO Contrato(Descricao, Arquivo) " & vbCrLf & _
                      " VALUES ('" & Descricao & "'," & BytesToString(Arquivo) & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Contrato " & vbCrLf & _
                      "    SET Descricao   ='" & Descricao & "'," & vbCrLf & _
                      "        Arquivo     = " & BytesToString(Arquivo) & vbCrLf & _
                      "  WHERE Contrato_Id = '" & Codigo.ToString & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Contrato " & vbCrLf & _
                      "  WHERE Contrato_Id = '" & Codigo.ToString & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region


End Class