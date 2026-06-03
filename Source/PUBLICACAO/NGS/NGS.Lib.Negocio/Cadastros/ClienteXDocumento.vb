Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio
Imports System.IO

'***********************************************************************************************************
'*********************************  Lista de Documentos do Cliente  ****************************************
'***********************************************************************************************************
<Serializable()> _
Public Class ListClienteXDocumento
    Inherits List(Of ClienteXDocumento)

#Region "Construtor"
    Public Sub New(ByVal pCliente As Negocio.Cliente)

        _Cliente = pCliente

        If String.IsNullOrWhiteSpace(_Cliente.Codigo) Then Exit Sub

        Dim Banco As New AcessaBanco
        Try
            Dim SQL As String = ""
            SQL = "SELECT Documento_Id, Descricao, NomeDoArquivo, Arquivo " & _
                  "  FROM ClientesXDocumentos " & _
                  " WHERE Cliente_Id    ='" & Me.Cliente.Codigo & "' " & _
                  "   AND EndCliente_Id = " & Me.Cliente.CodigoEndereco.ToString()

            Dim ds As DataSet = Banco.ConsultaDataSet(SQL, "Documentos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim cXd As New Negocio.ClienteXDocumento(Cliente)

                cXd.Codigo = row("Documento_Id")
                cXd.Descricao = row("Descricao")
                cXd.NomeDoArquivo = row("NomeDoArquivo")
                cXd.Arquivo = DirectCast(row("Arquivo"), Byte())

                Me.Add(cXd)
            Next
        Catch ex As Exception
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Field"
    Private _Cliente As Negocio.Cliente
#End Region

#Region "Property"
    Public Property Cliente As Negocio.Cliente
        Get
            Return _Cliente
        End Get
        Set(value As Negocio.Cliente)
            _Cliente = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each pXc As ClienteXDocumento In Me
            pXc.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

'***********************************************************************************************************
'**************************************   Contratos do Pedido   ********************************************
'***********************************************************************************************************
<Serializable()> _
Public Class ClienteXDocumento

#Region "Contrutor"
    Public Sub New(ByVal pCliente As Negocio.Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""

    Private _Cliente As Negocio.Cliente

    Private _Codigo As Integer
    Private _Descricao As String
    Private _NomeDoArquivo As String
    Private _Arquivo As Byte()


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

    Public ReadOnly Property Cliente As Negocio.Cliente
        Get
            Return _Cliente
        End Get
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

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

    Public Property NomeDoArquivo() As String
        Get
            Return _NomeDoArquivo
        End Get
        Set(ByVal value As String)
            _NomeDoArquivo = value
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
                Sql = "INSERT INTO ClientesXDocumentos " & vbCrLf &
                      "(Cliente_Id, EndCliente_Id, Documento_Id, Descricao, NomeDoArquivo, Arquivo)" & vbCrLf &
                      " VALUES('" & Me.Cliente.Codigo & "'," & Me.Cliente.CodigoEndereco & "," & vbCrLf &
                      Me.Codigo & ",'" & Me.Descricao.ToString() & "','" & Me.NomeDoArquivo.ToString() & "'," & FuncoesStrings.BytesToString(Arquivo) & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update ClientesXDocumentos Set" & vbCrLf &
                      "   Descricao     ='" & Me.Descricao.ToString() & "'" & vbCrLf &
                      "  ,NomeDoArquivo ='" & Me.NomeDoArquivo.ToString() & "'" & vbCrLf &
                      "  ,Arquivo       =" & FuncoesStrings.BytesToString(Arquivo) & vbCrLf &
                      " WHERE Cliente_Id     ='" & Me.Cliente.Codigo & "'" & vbCrLf &
                      "   AND EndCliente_Id  = " & Me.Cliente.CodigoEndereco & vbCrLf &
                      "   AND Documento_Id   = " & Me.Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE ClientesXDocumentos " & _
                      " WHERE Cliente_Id     ='" & Me.Cliente.Codigo & "'" & vbCrLf & _
                      "   AND EndCliente_Id  = " & Me.Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Documento_Id   = " & Me.Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class


