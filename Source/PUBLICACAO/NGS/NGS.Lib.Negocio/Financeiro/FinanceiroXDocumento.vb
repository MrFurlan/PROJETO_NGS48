Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio
Imports System.IO

'***********************************************************************************************************
'********************************  Lista de Documentos Financeiros  ****************************************
'***********************************************************************************************************
<Serializable()> _
Public Class ListFinanceiroXDocumento
    Inherits List(Of FinanceiroXDocumento)

#Region "Construtor"
    Public Sub New(ByVal pObjTitulo As Negocio.TituloV)
        _ObjTitulo = pObjTitulo
        If _ObjTitulo.Codigo = 0 Then Exit Sub

        Dim Banco As New AcessaBanco
        Try
            Dim SQL As String = ""
            SQL = "SELECT Titulo_Id, Sequencia_Id, Documento_Id, Descricao, NomeDoArquivo, Arquivo " & _
                  "  FROM FinanceiroXDocumentos " & _
                  " WHERE Titulo_Id     = " & Me.ObjTitulo.Codigo.ToString() & vbCrLf & _
                  "   AND Sequencia_Id   = " & Me.ObjTitulo.Sequencia.ToString()

            Dim ds As DataSet = Banco.ConsultaDataSet(SQL, "Contratos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim pXc As New Negocio.FinanceiroXDocumento(ObjTitulo)

                pXc.Codigo = row("Documento_Id")
                pXc.Descricao = row("Descricao")
                pXc.NomeDoArquivo = row("NomeDoArquivo")
                pXc.Arquivo = DirectCast(row("Arquivo"), Byte())

                Me.Add(pXc)
            Next
        Catch ex As Exception
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Field"
    Private _ObjTitulo As Negocio.TituloV
#End Region

#Region "Property"
    Public Property ObjTitulo As Negocio.TituloV
        Get
            Return _ObjTitulo
        End Get
        Set(value As Negocio.TituloV)
            _ObjTitulo = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each fXd As FinanceiroXDocumento In Me
            fXd.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class

'***********************************************************************************************************
'**************************************   Contratos do ObjTitulo   ********************************************
'***********************************************************************************************************
<Serializable()> _
Public Class FinanceiroXDocumento

#Region "Contrutor"
    Public Sub New(ByVal pObjTitulo As Negocio.TituloV)
        _ObjTitulo = pObjTitulo
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""

    Private _ObjTitulo As Negocio.TituloV

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

    Public ReadOnly Property ObjTitulo As Negocio.TituloV
        Get
            Return _ObjTitulo
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

    Public Property Arquivo() As Byte()
        Get
            Return _Arquivo
        End Get
        Set(ByVal value As Byte())
            _Arquivo = value
        End Set
    End Property

#End Region

#Region "Methods"

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
                Sql = "INSERT INTO FinanceiroXDocumentos " & vbCrLf & _
                      "(Titulo_Id, Sequencia_Id, Documento_Id, Descricao, NomeDoArquivo, Arquivo)" & vbCrLf & _
                      " VALUES(" & Me.ObjTitulo.Codigo.ToString() & "," & Me.ObjTitulo.Sequencia.ToString() & "," & Me.Codigo & ",'" & Me.Descricao.ToString() & "','" & Me.NomeDoArquivo.ToString() & "'," & BytesToString(Arquivo) & ")" & vbCrLf
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update FinanceiroXDocumentos Set" & vbCrLf & _
                      "   Descricao     ='" & Me.Descricao.ToString() & "'" & vbCrLf & _
                      "  ,NomeDoArquivo ='" & Me.NomeDoArquivo.ToString() & "'" & vbCrLf & _
                      "  ,Arquivo       =" & BytesToString(Arquivo) & vbCrLf & _
                      " WHERE Titulo_Id      = " & Me.ObjTitulo.Codigo & vbCrLf & _
                      "   AND Sequencia_Id   = " & Me.ObjTitulo.Sequencia & vbCrLf & _
                      "   AND Documento_Id   = " & Me.Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE FinanceiroXDocumentos " & _
                      " WHERE Titulo_Id      = " & Me.ObjTitulo.Codigo & vbCrLf & _
                      "   AND Sequencia_Id   = " & Me.ObjTitulo.Sequencia & vbCrLf & _
                      "   AND Documento_Id   = " & Me.Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class