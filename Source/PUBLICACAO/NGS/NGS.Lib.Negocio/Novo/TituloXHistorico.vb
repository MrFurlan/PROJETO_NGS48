Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class ListTituloXHistorico
        Inherits List(Of TituloXHistorico)

        Public Sub New(ByRef pTitulo As Novo.TituloNovo)
            Dim sql As String = ""

            sql = "SELECT Titulo_Id, Usuario, Acao, Data, Historico " & vbCrLf & _
                  "  FROM TitulosXHistorico" & vbCrLf & _
                  " WHERE Titulo_id = " & pTitulo.Codigo & vbCrLf & _
                  " ORDER BY Data DESC " & vbCrLf

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "TxH")

            For Each row In ds.Tables(0).Rows
                Dim His As New TituloXHistorico(pTitulo)
                His.Usuario = row("Usuario")
                His.Acao = row("Acao")
                His.Data = row("Data")
                His.Historico = row("Historico")
                Me.Add(His)
            Next
        End Sub

#Region "Fields"
        Private _Titulo As Novo.TituloNovo
#End Region

#Region "Property"
        Public ReadOnly Property Titulo As Novo.TituloNovo
            Get
                Return _Titulo
            End Get
        End Property

        Public ReadOnly Property SequenciaLivre As Integer
            Get
                Return Me.Max(Function(s) s.Sequencia)
            End Get
        End Property
#End Region

    End Class

    <Serializable()> _
    Public Class TituloXHistorico

#Region "Construtor"
        Public Sub New()

        End Sub

        Public Sub New(ByRef pTitulo As Novo.TituloNovo)
            _CodigoTitulo = pTitulo.Codigo
            _Titulo = pTitulo
        End Sub
#End Region

#Region "Fields"
        Private _IUD As String
        Private _CodigoTitulo As Integer
        Private _Titulo As Novo.TituloNovo
        Private _Sequencia As Integer
        Private _Usuario As String = ""
        Private _Acao As String = ""
        Private _Data As DateTime
        Private _Historico As String = ""
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

        Public Property CodigoTitulo As Integer
            Get
                Return _CodigoTitulo
            End Get
            Set(value As Integer)
                _CodigoTitulo = value
                _Titulo = Nothing
            End Set
        End Property

        Public Property Titulo As Novo.TituloNovo
            Get
                If _Titulo Is Nothing And _CodigoTitulo > 0 Then _Titulo = New Novo.TituloNovo(_CodigoTitulo)
                Return _Titulo
            End Get
            Set(value As Novo.TituloNovo)
                _Titulo = value
            End Set
        End Property

        Public Property Sequencia As Integer
            Get
                Return _Sequencia
            End Get
            Set(value As Integer)
                _Sequencia = value
            End Set
        End Property

        Public Property Usuario As String
            Get
                Return _Usuario
            End Get
            Set(value As String)
                _Usuario = value
            End Set
        End Property

        Public Property Acao As String
            Get
                Return _Acao
            End Get
            Set(value As String)
                _Acao = value
            End Set
        End Property

        Public Property Data As DateTime
            Get
                Return _Data
            End Get
            Set(value As DateTime)
                _Data = value
            End Set
        End Property

        Public Property Historico As String
            Get
                Return _Historico
            End Get
            Set(value As String)
                _Historico = value
            End Set
        End Property
#End Region

#Region "Methods"
        Public Function Salvar() As Boolean
            Dim Banco As New AcessaBanco
            Dim sqls As New ArrayList
            SalvarSql(sqls)
            Return Banco.GravaBanco(sqls)
        End Function

        Public Sub SalvarSql(ByRef sqls As ArrayList)
            Dim sql As String = String.Empty
            Select Case Me.IUD
                Case "I"
                    sql = " INSERT INTO TitulosXHistorico(Titulo_Id, Usuario, Acao, Data, Historico) " & vbCrLf & _
                  " VALUES (" & Me.CodigoTitulo & ",'" & Me.Usuario & "','" & Me.Acao & "','" & Data.ToString("yyyy-MM-dd HH:mm:ss") & "','" & Me.Historico & "')"
                Case "D"
                    sql = "DELETE FROM TitulosXHistorico WHERE Titulo_Id = " & Me.CodigoTitulo
            End Select
            If Not String.IsNullOrWhiteSpace(sql) Then sqls.Add(sql)
        End Sub
#End Region
    End Class
End Namespace
