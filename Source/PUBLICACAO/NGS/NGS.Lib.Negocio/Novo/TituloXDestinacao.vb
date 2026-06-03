Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class TituloXDestinacao
        Implements IBaseEntity

#Region "Contrutor"
        Public Sub New(ByVal pTitulo As Novo.TituloNovo)
            _Titulo = pTitulo
            Dim sql As String
            sql = "SELECT Titulo_Id, Procuracao, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao" & vbCrLf & _
                  "  FROM TituloXDestinacao " & vbCrLf & _
                  " Where Titulo_id = " & pTitulo.Codigo
            Dim banco As New AcessaBanco
            Dim ds As DataSet
            ds = banco.ConsultaDataSet(sql, "TituloXDestinacao")

            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            Dim row As DataRow = ds.Tables(0).Rows(0)
            CodigoProcuracao = row("Procuracao")
            CodigoDestinatario = row("Destinatario")
            EndDestinatario = row("EndDestinatario")
            NomeDoDestinatario = row("NomeDoDestinatario")
            Destinacao = row("Destinacao")
        End Sub
#End Region

#Region "Fields"
        Private _IUD As String = ""
        Private _Titulo As Novo.TituloNovo

        '** Procuracao *************************************************************
        Private _CodigoProcuracao As Integer
        Private _Procuracao As Procuracao

        '** Destinatario ***********************************************************
        Private _CodigoDestinatario As String = ""
        Private _EndDestinatario As Integer
        Private _Destinatario As Cliente
        Private _NomeDoDestinatario As String = ""
        Private _Destinacao As String = ""
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
        Public ReadOnly Property Titulo() As Novo.TituloNovo
            Get
                Return _Titulo
            End Get
        End Property

        '** Procuracao *************************************************************
        Public Property CodigoProcuracao() As Integer
            Get
                Return _CodigoProcuracao
            End Get
            Set(ByVal value As Integer)
                _CodigoProcuracao = value
                _Procuracao = Nothing
            End Set
        End Property
        Public Property Procuracao() As Procuracao
            Get
                If _Procuracao Is Nothing And _CodigoProcuracao > 0 Then _Procuracao = New Procuracao(Titulo.CodigoEmpresa, Titulo.EnderecoEmpresa, _CodigoProcuracao)
                Return _Procuracao
            End Get
            Set(ByVal value As Procuracao)
                _Procuracao = value
            End Set
        End Property

        '** Destinatario ***********************************************************
        Public Property CodigoDestinatario() As String
            Get
                Return _CodigoDestinatario
            End Get
            Set(ByVal value As String)
                _CodigoDestinatario = value
            End Set
        End Property
        Public Property EndDestinatario() As Integer
            Get
                Return _EndDestinatario
            End Get
            Set(ByVal value As Integer)
                _EndDestinatario = value
            End Set
        End Property
        Public Property Destinatario() As Cliente
            Get
                If _Destinatario Is Nothing And _CodigoDestinatario.Length > 0 Then _Destinatario = New Cliente(_CodigoDestinatario, _EndDestinatario)
                Return _Destinatario
            End Get
            Set(ByVal value As Cliente)
                _Destinatario = value
            End Set
        End Property
        Public Property NomeDoDestinatario() As String
            Get
                Return _NomeDoDestinatario
            End Get
            Set(ByVal value As String)
                _NomeDoDestinatario = value
            End Set
        End Property
        Public Property Destinacao() As String
            Get
                Return _Destinacao
            End Get
            Set(ByVal value As String)
                _Destinacao = value
            End Set
        End Property
#End Region

#Region "Methods"
        Public Sub SalvarSql(ByRef sqls As ArrayList)
            Dim strSQL As String = ""
            Select Case Me.IUD
                Case "I"
                    strSQL = "Insert into TituloXDestinacao(Titulo_Id, Procuracao, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao) " & vbCrLf & _
                             " Values(" & Titulo.Codigo & "," & Me.CodigoProcuracao & ",'" & Me.CodigoDestinatario & "'," & Me.EndDestinatario & ",'" & Me.NomeDoDestinatario & "','" & Me.Destinacao & "')"
                Case "U"
                    strSQL = "Update TituloXDestinacao set " & vbCrLf & _
                             "  Procuracao         = " & Me.CodigoProcuracao & vbCrLf & _
                             " ,Destinatario       ='" & Me.CodigoDestinatario & "'" & vbCrLf & _
                             " ,EndDestinatario    = " & Me.EndDestinatario & vbCrLf & _
                             " ,NomeDoDestinatario ='" & Me.NomeDoDestinatario & "'" & vbCrLf & _
                             " ,Destinacao         ='" & Me.Destinacao & "'" & vbCrLf & _
                             " Where Titulo_id = " & Titulo.Codigo
                Case "D"
                    strSQL = "Delete TituloXDestinacao" & vbCrLf & _
                             " Where Titulo_id = " & Titulo.Codigo
            End Select
            If strSQL.Length > 0 Then sqls.Add(strSQL)
        End Sub
#End Region

    End Class
End Namespace