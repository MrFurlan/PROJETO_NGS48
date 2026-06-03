Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

<Serializable()>
Public Class ListClientesXHistoricoOperacoes
    Inherits List(Of ClientesXHistoricoOperacoes)
    Implements IBaseEntity

    Public Erro As Exception

#Region "Contrutors"
    Public Sub New()

    End Sub
    Public Sub New(Cliente_Id As String, EndCliente_Id As String, Produto_Id As String, Operacao_Id As Integer, SubOperacao_Id As Integer, TipoDocumento As Integer)

        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String
            Sql = "SELECT Cliente_Id, " & vbCrLf &
                    "EndCliente_Id, " & vbCrLf &
                    "Produto_Id, " & vbCrLf &
                    "Operacao_Id, " & vbCrLf &
                    "SubOperacao_Id, " & vbCrLf &
                    "TipoDocumento " & vbCrLf &
                    "FROM ClientesXHistoricoOperacoes " & vbCrLf &
                    "WHERE 1=1 " & vbCrLf

            If Cliente_Id.Length > 0 Then

                Sql &= " AND Cliente_Id = '" & Cliente_Id & "'" & vbCrLf &
                        " AND EndCliente_Id = " & EndCliente_Id & vbCrLf

            End If

            If Produto_Id.Length > 0 Then

                Sql &= " AND Produto_Id = '" & Produto_Id & "'" & vbCrLf

            End If

            If Operacao_Id <> 0 Then

                Sql &= " AND Operacao_Id = " & Operacao_Id & vbCrLf &
                         " AND SubOperacao_Id = " & SubOperacao_Id & vbCrLf

            End If

            If TipoDocumento <> 0 Then

                Sql &= " AND TipoDocumento = " & TipoDocumento & vbCrLf

            End If



            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ClientesXHistoricoOperacoes")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            For Each row In ds.Tables(0).Rows

                Dim historicoOP As New ClientesXHistoricoOperacoes()

                historicoOP.Cliente_Id = row("Cliente_Id")
                historicoOP.EndCliente_Id = row("EndCliente_Id")
                historicoOP.Produto_Id = row("Produto_Id")
                historicoOP.Operacao_Id = row("Operacao_Id")
                historicoOP.SubOperacao_Id = row("SubOperacao_Id")
                historicoOP.TipoDocumento = row("TipoDocumento")

                Me.Add(historicoOP)


            Next
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Pr As ClientesXHistoricoOperacoes In Me
            If Pr.IUD <> "" Then
                Pr.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class


<Serializable()>
Public Class ClientesXHistoricoOperacoes
    Implements IBaseEntity


#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(Cliente_Id As String, EndCliente_Id As Integer, Produto_Id As String, Operacao_Id As Integer, Suboperacao_Id As Integer, TipoDocumento As Integer)
        Dim Banco As New AcessaBanco()
        Try
            Dim Sql As String
            Sql = "SELECT Cliente_Id, " & vbCrLf &
                    "EndCliente_Id, " & vbCrLf &
                    "Produto_Id, " & vbCrLf &
                    "Operacao_Id, " & vbCrLf &
                    "Suboperacao_Id, " & vbCrLf &
                    "TipoDocumento " & vbCrLf &
                    "FROM ClientesXHistoricoOperacoes" & vbCrLf &
                    "WHERE 1=1  " & vbCrLf

            If Cliente_Id.Length > 0 Then

                Sql &= " AND Cliente_Id = '" & Cliente_Id & "'" & vbCrLf &
                         " AND EndCliente_Id = " & EndCliente_Id & vbCrLf

            End If

            If Produto_Id.Length > 0 Then

                Sql &= " AND Produto_Id = '" & Produto_Id & "'" & vbCrLf

            End If

            If Operacao_Id <> 0 Then

                Sql &= " AND Operacao_Id = " & Operacao_Id & vbCrLf &
                         " AND SubOperacao_Id = " & Suboperacao_Id & vbCrLf

            End If

            If TipoDocumento <> 0 Then

                Sql &= " AND TipoDocumento = " & TipoDocumento & vbCrLf

            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "ClientesXHistoricoOperacoes")

            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.Cliente_Id = row("Cliente_Id")
            Me.EndCliente_Id = row("EndCliente_Id")
            Me.Produto_Id = row("Produto_Id")
            Me.Operacao_Id = row("Operacao_Id")
            Me.SubOperacao_Id = row("SubOperacao_Id")
            Me.TipoDocumento = row("TipoDocumento")

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"

    Public Erro As Exception
    Private _IUD As String = String.Empty
    Private _Cliente_Id As String = String.Empty
    Private _EndCliente_Id As Integer = 0
    Private _Produto_Id As String = String.Empty
    Private _Operacao_Id As Integer
    Private _SubOperacao_Id As Integer
    Private _TipoDocumento As Integer

#End Region

#Region "Properties"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Cliente_Id() As String
        Get
            Return _Cliente_Id
        End Get
        Set(ByVal value As String)
            _Cliente_Id = value
        End Set
    End Property

    Public Property EndCliente_Id() As Integer
        Get
            Return _EndCliente_Id
        End Get
        Set(ByVal value As Integer)
            _EndCliente_Id = value
        End Set
    End Property

    Public Property Produto_Id() As String
        Get
            Return _Produto_Id
        End Get
        Set(ByVal value As String)
            _Produto_Id = value
        End Set
    End Property

    Public Property Operacao_Id() As Integer
        Get
            Return _Operacao_Id
        End Get
        Set(ByVal value As Integer)
            _Operacao_Id = value
        End Set
    End Property

    Public Property SubOperacao_Id() As Integer
        Get
            Return _SubOperacao_Id
        End Get
        Set(ByVal value As Integer)
            _SubOperacao_Id = value
        End Set
    End Property

    Public Property TipoDocumento() As Integer
        Get
            Return _TipoDocumento
        End Get
        Set(ByVal value As Integer)
            _TipoDocumento = value
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"

                sql = "INSERT INTO ClientesXHistoricoOperacoes(Cliente_Id, " & vbCrLf
                sql &= "EndCliente_Id, " & vbCrLf
                sql &= "Produto_Id, " & vbCrLf
                sql &= "Operacao_Id, " & vbCrLf
                sql &= "SubOperacao_Id, " & vbCrLf
                sql &= "TipoDocumento " & vbCrLf
                sql &= ") " & vbCrLf

                sql &= "VALUES ('" & Me.Cliente_Id & "', " & vbCrLf
                sql &= "" & Me.EndCliente_Id & ",  " & vbCrLf
                sql &= "'" & Me.Produto_Id & "', " & vbCrLf
                sql &= "" & Me.Operacao_Id & ", " & vbCrLf
                sql &= "" & Me.SubOperacao_Id & ", " & vbCrLf
                sql &= "" & Me.TipoDocumento & " " & vbCrLf
                sql &= " )  " & vbCrLf

                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE ClientesXHistoricoOperacoes SET TipoDocumento = " & Me.TipoDocumento & " " & vbCrLf
                sql &= "WHERE Cliente_Id = '" & Me.Cliente_Id & "' " & vbCrLf
                sql &= "AND   EndCliente_Id = " & Me.EndCliente_Id & " " & vbCrLf
                sql &= "And   Produto_Id = '" & Me.Produto_Id & "' " & vbCrLf
                sql &= "AND   Operacao_Id = " & Me.Operacao_Id & " " & vbCrLf
                sql &= "AND   SubOperacao_Id = " & Me.SubOperacao_Id & "" & vbCrLf

                Sqls.Add(sql)
            Case "D"
                sql = "DELETE FROM ClientesXHistoricoOperacoes " & vbCrLf
                sql &= " WHERE Cliente_Id = '" & Me.Cliente_Id & "' " & vbCrLf
                sql &= " AND   EndCliente_Id = " & Me.EndCliente_Id & " " & vbCrLf
                sql &= " And   Produto_Id = '" & Me.Produto_Id & "' " & vbCrLf
                sql &= " AND   Operacao_Id = " & Me.Operacao_Id & " " & vbCrLf
                sql &= " And   SubOperacao_Id = " & Me.SubOperacao_Id & " " & vbCrLf
                sql &= " AND   TipoDocumento = " & Me.TipoDocumento & "" & vbCrLf

                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
