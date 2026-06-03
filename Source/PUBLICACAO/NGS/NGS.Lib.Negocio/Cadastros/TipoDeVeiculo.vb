Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class TipoDeVeiculo
    Implements IBaseEntity

#Region "Variáveis locais"

    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Descricao As String
    Private _Capacidade As Integer
    Private _TaraMinima As Integer
    Private _TaraMaxima As Integer
    Private _ViaDeTransporte As Integer
    Private _ViaDeTransporteDetalhes As ViaDeTransporte
    Private _CodigoPamcard As Integer

#End Region

#Region "Campos públicos"

    Public Erro As Exception

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Selecionar(Codigo)
    End Sub

#End Region

#Region "Propriedades"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
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

    Public Property Capacidade() As Integer
        Get
            Return _Capacidade
        End Get
        Set(ByVal value As Integer)
            _Capacidade = value
        End Set
    End Property

    Public Property TaraMinima() As Integer
        Get
            Return _TaraMinima
        End Get
        Set(ByVal value As Integer)
            _TaraMinima = value
        End Set
    End Property

    Public Property TaraMaxima() As Integer
        Get
            Return _TaraMaxima
        End Get
        Set(ByVal value As Integer)
            _TaraMaxima = value
        End Set
    End Property

    Public Property ViaDeTransporte() As Integer
        Get
            Return _ViaDeTransporte
        End Get
        Set(ByVal value As Integer)
            _ViaDeTransporte = value
        End Set
    End Property

    Public Property ViaDeTransporteDetalhes() As ViaDeTransporte
        Get
            If _ViaDeTransporteDetalhes Is Nothing And _ViaDeTransporte > 0 Then _ViaDeTransporteDetalhes = New ViaDeTransporte(_ViaDeTransporte)
            Return _ViaDeTransporteDetalhes
        End Get
        Set(ByVal value As ViaDeTransporte)
            _ViaDeTransporteDetalhes = value
        End Set
    End Property

    Public Property CodigoPamcard() As Integer
        Get
            Return _CodigoPamcard
        End Get
        Set(ByVal value As Integer)
            _CodigoPamcard = value
        End Set
    End Property

#End Region

#Region "Métodos"

    Dim Banco As New AcessaBanco

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True

        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Dim dsNum As DataSet
                Dim SqlNumerador As String = "SELECT Max(Codigo_Id) Codigo FROM TiposDeVeiculos "
                dsNum = Banco.ConsultaDataSet(SqlNumerador, "TiposDeVeiculos")
                If dsNum.Tables.Count > 0 Then
                    Me.Codigo = Convert.ToInt32(dsNum.Tables(0).Rows(0)("Codigo")) + 1
                End If
                dsNum.Dispose()

                Sql = " INSERT Into TiposDeVeiculos (Codigo_Id, Descricao, Capacidade, TaraMinima, TaraMaxima, ViaDeTransporte) " & vbCrLf & _
                      " Values(" & Me.Codigo & ", " & vbCrLf & _
                      "'" & UCase(Me.Descricao) & "', " & vbCrLf & _
                      "'" & UCase(Me.Capacidade) & "', " & vbCrLf & _
                      "'" & UCase(Me.TaraMinima) & "', " & vbCrLf & _
                      "'" & UCase(Me.TaraMaxima) & "', " & Me.ViaDeTransporte & ")"

                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE TiposDeVeiculos " & vbCrLf & _
                      "    SET Descricao     = '" & Me.Descricao & "'," & vbCrLf & _
                      "        Capacidade      = '" & Me.Capacidade & "'," & vbCrLf & _
                      "        TaraMinima      = '" & Me.TaraMinima & "'," & vbCrLf & _
                      "        TaraMaxima      = '" & Me.TaraMaxima & "'," & vbCrLf & _
                      "        ViaDeTransporte = " & Me.ViaDeTransporte & vbCrLf & _
                      "  WHERE Codigo_Id       = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE TiposDeVeiculos" & vbCrLf & _
                      "  WHERE Codigo_Id = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco

        Try
            Dim strSql As String = "SELECT Codigo_Id AS Codigo, Descricao, Capacidade, TaraMinima, TaraMaxima, ViaDeTransporte, isnull(CodigoPamcard,0) AS CodigoPamcard " & _
                                   "FROM TiposDeVeiculos " & _
                                   "WHERE Codigo_Id = " & Codigo

            Dim dsTiposDeVeiculos As DataSet = objBanco.ConsultaDataSet(strSql, "TiposDeVeiculos")

            With dsTiposDeVeiculos.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = Convert.ToInt32(.Item(0)("Codigo"))
                    Me.Descricao = .Item(0)("Descricao").ToString
                    Me.Capacidade = Convert.ToInt32(.Item(0)("Capacidade"))
                    Me.TaraMinima = Convert.ToInt32(.Item(0)("TaraMinima"))
                    Me.TaraMaxima = Convert.ToInt32(.Item(0)("TaraMaxima"))
                    Me.ViaDeTransporte = Convert.ToInt32(.Item(0)("ViaDeTransporte"))
                    Me.CodigoPamcard = Convert.ToInt32(.Item(0)("CodigoPamcard"))
                End If
            End With

            Return True


        Catch ex As Exception

            Debug.WriteLine(ex.Message)
            Return False

        End Try
    End Function

#End Region

End Class

<Serializable()> _
Public Class TipoDeVeiculos
    Inherits List(Of TipoDeVeiculo)

#Region "Construtor"
    Public Sub New(Optional ByVal CarregarTiposDeVeiculos As Boolean = False)
        If CarregarTiposDeVeiculos Then
            Dim objBanco As New AcessaBanco
            Dim strSql As String = "SELECT     Codigo_Id, Descricao, Capacidade, TaraMinima, TaraMaxima, ViaDeTransporte, isnull(CodigoPamcard,0) AS CodigoPamcard " & _
                                   "FROM         TiposDeVeiculos " & _
                                   "ORDER BY Codigo_Id "
            Dim ds As DataSet = objBanco.ConsultaDataSet(strSql, "TiposDeVeiculos")

            For Each rom As DataRow In ds.Tables(0).Rows
                Dim tipo As New TipoDeVeiculo
                tipo.Codigo = rom("Codigo_Id")
                tipo.Descricao = rom("Descricao")
                tipo.Capacidade = rom("Capacidade")
                tipo.TaraMinima = rom("TaraMinima")
                tipo.TaraMaxima = rom("TaraMaxima")
                tipo.ViaDeTransporte = rom("ViaDeTransporte")
                tipo.CodigoPamcard = rom("CodigoPamcard")

                Add(tipo)
            Next
        End If
    End Sub
#End Region

End Class