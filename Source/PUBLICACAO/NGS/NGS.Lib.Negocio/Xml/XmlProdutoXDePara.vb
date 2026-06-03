Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

<Serializable()> _
Public Class ListXmlProdutoXDePara
    Inherits List(Of XmlProdutoXDePara)
    Implements IBaseEntity

    Public Erro As Exception

#Region "Contrutors"
    Public Sub New()

    End Sub
    Public Sub New(ByVal CodigoCliente As String, ByVal EndCliente As Integer)

        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "SELECT Cliente_Id, EndCliente_Id, ProdutoXML_Id, Produto_Id, ClienteConsolidado, NomeProdutoXML, NCMProdutoXML, UnidadeProdutoXML " & vbCrLf & _
                  "  FROM XmlProdutoXDePara" & vbCrLf & _
                  " WHERE Cliente_Id  = '" & CodigoCliente & "'" & vbCrLf & _
                  "   AND Endereco_Id = " & EndCliente

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "XmlProdutoXDePara")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            For Each row In ds.Tables(0).Rows
                Dim Pr As New XmlProdutoXDePara()

                Pr.CodigoCliente = row("Cliente_Id")
                Pr.EndCliente = row("EndCliente_Id")
                Pr.CodigoProdutoXML = row("ProdutoXML_Id")
                Pr.CodigoProduto = row("Produto_Id")

                Pr.ClienteConsolidado = row("ClienteConsolidado")

                Pr.NomeProdutoXML = row("NomeProdutoXML")
                Pr.NCMProdutoXML = row("NCMProdutoXML")
                Pr.UnidadeProdutoXML = row("UnidadeProdutoXML")
            Next
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub

    Public Sub New(ByVal pNCMProdutoXML As String, ByVal CodigoCliente As String, ByVal EndCliente As Integer, ByVal ProdutoXML_Id As String)

        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "  SELECT Cliente_Id, EndCliente_Id, ProdutoXML_Id, Produto_Id, ClienteConsolidado, NomeProdutoXML, NCMProdutoXML, UnidadeProdutoXML " & vbCrLf &
                  "  FROM XmlProdutoXDePara" & vbCrLf &
                  "  WHERE NCMProdutoXML  = '" & pNCMProdutoXML & "'"

            If ProdutoXML_Id.Length > 0 Then
                Sql &= " OR ( Cliente_Id  = '" & CodigoCliente & "'" & vbCrLf &
                  "     AND EndCliente_Id = " & EndCliente & " AND ProdutoXML_Id = '" & ProdutoXML_Id & "' ) "
            End If

            Sql &= ";"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "XmlProdutoXDePara")

            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            For Each row In ds.Tables(0).Rows

                Dim Pr As New XmlProdutoXDePara()

                Pr.CodigoCliente = row("Cliente_Id")
                Pr.EndCliente = row("EndCliente_Id")
                Pr.CodigoProdutoXML = row("ProdutoXML_Id")
                Pr.CodigoProduto = row("Produto_Id")

                Pr.ClienteConsolidado = row("ClienteConsolidado")

                Pr.NomeProdutoXML = row("NomeProdutoXML")
                Pr.NCMProdutoXML = row("NCMProdutoXML")
                Pr.UnidadeProdutoXML = row("UnidadeProdutoXML")

                Me.Add(Pr)

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
        For Each Pr As XmlProdutoXDePara In Me
            If Pr.IUD <> "" Then
                Pr.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class


<Serializable()> _
Public Class XmlProdutoXDePara
    Implements IBaseEntity


#Region "Constructors"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoCliente As String, ByVal pEndCliente As Integer, Optional ByVal pCodigoProdutoDe As String = "", Optional ByVal pCodigoProdutoPara As String = "", Optional ByVal pNCMProdutoXML As String = "")
        Dim Banco As New AcessaBanco()
        Try
            Dim Sql As String

            Sql = " SELECT TOP 1 Cliente_Id, EndCliente_Id, ProdutoXML_Id, Produto_Id, ClienteConsolidado, NomeProdutoXML, NCMProdutoXML, UnidadeProdutoXML  " & vbCrLf &
                  " FROM XmlProdutoXDePara " & vbCrLf &
                  " WHERE 1=1 " & vbCrLf

            If Not String.IsNullOrWhiteSpace(pCodigoCliente) Then
                Sql &= "   AND Cliente_Id = '" & pCodigoCliente & "'" & vbCrLf
                Sql &= "   AND EndCliente_Id = '" & pEndCliente & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(pCodigoProdutoDe) Then
                Sql &= "   AND ProdutoXML_Id = '" & pCodigoProdutoDe & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(pCodigoProdutoPara) Then
                Sql &= "   AND Produto_Id= '" & pCodigoProdutoPara & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(pNCMProdutoXML) Then
                Sql &= "   AND NCMProdutoXML= '" & pNCMProdutoXML & "'" & vbCrLf
            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "XmlProdutoXDePara")

            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.CodigoCliente = row("Cliente_Id")
            Me.EndCliente = row("EndCliente_Id")
            Me.CodigoProdutoXML = row("ProdutoXML_Id")
            Me.CodigoProduto = row("Produto_Id")

            Me.ClienteConsolidado = row("ClienteConsolidado")

            Me.NomeProdutoXML = row("NomeProdutoXML")
            Me.NCMProdutoXML = row("NCMProdutoXML")
            Me.UnidadeProdutoXML = row("UnidadeProdutoXML")

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
    Private _CodigoCliente As String = String.Empty
    Private _EndCliente As Integer = 0
    Private _Cliente As [Lib].Negocio.Cliente

    Private _CodigoProdutoXML As String = String.Empty

    Private _CodigoProduto As String = String.Empty
    Private _Produto As Produto = Nothing

    Private _ClienteConsolidado As Boolean = False

    Private _NomeProdutoXML As String = String.Empty
    Private _NCMProdutoXML As String = String.Empty
    Private _UnidadeProdutoXML As String = String.Empty

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

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
        End Set
    End Property


    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing AndAlso CodigoCliente > 0 Then _Cliente = New Cliente(CodigoCliente, EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property


    Public Property ClienteConsolidado() As Boolean
        Get
            Return _ClienteConsolidado
        End Get
        Set(ByVal value As Boolean)
            _ClienteConsolidado = value
        End Set
    End Property


    Public Property CodigoProdutoXML() As String
        Get
            Return _CodigoProdutoXML
        End Get
        Set(ByVal value As String)
            _CodigoProdutoXML = value
        End Set
    End Property

    Public Property NomeProdutoXML() As String
        Get
            Return _NomeProdutoXML
        End Get
        Set(ByVal value As String)
            _NomeProdutoXML = value
        End Set
    End Property

    Public Property NCMProdutoXML() As String
        Get
            Return _NCMProdutoXML
        End Get
        Set(ByVal value As String)
            _NCMProdutoXML = value
        End Set
    End Property

    Public Property UnidadeProdutoXML() As String
        Get
            Return _UnidadeProdutoXML
        End Get
        Set(ByVal value As String)
            _UnidadeProdutoXML = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing AndAlso CodigoProduto > 0 Then _Produto = New Produto(CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
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
                sql = "IF NOT EXISTS  (" & vbCrLf &
                        "Select 1 " & vbCrLf &
                            " FROM XmlProdutoXDePara " & vbCrLf &
                            " WHERE Cliente_Id    = '" & Me.CodigoCliente & "'" & vbCrLf &
                            "  AND EndCliente_Id = " & Me.EndCliente & vbCrLf &
                            "  AND ProdutoXML_Id = '" & Me.CodigoProdutoXML & "'" & vbCrLf &
                            "  AND Produto_Id    = '" & Me.CodigoProduto & "')" & vbCrLf &
                        "BEGIN " & vbCrLf &
                            " INSERT INTO XmlProdutoXDePara(Cliente_Id, EndCliente_Id, ProdutoXML_Id, Produto_Id, " & vbCrLf &
                                                    "ClienteConsolidado, NomeProdutoXML, NCMProdutoXML, UnidadeProdutoXML) " & vbCrLf &
                            " VALUES ('" & Me.CodigoCliente & "', " & Me.EndCliente & ", '" & Me.CodigoProdutoXML & "','" & Me.CodigoProduto & "'," & vbCrLf &
                            IIf(Me.ClienteConsolidado, 1, 0) & ",'" & Me.NomeProdutoXML & "','" & Me.NCMProdutoXML & "','" & Me.UnidadeProdutoXML & "')" & vbCrLf &
                        " END "
                Sqls.Add(sql)

            Case "U"
                sql = "UPDATE XmlProdutoXDePara " & vbCrLf &
                      "   Set ClienteConsolidado = " & IIf(Me.ClienteConsolidado, 1, 0) & vbCrLf &
                      "     , NomeProdutoXML = '" & Me.CodigoProdutoXML & "'" & vbCrLf &
                      "     , NCMProdutoXML = '" & Me.NCMProdutoXML & "'" & vbCrLf &
                      "     , UnidadeProdutoXML = '" & Me.UnidadeProdutoXML & "'" & vbCrLf &
                      " WHERE Cliente_Id      = '" & Me.CodigoCliente & "'" & vbCrLf &
                      "   AND EndCliente_Id   = " & Me.EndCliente & vbCrLf &
                      "   AND ProdutoXML_Id   = '" & Me.CodigoProdutoXML & "'" & vbCrLf &
                      "   AND Produto_Id      = '" & Me.CodigoProduto & "'"

                Sqls.Add(sql)

            Case "D"
                sql = "DELETE XmlProdutoXDePara " & vbCrLf &
                      " WHERE Cliente_Id      = '" & Me.CodigoCliente & "'" & vbCrLf &
                      "   AND EndCliente_Id   = " & Me.EndCliente & vbCrLf &
                      "   AND ProdutoXML_Id   = '" & Me.CodigoProdutoXML & "'" & vbCrLf &
                      "   AND Produto_Id      = '" & Me.CodigoProduto & "'"
                Sqls.Add(sql)

        End Select
    End Sub
#End Region
End Class
