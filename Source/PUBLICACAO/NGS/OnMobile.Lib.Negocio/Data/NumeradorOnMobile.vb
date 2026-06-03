Imports System.Configuration

<Serializable()>
Public Class NumeradorOnMobile

#Region "Propriedades"

    Private _sEmpresa, _sDescricao, _sReduzido, _sNome, _sSerie As String
    Private _iEndEmpresa, _iNumerador, _iSequencia As Integer
    Private _IUD As String
    Public Property IUD() As Char
        Get
            Return _IUD
        End Get
        Set(ByVal value As Char)
            value = IUD
        End Set
    End Property

    Public Property Reduzido() As String
        Get
            Return _sReduzido
        End Get
        Set(ByVal value As String)
            _sReduzido = value
        End Set
    End Property

    Public Property Empresa() As String
        Get
            Return _sEmpresa
        End Get
        Set(ByVal Value As String)
            _sEmpresa = Value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _iEndEmpresa
        End Get
        Set(ByVal Value As Integer)
            _iEndEmpresa = Value
        End Set
    End Property

    Public Property Nome() As String
        Get
            Return _sNome
        End Get
        Set(ByVal value As String)
            _sNome = value
        End Set
    End Property

    Public Property Numerador() As Integer
        Get
            Return _iNumerador
        End Get
        Set(ByVal Value As Integer)
            _iNumerador = Value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _sSerie
        End Get
        Set(ByVal value As String)
            _sSerie = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _sDescricao
        End Get
        Set(ByVal Value As String)
            _sDescricao = Value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _iSequencia
        End Get
        Set(ByVal Value As Integer)
            _iSequencia = Value
        End Set
    End Property


#End Region
    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal Numerador As Integer)
        Dim Banco As New AcessaBancoOnMobile
        Dim Sql As String

        Sql = " SELECT c.Reduzido,n.Empresa_Id, n.EndEmpresa_Id, C.Fantasia + ' - ' + C.Cidade + '-' + C.Estado AS Nome," & vbCrLf &
              "        n.Numerador_Id, n.Descricao, n.Sequencia, isnull(n.Serie, '') AS Serie " & vbCrLf &
              "   FROM Numerador n" & vbCrLf &
              "  INNER JOIN Clientes c" & vbCrLf &
              "     ON n.Empresa_Id    = c.Cliente_Id " & vbCrLf &
              "    AND n.EndEmpresa_Id = c.Endereco_Id " & vbCrLf &
              "  WHERE n.Empresa_Id    ='" & Empresa & "'" & vbCrLf &
              "    AND n.EndEmpresa_Id = " & EndEmpresa & vbCrLf &
              "    AND n.Numerador_Id  = " & Numerador & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Nume")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim Row As DataRow = ds.Tables(0).Rows(0)
            Me.Reduzido = Row("Reduzido")
            Me.Empresa = Row("Empresa_Id")
            Me.EndEmpresa = Row("EndEmpresa_Id")
            Me.Numerador = Row("Numerador_Id")
            Me.Descricao = Row("Descricao")
            Me.Nome = Row("Nome")
            Me.Sequencia = Row("Sequencia")
            Me.Serie = Row("Serie")
        Else
            Throw New Exception("Não existe Numerador " & Numerador & " Cadastrado para a Empresa ")
        End If
    End Sub

    Public Function IncrementarNumeradorSql(Optional ByVal pGravaServidor As Boolean = False, Optional ByVal pIncremento As Integer = 0) As String
        Dim strSQL As String

        strSQL = "	  update Numerador set " & vbCrLf &
                 "	     Sequencia        = Sequencia + " & IIf(pIncremento = 0, "1", pIncremento.ToString) & vbCrLf
        If pGravaServidor Then
            strSQL &= " Where Empresa_id    ='" & ConfigurationManager.AppSettings("server").ToString() & "'" & vbCrLf &
                      "	  and EndEmpresa_id = 0 " & vbCrLf
        Else
            strSQL &= " Where Empresa_id    ='" & Me.Empresa & "'" & vbCrLf &
                      "	  and EndEmpresa_id = " & Me.EndEmpresa & vbCrLf
        End If
        strSQL &= " and Numerador_id  = " & Me.Numerador
        Return strSQL
    End Function

    Public Sub New(ByVal Numerador As Integer)
        Dim Banco As New AcessaBancoOnMobile
        Dim strSQL As String
        Dim Servidor As String = ConfigurationManager.AppSettings("server").ToString()

        strSQL = " SELECT Empresa_Id, EndEmpresa_Id, Numerador_Id, Descricao, Sequencia, isnull(Serie, '') as Serie " & vbCrLf &
                 "  FROM Numerador " & vbCrLf &
                 " WHERE Empresa_Id    ='" & Servidor & "'" & vbCrLf &
                 "   AND EndEmpresa_Id = 0" & vbCrLf &
                 "   AND Numerador_Id  = " & Numerador & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "numer")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            Me.Empresa = row("Empresa_Id")
            Me.EndEmpresa = row("EndEmpresa_Id")
            Me.Numerador = row("Numerador_Id")
            Me.Descricao = row("Descricao")
            Me.Sequencia = row("Sequencia")
            Me.Serie = row("Serie")
        Else
            Throw New Exception("Não existe Numerador " & Numerador & " Cadastrado para o servidor " & Servidor)
        End If
    End Sub
End Class
