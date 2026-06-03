Imports NGS.Lib.Uteis

Public Class ValidationKey

#Region "builder"

    Public Sub New()

    End Sub

    Public Sub New(ByVal empresa As String)
        Me.CodigoEmpresa = empresa
        Dim existechavelocal As Boolean = False
        Dim sql As String = "Select empresa, keyCode from ValidationKey where empresa = '" & empresa & "'"
        Dim ds = New AcessaBanco().ConsultaDataSet(sql, "ValidationKey")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Me.KeyCode = ds.Tables(0).Rows(0)(1).ToString()
            existechavelocal = True
        End If

        If Not Me.AtivoLocal AndAlso Funcoes.VerificaConexaoInternet() Then
            sql = "select qtdedias from ValidationKey where empresa = '" & empresa & "' and ativo = 1"
            ds = New AcessaBancoMySql().ConsultaDataSet(sql, "ValidationKey")

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                Me.QtdeDias = ds.Tables(0).Rows(0)(0).ToString()
                Dim dataValidade As Date = Now.AddDays(Me.QtdeDias - 1)
                Me.KeyCode = FuncoesStrings.CodificarPara64Bits(empresa & DateTime.Now.ToSqlDate() & dataValidade.ToSqlDate())
                UsuarioServidor.KeyCodeActive = True
                Me.salvar(existechavelocal)
            End If
        End If
    End Sub

#End Region

#Region "Fields"

    Private _CodigoEmpresa As String
    Private _QtdeDias As Integer
    Private _KeyCode As String
    Private _AtivoLocal As Boolean

#End Region

#Region "Properties"

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property QtdeDias() As Integer
        Get
            Return _QtdeDias
        End Get
        Set(ByVal value As Integer)
            _QtdeDias = value
        End Set
    End Property

    Public Property KeyCode() As String
        Get
            Return _KeyCode
        End Get
        Set(ByVal value As String)
            _KeyCode = value
        End Set
    End Property

    Public ReadOnly Property AtivoLocal() As Boolean
        Get
            If Not String.IsNullOrWhiteSpace(KeyCode) Then
                Dim textoDecodificado As String = FuncoesStrings.DecodificarDe64Bits(KeyCode)
                Dim DataAtualizacao As Date = textoDecodificado.Substring(8, 10)
                Dim dataValidade As Date = Right(textoDecodificado, 10)
                Dim dataatual As Date = New DateTime(Now.Year, Now.Month, Now.Day)
                If dataatual >= DataAtualizacao AndAlso dataatual <= dataValidade Then
                    Return True
                End If
            End If
            Return False
        End Get
    End Property

#End Region

#Region "Metods"

    Private Function salvar(ByVal existkeylocal As Boolean) As Boolean
        Dim sql As String = ""
        If existkeylocal Then
            sql = String.Format("Update ValidationKey set keycode = '{0}' Where empresa = '{1}'", Me.KeyCode, Me.CodigoEmpresa)
        Else
            sql = String.Format("Insert into ValidationKey (empresa, keycode) values ('{0}', '{1}')", Me.CodigoEmpresa, Me.KeyCode)
        End If
        Return New AcessaBanco().GravaBanco(sql)
    End Function

#End Region



End Class
