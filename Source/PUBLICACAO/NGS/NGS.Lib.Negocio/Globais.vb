Imports Microsoft.VisualBasic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Globais

    Public Shared GCodEmpresa As String
    Public Shared GSenhaUser As String
    Public Shared CaminhoImagem As String
    Public Shared GPais As String

    Public Shared GSelCNPJ As String
    Public Shared GSelNome As String
    Public Shared GSelCidade As String
    Public Shared GSelEstado As String
    Public Shared GTituloRelatorio As String

    Public Shared GProcesso As String

    Public Shared GPermiteGravar As String = "S"
    Public Shared GPermiteAlterar As String = "S"
    Public Shared GPermiteExcluir As String = "S"
    Public Shared GPermiteLeitura As String = "S"
    Public Shared GPermiteAcesso As String = "S"
    Public Shared GPermiteLiberar As String = "S"
    Public Shared GTemImagem As String = "S"

    'Verificação se usuario consulta banco replica
    Public Shared GPermiteConsultaReplica As String

    Public Shared W As Integer
    Public Shared H As Integer
    Public Shared L As Integer
    Public Shared T As Integer

    Public Shared TipoBancoLocal As String
    Public Shared TipobancoReplica As String
    'Public Shared EnderecoBDLocal As String
    'Public Shared TipoBDLocal As String
    'Public Shared EnderecoLocal As String

    'Public Shared EnderecoReplica As String
    'Public Shared NomeServidor As String
    'Public Shared flagReplica As Boolean
    'Public Shared flagLocal As Boolean = True

    'Public Shared AutorizaBancoReplica As Boolean
    'Public Shared AutorizaBancoLocal As Boolean = True

    'Public Shared AutorizaGravarReplica As Boolean
    'Public Shared AutorizaGravarLocal As Boolean = True

    Public Shared ProgramaImportacao As Boolean = False

    Public Shared flagAutorizacao As Boolean
    Public Shared flagcancel As Boolean
    Public Shared NumeroBytes As New Byte

    Public Shared mensagemerro As Boolean = False

End Class