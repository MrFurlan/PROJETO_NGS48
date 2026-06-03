''' <summary>
''' Lista de enumeradores de acontecimentos que podem acontecer no sistema
''' </summary>
''' <remarks>
''' Este enumerador tem os valores de acontecimentos que podem acontecer no sistema. Esses valores indicam que a tela principal deve mostrar uma mensagem genérica para 
''' indicar algo que aconteceu numa tela anterior. Geralmente esse valor irá ser passado através da seguinte sintaxe:
''' <example>
''' HttpContext.Current.Items("nomeDoItem") = AcontecimentosGerais.Valor
''' </example>
''' Na tela principal, a chamada será feita assim, como neste exemplo:
''' <example>
''' Select Case CType(HttpContext.Current.Items("Acontecimento"), NGS.Lib.Uteis.AcontecimentosGerais)
'''      Case SenhaModificada:
'''           lblMensagem.Text = "A senha foi modificada com sucesso."
''' End Select
''' </example>
''' </remarks>
Public Enum AcontecimentosGerais
    SenhaModificada
End Enum