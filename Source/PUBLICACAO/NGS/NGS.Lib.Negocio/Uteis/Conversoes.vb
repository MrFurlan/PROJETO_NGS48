Imports System.ComponentModel
Imports System.IO
Imports System.Reflection

Public Class Conversoes

    Public Shared Function ConverterClasseOperacao(ByVal Classe As String) As eClassesOperacoes
        If Classe = "" Then
            Return eClassesOperacoes.NENHUMA
        Else
            Dim strClasse As String = Classe.Replace(" ", "")
            Return [Enum].Parse(GetType(eClassesOperacoes), strClasse.ToUpper())
        End If
    End Function

    Public Shared Function ConverterOperadoresBancoParaString(ByVal Operador As eOperadoresSelecao) As String
        Select Case Operador
            Case eOperadoresSelecao.Como : Return "LIKE"
            Case eOperadoresSelecao.E : Return "IS"
            Case eOperadoresSelecao.Entre : Return "BETWEEN"
            Case eOperadoresSelecao.Igual : Return "="
            Case eOperadoresSelecao.MaiorOuIgualQue : Return ">="
            Case eOperadoresSelecao.MaiorQue : Return ">"
            Case eOperadoresSelecao.MenorOuIgualQue : Return "<="
            Case eOperadoresSelecao.MenorQue : Return "<"
            Case eOperadoresSelecao.NaoE : Return "NOT IS"
            Case Else : Return ""
        End Select
    End Function

    Public Shared Function ConverterSinal(ByVal Sinal As String) As eSinal
        Select Case Sinal
            Case "+" : Return eSinal.Positivo
            Case "-" : Return eSinal.Negativo
            Case Else : Return eSinal.Igual
        End Select
    End Function

    Public Shared Function ConverterTipoLancamento(ByVal Tipo As String) As eTiposLancamentosPedidos
        Select Case Tipo
            Case "N" : Return eTiposLancamentosPedidos.Normal
            Case "E" : Return eTiposLancamentosPedidos.Estorno
            Case "C" : Return eTiposLancamentosPedidos.Complemento
        End Select
        Return eTiposLancamentosPedidos.Normal
    End Function

    Public Shared Function ConverterTipoMoeda(ByVal Tipo As String) As eTiposMoeda
        Select Case Tipo
            Case "O" : Return eTiposMoeda.Oficial
            Case "M" : Return eTiposMoeda.MoedaEstrangeira
        End Select
        Return eTiposMoeda.Oficial
    End Function

    Public Shared Function ConverterImagemEmByte(ByVal Arquivo As String) As Byte()
        If File.Exists(Arquivo) Then
            Dim fileImagem As New FileStream(Arquivo, FileMode.Open)
            Dim objLeitorImagem As New BinaryReader(fileImagem)
            Dim byteImagem As Byte() = objLeitorImagem.ReadBytes(Convert.ToInt32(objLeitorImagem.BaseStream.Length))
            fileImagem.Close()
            objLeitorImagem.Close()

            Return byteImagem
        Else : Return Nothing
        End If
    End Function

    Public Shared Function GetValueFromDescription(Of T)(ByVal description As String) As T
        Dim type = GetType(T)

        If (Not type.IsEnum) Then
            Throw New ArgumentException("Enum não encontrado")
        End If

        For Each attribute In System.Enum.GetValues(type)
            Dim enumerador = [Enum].Parse(GetType(T), attribute.ToString())
            If GetEnumDescription(enumerador) = description Then
                Return enumerador
            End If
        Next
    End Function

    Shared Function GetEnumDescription(ByVal EnumConstant As [Enum]) As String
        Dim fi As FieldInfo = EnumConstant.GetType().GetField(EnumConstant.ToString())
        Dim attr() As DescriptionAttribute = DirectCast(
       fi.GetCustomAttributes(GetType(DescriptionAttribute), False),
          DescriptionAttribute())
        If attr.Length > 0 Then
            Return attr(0).Description
        Else
            Return EnumConstant.ToString()
        End If
    End Function


End Class