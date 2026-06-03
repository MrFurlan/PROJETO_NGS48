''' <summary>
''' 
''' </summary>
''' <remarks></remarks>
Public Class RegistroBlocoK001
    Private _REG As String = "K001"
    Private _IND_MOV As Integer

    ''' <summary>
    ''' Contém o valor fixo do registro.
    ''' </summary>
    ''' <value>K001</value>
    ''' <returns>K001</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property REG() As String
        Get
            Return _REG
        End Get
    End Property

    ''' <summary>
    ''' Indicador de movimento
    ''' </summary>
    ''' <value>0 ou 1</value>
    ''' <returns>0 - Bloco de dados informados / 1 - Bloco sem dados informados</returns>
    ''' <remarks>Se preenchido com 1 (um), devem ser informados os registros K001 e K990 (encerramento do bloco), significando que não há informação do controle da produção e do estoque.
    ''' Se preenchido com 0 (zero), então deve ser informado pelo menos um registro K100 e seus respectivos registros filhos, além do registro K990 (encerramento do bloco). </remarks>
    Public Property IND_MOV() As Integer
        Get
            Return _IND_MOV
        End Get
        Set(ByVal value As Integer)
            _IND_MOV = value
        End Set
    End Property
End Class

''' <summary>
''' Período de Apuração do ICMS/IPI
''' </summary>
''' <remarks>Este registro tem o objetivo de informar o período de apuração do ICMS ou do IPI, prevalecendo os períodos mais curtos.
'''          Contribuintes com mais de um período de apuração no mês declaram um registro K100 para cada período no mesmo arquivo.
'''          Não podem ser informados dois ou mais registros com os mesmos campos DT_INI e DT_FIN.</remarks>
Public Class RegistroBlocoK100
    Private _REG As String = "k100"
    Public ReadOnly Property REG() As String
        Get
            Return _REG
        End Get
    End Property

    Private _DT_INI As Date
    Public Property DT_INI() As Date
        Get
            Return _DT_INI
        End Get
        Set(ByVal value As Date)
            _DT_INI = value
        End Set
    End Property

    Private _DT_FIN As Date
    Public Property DT_FIN() As Date
        Get
            Return _DT_FIN
        End Get
        Set(ByVal value As Date)
            _DT_FIN = value
        End Set
    End Property



End Class

Public Class RegistroBlocoK200

End Class

Public Class RegistroBlocoK220

End Class

Public Class RegistroBlocoK230

End Class

Public Class RegistroBlocoK235

End Class

Public Class RegistroBlocoK250

End Class

Public Class RegistroBlocoK255

End Class

Public Class RegistroBlocoK990

End Class

