Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class MDTwoValues(Of TKey, TValue1, TValue2)
    Private Property List1() As Dictionary(Of TKey, TValue1)
        Get
            Return m_List1
        End Get
        Set(value As Dictionary(Of TKey, TValue1))
            m_List1 = value
        End Set
    End Property
    Private m_List1 As Dictionary(Of TKey, TValue1)
    Private Property List2() As Dictionary(Of TKey, TValue2)
        Get
            Return m_List2
        End Get
        Set(value As Dictionary(Of TKey, TValue2))
            m_List2 = value
        End Set
    End Property
    Private m_List2 As Dictionary(Of TKey, TValue2)

    Public ReadOnly Property Keys() As Dictionary(Of TKey, TValue1).KeyCollection
        Get
            Return List1.Keys
        End Get
    End Property
    Public ReadOnly Property Values1() As Dictionary(Of TKey, TValue1).ValueCollection
        Get
            Return List1.Values
        End Get
    End Property
    Public ReadOnly Property Values2() As Dictionary(Of TKey, TValue2).ValueCollection
        Get
            Return List2.Values
        End Get
    End Property
    'public TValue1 this[TKey key]
    '{
    '    get { return GetValue1(key); }
    '    set { SetValue1(key, value); }
    '}
    Default Public Property Item(key As TKey) As [Object]()
        Get
            Return New [Object]() {GetValue1(key), GetValue2(key)}
        End Get
        Set(value As [Object]())
            Dim array = TryCast(value, [Object]())
            SetValue1(key, DirectCast(array(0), TValue1))
            SetValue2(key, DirectCast(array(1), TValue2))
        End Set
    End Property

    Public Sub New()
        List1 = New Dictionary(Of TKey, TValue1)()
        List2 = New Dictionary(Of TKey, TValue2)()
    End Sub

    Public Sub Add(key As TKey, value1 As TValue1, value2 As TValue2)
        List1.Add(key, value1)
        List2.Add(key, value2)
    End Sub

    Public Sub SetValue1(key As TKey, value As TValue1)
        List1(key) = value
    End Sub

    Public Sub SetValue2(key As TKey, value As TValue2)
        List2(key) = value
    End Sub

    Public Function GetValue1(key As TKey) As TValue1
        Return List1(key)
    End Function

    Public Function GetValue2(key As TKey) As TValue2
        Return List2(key)
    End Function

    Public Sub Remove(key As TKey)
        List1.Remove(key)
        List2.Remove(key)
    End Sub

    Public Function ContainsKey(key As TKey) As [Boolean]
        Return List1.ContainsKey(key)
    End Function

    Public Function GetDictionary1() As Dictionary(Of TKey, TValue1)
        Return List1
    End Function

    Public Function GetDictionary2() As Dictionary(Of TKey, TValue2)
        Return List2
    End Function
End Class
