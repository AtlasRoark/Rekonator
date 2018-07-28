<Serializable()>
Public Class Comparision
    Property LeftColumn As String
    Property RightColumn As String
    Property ComparisionTest As ComparisionType
    Property Percision As Integer = 0
    'Property ComparisionOption As Integer = 0
    'Property ComparisionMethod As CompareMethod

End Class

Public Enum ComparisionType
    TextEquals
    TextCaseEquals
    NumberEquals
    DateEquals
    DateTimeEquals
End Enum