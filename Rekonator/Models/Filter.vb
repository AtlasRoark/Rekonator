Public Class Filter
    Property DataSourceName As String
    Property FilterColumns As String()
    Property FilterOption As FilterOption
End Class

Public Enum FilterOption
    EqualsZero
    NonZero
    NonBlankOrZero
    IsNegative
    IsPosition
    Equals
    IsGreaterThan
    IsGreaterThanOrEqualTo
    IsLessThan
    IsLessThanOrEqualTo
End Enum
