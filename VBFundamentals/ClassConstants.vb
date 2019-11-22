Public Class ClassConstants
    ' Using Class constants is best practice
    Public Const DEFAULT_ACTIVE As Boolean = True
    Public Const DEFAULT_LIST_PRICE As Decimal = 999.99D
    ' Public constants respect the class boundary, they are not accessible outside without calling them directly
    Private Const NOT_SEEN_OUTSIDE As String = "Not Seen"

End Class
