Public Class Customer
    'Inherit information from CommonBase class:
    Inherits CommonBase

    Sub New()
        'MyBase keyword allows you to access any public property or method from the inherited class, in VB
        'This action initialises the Sub-section called "New()" in that inherited class and so we now have initialised variables in our child class
        'Essentially just a shortcut to writing out what the default values are for this child class, and makes sure we don't have some sort of conflict where they've been defined differently in each class instance
        MyBase.New()

        CustomerID = 1
    End Sub

    Public Property CustomerID As Integer
    Public Property FirstName As String
    Public Property LastName As String
    Public Property CompanyName As String

    Protected Overrides Function GetClassData() As String 'function returns a string
        'define an instance of the class "StringBuilder", of length 1024. This is efficient method of concatenating string together
        Dim sb As New Text.StringBuilder(1024)

        'take string builder and append each of the following lines, which contain information stored in this class
        sb.AppendLine("CustomerID: " + CustomerID.ToString())
        sb.AppendLine("Company Name: " + CompanyName)
        sb.AppendLine("First Name: " + FirstName)
        sb.AppendLine("Last Name: " + LastName)
        'The following are inherited from CommonBase Class:
        'They can be removed when using the syntax of overridable functions (and also adding the keyword "OVERRIDES" at the start of the function definition)
        'sb.AppendLine("Is Active: " + IsActive.ToString())
        'sb.AppendLine("Modified Date: " + ModifiedDate.ToLongDateString())
        'sb.AppendLine("Created By: " + CreatedBy)
        'They are replaced by the following line, which accesses that function from the inherited class
        sb.AppendLine(MyBase.GetClassData())

        Return sb.ToString()
    End Function
End Class
