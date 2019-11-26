Module MainModule
    Dim Name As String = "10 Speed Bike"
    Sub Main()
        'Objects()
        'Main4()
        'Console.ReadKey()
        'BuiltInString()
        'NumericDataType()
        'DateAndTime()
        'ClassProduct()
        ClassFunction()
    End Sub
    Sub Main0()
        Dim Name As String = "10 Speed Bike"
        ' local variables inside sub "shadow" global variables that are further away 
        Console.WriteLine(Name)

        Console.ReadKey()
    End Sub
    Sub Main1()
        Console.WriteLine(Name)
    End Sub

    Sub Main2()
        If True Then
            Dim Name As String = "Tricycle"
            Dim ListPrice As Decimal = 59.99D
        End If

        Console.WriteLine(Name)
        ' The following does not compile properly, because the if statement above is seen a different block of code.
        ' The variable is initialised inside the if statement and expires when the if statements ends
        ' Console.WriteLine(ListPrice)

        Console.ReadKey()
    End Sub
    Sub Main3()
        IncrementListPrice()
        IncrementListPrice()
        IncrementListPrice()

        Console.ReadKey()
    End Sub

    Sub IncrementListPrice()
        Static ListPrice As Decimal = 0

        ListPrice = ListPrice + 1

        Console.WriteLine(ListPrice)
    End Sub
    Sub Objects()
        Dim theData As Object
        ' This is a rarely the right data type to be using. It is very memory intensive and is slow to access and slow to use, avoid where you can but know it is available.

        theData = "10 Speed Bike"
        Console.WriteLine(theData)

        theData = 999.99
        Console.WriteLine(theData)

        theData = DateTime.Now
        Console.WriteLine(theData)

        theData = vbEmpty
        Console.WriteLine(theData)

        theData = DBNull.Value
        Console.WriteLine(theData)
    End Sub
    Sub Main4()
        Dim isActive As Boolean = ClassConstants.DEFAULT_ACTIVE
        Dim Name As String = "10 Speed Bike"
        Dim ListPrice As Decimal
        Console.WriteLine(isActive)

        ListPrice = ClassConstants.DEFAULT_LIST_PRICE
        Console.WriteLine(ListPrice)
        Console.ReadKey()
    End Sub
    Sub Maths()
        Dim ListPrice As Decimal = 999.99D

        ListPrice += 1
        Console.WriteLine(ListPrice)
        Console.ReadKey()
    End Sub
    Sub BuiltInString()
        Dim Name As String = "10 Speed Bike"

        Console.WriteLine("Built-in String Methods")
        ' The following method calls on the string variable "Name" are all string-based built in method calls on objects of type string.
        ' They return a new manipulated string output, but do not change the existing string that is saved to the variable "Name"
        ' To do that, we would have to reassign the variable e.g Name = Name.Replace("you", "who")
        Console.WriteLine("Show the Length of the Bike Name: ")
        Console.WriteLine(Name.Length)
        Console.WriteLine("Show the index position of the first space: ")
        Console.WriteLine(Name.IndexOf(" "))
        Console.WriteLine("Show the index position of the last space: ")
        Console.WriteLine(Name.LastIndexOf(" "))
        Console.WriteLine("Test whether or not the referenced string ends with the letter 'e' or not, and output a boolean result:")
        Console.WriteLine(Name.EndsWith("e"))
        Console.WriteLine("Test whether or not the referenced string ends with the letter 'q' or not, and output a boolean result:")
        Console.WriteLine(Name.EndsWith("q"))
        Console.WriteLine("Insert the word 'Mountain' into the string: ")
        Console.WriteLine(Name.Insert(9, "Mountain "))
        Console.WriteLine("Remove the characters in positions 0 to 9 inclusive: ")
        Console.WriteLine(Name.Remove(0, 9))
        Console.WriteLine("Replace the string '10' with the string'12': ")
        Console.WriteLine(Name.Replace("10", "12"))
        Console.WriteLine("Change the case of the whole string to upper case!: ")
        Console.WriteLine(Name.ToUpper())
        Console.WriteLine("Change the case of the whole string to lower case!: ")
        Console.WriteLine(Name.ToLower())

        Console.ReadKey()
    End Sub
    Sub NumericDataType()
        Dim ListPrice As Decimal = 999.99D
        ' The following method calls on the decimal variable ListPrice are all built-in method calls 
        ' They return a new manipulated decimal output, but do not change the actual variable value

        Console.WriteLine("Built-in Numeric Methods")
        Console.WriteLine("Check to see if the variable ListPrice is equal to 999.99 as a decimal")
        Console.WriteLine(ListPrice.Equals(999.99D))
        Console.WriteLine("Check to see what the minimum possible value for this class of objects, decimals, is ")
        Console.WriteLine(Decimal.MinValue)
        Console.WriteLine("Check to see what the mximum possible value for this class of objects, decimals, is ")
        Console.WriteLine(Decimal.MaxValue)
        Console.WriteLine("Check the rounded up value of the variable")
        Console.WriteLine(Decimal.Ceiling(ListPrice))
        Console.WriteLine("Check the rounded down value of the variable")
        Console.WriteLine(Decimal.Floor(ListPrice))

        ' The following attempts to parse certain values from a string to a decimal data type
        Console.WriteLine("Display the following depending on a syntax of 'Decimal.TryParse(input, into)', where the aim is to try and parse the input into a variable. ")
        Console.WriteLine("In this case, we tried to parse the STRING '55.01' into the variable as a decimal, and it worked because the string only contained numbers.")
        ' takes the inputted value, here 55.01m, tries to make it a decimal (which can be done: something VB does automatically), and add it to that variable "ListPrice"
        Decimal.TryParse("55.01", ListPrice)
        Console.WriteLine(ListPrice)

        Console.WriteLine("In this case we tried to parse the STRING 'fifty-five' into the variable, but it did not work because the string contained letters, not the numerical value required to become a decimal data type.")
        ' takes the inputted value, here 55 feet, tries to make it a decimal (which can't be done, because the string only contains letters), and add it to that variable "ListPrice".
        ' since it can't be doney, the result just shows as 0
        Decimal.TryParse("Fifty-five", ListPrice)
        Console.WriteLine(ListPrice)

        Console.ReadKey()
    End Sub
    Sub DateAndTime()
        Dim SellDate As DateTime = #1/1/2019 12:25:37 PM#

        ' The following are build in methods for the DateTime data type, used for reading and manipulating the values for the various time scales stored in a given DateTime data object
        Console.WriteLine("Built-in DateTime Methods")
        Console.WriteLine("Add and Subtract days from a Date")
        Console.WriteLine(SellDate.AddDays(10))
        Console.WriteLine(SellDate.AddDays(-10))
        Console.WriteLine("Add and Subtract years from a Date")
        Console.WriteLine(SellDate.AddYears(2))
        Console.WriteLine(SellDate.AddYears(-3))

        Console.WriteLine("Print the various components of a Date: ")
        Console.WriteLine("Date (time set to midnight):")
        Console.WriteLine(SellDate.Date)
        Console.WriteLine("Day:")
        Console.WriteLine(SellDate.Day)
        Console.WriteLine("Day of Week (number given, e.g. 2 =  tuesday):")
        Console.WriteLine(SellDate.DayOfWeek)
        Console.WriteLine("Year:")
        Console.WriteLine(SellDate.Year)
        Console.WriteLine("Day of Year:")
        Console.WriteLine(SellDate.DayOfYear)

        Console.WriteLine("Hour, Minute, Second, respectively:")
        Console.WriteLine(SellDate.Hour)
        Console.WriteLine(SellDate.Minute)
        Console.WriteLine(SellDate.Second)

        Console.ReadKey()

    End Sub
    Sub ClassProduct()
        'Create a new instance of this product class, take that object and save it to the variable "prod":
        Dim prod As New Product
        'Create the variable that needs to be passed through the method call of CalculateSellEndDate:
        Dim sellDate As DateTime
        'Set the SellStartDate property to a date:
        prod.SellStartDate = #1/1/2019#
        'Call the sub-procedure to set the SellEndDate, which is SellStartDate + 20.
        'In this iteration, we also need to pass it an argument (ideally a variable), this does not have to be called "sellDate", I have just used this again for consistency and making it obvious. 
        'This action is according to the ByRef keyword given when defining this method in the class file
        'The next line is commented out to abide by new function designed for next module, which has fewer arguments. It would still be correct syntax if we were using the ByRef tactic 
        'prod.CalculateSellEndDate(30, sellDate)

        'Output the SellEndDate
        Console.WriteLine(prod.SellEndDate)
        'Write out the newly defined SellEndDate, defined by its own variable:
        Console.WriteLine(sellDate)

        Console.ReadKey()
    End Sub
    Sub ClassFunction()
        'Set-up your class instances and/or variables
        Dim prod As New Product
        Dim sellDate As DateTime
        'Set the value of SellStartDate for this instance of Products: "prod"
        prod.SellStartDate = #5/1/2019#

        'Calculate the SellEndDate using the built in function
        'Doesn't print the sell date, despite the method including "return"
        sellDate = prod.CalculateSellEndDate(25)

        'Print the sell date to the console
        Console.WriteLine(sellDate)
        Console.ReadKey()
    End Sub
End Module