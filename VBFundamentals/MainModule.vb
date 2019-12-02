Module MainModule
    'We define a dictionary, which is of Generic Type
    'We define it to have some sort of data type with the keyword "Of"
    'A dictionary has two things, it has a key value and a value itself that is reference by its corresponding key value
    'We use an integer as they key, like a primary key in an SQL table. In theory, we could use whatever data type we want as a key value
    'We use the Product object as the value object type, again it could be any data type 
    Function LoadProducts() As Dictionary(Of Integer, Product)
        'Dimensionalise a new dictionary called "products", with the same key and values as the function output data types
        Dim products As New Dictionary(Of Integer, Product)
        'Dimensionalise a new variable as a Product class type
        Dim prod As Product

        'now we assign a new product class instance to the variable "prod", with the data you require
        prod = New Product() With {.ProductID = 1, .Name = "10 Speed Bike", .ProductNumber = "10-SP", .ListPrice = 1469.16D}
        'assign this product type to the dictionary, with key value as the productID (because it is unique), and the value as the product class object assigned to prod
        'the syntax for the assignment is a bit different as ":=", we use this to assign specific key values, but really it is not necessary as long as you add things in the correct order
        products.Add(key:=prod.ProductID, value:=prod)

        'We can repeat this process for two more product class instances defined below, assinging them to the dictionary
        'We can use the same variable as each time we use it, we just assign it a new value and we don't have to worry about what it's replacing because the previous value has already been assigned to the dictionary
        prod = New Product() With {.ProductID = 2, .Name = "Bike Helmet", .ProductNumber = "BIKE-HE", .ListPrice = 994.47D}
        'Notice the difference here, we don't have to use the := syntax if we pass the parameters in the exact order that they are listed in the method
        products.Add(prod.ProductID, prod)

        prod = New Product() With {.ProductID = 3, .Name = "Inner Tube", .ProductNumber = "BIKE-IN-TU", .ListPrice = 756.81D}
        products.Add(prod.ProductID, prod)

        Return products

    End Function
    'Now we're returning a list
    'This time there are no key values like there were with the dictionary function
    'Once you "type" the Generic List using (Of __), that is the only data type that can go into that list
    Function LoadProductsList() As List(Of Product)
        Dim products As New List(Of Product) From {
            New Product() With {.ProductID = 680, .Name = "HL Road Frame - Black, 58", .ProductNumber = "FR-R92B-58", .Colour = "Black", .Size = "58", .Weight = 1016.04D, .StandardCost = 1059.31D, .ListPrice = 1431.5D, .SellStartDate = #6/1/1998 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#},
            New Product() With {.ProductID = 706, .Name = "HL Road Frame - Red, 58", .ProductNumber = "FR-R92R-58", .Colour = "Red", .Size = "58", .Weight = 1016.04D, .StandardCost = 1059.31D, .ListPrice = 1431.5D, .SellStartDate = #6/1/1998 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#},
            New Product() With {.ProductID = 707, .Name = "Sport-100 Helmet, Red", .ProductNumber = "HL-U509-R", .Colour = "Red", .Size = "L", .Weight = 15.1D, .StandardCost = 13.0863D, .ListPrice = 34.99D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#},
            New Product() With {.ProductID = 708, .Name = "Sport-100 Helmet, Black", .ProductNumber = "HL-U509", .Colour = "Black", .Size = "L", .Weight = 15.1D, .StandardCost = 13.0863D, .ListPrice = 34.99D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#},
            New Product() With {.ProductID = 709, .Name = "Mountain Bike Socks, Black", .ProductNumber = "SO-B909-M", .Colour = "White", .Size = "M", .Weight = 10, .StandardCost = 3.3963D, .ListPrice = 9.5D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #6/30/2002 12:00:00 AM#},
            New Product() With {.ProductID = 710, .Name = "Mountain Bike Socks, Red", .ProductNumber = "SO-B909-L", .Colour = "White", .Size = "L", .Weight = 10, .StandardCost = 3.3963D, .ListPrice = 9.5D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #6/30/2002 12:00:00 AM#},
            New Product() With {.ProductID = 711, .Name = "Sport-100 Helmet, Blue", .ProductNumber = "HL-U509-B", .Colour = "Blue", .Size = "M", .Weight = 10, .StandardCost = 13.0863D, .ListPrice = 34.99D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#},
            New Product() With {.ProductID = 712, .Name = "AWC Logo Cap", .ProductNumber = "CA-1098", .Colour = "Multi", .Size = "L", .Weight = 5.5D, .StandardCost = 6.9223D, .ListPrice = 8.99D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#},
            New Product() With {.ProductID = 713, .Name = "Long-Sleeve Logo Jersey, S", .ProductNumber = "LJ-0192-S", .Colour = "Multi", .Size = "S", .Weight = 3.5D, .StandardCost = 38.4923D, .ListPrice = 49.99D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#},
            New Product() With {.ProductID = 714, .Name = "Long-Sleeve Logo Jersey, M", .ProductNumber = "LJ-0192-M", .Colour = "Multi", .Size = "M", .Weight = 3.5D, .StandardCost = 38.4923D, .ListPrice = 49.99D, .SellStartDate = #7/1/2001 12:00:00 AM#, .SellEndDate = #12/31/2021 12:00:00 AM#}
        }
        Return products
    End Function
    Sub Main()
        'Objects()
        'Main4()
        'Console.ReadKey()
        'BuiltInString()
        'NumericDataType()
        'DateAndTime()
        'ClassProduct()
        'ClassFunction()
        'CalculateProfit()
        'CalculateTheProfit()
        'Inheritance()
        'Inheritance2()
        'Arrays()
        'ArrayList()
        'ArrayList2()
        'Dictionaries()
        'DictionaryKeys()
        'DictionaryMethods()
        'LINQ()
        'ListOfT()
        'ListOfTMethods()
        'DoWhileLoop()
        'DoUntilLoop()
        'ForNext()
        'ForStepNext()
        'ForEachNext()
        'IfElseElseIfStatements()
        'SelectCase()
        CompilerConstants()
    End Sub
    Sub Main0()
        Dim Name As String = "10 Speed Bike"
        ' local variables inside sub "shadow" global variables that are further away 
        Console.WriteLine(Name)

        Console.ReadKey()
    End Sub
    'Sub Main1()
    '    Console.WriteLine(Name)
    'End Sub

    Sub Main2()
        If True Then
            Dim Name As String = "Tricycle"
            Dim ListPrice As Decimal = 59.99D
        End If

        'the following still compiles properly if the value "Name" is defined outside of this Sub, in the Main module
        'Console.WriteLine(Name)
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
    Sub CalculateProfit()
        'new product class intance
        Dim prod As New Product

        'Set the variables you need for this instance
        'These two initialisations of code are no longer needed when you use constructors in the class file to set up defaults of variables
        'But we can still have them in there to override the defaults
        'prod.StandardCost = 250
        'prod.ListPrice = 500

        'Two cases for setting the variable input or not - it is optional and behaves differently depending on it.
        'Output without setting the newCost variable:
        Console.WriteLine(prod.CalculateProfit())
        'Output setting the newCost variable:
        Console.WriteLine(prod.CalculateProfit(1100))
        'What happens if we try to parse the argument as 0? It just returns the value of ListPrice, because it is taking away 0 (the StandardCost)
        Console.WriteLine(prod.CalculateProfit(0))
        Console.ReadKey()
    End Sub
    Sub CalculateTheProfit()
        'We can call this new shared method on just the Product class name, it doesn't even need an instance of a class to be created!
        'Dim prod As New Product
        Console.WriteLine(Product.CalculateTheProfit(900, 1500)) '=600 always, regardless of what you've defined as properties for the class instance
        'Console.WriteLine(prod.CalculateTheProfit(900, 1500)) would work if you initialised an instance, but it throws a warning in the console
        Console.ReadKey()
    End Sub
    Sub Inheritance()
        Dim prod As New Product
        Console.WriteLine(Product.CalculateTheProfit(900, 1400))
        Console.ReadKey()
    End Sub
    Sub Inheritance2()
        'We can pre-initialise properties values when we initialise an instance of a class, using the WITH keyword
        'IntelliSense is very clever as it can help by giving us a list of properties we haven't defined yet when we type a .
        Dim prod As New Product With {
            .ProductID = 700,
            .Name = "10 Speed Bike",
            .ProductNumber = "10-SP"
        }
        'Let's see our new method for getting class data in action
        'We now call the ToString Method which calls it in CommonBase class, which is overriding a ToString function and hence collects and displays the data we require
        Console.WriteLine(prod.ToString())

        'We can do the same pre-initialisation with the customer class instance
        Dim cust As New Customer With {
            .CustomerID = 349,
            .CompanyName = "Beach Computer Consulting",
            .FirstName = "Bruce",
            .LastName = "Jones"
        }

        'Same again, calling the ToString function because we cannot call the Protected function GetClassData from outside the inheritance chain
        Console.WriteLine(cust.ToString())

        Console.ReadKey()
    End Sub
    Sub ProtectedFunctions()
        'Protected methods can't be seen by an instance variable, they can only be used within inherited classes
    End Sub
    Sub Arrays()
        'Defining an array of a single type

        'Dim products(3) As String
        'products(0) = "10 Speed Bike"
        'products(1) = "Bike Helmet"
        'products(2) = "Inner Tube"

        'Alternative way Of defining an array, in fewer lines:
        Dim products As String() = {"10 Speed Bike", "Bike Helmet", "Inner Tube"}

        Console.WriteLine(products(0))
        Console.WriteLine(products(1))
        Console.WriteLine(products(2))

        Console.WriteLine(products.Count)

        Console.ReadKey()
    End Sub
    Sub ArrayList()
        Dim products As New ArrayList From {
            "10 Speed Bike",
            "Bike Helmet",
            "Inner Tube",
            1,
            3.35D,
            New Product With {.ProductID = 1}
        }
        Dim product As New Product

        Console.WriteLine(products(0))
        Console.WriteLine(products(1))
        Console.WriteLine(products(2))
        Console.WriteLine(products(3))
        Console.WriteLine(products(4))
        Console.WriteLine(products(5))

        Console.WriteLine(products.Count)

        Console.ReadKey()
    End Sub
    Sub ArrayList2()
        'Define an array list with multiple class instances contained in it
        Dim products As New ArrayList From {
            New Product() With {.ProductID = 1, .Name = "10 Speed Bike", .ProductNumber = "10-SP"},
            New Product() With {.ProductID = 1, .Name = "Bike Helmet", .ProductNumber = "BIKE-HE"},
            New Product() With {.ProductID = 1, .Name = "Inner Tube", .ProductNumber = "BIKE-IN-TU"}
        }

        'The only way to access the property values of the class instances is to cast them, then we can show these properties on the console
        Console.WriteLine(DirectCast(products(0), Product).Name)
        Console.WriteLine(DirectCast(products(1), Product).ProductID)
        Console.WriteLine(DirectCast(products(2), Product).ProductNumber)

        Console.ReadKey()
    End Sub

    Sub Dictionaries()
        Dim products = LoadProducts()

        'We pass in the key for that key-value pair (as opposed to the index values for an array) to access the corresponding value
        'The good thing about these dictionaries is that, unlike with arrays, we don't have to do a DirectCast to access the class properties and write them to the line
        'Because the dictionary is defines as having values as the product class object type, it expects the outputs to be the class object type - whereas the arrays we didn't have this
        Console.WriteLine(products(1).Name)
        Console.WriteLine(products(2).ProductID)
        Console.WriteLine(products(3).ProductNumber)

        Console.ReadKey()
    End Sub
    Sub DictionaryKeys()
        'Dim the dictionary using the function
        Dim products = LoadProducts()

        'TRUE/FALSE for whether the dictionary contains certain keys
        Console.WriteLine(products.ContainsKey(1)) 'True - it does contain the key 1
        Console.WriteLine(products.ContainsKey(99)) 'False - it doesn't contian the key 99

        Console.ReadKey()
    End Sub
    Sub DictionaryMethods()
        'dim the dictionary using the function
        Dim products = LoadProducts()
        'Display the total number of items in the dictionary
        Console.WriteLine(products.Count)
        'Remove an item by referring to it with a key
        products.Remove(1)
        Console.WriteLine(products.Count)
        'Remove all items from the dictionary
        products.Clear()
        Console.WriteLine(products.Count)

        Console.ReadKey()
    End Sub
    Sub LINQ()
        Dim products = LoadProducts()

        'Display sum of all list prices
        'These LINQ expressions iterate over the collection and pass each item in the collection to the function in the LINQ method
        Console.WriteLine(
            products.Sum(Function(p)
                             Return p.Value.ListPrice 'Access the value of the dictionary with p.Value, then again access the list price with .ListPrice, pass that value to the loop
                         End Function).ToString("c")) 'return the ocurrence with "c"

        'Display the average of all list prices
        'We can short hand the function part so we don't have to put the return or end the function  
        Console.WriteLine(
            products.Average(Function(p) p.Value.ListPrice).ToString("c"))

        'Display the minimum of all list prices
        Console.WriteLine(
            products.Min(Function(p) p.Value.ListPrice).ToString("c"))

        'Display the maximum of all list price
        Console.WriteLine(
            products.Max(Function(p) p.Value.ListPrice).ToString("c"))

        Console.ReadKey()
    End Sub
    Sub ListOfT()
        'Dim the list of products
        Dim products = LoadProductsList()
        'print the information
        Console.WriteLine(products(0).Name)
        Console.WriteLine(products(1).ProductID)
        Console.WriteLine(products(2).ListPrice)
        Console.WriteLine(products(3).ProductNumber)
        Console.WriteLine(products(4).SellEndDate)
        Console.WriteLine(products(5).SellStartDate)
        Console.WriteLine(products(6).Size)
        Console.WriteLine(products(7).StandardCost)
        Console.WriteLine(products(8).Colour)
        Console.WriteLine(products(9).Weight)

        'See if a specific key exists in the list; True/False
        Console.WriteLine(
            products.Exists(Function(p) p.ProductID = 706)) 'True
        Console.WriteLine(
            products.Exists(Function(p) p.ProductID = 99)) 'False
        Console.WriteLine(
            products.Exists(Function(p) p.Colour = "Black")) 'True (it is case sensitive)

        Console.ReadKey()
    End Sub
    Sub ListOfTMethods()
        Dim products = LoadProductsList()

        'Display the total number of items in the list
        Console.WriteLine(products.Count)

        'Remove an item by index
        products.RemoveAt(0)
        products.RemoveAt(1)
        Console.WriteLine(products.Count)

        'Remove an item by a product object, using the LINQ expression looping through passing p to the function to find the item based on a criteria
        products.Remove(products.Find(Function(p) p.ProductID = 708))
        Console.WriteLine(products.Count)

        'Remove all items
        products.Clear()
        Console.WriteLine(products.Count)

        Console.ReadKey()
    End Sub
    Sub LINQOnListOfT()
        Dim products = LoadProductsList()

        'display the sum of all list prices
        Console.WriteLine(
            products.Sum(Function(p) p.ListPrice).ToString("c"))

        'display the average list price
        Console.WriteLine(
            products.Average(Function(p) p.ListPrice).ToString("c"))

        'display the minimum of the list prices
        Console.WriteLine(
            products.Min(Function(p) p.ListPrice).ToString("c"))

        'display the maximum of the list prices
        Console.WriteLine(
            products.Max(Function(p) p.ListPrice).ToString("c"))
    End Sub

    Sub DoWhileLoop()
        'gives us the generic list of product objects:
        Dim products = LoadProductsList()
        'two variables to keep track of the point in the list we're looping through:
        Dim index As Integer = 0
        'one to keep track of the sum of the 
        Dim sum As Decimal = 0

        'Start our loop
        'While index is less than the amount of the products minus 1...
        Do While index < (products.Count - 1)
            '... do the following
            'output each product in question for this given index number
            'Overriding function kicks in here
            Console.WriteLine(products(index).ToString())
            'add the list price for the given product we are on to the sum variable
            'by the end of the loop this will be the total sum of all the list prices of all the products
            sum += products(index).ListPrice

            'jump to the next index in the list, then go to the start of the while loop again. 
            'If theThen condition Is still True, re-run the block Of code For the Next index value
            'Repeat until end of the list
            index += 1

            'go back up to the start of the while loop:
        Loop

        'once the loop is done, we can output the final sum output amount to the console
        Console.WriteLine("Sum: " & sum.ToString("c"))

        Console.ReadKey()
    End Sub

    Sub DoLoopWhile()
        'set up variables as before
        Dim products = LoadProductsList()
        Dim index As Integer = 0
        Dim sum As Decimal = 0

        'Now we can define a loop that we know we want to enter at least once, because the conditional statement is checked AT THE END
        Do
            'same code as before ^^^
            Console.WriteLine(products(index).ToString())
            sum += products(index).ListPrice

            'If you want to exit the code early, use:
            'Exit Do
            'this is perhaps best used with an if statement

            index += 1

            'If the condition is still true after a cycle, do it again
            'Perhaps the logic is easier to follow in terms of length of lists in this case. 
            '"Have we done all the items in the list by now? Yes? Ok, end the loop. No? Ok, go again":
        Loop While index < (products.Count)

        Console.WriteLine("Sum: " & sum.ToString("c"))

        Console.ReadKey()
    End Sub

    Sub DoUntilLoop()
        'Set up our variables
        Dim products = LoadProductsList()
        Dim index As Integer = 0
        'this time we're going to define a min variable, which starts as the maximum possible decimal value allowed in VB (we hope something will be smaller than that)
        Dim min As Decimal = Decimal.MaxValue

        'Do everything in the following code block, until the condition following becomes true
        'Do everything following, until the index becomes greater than the amount of products less one
        Do Until index > (products.Count - 1)
            'Write out the product elements
            Console.WriteLine(products(index).ToString())

            'Set the minimum value to a decimal type, depending on a condition
            min = Convert.ToDecimal(
                IIf(products(index).ListPrice < min, products(index).ListPrice, min))
            'If the list price of the product for the current index is less than the current minimum value: 
            'set the minimum to that product list value (it becomes the new minimum), 
            'If Not, leave it as minimum

            'syntax is as followes: IIf(condition(true/false?), return this if true, return this if false)

            'Go to next index
            index += 1
        Loop

        'Output the minimum value, that is decided since all products list prices have been cycled through and the lowest value will remain as the min variable
        Console.WriteLine("Min: " & min.ToString())

        Console.ReadKey()
    End Sub

    Sub DoLoopUntil()
        'Set up the variables in the same way as previously ^^^
        Dim products = LoadProductsList()
        Dim index As Integer = 0
        Dim max As Decimal = Decimal.MinValue

        'Start the loop, bear in mind with this syntax, it will always do one loop first, before checking a given condition
        Do
            'output the products
            Console.WriteLine(products(index).ToString())

            'check to see if the current index of product list price is bigger than the current maximum variable saved
            'if it is, replace the variable with this new larger amount
            'If it is not, keep the same variable value max as it is known to be the biggest so far
            max = Convert.ToDecimal(
                IIf(products(index).ListPrice > max, products(index).ListPrice, max))

            index += 1

            'checking whether or not to continue to the next loop, only stop if the following condition is true:
        Loop Until index > (products.Count - 1)
    End Sub

    Sub ForNext()
        Dim products = LoadProductsList()

        'Give a range of values over which to iterate a process
        'List index values start at zero and end at one less than their length 
        For index As Integer = 0 To (products.Count - 1)
            'do the following for each of the placeholder value "index"
            'index is a range of values and the products list can be referenced for each of these range of values
            Console.WriteLine(products(index).ToString())

            'incremement to the next value in the loop
        Next

        Console.ReadKey()

    End Sub

    Sub ForStepNext()
        Dim products = LoadProductsList()

        'We can use the Step keyword to alter this syntax to tell the increment to loop in a certain length each time. 
        'By default, the step to iterate is 1, meaning it just jumps to the next item if it's not specified
        'Here we have specified the list of indexes to iterate over to be from the largest number to the smallest number, and made sure to Step as -1
        'In essence, we print the list of products backwards because of the set-up line
        For index As Integer = (products.Count - 1) To 0 Step -1
            Console.WriteLine(products(index).ToString())
        Next

        Console.ReadKey()
    End Sub

    Sub ForEachNext()
        'Set up list of products
        Dim products = LoadProductsList()

        'We can be a bit more vague about the length with this syntax, as in we don't need to know which index values the list runs from and to, or the increment for this
        'Instead, we can just tell the system to only do the following block of code on the items specified in this line
        'based of this, it takes care of the indexing for us, once it's works out how many of the specified items are in the list
        For Each prod As Product In products
            Console.WriteLine(prod.ToString())

            'We can exit the for loop at any time using
            'Exit For
        Next

        Console.ReadKey()
    End Sub

    Sub IfElseElseIfStatements()
        Dim products = LoadProductsList()

        'An example of an If, ElseIf and Else structure for operating based on conditional operations (true/false)
        For index As Integer = 0 To (products.Count - 1)
            If products(index).Colour = "Red" Then
                Console.WriteLine("Red colour product: " & products(index).ToString())
            ElseIf products(index).Colour = "Black" Then
                Console.WriteLine("Black colour product: " & products(index).ToString())
            Else
                Console.WriteLine("Other colour product: " & products(index).ToString())
            End If
        Next

        Console.ReadKey()
    End Sub

    Sub SelectCase()
        Dim products = LoadProductsList()

        'iterate through the list
        For index As Integer = 0 To (products.Count - 1)
            'case syntax that registers for a given value of each product (cycled through in the For loop), operate the following code.
            'this is more comprehensive than the If statement set up as it can eliminate monotony

            'For the case of each product's size:
            Select Case products(index).Size
                'for the case where the size is equal to 58...
                Case "58"
                    '...print the following about that product
                    Console.WriteLine("Size 58: " & products(index).Name)

                'for the case where the size is equal to L...
                Case "L"
                    '...print the following about that product
                    Console.WriteLine("Size L: " & products(index).Name)

                'for the case where the size is equal to M...
                Case "M"
                    '...print the following about that product
                    Console.WriteLine("Size M: " & products(index).Name)

                    'Anything that does not match any of the cases in the above code, will fall into the else statement:
                Case Else
                    Console.WriteLine(products(index).Name & " has size " & products(index).Size)
            End Select
        Next

        Console.ReadKey()
    End Sub

    Sub CompilerConstants()
        'The following is decision making structure that happens at compile time, not run time:
        'A compiler construct not a run time construct
        'DEBUG is a compile construct
#If DEBUG Then
        Console.WriteLine("In DEBUG mode")
#Else
    Console.WriteLine("In RELEASE mode")
#End If

        'We can define our own compiler constants that are defined on compiling
        '#Const conLANGUAGE = "ENGLISH"
#Const conLANGUAGE = "GERMAN"

#If conLANGUAGE = "ENGLISH" Then
        Console.WriteLine("Good Morning Mr. Jones")
#Else
        Console.WriteLine("Guten Morgen Herr Jones")
#End If

        Console.ReadKey()
    End Sub

    ' We can add a "#Region" syntax with a matching "#End Region" syntax to make it easier to collapse areas and focus better on the work we're doing
#Region "Sub - With End With"
    Sub WithEndWith()
        Dim prod As New Product
        'We can eliminate a little bit of typing with this With keyword
        With prod
            .ProductID = 1
            .Name = "A new product"
            .ProductNumber = "A101"
            .StandardCost = 50
            .ListPrice = 100
        End With

        Console.WriteLine(prod.ToString())

        Console.ReadKey()
    End Sub
#End Region

End Module