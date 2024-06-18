WpfWindow("OpenText MyFlight Sample Application").WpfEdit("passengerName").Set "John Smith"
WpfWindow("OpenText MyFlight Sample Application").WpfButton("ORDER").Click

'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
'Start monitor transaction time
services.StartTransaction "Process Order"
WpfWindow("OpenText MyFlight Sample Application").WpfProgressBar("progBar").WaitProperty "value", 100
services.EndTransaction "Process Order"
'End monitor transaction time

'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
'Checkpoint string "Order #* completed" is displayed.
WpfWindow("OpenText MyFlight Sample Application").WpfObject("Order # completed").Check CheckPoint("Order # completed") @@ hightlight id_;_2142499752_;_script infofile_;_ZIP::ssf11.xml_;_
'Output "Order Price" to dataTable
WpfWindow("OpenText MyFlight Sample Application").WpfObject("Order $ Price").Output CheckPoint("Order $ Price")

'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
'DB SQLite3 Output Wizard
' the database connection string must be parameterized to match who is running this test so
whoIsLoggedIn = Environment.Value("UserName") ' get the name
DataTable.Value("userLoggedIn", dtLocalSheet) = whoIsLoggedIn ' store it into the data table
' now, in this connection string, it can concatentate the string
DbTable("DbTable").Output CheckPoint("DbTable")

'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
'Capture "Order number" in Output Parameter
WpfWindow("OpenText MyFlight Sample Application").WpfObject("Order # completed").Output CheckPoint("Order #")
'Parse "Order Number" with Regular Expression
str = Parameter.Item("Order_Number")
Set RegEx = New RegExp
RegEx.Pattern = "[^\d*]"
RegEx.IgnoreCase = True 
RegEx.Global = True 
Parameter.Item("Order_Number") = RegEx.Replace(str, "")

'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
WpfWindow("OpenText MyFlight Sample Application").WpfButton("NEW SEARCH").Click
