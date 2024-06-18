'DB SQLite3 Checkpoint Wizard 
'CheckPoint If DB matches passed parameter "OrderNumber" & Static "CustomerName" exists
'msgbox Parameter.Item("Order_Number") ' this halts execution until you dismiss the msgbox
print Parameter.Item("Order_Number")  ' this concatenates it to the end of the output window

' the database connection string must be parameterized to match who is running this test so
whoIsLoggedIn = Environment.Value("UserName") ' get the name
DataTable.Value("userLoggedIn", dtLocalSheet) = whoIsLoggedIn ' store it into the data table
' now, in the connection string in the cell, it can concatentate the string
DbTable("DbTable_Checkpoint").Check CheckPoint("DbTable_Checkpoint")

'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
WpfWindow("OpenText MyFlight Sample Application").WpfTabStrip("WpfTabStrip").Select "SEARCH ORDER"
WpfWindow("OpenText MyFlight Sample Application").WpfRadioButton("byNumberRadio").Set
WpfWindow("OpenText MyFlight Sample Application").WpfEdit("byNumberWatermark").Set Parameter("Order_Number") @@ hightlight id_;_1926714864_;_script infofile_;_ZIP::ssf6.xml_;_
WpfWindow("OpenText MyFlight Sample Application").WpfButton("SEARCH").Click

'-----------------------------------------------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------------------------------------------
'Using the DB output parameters from action 5 to perform GUI validation

WpfWindow("OpenText MyFlight Sample Application").Check CheckPoint("FlightNum") @@ hightlight id_;_1445576_;_script infofile_;_ZIP::ssf6.xml_;_
WpfWindow("OpenText MyFlight Sample Application").WpfComboBox("numOfTicketsCombo").Check CheckPoint("NumOfTicketsCombo")
WpfWindow("OpenText MyFlight Sample Application").WpfObject("Order $ Price").Check CheckPoint("Order $ Price") @@ hightlight id_;_1997777216_;_script infofile_;_ZIP::ssf11.xml_;_

'Application object highlight
WpfWindow("OpenText MyFlight Sample Application").WpfObject("Order $ Price").highlight

foo = 1 ' to make it easy to set a breakpoint so that variable values can be examined

