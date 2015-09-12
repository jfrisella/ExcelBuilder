#Excel Builder


Basic usage of Excel classes.

```javascript

//Create an Excel Workbook
var workbook = new Excel.Workbook({
	name: "MyWorkbook"
});

//Could also add/change name after instantiation
workbook.setName("NewName");


//Create Worksheet
var worksheet1 = new Excel.Worksheet({
	name: "Worksheet1"
});

//Could also add/change name after instantiation
worksheet1.setName("NewWorksheetName");

//Have some data as two dimensional array
var data = [
	[1,2,3],
	[4,5,6],
	[7,8,9]
];

//Convenience method for filling worksheet
//must be two dimensional array
worksheet1.addAllData(data);

//Add Worksheet to Workbook
workbook.add(worksheet1);


//Render Workbook
var xml = workbook.render();

//Do something with XML

//Client Side you can automatically fire the download
//this will create a link, click it, and remove the link after click
//basically firing the download
workbook.download(); 

//Just get the url
var url = workbook.toUrl();


```


There are a few more Classes you have access to, if you want to customize a little more. Below is a list of 
all the classes available in the Excel Library.

- Workbook
- Worksheet
- Row
- Cell
- Data



Workbook:

```javascript

//Instantiate new Workbook
//and pass in Workbook name
var workbook = new Excel.Workbook({
	name: "MyWorkBook"
});

//Set Name of Workbook
workbook.setName("MyWorkbookName");

//Get Name of Workbook
var name = workbook.getName();		//--> "MyWorkbookName"

//Add Worksheet to Workbook
workbook.add(MyWorksheet);

//Render Workbook
var xml = workbook.render();		//--> Rendered XML

//Initiate Download of Excel Document
//Client side
workbook.download();

//Returns url of XML
var link = workbook.toUrl();		//--> XML Link

```


Worksheet: 

```javascript

//Instantiate new Worksheet
var worksheet = new Excel.Worksheet({
	name: "MyWorkSheet"
});

//Set Name of Worksheet
worksheet.setName("MyWorkSheetName");

//Get Name of Worksheet
var name = worksheet.getName();		//--> "MyWorkSheetName"

//Add Row to Worksheet
worksheet.add(Row);

//Add Full list of Data to Worksheet
//Convenience method to add two dimensional array to Worksheet
worksheet.addAllData([
	["some", "data", "goes"],
	["here", 4, 5]
]);

//Render Worksheet
var xml = worksheet.render();		//--> Rendered XML

```


Row:

```javascript

//Initiate new Row
var row = new Row();

//Add Cell to Row
row.add(Cell);

//Render Row
var xml = row.render();			//--> Rendered XML

```


Cell:

```javascript

//Initiate new Cell
var cell = new Cell();

//Add Data to Cell
cell.add(Data);

//Render Cell
var xml = cell.render();		//--> Rendered XML

```


Data:

```javascript

//Initiate new Data
var data = new Data();

//Add text/number to Data
data.add("Some Text" || 5);

//Render Data
var xml = data.render();		//--> Rendered XML

```