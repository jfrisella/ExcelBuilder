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
workbook.addWorksheet(worksheet1);


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