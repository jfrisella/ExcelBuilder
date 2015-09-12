/**
    Excel Builder

    The MIT License (MIT)

    Copyright (c) 2015 James Frisella

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.

*/


/**
*   Create Excel Object
*/
if ("object" !== typeof Excel) {
    Excel = {};
}

(function (self) {

    /**
    *    Workbook class
    */
    self.Workbook = function (options) {
        
        options = (typeof options === "object") ? options : {};

        //Worksheets
        var worksheets = [];

        //Name
        var name = options.name || "WorkBook";

        var workbookXML = '<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40"><ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel"><ActiveSheet>0</ActiveSheet></ExcelWorkbook><Styles><Style ss:ID="Default" ss:Name="Normal"><Alignment ss:Vertical="Bottom"/><Borders/><Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/><Interior/><NumberFormat/><Protection/></Style></Styles>';
        var workbookXMLEnd = '</Workbook>';


        //Add Worksheet
        this.addWorksheet = function (sheet) {
            if (!sheet) return;
            worksheets.push(sheet);
        }

        //Render Workbook
        this.render = function () {
            var final = [];
            for (var i = 0; i < worksheets.length; i += 1) {
                final.push(worksheets[i].render());
            }
            final.unshift(workbookXML);
            final.push(workbookXMLEnd);
            return final.join("");
        }

        //To Url - creates client side url for .xls download
        this.toUrl = function(){
            return toUrl(
                this.render()
            );
        }

        //Download - creates/follows/deletes link to .xls file
        this.download = function(){
            var link;
            link = document.createElement("a");
            link.href = this.toUrl();
            link.download = name;
            link.target = '_blank';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        }

        //Get Name
        this.getName = function(){
            return name;
        }

        //Set Name
        this.setName = function(n){
            name = n;
        }
    }

    /**
    *   Worksheet Class
    */
    self.Worksheet = function (options) {

        //Options
        options = (typeof options === "object") ? options : {};

        //Rows
        var rows = [];

        //Sheet XML
        var sheetXML = '<Worksheet ss:Name="{{SHEET_NAME}}"><Table>';
        var sheetXMLEnd = "</Table></Worksheet>";

        var name = options.name || 
            (function(){
                //If using default name
                //increment worksheet counter
                //everytime a new worksheet is created
                var count = worksheet_counter;
                worksheet_counter += 1;
                return "Worksheet" + count;
            }());

        //Add Rows
        this.addRow = function (row) {
            if (!row) return;
            rows.push(row);
        }

        //Add All Data
        //Convenience method for population worksheet
        this.addAllData = function(items){
            
            for (var i = 0; i < items.length; i += 1) {
                row = new self.Row();
                for (var j = 0; j < items[i].length; j += 1) {
                    data = new self.Data();
                    data.addText(items[i][j]);
                    cell = new self.Cell();
                    cell.addData(data);
                    row.addCell(cell);
                }
                this.addRow(row);
            }
        }

        //Render Worksheet
        this.render = function () {
            var final = [];
            for (var i = 0; i < rows.length; i += 1) {
                final.push(rows[i].render());
            }
            final.unshift(buildSheet());
            final.push(sheetXMLEnd);
            return final.join("");
        }


        //Build Sheet
        var buildSheet = function () {
            return sheetXML.replace(/({{SHEET_NAME}})/i, name);
        }

        //Get Name
        this.getName = function(){
            return name;
        }

        //Set Name
        this.setName = function(n){
            name = n;
        }

    }

    /**
    *   Row Class
    */
    self.Row = function (options) {

        //Cell
        var cells = [];

        //Add Cell
        this.addCell = function (cell) {
            if (!cell) return;
            cells.push(cell);
        }

        //Render Row
        this.render = function () {
            var final = [];
            for (var i = 0; i < cells.length; i += 1) {
                final.push(cells[i].render());
            }
            final.unshift("<Row>");
            final.push("</Row>");
            return final.join("");
        }
    }

    /**
    *   Cell Class
    */
    self.Cell = function (options) {

        //Rows
        var datas = [];

        //Add Data
        this.addData = function (data) {
            if (!data) return;
            datas.push(data);
        }

        //Render Cell
        this.render = function () {
            var final = [];
            for (var i = 0; i < datas.length; i += 1) {
                final.push("<Cell>");
                final.push(datas[i].render());
                final.push("</Cell>");
            }
            return final.join("");
        }
    }

    /**
    *   Data Class
    */
    self.Data = function (options) {

        //Text
        var items = [];

        //Add Text
        this.addText = function (text) {
            if (!text && parseInt(text) !== 0) return;
            items.push(text);
        }

        //Render Data
        this.render = function () {
            var final = [];
            for (var i = 0; i < items.length; i += 1) {
                final.push("<Data ss:Type='String'>");
                final.push(items[i]);
                final.push("</Data>");
            }
            return final.join("");
        }

    }


    /**
    *   Build Worksheet
    *       - static helper function since building the worksheet is basically repeated code
    *       @param worksheet {Object} - the worksheet instance to add data
    *       @param items {Array(2D)} - a two dimensional array of data to add to worksheet
    *       @return no need since items are added to worksheet directly
    */
    self.buildWorksheet = function (worksheet, items) {

        for (var i = 0; i < items.length; i += 1) {
            row = new self.Row();
            for (var j = 0; j < items[i].length; j += 1) {
                data = new self.Data();
                data.addText(items[i][j]);
                cell = new self.Cell();
                cell.addData(data);
                row.addCell(cell);
            }
            worksheet.addRow(row);
        }

    }


    //Private Functions
    var toUrl = function(renderedXml){
        return download_uri + 
            window.btoa(
                unescape(
                    encodeURIComponent(
                        renderedXml
                    )
                )
            );
    }
    
    //Only used when building uri for client side applications
    var download_uri = "data:application/vnd.ms-excel;base64,";

    //Workbook counter
    //Needed if not naming worksheets
    //to make sure they do not repeat names
    var worksheet_counter = 1;

} (Excel));