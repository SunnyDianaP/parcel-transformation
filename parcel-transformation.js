
var Excel =  require('exceljs');

var workbook = new Excel.Workbook();
workbook.xlsx.readFile('taco.test.xlsx')
	   .then(function() 
		{
 
	    	var worksheet = workbook.getWorksheet('Taco'); // fetch sheet by name
	    	var firstCol = worksheet.getColumn(1); // Access an individual columns by num
	    	// var cell = worksheet.getCell('A3').value // Get string within an object

				worksheet.getCell('A3').value = 5;

	    	firstCol.eachCell (function(cell, rowNumber) 
	    	{
	    		var strings = cell.value.toString(); //stringify
	    		var cleaned = strings.replace(/\W/g, '') //remove all non-alphanumeric characters 
	    		cell.value = cleaned 
	    		console.log (cleaned); 
	    	});
    // console.log (cell);
   		 

// write to a csv file
			workbook.xlsx.writeFile('cleaned-copy.xlsx')
		    .then(function() 
		    {
		        // done
		    });

		});

