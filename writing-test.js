var Excel =  require('exceljs');


    // write to a file
var workbook = createAndFillWorkbook();
workbook.xlsx.writeFile(cleanedCopy)
    .then(function() {
        // done
    });
