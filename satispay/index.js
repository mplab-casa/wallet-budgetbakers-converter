// XLSX is a global from the standalone script

async function analyze() {
  var output = [];

  const file = document.getElementById("inputFile").files[0];
  const data = await file.arrayBuffer();
  const input = XLSX.read(data);
  var sheet = input.SheetNames[0];
  output.push(GenerateRow("Date", "Description", "Expense", "Entrance", "Type"))
  var sheetLenght = input.Sheets[sheet]["!ref"].split(":")[1].substring(1);
  var fromDate = document.getElementById("inputFromDate").value?new Date(document.getElementById("inputFromDate").value):undefined
  if(fromDate){fromDate.setHours(0,0,0,0);}
  for (i = 2; i <= sheetLenght; i++) {
    // La data è nella colonna A con formato 05/01/2026 17:43:33
    var dateValue = input.Sheets[sheet]["A" + i] ? input.Sheets[sheet]["A" + i].v : null;
    var transactionDate;
    
    // Prova diversi formati per la data
    if (dateValue) {
      if (typeof dateValue === 'number') {
        // Se è un numero seriale di Excel
        transactionDate = XLSX.SSF.parse_date_code(dateValue);
        transactionDate = new Date(transactionDate.y, transactionDate.m - 1, transactionDate.d, transactionDate.H, transactionDate.M, transactionDate.S);
      } else {
        // Se è una stringa, prova il parsing con moment
        transactionDate = moment(dateValue, "DD/MM/YYYY HH:mm:ss").toDate();
      }
    }
    
    // Amount è nella colonna D
	var amount = input.Sheets[sheet]["D" + i] ? parseFloat(input.Sheets[sheet]["D" + i].v) : 0;
	var type = input.Sheets[sheet]["E" + i] ? input.Sheets[sheet]["E" + i].v : "";
	var description = input.Sheets[sheet]["B" + i] ? input.Sheets[sheet]["B" + i].v : "";
	
	// Rimuovi tutto quello che non è lettere, numeri o caratteri speciali comuni e trim spazi
	description = description.replace(/[^\w\s\.,;:!?\-\(\)\[\]{}@#$%&*+=<>\/\\'"]/g, '').trim();
	type = type.replace(/[^\w\s\.,;:!?\-\(\)\[\]{}@#$%&*+=<>\/\\'"]/g, '').trim();
	
	if(!fromDate || !transactionDate || fromDate <= transactionDate)
    output.push(GenerateRow(
      transactionDate ? (transactionDate.toLocaleDateString('it-IT') + " " + transactionDate.toLocaleTimeString('it-IT')) : "",
      description,
	  amount < 0 ? -amount : 0,
      amount > 0 ? amount : 0,
	  type
	  ));
  }
  var wb = aoa_to_workbook(output);
  // Genera nome file con data e ora attuali
  var now = new Date();
  var timestamp = now.getFullYear() + 
    String(now.getMonth() + 1).padStart(2, '0') + 
    String(now.getDate()).padStart(2, '0') + '_' +
    String(now.getHours()).padStart(2, '0') + 
    String(now.getMinutes()).padStart(2, '0') + 
    String(now.getSeconds()).padStart(2, '0');
  XLSX.writeFile(wb, "satispay_export_" + timestamp + ".xlsx");
}

function GenerateRow(Date, Description, Expense, Entrance, Type) {
  var row = [];
  row[0] = String(Date || "");
  row[1] = String(Description || "");
  row[2] = String(Expense || 0);
  row[3] = String(Entrance || 0);
  row[4] = String(Type || "");
  return row;
}

function aoa_to_workbook(data/*:Array<Array<any> >*/, opts)/*:Workbook*/ {
  return sheet_to_workbook(XLSX.utils.aoa_to_sheet(data, opts), opts);
}
function sheet_to_workbook(sheet/*:Worksheet*/, opts)/*:Workbook*/ {
  var n = opts && opts.sheet ? opts.sheet : "Sheet1";
  var sheets = {}; sheets[n] = sheet;
  return { SheetNames: [n], Sheets: sheets };
}