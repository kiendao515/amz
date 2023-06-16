const XLSX = require('xlsx');
const il = XLSX.readFile("Inventory-Ledger-06.05.23-07.06.23.xlsx")
const workbook = XLSX.readFile('output.xlsx');
const { getJsDateFromExcel } = require("excel-date-to-js");
class InventoryLedger {
    constructor(date, fnsku, msku, quantity, referenceID, disposition, event_type, date_time) {
        this.date = date;
        this.fnsku = fnsku;
        this.msku = msku;
        this.quantity = quantity;
        this.referenceID = referenceID;
        this.disposition = disposition;
        this.event_type = event_type;
        this.date_time = date_time;
    }
}
class Transaction {
    constructor(date, sku, fnsku, type, quantity, disposition, shipmentID) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.type = type;
        this.quantity = quantity;
        this.disposition = disposition;
        this.shipmentID = shipmentID;
    }
}
function getListTransaction(inventoryLedger) {
    const transactions = [];
    inventoryLedger.forEach(element => {
        const { date_time, fnsku, msku, quantity, disposition, event_type } = element;
        if (disposition === 'SELLABLE' && (event_type === 'Shipments' || event_type === 'CustomerReturns' ||
            event_type === 'Adjustments' || event_type === 'VendorReturns')) {
            transactions.push({
                date: date_time,
                sku: msku,
                fnsku,
                type: event_type,
                quantity: quantity,
                disposition: disposition,
                shipmentID: null
            });
        }
    });

    return transactions.map(t => new Transaction(
        t.date,
        t.sku,
        t.fnsku,
        t.type,
        t.quantity,
        t.disposition,
        t.shipmentID
    ));
}
function getListRecepts(inventoryLedger){
    let recepts =[]
    inventoryLedger.forEach(element => {
        const { date_time, fnsku, msku, quantity, disposition, event_type } = element;
        if(disposition === 'SELLABLE' && event_type === 'Receipts'){
            recepts.push({
                date: date_time,
                sku: msku,
                fnsku,
                type: event_type,
                quantity: quantity,
                disposition: disposition,
                shipmentID: element.referenceID
            })
        }                                                                                                                                                                                                                                                                                                                        
    });
    return recepts.map(t => new Transaction(
        t.date,
        t.sku,
        t.fnsku,
        t.type,
        t.quantity,
        t.disposition,
        t.shipmentID
    ));
}
GenerateFile = async () => {
    const ws1 = il.Sheets["Sheet1"]
    const worksheet = workbook.Sheets['Danh sách giao dịch bổ sung']; // Replace 'Sheet1' with the actual sheet name
    // Use XLSX.utils.sheet_to_json() to convert the worksheet to a JSON array
    const arr8 = XLSX.utils.sheet_to_json(ws1)
    let inventoryLedger = arr8.map((row) => {
        return new InventoryLedger(
            row['Date'],
            row['FNSKU'],
            row['MSKU'],
            row['Quantity'],
            row['Reference ID'],
            row['Disposition'],
            row['Event Type'],
            row['Date and Time']
        )
    })

    let transations = getListTransaction(inventoryLedger)
    let recepts = getListRecepts(inventoryLedger);
    const receptSheet = XLSX.utils.json_to_sheet(recepts)
    XLSX.utils.book_append_sheet(workbook, receptSheet, "Danh sách RECEPT");
    const existingData = XLSX.utils.sheet_to_json(worksheet);
    const mergedData = [ ...transations, ...existingData ];
    const newSheet = XLSX.utils.json_to_sheet(mergedData);
    workbook.Sheets['Danh sách giao dịch bổ sung'] = newSheet; // Replace 'Sheet1' with the actual sheet name
    XLSX.writeFile(workbook, 'news.xlsx');

}

GenerateFile()