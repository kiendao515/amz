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
    constructor(date, sku, fnsku, type, quantity, disposition, shipmentID, shipment_recept, cogs) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.type = type;
        this.quantity = quantity;
        this.disposition = disposition;
        this.shipmentID = shipmentID;
        this.shipment_recept = shipment_recept;
        this.cogs = cogs;
    }
}
class Cog {
    constructor(sku, fnsku, current_shipment, current_shipment_cog, date, to_date, remainder,
        next_shipment, next_shipment_cog) {
        this.sku = sku;
        this.fnsku = fnsku;
        this.current_shipment = current_shipment;
        this.current_shipment_cog = current_shipment_cog;
        this.date = date;
        this.to_date = to_date;
        this.remainder = remainder;
        this.next_shipment = next_shipment;
        this.next_shipment_cog = next_shipment_cog
    }
}
class Result {
    constructor(date, sku, fnsku, shipmentID, nextShipmentID, sale_quantity, total_inventory, data, listShipmentID,
        listQuantityOfShipment) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.shipmentID = shipmentID;
        this.nextShipmentID = nextShipmentID;
        this.sale_quantity = sale_quantity;
        this.total_inventory = total_inventory;
        this.data = data;
        this.listShipmentID = listShipmentID;
        this.listQuantityOfShipment = listQuantityOfShipment;
    }
}
function getListTransaction(inventoryLedger) {
    const transactions = [];
    inventoryLedger.forEach(element => {
        const { date_time, fnsku, msku, quantity, disposition, event_type, referenceID } = element;
        if (disposition === 'SELLABLE' && (event_type === 'Shipments' || event_type === 'CustomerReturns' ||
            event_type === 'Adjustments' || event_type === 'VendorReturns' || event_type === 'Receipts')) {
            transactions.push({
                date: date_time,
                sku: msku,
                fnsku,
                type: event_type,
                quantity: quantity,
                disposition: disposition,
                shipmentID: null,
                shipment_recept: event_type === 'Receipts' ? referenceID : null,
                cogs: null
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
        t.shipmentID,
        t.shipment_recept,
        t.cogs
    ));
}
let findFutureDate = async (skuData, transations) => {
    transation_recepts = transations.filter(t => t.type === 'Receipts')
    skuData.forEach(sku => {
        let futureTransactions = transation_recepts.filter(function (element) {
            return element.sku === sku.sku;
        })
        sku.listShipmentID = sku.listShipmentID.split(',')
        sku.listQuantityOfShipment = sku.listQuantityOfShipment.split(',')
        futureTransactions.forEach(transaction => {
            let beforeCurrentShipment = sku.listShipmentID.filter(value => value !== sku.shipmentID).slice(0, sku.listShipmentID.indexOf(sku.shipmentID));
            let afterCurrentShipment = sku.listShipmentID.slice(sku.listShipmentID.indexOf(sku.shipmentID));
            if (afterCurrentShipment.includes(transaction.shipment_recept)) {
                const index = sku.listShipmentID.indexOf(sku.shipmentID);
                sku.listQuantityOfShipment[index] = parseInt(transaction.quantity) + parseInt(sku.listQuantityOfShipment[index]);
            } else if (beforeCurrentShipment.includes(transaction.shipment_recept)) {
                const index = sku.listShipmentID.indexOf(transaction.shipment_recept);
                sku.listQuantityOfShipment[index] = parseInt(transaction.quantity) + parseInt(sku.listQuantityOfShipment[index]);
            } else {
                sku.listShipmentID.unshift(transaction.shipment_recept);
                sku.listQuantityOfShipment.unshift(transaction.quantity);
            }
        });
        sku.listShipmentID = sku.listShipmentID.join(',');
        sku.listQuantityOfShipment = sku.listQuantityOfShipment.join(',');
    })
    return skuData
}
const findDate = async (skuData, transactions) => {
    const cogs = [];
    for (let i = 0; i < skuData.length; i++) {
        const element = skuData[i];
        element.listShipmentID = element.listShipmentID.split(',');
        element.listQuantityOfShipment = element.listQuantityOfShipment.split(',');
        const index = element.listShipmentID.indexOf(element.shipmentID);
        if (element.nextShipmentID != null && element.listQuantityOfShipment[index] > 0) {
            const matchingTransaction = transactions.find(t => (element.shipmentID === t.shipmentID && t.sku === element.sku));
            if (matchingTransaction) {
                let total = 0;
                let tmp = transactions.filter(t => t.sku === element.sku && t.type != 'Receipts')
                const matchingTransactionIndex = tmp.indexOf(matchingTransaction);
                for (let j = matchingTransactionIndex - 1; j >= 0; j--) {
                    const t = tmp[j];
                    total += t.quantity;
                    // if(element.sku === 'Template-set3'){
                    //     console.log(t);
                    //     console.log(-parseInt(element.listQuantityOfShipment[index]),total);
                    // }
                    if (-parseInt(element.listQuantityOfShipment[index]) >= total) {
                        const d = new Date(t.date);
                        d.setFullYear(d.getFullYear() + 3);
                        t.shipmentID = element.listShipmentID[index - 1];
                        cogs.push(new Cog(
                            t.sku,
                            element.fnsku,
                            element.listShipmentID[index - 1],
                            null,
                            new Date(t.date),
                            new Date(d),
                            parseInt(total) + parseInt(element.listQuantityOfShipment[index]),
                            element.listShipmentID[index - 2],
                            null
                        ));
                        let rs = await findNextDate(element, transactions, parseInt(total) + parseInt(element.listQuantityOfShipment[index]))
                        console.log(rs);
                        cogs.push(...rs)
                        break;
                    }
                }
            }
        }
    }

    return [cogs, transactions];
};

const findNextDate = async (skuData, transactions, remainder) => {
    console.log(skuData);
    const cogs = [];
    const listTransactionOfSku = transactions.filter(t => t.sku === skuData.sku && t.type !== 'Receipts');
    let tmp = [];
    let rs = [];
    const index = skuData.listShipmentID.indexOf(skuData.shipmentID);
    skuData.listQuantityOfShipment[index - 1] = parseInt(skuData.listQuantityOfShipment[index - 1]);
    skuData.listQuantityOfShipment[index - 1] += parseInt(remainder);
    if (index > 0) {
        for (let i = index - 1; i >= 0; i--) {
            tmp.push({
                shipmentID: skuData.listShipmentID[i],
                quantityOfShipment: parseInt(skuData.listQuantityOfShipment[i]),
                date: null
            });
        }
    }
    let filteredTransactions = listTransactionOfSku;
    const startIndex = listTransactionOfSku.findIndex(t => t.shipmentID === skuData.nextShipmentID);
    filteredTransactions = listTransactionOfSku.slice(0, startIndex).reverse();
    // if(skuData.sku === 'Template-set3'){
    //     console.log(filteredTransactions);
    // }

    const processNextTmpElement = (index) => {
        if (index >= tmp.length) {
            return;
        }
        let total = 0;
        let stopIndex = -1;
        console.log(filteredTransactions);
        for (let j = 0; j < filteredTransactions.length; j++) {
            const t = filteredTransactions[j];
            console.log("chay vao day", t);
            if (t.sku === skuData.sku) {
                total += t.quantity;
                if (skuData.sku === 'Template-set3') {
                    console.log(t, total);
                }
                if (-tmp[index].quantityOfShipment >= total) {

                    t.shipmentID = tmp[index + 1]?.shipmentID;
                    tmp[index].date = new Date(t.date);
                    if (tmp[index + 1]?.shipmentID != undefined) {
                        const d = new Date(t.date);
                        d.setFullYear(d.getFullYear() + 3);
                        cogs.push(new Cog(
                            t.sku,
                            t.fnsku,
                            tmp[index + 1]?.shipmentID,
                            null,
                            new Date(t.date),
                            new Date(d),
                            total + tmp[index].quantityOfShipment,
                            tmp[index + 2]?.shipmentID,
                            null
                        ));
                    }
                    stopIndex = j;
                    break;
                }
            }
        }

        // Update filteredTransactions based on the stopIndex
        if (stopIndex !== -1) {
            filteredTransactions = filteredTransactions.slice(stopIndex + 1);
        }

        // Process the next element in tmp array recursively
        // if (tmp[index + 2]?.shipmentID !== undefined) {
        processNextTmpElement(index + 1);
        //}
    };

    // Start the recursive processing from the first element of tmp array
    processNextTmpElement(0);

    // Return the result array cogs or perform additional operations if needed
    return cogs;
};
function ExcelDateToJSDate(serial) {
    var utc_days  = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;                                        
    var date_info = new Date(utc_value * 1000);
 
    var fractional_day = serial - Math.floor(serial) + 0.0000001;
 
    var total_seconds = Math.floor(86400 * fractional_day);
 
    var seconds = total_seconds % 60;
 
    total_seconds -= seconds;
 
    var hours = Math.floor(total_seconds / (60 * 60));
    var minutes = Math.floor(total_seconds / 60) % 60;
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
 };
const findFinalDate =async (result, skuData) => {
    let d= []
    skuData.forEach(s => {
        let rs = result.filter(r => r.sku === s.sku)
        for(var i = 0; i < rs.length; i++){
            if(rs[i].date.toString().includes('.')){
                rs[i].date = ExcelDateToJSDate(rs[i].date)
            }
            if(rs[i].to_date?.toString().includes('.')){
                rs[i].to_date= ExcelDateToJSDate(rs[i].to_date)
            }
        }
        rs = rs.sort((a, b) => new Date(a.date) - new Date(b.date));
        if(rs.length != 0){
            for(var i = 0; i < rs.length-1; i++){
                rs[i].to_date =rs[i+1].date;
            }
            // let currentDate = rs[rs.length-1].date;
            // currentDate.setFullYear(currentDate.getFullYear() + 3);
            // rs[rs.length-1].to_date = new Date(currentDate.toISOString())
        }
        for(var i = rs.length-1 ; i >=0; i--){
            d.push(rs[i]);
        }
    })
    return d;
}

GenerateFile = async () => {
    const ws1 = il.Sheets["Sheet1"]
    const worksheet = workbook.Sheets['Danh sách giao dịch bổ sung']; // Replace 'Sheet1' with the actual sheet name
    const ws2 = workbook.Sheets['Ngày chuyển giao'];
    const skuData = workbook.Sheets['Giao dịch phát sinh']
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
    const arr9 = XLSX.utils.sheet_to_json(skuData);
    let skus = arr9.map((row) => {
        return new Result(
            row['date'],
            row['sku'],
            row['fnsku'],
            row['shipmentID'],
            row['nextShipmentID'],
            row['sale_quantity'],
            row['total_inventory'],
            row['data'],
            row['listShipmentID'],
            row['listQuantityOfShipment']
        )
    })

    let transations = getListTransaction(inventoryLedger)
    let futureDate = await findFutureDate(skus, transations);
    const futureDataSheets = XLSX.utils.json_to_sheet(futureDate)
    XLSX.utils.book_append_sheet(workbook, futureDataSheets, "Giao dịch phát sinh(new)");
    const existingData = XLSX.utils.sheet_to_json(worksheet);
    const existingDate = XLSX.utils.sheet_to_json(ws2)
    const mergedData = [...transations, ...existingData];
    let currentDate = await findDate(skus, mergedData)
    const mergedDate = [...currentDate[0], ...existingDate];
    const newSheet = XLSX.utils.json_to_sheet(currentDate[1]);
    //const newSheetDate = XLSX.utils.json_to_sheet(mergedDate)
    const finalDate = await findFinalDate(mergedDate, skus)
    console.log("date cuoi cug day",finalDate);
    const newSheetDate = XLSX.utils.json_to_sheet(finalDate)
    workbook.Sheets['Danh sách giao dịch bổ sung'] = newSheet; // Replace 'Sheet1' with the actual sheet name
    workbook.Sheets['Ngày chuyển giao'] = newSheetDate;
    XLSX.writeFile(workbook, 'news.xlsx');

}

GenerateFile()