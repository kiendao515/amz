const XLSX = require('xlsx');

const workbook = XLSX.readFile('Payment 12.07.22 - 05.05.23 (Thành).xlsx');
const il = XLSX.readFile("Inventory-Ledger-18.05.22-18.05.23.xlsx")
const wb_inventory = XLSX.readFile("Inventory-03.05.xlsx")
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
class Inventory {
    constructor(sku, fnsku, afn_fulfillable_quantity, afn_reserved_quantity) {
        this.sku = sku;
        this.fnsku = fnsku;
        this.afn_fulfillable_quantity = afn_fulfillable_quantity;
        this.afn_reserved_quantity = afn_reserved_quantity

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

class Transaction {
    constructor(date, sku, fnsku, type, quantity, disposition, shipmentID, shipment_recept) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.type = type;
        this.quantity = quantity;
        this.disposition = disposition;
        this.shipmentID = shipmentID;
        this.shipment_recept = shipment_recept;
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
class OriginalData {
    constructor(date, sku, fnsku, shipmentID, nextShipmentID, sale_quantity, total_inventory, calculated_inventory, actual_inventory,
        difference) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.shipmentID = shipmentID;
        this.nextShipmentID = nextShipmentID;
        this.sale_quantity = sale_quantity;
        this.total_inventory = total_inventory;
        this.calculated_inventory = calculated_inventory
        this.actual_inventory = actual_inventory
        this.difference = difference;
    }
}

function getSKUData(inventoryLedger, inventory) {
    const inventoryLedgerList = [];
    let skuData = {}
    inventoryLedger.forEach(element => {
        const { date, fnsku, msku, quantity, disposition, event_type, referenceID, date_time } = element;
        if (event_type === 'Receipts' && referenceID != undefined && new Date(date_time) < new Date("05/03/2023")) {
            inventoryLedgerList.push({
                date: date_time, msku, fnsku,
                shipmentID: referenceID,
                quantity: quantity
            })
            if (!skuData[msku]) {
                skuData[msku] = {
                    date: date_time, msku, fnsku,
                    shipmentID: referenceID,
                    nextShipmentID: null,
                    sale_quantity: 0,
                    total_inventory: 0,
                    listShipmentID: null,
                    listQuantityOfShipment: null
                }
            }
        }
    });
    Object.values(skuData).forEach(v => {
        inventory.forEach((i) => {
            if (v.msku === i.sku) {
                let t = i.afn_fulfillable_quantity + i.afn_reserved_quantity
                v.total_inventory += t;
            }
        })
    });

    Object.values(skuData).forEach(v => {
        let filteredData = inventoryLedgerList.filter(element => v.msku === element.msku);
        const groupedRecords = filteredData.reduce((groups, record) => {
            const referenceID = record.shipmentID;
            if (!groups[referenceID]) {
                groups[referenceID] = [];
            }
            groups[referenceID].push(record);
            return groups;
        }, {});

        // Step 2: Sort records within each group by date in ascending order
        for (const referenceID in groupedRecords) {
            groupedRecords[referenceID].sort((a, b) => {
                return new Date(a.date) - new Date(b.date);
            });
        }

        // Step 3: Combine groups into a new sorted list
        const sortedRecords = Object.values(groupedRecords).flat();
        const distinctRecords = [];
        let previousReferenceID = null;

        for (const record of sortedRecords) {
            const currentReferenceID = record.shipmentID;

            if (currentReferenceID !== previousReferenceID) {
                distinctRecords.push(record);
                previousReferenceID = currentReferenceID;
            }
        }


        distinctRecords.sort((a, b) => {
            if (a.shipmentID !== b.shipmentID) {
                // Sắp xếp các bản ghi có cùng sku và cùng referenceID theo date tăng dần
                return new Date(b.date) - new Date(a.date);
            }
        });
        const referenceIDs = distinctRecords.map(item => item.shipmentID);
        v.shipmentID = referenceIDs;
        v.listShipmentID = referenceIDs;
    });
    Object.values(skuData).forEach(sku => {
        let tmp = []
        sku.shipmentID.forEach(shipmentID => {
            const filteredInventory = inventoryLedgerList.filter(item => item.shipmentID === shipmentID && sku.msku === item.msku);
            const saleQuantity = filteredInventory.reduce((total, item) => total + item.quantity, 0);
            tmp.push(saleQuantity)
        });
        sku.sale_quantity = tmp;
        sku.listQuantityOfShipment = tmp;
    });
    Object.values(skuData).forEach(sku => {
        let count = 0;
        let total = 0;
        for (let i = 0; i < sku.shipmentID.length; i++) {
            total += sku.sale_quantity[i];
            count = count + 1;
            if (total >= sku.total_inventory) {
                sku.sale_quantity = total;
                if (sku.shipmentID[i - 1]) {
                    sku.nextShipmentID = sku.shipmentID[i - 1];
                }
                sku.shipmentID = sku.shipmentID[i]
            }
        }
    });
    return Object.values(skuData).map(sku => new Result(
        sku.date,
        sku.msku,
        sku.fnsku,
        sku.shipmentID,
        sku.nextShipmentID,
        sku.sale_quantity,
        sku.total_inventory,
        sku.sale_quantity - sku.total_inventory,
        sku.listShipmentID,
        sku.listQuantityOfShipment
    ));
}
function parseDate(input) {
    if (input.length > 7) { // input is likely in MM/DD/YYYY format
        return new Date(input);
    } else { // input is likely a serial date
        const serialDate = parseInt(input);
        if (!isNaN(serialDate)) {
            const date = new Date((serialDate - 25569) * 86400 * 1000);
            date.setUTCHours(0, 0, 0, 0);
            return date;
        }
    }
    return null; // input format not recognized
}
function getListTransaction(inventoryLedger) {
    const transactions = [];
    inventoryLedger.forEach(element => {
        const { date_time, fnsku, msku, quantity, disposition, event_type, referenceID } = element;
        if (new Date(date_time) < new Date("05/03/2023") && disposition === 'SELLABLE' && (event_type === 'Shipments' || event_type === 'CustomerReturns' ||
            event_type === 'Adjustments' || event_type === 'VendorReturns' || event_type === "Receipts")) {
            transactions.push({
                date: date_time,
                sku: msku,
                fnsku,
                type: event_type,
                quantity: quantity,
                disposition: disposition,
                shipmentID: null,
                shipment_recept: event_type === 'Receipts' ? referenceID : null
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
        t.shipment_recept
    ));
}
const findPreviousDate = async (skuData, transaction, remainder, currentDate) => {
    //transaction = transaction.filter(t => t.type !== "Receipts")
    let listTransactionOfSku = transaction.filter(t => t.sku === skuData.sku);
    let tmp = [];
    let rs = [];
    let index = skuData.listShipmentID.indexOf(skuData.shipmentID);
    skuData.listQuantityOfShipment[index + 1] += remainder;
    for (let i = index + 1; i < skuData.listShipmentID.length; i++) {
        tmp.push({
            shipmentID: skuData.listShipmentID[i],
            quantityOfShipment: skuData.listQuantityOfShipment[i],
            date: null
        });
    }
    let filteredTransactions = listTransactionOfSku;

    const startIndex = listTransactionOfSku.findIndex(t => t.shipmentID === skuData.shipmentID);
    filteredTransactions = listTransactionOfSku.slice(startIndex + 1);
    const processNextTmpElement = (index) => {
        if (index >= tmp.length) {
            return;
        }

        let total = 0;
        filteredTransactions = filteredTransactions.filter(t => t.sku === skuData.sku)
        for (let j = 0; j < filteredTransactions.length; j++) {
            const t = filteredTransactions[j];
            total += t.quantity;
            if (-tmp[index].quantityOfShipment >= total) {
                t.shipmentID = tmp[index].shipmentID;
                tmp[index].date = new Date(t.date)
                if (index == 0) {
                    var currentDateTime = new Date(currentDate);
                    var previousDateTime = new Date(currentDateTime.getTime() - (24 * 60 * 60 * 1000));
                    var previousDateTimeString = previousDateTime.toISOString();
                    rs.push(new Cog(t.sku, t.fnsku, tmp[index].shipmentID, null, new Date(t.date), new Date(previousDateTimeString), total + tmp[index].quantityOfShipment, skuData.shipmentID, null));
                } else {
                    var currentDateTime = new Date(tmp[index - 1].date);
                    var previousDateTime = new Date(currentDateTime.getTime() - (24 * 60 * 60 * 1000));
                    var previousDateTimeString = previousDateTime.toISOString();
                    rs.push(new Cog(t.sku, t.fnsku, tmp[index].shipmentID, null, new Date(t.date), new Date(previousDateTimeString), total + tmp[index].quantityOfShipment, tmp[index - 1]?.shipmentID, null));
                }
                break;
            }
            if(j == (filteredTransactions.length -1) &&  -tmp[index].quantityOfShipment < total){
                var currentDateTime = new Date(tmp[index - 1]?.date);
                t.shipmentID = tmp[index].shipmentID;
                rs.push(new Cog(t.sku, t.fnsku, tmp[index].shipmentID, null, new Date(t.date), 
                currentDateTime, total + tmp[index].quantityOfShipment, tmp[index - 1]?.shipmentID, null));
            }
        }

        // Update filteredTransactions based on the updated shipmentID
        let s = filteredTransactions.findIndex(t => t.shipmentID === tmp[index].shipmentID);
        filteredTransactions = filteredTransactions.slice(s + 1);

        // Process the next element in tmp array recursively
        processNextTmpElement(index + 1);
    };

    // Start the recursive processing from the first element of tmp array
    processNextTmpElement(0);

    // Return the result array rs or perform additional operations if needed
    return rs;
};

const findDate = async (skuData, transaction) => {
    transaction = transaction.filter(t => t.type !== "Receipts")
    const cogs = [];
    for (let i = 0; i < skuData.length; i++) {
        const element = skuData[i];
        let total = 0;
        if (element.data > 0) {
            for (let j = 0; j < transaction.length; j++) {
                const t = transaction[j];
                if (t.sku === element.sku) {
                    total += t.quantity;
                    if (-element.data >= total) {
                        let d = new Date(t.date)
                        d.setFullYear(d.getFullYear() + 3)
                        t.shipmentID = element.shipmentID
                        cogs.push(new Cog(t.sku, element.fnsku, element.shipmentID, null, new Date(t.date), new Date(d), total + element.data,
                            element.nextShipmentID, null));
                        let rs = await findPreviousDate(element, transaction, total + element.data, t.date);
                        cogs.push(...rs)
                        break;
                    }
                }
            }
        } else {
            let tmp = transaction.filter(t => t.sku === element.sku)
            let rs = await findPreviousDate(element, transaction, 0, new Date(tmp[0]?.date));
            cogs.push(...rs)
        }
    }
    return [cogs, transaction];
};

const handleWriteAllShipment = async (rs) => {
    rs.forEach(s => {
        s.listShipmentID = s.listShipmentID.join(',');
        s.listQuantityOfShipment = s.listQuantityOfShipment.map(qty => {
            return Number.isNaN(qty) ? '' : qty;
        }).join(',');
    })
    return rs;
}
const findOriginalData = async (data, inventory) => {
    let result = data.map(obj => Object.assign({}, obj));
    result.forEach(d => {
        d.date = undefined
        d.listShipmentID = 0
        d.listQuantityOfShipment = 0
        let tmp = inventory.filter(i => i.sku === d.sku)
        for (var i = 0; i < tmp.length; i++) {
            d.listShipmentID += (tmp[i].afn_fulfillable_quantity + tmp[i].afn_reserved_quantity)
        }
        d.listQuantityOfShipment = d.total_inventory - d.listShipmentID;
    })
    result[0].date = "03/05/2023";
    return result.map(obj => new OriginalData(
        obj.date,
        obj.sku,
        obj.fnsku,
        obj.shipmentID,
        obj.nextShipmentID,
        obj.sale_quantity,
        obj.data,
        obj.total_inventory,
        obj.listShipmentID,
        obj.listQuantityOfShipment
    ));
}

GenerateFile = async () => {
    const ws1 = il.Sheets["240848019495"]
    const ws2 = wb_inventory.Sheets["Inventory 03.05"]
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

    const arr3 = XLSX.utils.sheet_to_json(ws2);
    let inventory = arr3.map((row) => {
        return new Inventory(
            row['sku'],
            row['fnsku'],
            row['afn-fulfillable-quantity'],
            row['afn-reserved-quantity']
        )
    })
    let transations = getListTransaction(inventoryLedger)
    let transaction_receipts = transations.filter(t => t.type === "Receipts")
    let rs = getSKUData(inventoryLedger, inventory);
    let result = await findDate(rs, transations)
    let date = result[0];
    let transaction_shipment = [...result[1], ...transaction_receipts];

    let newRusult = await handleWriteAllShipment(rs);
    let orignalData = await findOriginalData(newRusult, inventory);
    console.log(newRusult);
    const newWorksheet = XLSX.utils.json_to_sheet(newRusult);
    const nw2 = XLSX.utils.json_to_sheet(transaction_shipment);
    const nw3 = XLSX.utils.json_to_sheet(date)
    const nw4 = XLSX.utils.json_to_sheet(orignalData)
    XLSX.utils.book_append_sheet(workbook, newWorksheet, "Giao dịch phát sinh");
    XLSX.utils.book_append_sheet(workbook, nw2, "Danh sách giao dịch bổ sung");
    XLSX.utils.book_append_sheet(workbook, nw3, "Ngày chuyển giao");
    XLSX.utils.book_append_sheet(workbook, nw4, "original inventory statistics")
    XLSX.utils.sheet_add_aoa(nw4, [["date", "sku", "fnsku", "shipment id", "next shipment id", "receipts (total received quantity calculate from current shipment)", "transaction (total transaction calculate from current shipment)",
        "calculated inventory (inventory quantity on 08/06 according to calculation)", "actual inventory (inventory quantity on 08/06 according to actual)",
        "difference"]], { origin: "A1" });
    XLSX.writeFile(workbook, 'output.xlsx');
}

GenerateFile()




