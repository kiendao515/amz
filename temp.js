const XLSX = require('xlsx');

const workbook = XLSX.readFile('Payment 12.07.22 - 05.05.23 (Thành).xlsx');
const rm = XLSX.readFile("Removal Order Detail 01.01.22 - 06.05.23.xlsx")
const il = XLSX.readFile("Inventory-Ledger-18.05.22-18.05.23.xlsx")
const wb_inventory = XLSX.readFile("Inventory 06.05.xlsx")
const { getJsDateFromExcel } = require("excel-date-to-js");

class Payment {
    constructor(date, settlementID, type, orderID, group, sku, description, quantity, market_place, account_type, fullfillment, order_city, order_state, order_postal, tax_collection,
        product_sales, product_sales_tax, shipping_credits, shipping_credit_tax, gift_wrap_credits, gift_wrap_credits_tax, regulatory_fee, regulatory_fee_tax, promotional_rebates,
        promotional_rebates_tax, marketplace_withheld_tax, selling_fee, fba_fee, other_transaction_fee, other, total) {
        this.date = date;
        this.settlementID = settlementID;
        this.type = type;
        this.orderID = orderID;
        this.group = group;
        this.sku = sku;
        this.description = description;
        this.quantity = quantity;
        this.market_place = market_place;
        this.account_type = account_type;
        this.fullfillment = fullfillment;
        this.order_city = order_city;
        this.order_state = order_state;
        this.order_postal = order_postal;
        this.tax_collection = tax_collection;
        this.product_sales = product_sales;
        this.product_sales_tax = product_sales_tax;
        this.shipping_credits = shipping_credits;
        this.shipping_credit_tax = shipping_credit_tax;
        this.gift_wrap_credits = gift_wrap_credits;
        this.gift_wrap_credits_tax = gift_wrap_credits_tax;
        this.regulatory_fee = regulatory_fee;
        this.regulatory_fee_tax = regulatory_fee_tax;
        this.promotional_rebates = promotional_rebates;
        this.promotional_rebates_tax = promotional_rebates_tax;
        this.marketplace_withheld_tax = marketplace_withheld_tax;
        this.selling_fee = selling_fee;
        this.fba_fee = fba_fee;
        this.other_transaction_fee = other_transaction_fee;
        this.other = other;
        this.total = total;
    }
}


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
    constructor(date, sku, fnsku, shipmentID,nextShipmentID, sale_quantity, total_inventory, data) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.shipmentID = shipmentID;
        this.nextShipmentID = nextShipmentID;
        this.sale_quantity = sale_quantity;
        this.total_inventory = total_inventory;
        this.data = data;
    }
}
class CustomerReturn {
    constructor(date, sku, fnsku, disposition, order_type, order_status, shipped_quantity, disposed_quantity, removal_fee) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.disposition = disposition;
        this.order_type = order_type
        this.order_status = order_status;
        this.shipped_quantity = shipped_quantity;
        this.disposed_quantity = disposed_quantity;
        this.removal_fee = removal_fee;
    }
}
class Transaction {
    constructor(date, sku, fnsku, type, quantity, disposition) {
        this.date = date;
        this.sku = sku;
        this.fnsku = fnsku;
        this.type = type;
        this.quantity = quantity;
        this.disposition = disposition;
    }
}

class Cog {
    constructor(sku, fnsku,current_shipment,current_shipment_cog, date,to_date, remainder,
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

function getSKUData(paymentList, inventoryLedger, inventory) {
    const inventoryLedgerList = [];
    let skuData = {}
    inventoryLedger.forEach(element => {
        const { date, fnsku, msku, quantity, disposition, event_type, referenceID, date_time } = element;
        if (event_type === 'Receipts' && referenceID != undefined && new Date(date_time) < new Date("05/06/2023")) {
            inventoryLedgerList.push({
                date: date_time, msku, fnsku,
                shipmentID: referenceID,
                quantity: quantity
            })
            if (!skuData[msku]) {
                skuData[msku] = {
                    date: date_time, msku, fnsku,
                    shipmentID: referenceID,
                    nextShipmentID:null,
                    sale_quantity: 0,
                    total_inventory: 0
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
    });
    Object.values(skuData).forEach(sku => {
        let tmp = []
        sku.shipmentID.forEach(shipmentID => {
            const filteredInventory = inventoryLedgerList.filter(item => item.shipmentID === shipmentID && sku.msku === item.msku);
            const saleQuantity = filteredInventory.reduce((total, item) => total + item.quantity, 0);
            tmp.push(saleQuantity)
        });
        sku.sale_quantity = tmp;
    });
    Object.values(skuData).forEach(sku => {
        let count = 0;
        let total = 0;
        for (let i = 0; i < sku.shipmentID.length; i++) {
            total += sku.sale_quantity[i];
            count = count + 1;
            if (total >= sku.total_inventory) {
                sku.sale_quantity = total;
                if(sku.shipmentID[i-1]){
                    sku.nextShipmentID = sku.shipmentID[i-1];
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
        sku.sale_quantity - sku.total_inventory
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
function getListTransaction(inventoryLedger, removeArr) {
    const transactions = [];
    inventoryLedger.forEach(element => {
        const { date_time, fnsku, msku, quantity, disposition, event_type } = element;
        if (new Date(date_time) < new Date("05/06/2023") && disposition === 'SELLABLE' && (event_type === 'Shipments' || event_type === 'CustomerReturns' ||
            event_type === 'Adjustments' || event_type === 'VendorReturns')) {
            transactions.push({
                date: date_time,
                sku: msku,
                fnsku,
                type: event_type,
                quantity: quantity,
                disposition: disposition
            });
        }
    });

    removeArr.forEach(c => {
        const { date, sku, fnsku, disposition, order_type, order_status, shipped_quantity, disposed_quantity, removal_fee } = c;
        if (order_status === 'Completed' && disposition === 'Sellable' && new Date(date) < new Date("05/06/2023")) {
            if (disposed_quantity !== undefined && disposed_quantity !== 0 && order_type === 'Disposal') {
                transactions.push({
                    date: parseDate(date),
                    sku, fnsku,
                    type: order_type,
                    quantity: disposed_quantity,
                    disposition: disposition
                });
            }
            if ((order_type === 'Return' || order_type === 'Liquidations') && shipped_quantity !== 0) {
                transactions.push({
                    date, sku, fnsku,
                    type: order_type,
                    quantity: shipped_quantity,
                    disposition: disposition
                });
            }
        }
    });

    return transactions.map(t => new Transaction(
        t.date,
        t.sku,
        t.fnsku,
        t.type,
        t.quantity,
        t.disposition
    ));
}

const findDate = (skuData, transaction) => {
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
                        cogs.push(new Cog(t.sku, element.fnsku,element.shipmentID,null, t.date,new Date(d), total + element.data,
                            element.nextShipmentID,null));
                        break;
                    }
                }
            }
        }
    }

    return cogs;
};


function GenerateFile() {
    const worksheet = workbook.Sheets[0];
    const ws1 = il.Sheets["240848019495"]
    const ws2 = wb_inventory.Sheets["Inventory 06.05"]
    const ws3 = rm.Sheets["Thành"]

    // Use XLSX.utils.sheet_to_json() to convert the worksheet to a JSON array
    const jsonArray = XLSX.utils.sheet_to_json(worksheet);
    let payments = []
    payments = jsonArray.map((row) => {
        return new Payment(
            row['date/time'],
            row['settlement id'],
            row.type,
            row['order id'],
            row.group,
            row.sku,
            row.description,
            row.quantity,
            row.marketplace,
            row['account type'],
            row.fulfillment,
            row['order city'],
            row['order state'],
            row['order postal'],
            row['tax collection model'],
            row['product sales'],
            row['product sales tax'],
            row['shipping credits'],
            row['shipping credits tax'],
            row['gift wrap credits'],
            row['giftwrap credits tax'],
            row['Regulatory Fee'],
            row['Tax On Regulatory Fee'],
            row['promotional rebates'],
            row['promotional rebates tax'],
            row['marketplace withheld tax'],
            row['selling fees'],
            row['fba fees'],
            row['other transaction fees'],
            row.other,
            row.total
        );
    });

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

    const removeArr = XLSX.utils.sheet_to_json(ws3);
    let removeOrder = removeArr.map((row) => {
        return new CustomerReturn(
            row['request-date'],
            row['sku'],
            row['fnsku'],
            row['disposition'],
            row['order-type'],
            row['order-status'],
            row['shipped-quantity'],
            row['disposed-quantity'],
            row['removal-fee']
        )
    })
    let transations = getListTransaction(inventoryLedger, removeOrder)
    //console.log(transations);
    let rs = getSKUData(payments, inventoryLedger, inventory);
    let date = findDate(rs, transations);

    
    const newWorksheet = XLSX.utils.json_to_sheet(rs);
    const nw2 = XLSX.utils.json_to_sheet(transations);
    const nw3 = XLSX.utils.json_to_sheet(date)
    XLSX.utils.book_append_sheet(workbook, newWorksheet, "Giao dịch phát sinh");
    XLSX.utils.book_append_sheet(workbook, nw2, "Danh sách giao dịch bổ sung");
    XLSX.utils.book_append_sheet(workbook, nw3, "Ngày chuyển giao");
    XLSX.writeFile(workbook, 'temp.xlsx');
}

GenerateFile()




