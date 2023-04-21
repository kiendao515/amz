const XLSX = require('xlsx');

const workbook = XLSX.readFile('P&L.xlsx');

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

const worksheet = workbook.Sheets["Payment T3"];

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

function getSKUData(paymentList) {
  const skuData = {};
  
  paymentList.forEach(payment => {
    const { sku, type, quantity, product_sales, total } = payment;
    
    if (type === 'Order') {
      if (!skuData[sku]) {
        skuData[sku] = {
          sku,
          sale_quantity: 0,
          refund_quantity: 0,
          product_sales_order: 0,
          product_sales_refund: 0,
          refund_amount_order: 0,
          refund_amount_refund: 0
        };
      }
      
      skuData[sku].sale_quantity += quantity;
      skuData[sku].product_sales_order += product_sales;
      skuData[sku].refund_amount_order += total;
    }
    
    if (type === 'Refund') {
      if (!skuData[sku]) {
        skuData[sku] = {
          sku,
          sale_quantity: 0,
          refund_quantity: 0,
          product_sales_order: 0,
          product_sales_refund: 0,
          refund_amount_order: 0,
          refund_amount_refund: 0
        };
      }
      
      skuData[sku].refund_quantity += quantity;
      skuData[sku].product_sales_refund += product_sales;
      skuData[sku].refund_amount_refund += total;
    }
  });
  
  return skuData;
}

console.log(getSKUData(payments));





