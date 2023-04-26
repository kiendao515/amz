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

class CostOfGods {
  constructor(sku, fnsku, product_group, tags, from_date, to_date, total_amount) {
    this.sku = sku;
    this.fnsku = fnsku;
    this.product_group = product_group;
    this.tags = tags;
    this.from_date = from_date;
    this.to_date = to_date;
    this.total_amount = total_amount;
  }
}
class AdsPortfolio {
  constructor(sku, portfolio) {
    this.sku = sku;
    this.portfolio = portfolio;
  }
}
class AdsT3 {
  constructor(portfolio, spend) {
    this.portfolio = portfolio;
    this.spend = spend;
  }
}
class StorageFee {
  constructor(fnsku, estimated_monthly_storage_fee) {
    this.fnsku = fnsku;
    this.estimated_monthly_storage_fee = estimated_monthly_storage_fee
  }
}
class RemovalFee {
  constructor(sku, order_type, removal_fee) {
    this.sku = sku;
    this.order_type = order_type;
    this.removal_fee = removal_fee;
  }
}
class SurchargeFee {
  constructor(sku, short_time_range_long_term_storage_fee) {
    this.sku = sku;
    this.short_time_range_long_term_storage_fee = short_time_range_long_term_storage_fee
  }
}
class Result {
  constructor(sku, fnsku, sale_quantity, refund_quantity, product_sales, refund_amount, liquidations, gross_sales,
    product_sales_tax, shipping_credits, shipping_credit_tax, gift_wrap_credits, gift_wrap_credits_tax, regulatory_fee,
    regulatory_fee_tax, promotional_rebates, promotional_rebates_tax, marketplace_withheld_tax, referral_fees, fullfillment_fees, refund_commission, other_transaction_fee,
    other_adjustment, gross_profits, ads, storage_fee, disposal_fee, aged_inventory_surcharge, gross_profits_overall) {
    this.sku = sku;
    this.fnsku = fnsku;
    this.sale_quantity = sale_quantity;
    this.refund_quantity = refund_quantity;
    this.product_sales = product_sales;
    this.refund_amount = refund_amount;
    this.liquidations = liquidations;
    this.gross_sales = gross_sales;
    this.product_sales_tax = product_sales_tax
    this.shipping_credits = shipping_credits;
    this.shipping_credit_tax = shipping_credit_tax;
    this.gift_wrap_credits = gift_wrap_credits;
    this.gift_wrap_credits_tax = gift_wrap_credits_tax;
    this.regulatory_fee = regulatory_fee;
    this.regulatory_fee_tax = regulatory_fee_tax;
    this.promotional_rebates = promotional_rebates;
    this.promotional_rebates_tax = promotional_rebates_tax;
    this.marketplace_withheld_tax = marketplace_withheld_tax;
    this.referral_fees = referral_fees;
    this.fullfillment_fees = fullfillment_fees;
    this.refund_commission = refund_commission;
    this.other_transaction_fee = other_transaction_fee;
    this.other_adjustment = other_adjustment;
    this.gross_profits = gross_profits;
    this.ads = ads;
    this.storage_fee = storage_fee;
    this.disposal_fee = disposal_fee;
    this.aged_inventory_surcharge = aged_inventory_surcharge;
    this.gross_profits_overall = gross_profits_overall
  }
}

function getSKUData(paymentList, costOfGods, adsPortfolio, adsT3, storageFee, removalFee, surChargeFee) {
  const skuData = {};
  paymentList.forEach(payment => {
    const { sku, type, quantity, product_sales, product_sales_tax, shipping_credits, shipping_credit_tax,
      gift_wrap_credits, gift_wrap_credits_tax, regulatory_fee, regulatory_fee_tax, promotional_rebates,
      promotional_rebates_tax, marketplace_withheld_tax, selling_fee, fba_fee, other_transaction_fee, other,
      total } = payment;
    if (!skuData[sku]) {
      skuData[sku] = {
        sku,
        sale_quantity: 0,
        refund_quantity: 0,
        product_sales_order: 0,
        product_sales_refund: 0,
        liquidations: 0,
        gross_sales: 0,
        product_sales_tax: 0,
        shipping_credits: 0,
        shipping_credit_tax: 0,
        gift_wrap_credits: 0,
        gift_wrap_credits_tax: 0,
        regulatory_fee: 0,
        regulatory_fee_tax: 0,
        promotional_rebates: 0,
        promotional_rebates_tax: 0,
        marketplace_withheld_tax: 0,
        referral_fees: 0,
        fullfillment_fees: 0,
        refund_commission: 0,
        other_transaction_fee: 0,
        other: 0,
        gross_profits: 0,
        storage_fee: 0,
        disposal_fee: 0,
        aged_inventory_surcharge: 0,
        gross_profits_overall: 0
      };
    }
    if (type === 'Order') {
      skuData[sku].sale_quantity += quantity;
      skuData[sku].product_sales_order += product_sales;
      skuData[sku].gross_sales += product_sales;
      skuData[sku].referral_fees += selling_fee;
    }
    if (type === 'Refund') {
      skuData[sku].refund_quantity += quantity;
      skuData[sku].product_sales_refund += product_sales;
      skuData[sku].gross_sales += product_sales;
      skuData[sku].refund_commission += selling_fee;
    }
    if (type === 'Liquidations') {
      skuData[sku].liquidations += product_sales;
      skuData[sku].gross_sales += product_sales
    }
    skuData[sku].product_sales_tax += product_sales_tax;
    skuData[sku].shipping_credits += shipping_credits;
    skuData[sku].shipping_credit_tax += shipping_credit_tax;
    skuData[sku].gift_wrap_credits += gift_wrap_credits;
    skuData[sku].gift_wrap_credits_tax += gift_wrap_credits_tax;
    skuData[sku].regulatory_fee += regulatory_fee;
    skuData[sku].regulatory_fee_tax += regulatory_fee_tax;
    skuData[sku].promotional_rebates += promotional_rebates;
    skuData[sku].promotional_rebates_tax += promotional_rebates_tax;
    skuData[sku].marketplace_withheld_tax += marketplace_withheld_tax;
    skuData[sku].fullfillment_fees += fba_fee;
    skuData[sku].other_transaction_fee += other_transaction_fee;
    skuData[sku].other += other;
    skuData[sku].gross_profits += total;
  });

  costOfGods.forEach(skuInfo => {
    const { sku, fnsku } = skuInfo;
    if (skuData[sku]) {
      skuData[sku].fnsku = fnsku;
    }
  });
  adsPortfolio.forEach(ads => {
    const { sku, portfolio } = ads;
    if (skuData[sku]) {
      adsT3.forEach(ad => {
        if (ad.portfolio === portfolio) {
          skuData[sku].ads = parseFloat(ad.spend).toFixed(3) 
        }
      });
    }
  });

  removalFee.forEach(r => {
    const { sku, removal_fee, order_type } = r;
    if (skuData[sku]) {
      if (order_type === 'Disposal' || order_type === 'Return') {
        skuData[sku].disposal_fee += removal_fee
      }
    }
  });

  surChargeFee.forEach(s => {
    const { sku, short_time_range_long_term_storage_fee } = s;
    if (skuData[sku]) {
      skuData[sku].aged_inventory_surcharge += short_time_range_long_term_storage_fee;
    }
  });
  storageFee.forEach(s=>{
    Object.values(skuData).forEach(v => {
      if(s.fnsku === v.fnsku){
        v.storage_fee += s.estimated_monthly_storage_fee;
      }
    });
  })
  
  return Object.values(skuData).map(sku => new Result(
    sku.sku,
    sku.fnsku,
    sku.sale_quantity,
    sku.refund_quantity,
    sku.product_sales_order,
    sku.product_sales_refund,
    sku.liquidations,
    sku.gross_sales,
    sku.product_sales_tax,
    sku.shipping_credits,
    sku.shipping_credit_tax,
    sku.gift_wrap_credits,
    sku.gift_wrap_credits_tax,
    sku.regulatory_fee,
    sku.regulatory_fee_tax,
    sku.promotional_rebates,
    sku.promotional_rebates_tax,
    sku.marketplace_withheld_tax,
    sku.referral_fees,
    sku.fullfillment_fees,
    sku.refund_commission,
    sku.other_transaction_fee,
    sku.other,
    sku.gross_profits,
    sku.ads ? Number(sku.ads): null,
    sku.storage_fee,
    sku.disposal_fee,
    sku.aged_inventory_surcharge,
    sku.gross_profits + sku.aged_inventory_surcharge
  ));
}
function GenerateFile() {
  const worksheet = workbook.Sheets["Payment T3"];
  const ws2 = workbook.Sheets["Cost of Goods"]
  const ws3 = workbook.Sheets["Ads Portfolio"]
  const ws4 = workbook.Sheets["Ads T3"]
  const ws5 = workbook.Sheets["Storage Fee T3"]
  const ws6 = workbook.Sheets["Removal Fee T3"]
  const ws7 = workbook.Sheets["Surcharge Fee T3"]

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

  const arr2 = XLSX.utils.sheet_to_json(ws2);
  let costOfGods = []
  costOfGods = arr2.map((row) => {
    return new CostOfGods(
      row['Sku'],
      row['Fnsku'],
      row['Product Group'],
      row['Tags'],
      row['From Date'],
      row['To Date'],
      row['Total Amount']
    );
  });

  const arr3 = XLSX.utils.sheet_to_json(ws3);
  let adsPortfolio = arr3.map((row) => {
    return new AdsPortfolio(
      row['Sku'],
      row['Portfolio']
    );
  });

  const arr4 = XLSX.utils.sheet_to_json(ws4);
  let adsT3 = arr4.map((row) => {
    return new AdsT3(
      row['Portfolio'],
      row['Spend(USD)']
    )
  })

  const arr5 = XLSX.utils.sheet_to_json(ws5);
  let storageFee = arr5.map((row) => {
    return new StorageFee(
      row['fnsku'],
      row['estimated_monthly_storage_fee']
    )
  })

  const arr6 = XLSX.utils.sheet_to_json(ws6)
  let removalFee = arr6.map((row) => {
    return new RemovalFee(
      row['sku'],
      row['order-type'],
      row['removal-fee']
    )
  })

  const arr7 = XLSX.utils.sheet_to_json(ws7)
  let surChargeFee = arr7.map((row) => {
    return new SurchargeFee(
      row['sku'],
      row['short-time-range-long-term-storage-fee']
    )
  })

  let rs = getSKUData(payments, costOfGods, adsPortfolio, adsT3, storageFee, removalFee, surChargeFee);
  console.log(rs);
  const newWorksheet = XLSX.utils.json_to_sheet(rs);
  XLSX.utils.book_append_sheet(workbook, newWorksheet, "File hoan thanh");
  XLSX.writeFile(workbook, 'example.xlsx');
}

GenerateFile()




