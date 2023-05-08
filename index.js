const XLSX = require('xlsx');

const workbook = XLSX.readFile('P&L.xlsx');
const wb2 = XLSX.readFile("Removal Order Detail.xlsx")

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
  constructor(sku, order_type, removal_fee, order_status, disposition, shipped_quantity,disposed_quantity) {
    this.sku = sku;
    this.order_type = order_type;
    this.removal_fee = removal_fee;
    this.order_status = order_status;
    this.disposition = disposition;
    this.shipped_quantity = shipped_quantity;
    this.disposed_quantity = disposed_quantity;

  }
}
class SurchargeFee {
  constructor(sku, short_time_range_long_term_storage_fee) {
    this.sku = sku;
    this.short_time_range_long_term_storage_fee = short_time_range_long_term_storage_fee
  }
}
class InventoryLedger{
  constructor(date,fnsku,msku,quantity,disposition,event_type ) {
    this.date = date;
    this.fnsku= fnsku;
    this.msku = msku;
    this.quantity = quantity;
    this.disposition= disposition;
    this.event_type = event_type;
  }
}
class CustomerReturn{
  constructor(sku,quantity,detailed_disposition,status ) {
    this.sku = sku;
    this.quantity = quantity;
    this.detailed_disposition= detailed_disposition;
    this.status = status;
  }
}
class Result {
  constructor(sku, fnsku, sale_quantity, refund_quantity, product_sales, refund_amount, liquidations, gross_sales,
    product_sales_tax, shipping_credits, shipping_credit_tax, gift_wrap_credits, gift_wrap_credits_tax, regulatory_fee,
    regulatory_fee_tax, promotional_rebates, promotional_rebates_tax, marketplace_withheld_tax, referral_fees, fullfillment_fees, refund_commission, other_transaction_fee,
    other_adjustment, gross_profits, ads, storage_fee, disposal_fee, aged_inventory_surcharge, gross_profits_overall, mcf_quantity, lost_quantity_by_aw,
    adjusted_quantity_by_aw,removal_liquidations, removal_return, removal_disposal, customer_return_sellable,
    customer_return_unsellable,sellable_return_percent) {
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
    this.gross_profits_overall = gross_profits_overall;
    this.mcf_quantity = mcf_quantity;
    this.lost_quantity_by_aw = lost_quantity_by_aw;
    this.adjusted_quantity_by_aw = adjusted_quantity_by_aw;
    this.removal_liquidations = removal_liquidations;
    this.removal_return= removal_return;
    this.removal_disposal = removal_disposal;
    this.customer_return_sellable = customer_return_sellable;
    this.customer_return_unsellable = customer_return_unsellable
    this.sellable_return_percent = sellable_return_percent;
  }
}

function getSKUData(paymentList, costOfGods, adsPortfolio, adsT3, storageFee, removalFee, surChargeFee, adjustment,
  customerReturn) {
  const skuData = {};
  paymentList.forEach(payment => {
    const { sku, type, quantity, product_sales, product_sales_tax, shipping_credits, shipping_credit_tax,
      gift_wrap_credits, gift_wrap_credits_tax, regulatory_fee, regulatory_fee_tax, promotional_rebates,
      promotional_rebates_tax, marketplace_withheld_tax, selling_fee, fba_fee, other_transaction_fee, other,
      total, market_place, description } = payment;
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
        gross_profits_overall: 0,
        mcf_quantity: 0,
        lost_quantity_by_aw: 0,
        adjusted_quantity_by_aw: 0,
        removal_liquidations:0,
        removal_return:0, 
        removal_disposal:0,
        customer_return_sellable:0,
        customer_return_unsellable:0,
        sellable_return_percent:0 
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
    if(market_place?.slice(0, 3) === "sim" || market_place?.slice(1, 4) === "sim"){
      skuData[sku].mcf_quantity += quantity
    }
    if(description === 'FBA Inventory Reimbursement - General Adjustment'){
      skuData[sku].adjusted_quantity_by_aw += quantity;
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
    const { sku, removal_fee, order_type, order_status, disposition, shipped_quantity, disposed_quantity } = r;
    if (skuData[sku]) {
      if (order_type === 'Disposal' || order_type === 'Return') {
        skuData[sku].disposal_fee += removal_fee
      }
      if(order_status ==='Completed' && disposition ==='Sellable'){
        if(order_type === 'Liquidations'){
          skuData[sku].removal_liquidations += shipped_quantity
        }
        if(order_type === 'Return'){
          skuData[sku].removal_return += shipped_quantity
        }
        skuData[sku].removal_disposal += disposed_quantity
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
      else {
        costOfGods.forEach(skuInfo => {
          const { sku, fnsku } = skuInfo;
          if (s.fnsku === fnsku) {
            if (!skuData[sku]) {
              skuData[sku] = {
                sku,
                fnsku,
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
                gross_profits_overall: 0,
                mcf_quantity: 0,
                lost_quantity_by_aw: 0,
                adjusted_quantity_by_aw: 0,
                removal_liquidations:0,
                removal_return:0, 
                removal_disposal:0,
                customer_return_sellable:0,
                customer_return_unsellable:0,
                sellable_return_percent:0 
              };
            } 
          }
        });
      }
    });
  })
  adjustment.forEach(a => {
    Object.values(skuData).forEach(v => {
      if(a.msku === v.sku && a.event_type === "Adjustments" && a.quantity < 0 && new Date("03/01/2023") <=
      new Date(a.date) && new Date(a.date)<= new Date("03/31/2023")){
        v.lost_quantity_by_aw += a.quantity;
      }
      if(a.msku === v.sku && a.event_type === "SELLABLE" && a.quantity > 0 && new Date("03/01/2023") <=
      new Date(a.date) && new Date(a.date)<= new Date("03/31/2023")){
        v.adjusted_quantity_by_aw += a.quantity;
      }
    });
  });

  customerReturn.forEach(c=>{
    Object.values(skuData).forEach(v => {
      if(c.sku === v.sku){
        if(c.detailed_disposition === 'SELLABLE' && c.status === "Unit returned to inventory"){
          v.customer_return_sellable += c.quantity;
        }else{
          v.customer_return_unsellable += c.quantity;
        }
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
    sku.gross_profits + sku.aged_inventory_surcharge,
    sku.mcf_quantity,
    sku.lost_quantity_by_aw,
    sku.adjusted_quantity_by_aw,
    sku.removal_liquidations,
    sku.removal_return,
    sku.removal_disposal,
    sku.customer_return_sellable,
    sku.customer_return_unsellable,
    (sku.customer_return_sellable+ sku.customer_return_unsellable) !=0 ? 
    sku.customer_return_sellable / (sku.customer_return_sellable+ sku.customer_return_unsellable) * 100 +"%" : "0%"
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
  const ws8 = workbook.Sheets["Adjustments T3"]
  const ws9 = workbook.Sheets["Customer Return T3"]

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
      row['removal-fee'],
      row['order-type'],
      row['disposition'],
      row['shipped_quantity'],
      row['disposed-quantity']
    )
  })

  const arr7 = XLSX.utils.sheet_to_json(ws7)
  let surChargeFee = arr7.map((row) => {
    return new SurchargeFee(
      row['sku'],
      row['short-time-range-long-term-storage-fee']
    )
  })

  const arr8 = XLSX.utils.sheet_to_json(ws8)
  let adjustment = arr8.map((row) => {
    return new InventoryLedger(
      row['Date'],
      row['FNSKU'],
      row['MSKU'],
      row['Quantity'],
      row['Disposition'],
      row['Event Type']
    )
  })

  const arr9 = XLSX.utils.sheet_to_json(ws9);
  let customerReturn = arr9.map((row)=>{
    return new CustomerReturn(
      row['sku'],
      row['quantity'],
      row['detailed-disposition'],
      row['status']
    )
  })

  let rs = getSKUData(payments, costOfGods, adsPortfolio, adsT3, storageFee, removalFee, surChargeFee, adjustment,
    customerReturn);
  console.log(rs);
  // for (let i = 0; i < rs.length; i++) {
  //   let obj = rs[i];
  //   let sale_quantity = obj.sale_quantity;
  //   let refund_quantity= obj.refund_quantity;
  //   let product_sales_order= obj.product_sales;
  //   let product_sales_refund= obj.refund_amount;
  //   let liquidations= obj.liquidations;
  //   let gross_sales = obj.gross_sales;
  //   let product_sales_tax= obj.product_sales_tax;
  //   let  shipping_credits= obj.shipping_credits
  //   let shipping_credit_tax= obj.shipping_credit_tax
  //   let gift_wrap_credits= obj.gift_wrap_credits
  //   let gift_wrap_credits_tax= obj.gift_wrap_credits_tax
  //   let regulatory_fee= obj.regulatory_fee
  //   let regulatory_fee_tax= obj.regulatory_fee_tax
  //   let  promotional_rebates= obj.promotional_rebates
  //   let  promotional_rebates_tax= obj.promotional_rebates_tax
  //   let  marketplace_withheld_tax= obj.marketplace_withheld_tax
  //   let  referral_fees= obj.referral_fees
  //   let  fullfillment_fees= obj.fullfillment_fees
  //   let  refund_commission= obj.refund_commission
  //   let other_transaction_fee= obj.other_transaction_fee
  //   let other= obj.other_adjustment
  //   let  gross_profits= obj.gross_profits
  //   let ads= obj.ads
  //   let  storage_fee= obj.storage_fee
  //   let  disposal_fee= obj.disposal_fee
  //   let aged_inventory_surcharge= obj.aged_inventory_surcharge
  //   let  gross_profits_overall= obj.gross_profits_overall
  //   let j = i + 1;
  //   while (j < rs.length) {
  //     let otherObj = rs[j];
  //     if (otherObj.sku === obj.fnsku) {
  //       console.log(otherObj);
  //       liquidations += otherObj.liquidations;
  //       sale_quantity+= otherObj.sale_quantity;
  //       refund_quantity += otherObj.refund_quantity;
  //       product_sales_order += otherObj.product_sales;
  //       product_sales_refund+= otherObj.refund_amount;
  //       gross_sales += otherObj.gross_sales;
  //       product_sales_tax += otherObj.product_sales_tax
  //       shipping_credits += otherObj.shipping_credits
  //       shipping_credit_tax+= otherObj.shipping_credit_tax
  //       gift_wrap_credits+= otherObj.gift_wrap_credits
  //       gift_wrap_credits_tax+= otherObj.gift_wrap_credits_tax
  //       regulatory_fee+= otherObj.regulatory_fee
  //       regulatory_fee_tax+= otherObj.regulatory_fee_tax
  //       promotional_rebates+= otherObj.promotional_rebates
  //       promotional_rebates_tax+= otherObj.promotional_rebates_tax
  //       marketplace_withheld_tax += otherObj.marketplace_withheld_tax
  //       referral_fees+= otherObj.referral_fees
  //       fullfillment_fees+= otherObj.fullfillment_fees,
  //       refund_commission+= otherObj.refund_commission,
  //       other_transaction_fee+= otherObj.other_transaction_fee,
  //       other+= otherObj.other_adjustment
  //       gross_profits+= otherObj.gross_profits
  //       ads+= otherObj.ads;
  //       storage_fee += otherObj.storage_fee
  //       disposal_fee += otherObj.disposal_fee
  //       aged_inventory_surcharge+= otherObj.aged_inventory_surcharge
  //       gross_profits_overall += otherObj.gross_profits_overall
  //       rs.splice(j, 1);
  //     } else {
  //       j++;
  //     }
  //   }
  //   if (liquidations > obj.liquidations) {
  //     let newObj = { sku: obj.sku, fnsku: obj.fnsku,
  //       sale_quantity: sale_quantity,
  //       refund_quantity: refund_quantity,
  //       product_sales :product_sales_order,
  //       refund_amount : product_sales_refund,
  //       liquidations: liquidations,
  //       gross_sales: gross_sales,
  //       product_sales_tax,
  //       shipping_credits,
  //       shipping_credit_tax,
  //       gift_wrap_credits,
  //       gift_wrap_credits_tax,
  //       regulatory_fee,
  //       regulatory_fee_tax,
  //       promotional_rebates,
  //       promotional_rebates_tax,
  //       marketplace_withheld_tax,
  //       referral_fees,
  //       fullfillment_fees,
  //       refund_commission,
  //       other_transaction_fee,
  //       other_adjustment: other,
  //       gross_profits,
  //       ads,
  //       storage_fee,
  //       disposal_fee,
  //       aged_inventory_surcharge,
  //       gross_profits_overall:gross_profits_overall};
  //     rs.splice(i, 1, newObj);
  //   }
  //  // console.log(rs);
  // }
  const newWorksheet = XLSX.utils.json_to_sheet(rs);
  XLSX.utils.book_append_sheet(workbook, newWorksheet, "File hoan thanh");
  XLSX.writeFile(workbook, 'data.xlsx');
}

GenerateFile()




