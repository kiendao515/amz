const moment = require('moment/moment');
const XLSX = require('xlsx');

const workbook = XLSX.readFile('P&L.xlsx');
const wb_payment = XLSX.readFile('Payment-T3.xlsx')
const wb_inventory_ledger = XLSX.readFile('Inventory-Ledger-t3.xlsx')
const wb_removal = XLSX.readFile('Removal-Fee-T3.xlsx')
const wb_storage_fee = XLSX.readFile('Storage-Fee-T3.xlsx')
const wb_surcharge = XLSX.readFile('Surcharge-Fee-T3.xlsx')
const rm = XLSX.readFile("Removal Order Detail.xlsx")
const il = XLSX.readFile("Inventory Ledger.xlsx")
const { getJsDateFromExcel } = require("excel-date-to-js");
const cogs = XLSX.readFile("final.xlsx")
const axios = require('axios')
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
  constructor(sku, fnsku, total_amount, portfolio) {
    this.sku = sku;
    this.fnsku = fnsku;
    this.total_amount = total_amount;
    this.portfolio = portfolio
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
  constructor(sku, order_type, removal_fee, order_status, disposition, shipped_quantity, disposed_quantity) {
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
  constructor(sku, fnsku, short_time_range_long_term_storage_fee) {
    this.sku = sku;
    this.fnsku = fnsku;
    this.short_time_range_long_term_storage_fee = short_time_range_long_term_storage_fee
  }
}
class Transaction {
  constructor(date, sku, type, quantity, disposition, cogs) {
    this.date = date
    this.sku = sku;
    this.type = type;
    this.quantity = quantity;
    this.disposition = disposition;
    this.cogs = cogs
  }
}
class InventoryLedger {
  constructor(date, fnsku, msku, quantity, disposition, event_type) {
    this.date = date;
    this.fnsku = fnsku;
    this.msku = msku;
    this.quantity = quantity;
    this.disposition = disposition;
    this.event_type = event_type;
  }
}
class CustomerReturn {
  constructor(date, sku, fnsku, dispotition, order_type, order_status, shipped_quantity, disposed_quantity, removal_fee) {
    this.date = date;
    this.sku = sku;
    this.fnsku = fnsku;
    this.disposition = dispotition;
    this.order_type = order_type
    this.order_status = order_status;
    this.shipped_quantity = shipped_quantity;
    this.disposed_quantity = disposed_quantity;
    this.removal_fee = removal_fee;
  }
}
class Ads {
  constructor(sku, spend) {
    this.sku = sku;
    this.spend = spend;
  }
}
class Result {
  constructor(sku, fnsku, sale_quantity, refund_quantity, product_sales, refund_amount, liquidations, gross_sales,
    product_sales_tax, shipping_credits, shipping_credit_tax, gift_wrap_credits, gift_wrap_credits_tax, regulatory_fee,
    regulatory_fee_tax, promotional_rebates, promotional_rebates_tax, marketplace_withheld_tax, referral_fees, fullfillment_fees, refund_commission, other_transaction_fee,
    other_adjustment, gross_profits, subscription, ads, storage_fee, disposal_fee, vine_fee, aged_inventory_surcharge, gross_profits_overall, mcf_quantity, lost_quantity_by_aw,
    adjusted_quantity_by_aw, removal_liquidations, removal_return, removal_disposal, customer_return_sellable,
    customer_return_unsellable, sellable_return_percent, cogs_shipped, cogs_return, cogs_lost, cogs_adjusted, cogs_removal, tcogs) {
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
    this.subscription = subscription
    this.ads = ads;
    this.storage_fee = storage_fee;
    this.disposal_fee = disposal_fee;
    this.vine_fee = vine_fee
    this.aged_inventory_surcharge = aged_inventory_surcharge;
    this.gross_profits_overall = gross_profits_overall;
    this.mcf_quantity = mcf_quantity;
    this.lost_quantity_by_aw = lost_quantity_by_aw;
    this.adjusted_quantity_by_aw = adjusted_quantity_by_aw;
    this.removal_liquidations = removal_liquidations;
    this.removal_return = removal_return;
    this.removal_disposal = removal_disposal;
    this.customer_return_sellable = customer_return_sellable;
    this.customer_return_unsellable = customer_return_unsellable
    this.sellable_return_percent = sellable_return_percent;
    this.cogs_shipped = cogs_shipped;
    this.cogs_return = cogs_return;
    this.cogs_lost = cogs_lost;
    this.cogs_adjusted = cogs_adjusted;
    this.cogs_removal = cogs_removal;
    this.tcogs = tcogs;
  }
}

function getSKUData(paymentList, costOfGods, ads_brand, ads_display, ads_product, storageFee, surChargeFee, adjustment,
  customerReturn) {
  const skuData = {};
  costOfGods.forEach(skuInfo => {
    const { sku, fnsku } = skuInfo;
    if (!skuData[sku]) {
      skuData[sku] = {
        sku, fnsku,
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
        subscription: 0,
        ads: 0,
        storage_fee: 0,
        disposal_fee: 0,
        vine_fee: 0,
        aged_inventory_surcharge: 0,
        gross_profits_overall: 0,
        mcf_quantity: 0,
        lost_quantity_by_aw: 0,
        adjusted_quantity_by_aw: 0,
        removal_liquidations: 0,
        removal_return: 0,
        removal_disposal: 0,
        customer_return_sellable: 0,
        customer_return_unsellable: 0,
        sellable_return_percent: 0
      }
    }
  });
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
        subscription: 0,
        ads:0,
        storage_fee: 0,
        disposal_fee: 0,
        vine_fee: 0,
        aged_inventory_surcharge: 0,
        gross_profits_overall: 0,
        mcf_quantity: 0,
        lost_quantity_by_aw: 0,
        adjusted_quantity_by_aw: 0,
        removal_liquidations: 0,
        removal_return: 0,
        removal_disposal: 0,
        customer_return_sellable: 0,
        customer_return_unsellable: 0,
        sellable_return_percent: 0
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
    if (market_place?.slice(0, 3) === "sim" || market_place?.slice(1, 4) === "sim") {
      skuData[sku].mcf_quantity += quantity
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

  costOfGods.forEach(ads => {
    const { sku, portfolio } = ads;
    if(skuData[sku]){
      ads_brand.forEach(ad => {
        if (ad.portfolio == portfolio?.replace(/\r/g, "")) {
          skuData[sku].ads += parseFloat(ad.spend)
          console.log(skuData[sku].ads);
        }
      });
      ads_display.forEach(ads=>{
        if( sku === ads.sku){
          skuData[sku].ads += ads.spend;
        }
      })
      ads_product.forEach(ads=>{
        if( sku === ads.sku){
          skuData[sku].ads += ads.spend;
        }
      })
    }
  });
  // removalFee.forEach(r => {
  //   const { sku, removal_fee, order_type } = r;
  //   if (skuData[sku]) {
  //     if (order_type === 'Disposal' || order_type === 'Return') {
  //       skuData[sku].disposal_fee += removal_fee
  //     }
  //   }
  // });

  surChargeFee.forEach(s => {
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
            subscription: 0,
            storage_fee: 0,
            disposal_fee: 0,
            vine_fee: 0,
            aged_inventory_surcharge: 0,
            gross_profits_overall: 0,
            mcf_quantity: 0,
            lost_quantity_by_aw: 0,
            adjusted_quantity_by_aw: 0,
            removal_liquidations: 0,
            removal_return: 0,
            removal_disposal: 0,
            customer_return_sellable: 0,
            customer_return_unsellable: 0,
            sellable_return_percent: 0
          };
        }
      }
    });
    Object.values(skuData).forEach(v => {
      if (s.fnsku === v.fnsku) {
        v.aged_inventory_surcharge += s.short_time_range_long_term_storage_fee;
      }
    });
    // const { sku, short_time_range_long_term_storage_fee } = s;
    // if (skuData[sku]) {
    //   console.log("surcharge fee", skuData[sku]);
    //   skuData[sku].aged_inventory_surcharge += short_time_range_long_term_storage_fee;
    // }
  });
  storageFee.forEach(s => {
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
            subscription: 0,
            storage_fee: 0,
            disposal_fee: 0,
            aged_inventory_surcharge: 0,
            vine_fee: 0,
            gross_profits_overall: 0,
            mcf_quantity: 0,
            lost_quantity_by_aw: 0,
            adjusted_quantity_by_aw: 0,
            removal_liquidations: 0,
            removal_return: 0,
            removal_disposal: 0,
            customer_return_sellable: 0,
            customer_return_unsellable: 0,
            sellable_return_percent: 0
          };
        }
      }
    });
    Object.values(skuData).forEach(v => {
      if (s.fnsku === v.fnsku) {
        v.storage_fee += s.estimated_monthly_storage_fee;
      }
    });
  })
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
  adjustment.forEach(a => {
    var startDate = parseDate("03/01/2023")
    var endDate = parseDate("03/31/2023")
    var aDate = a.date.length > 7 ? parseDate(a.date) : parseDate(getJsDateFromExcel(a.date).getDate() + "/" + Number(getJsDateFromExcel(a.date).getMonth() + 1) + "/" + getJsDateFromExcel(a.date).getFullYear())
    costOfGods.forEach(skuInfo => {
      const { sku, fnsku } = skuInfo;
      if (a.fnsku === fnsku) {
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
            subscription: 0,
            storage_fee: 0,
            disposal_fee: 0,
            aged_inventory_surcharge: 0,
            gross_profits_overall: 0,
            mcf_quantity: 0,
            lost_quantity_by_aw: 0,
            adjusted_quantity_by_aw: 0,
            removal_liquidations: 0,
            removal_return: 0,
            removal_disposal: 0,
            customer_return_sellable: 0,
            customer_return_unsellable: 0,
            sellable_return_percent: 0
          };
        }
      }
    })
    Object.values(skuData).forEach(v => {
      if (startDate <=
        aDate && aDate <= endDate) {
        if (a.msku === v.sku && a.event_type === "Adjustments" && a.quantity < 0 && a.disposition === "SELLABLE") {
          v.lost_quantity_by_aw += a.quantity;
        }
        if (a.msku === v.sku && a.event_type === "Adjustments" && a.quantity > 0 && a.disposition === "SELLABLE") {
          v.adjusted_quantity_by_aw += a.quantity;
        }
        if (a.msku === v.sku && a.event_type === "CustomerReturns") {
          if (a.disposition === "SELLABLE") {
            v.customer_return_sellable += a.quantity;
          } else {
            v.customer_return_unsellable += a.quantity;
          }
        }
      }
    });
  });
  customerReturn.forEach(c => {
    var startDate = parseDate("03/01/2023")
    var endDate = parseDate("03/31/2023")
    var aDate = c.date.length > 7 ? parseDate(c.date.substring(0, 10)) : parseDate(getJsDateFromExcel(c.date).getDate() + "/" + Number(getJsDateFromExcel(c.date).getMonth() + 1) + "/" + getJsDateFromExcel(c.date).getFullYear())
    costOfGods.forEach(skuInfo => {
      const { sku, fnsku } = skuInfo;
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
          subscription: 0,
          storage_fee: 0,
          disposal_fee: 0,
          vine_fee: 0,
          aged_inventory_surcharge: 0,
          gross_profits_overall: 0,
          mcf_quantity: 0,
          lost_quantity_by_aw: 0,
          adjusted_quantity_by_aw: 0,
          removal_liquidations: 0,
          removal_return: 0,
          removal_disposal: 0,
          customer_return_sellable: 0,
          customer_return_unsellable: 0,
          sellable_return_percent: 0
        };
      }
    })
    const { order_status, disposition, shipped_quantity, disposed_quantity, sku, order_type, date, removal_fee } = c;
    Object.values(skuData).forEach(v => {
      if (sku === v.sku && startDate <= aDate && aDate <= endDate) {
        if (order_status === 'Completed' && disposition === 'Sellable') {
          if (order_type === 'Liquidations') {
            v.removal_liquidations += shipped_quantity
          }
          if (order_type === 'Return') {
            v.removal_return += shipped_quantity
          }
          v.removal_disposal += disposed_quantity != undefined ? disposed_quantity : 0
        }
        v.disposal_fee += removal_fee != undefined ? removal_fee : 0;
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
    -sku.subscription,
    sku.ads ? Number(-sku.ads) : null,
    -sku.storage_fee,
    -sku.disposal_fee,
    -sku.vine_fee,
    -sku.aged_inventory_surcharge,
    sku.gross_profits - (sku.subscription + Number(sku.ads ? sku.ads : 0) + sku.storage_fee + sku.disposal_fee + sku.vine_fee + sku.aged_inventory_surcharge),
    sku.mcf_quantity,
    sku.lost_quantity_by_aw,
    sku.adjusted_quantity_by_aw,
    sku.removal_liquidations,
    sku.removal_return,
    sku.removal_disposal,
    sku.customer_return_sellable,
    sku.customer_return_unsellable,
    (sku.customer_return_sellable + sku.customer_return_unsellable) != 0 ?
      sku.customer_return_sellable / (sku.customer_return_sellable + sku.customer_return_unsellable) * 100 + "%" : "0%"
  ));
}
let findCogs = async (rs, cogs_data) => {
  rs.forEach(sku => {
    let skuData = cogs_data.filter(t => t.sku === sku.sku && (new Date(t.date) >= new Date("03/01/2023")) && (new Date(t.date) < new Date("04/01/2023")) && t.disposition == "SELLABLE")
    if (skuData.length > 0) {
      sku.cogs_shipped = 0;
      sku.cogs_removal = 0;
      sku.cogs_adjusted = 0;
      sku.cogs_lost = 0;
      sku.cogs_return = 0;
      for (var i = 0; i < skuData.length; i++) {
        if (skuData[i].type == 'Shipments' && skuData[i].cogs != undefined) {
          sku.cogs_shipped += parseFloat(skuData[i].cogs)
        }
        else if (skuData[i].type == "CustomerReturns" && skuData[i].cogs != undefined) {
          sku.cogs_return += parseFloat(skuData[i].cogs)
        }
        else if (skuData[i].type == "Adjustments" && skuData[i].cogs != undefined) {
          if (skuData[i].quantity < 0) {
            sku.cogs_lost += parseFloat(skuData[i].cogs)
          } else if (skuData[i].quantity > 0) {
            sku.cogs_adjusted += parseFloat(skuData[i].cogs)
          }
        } else if (skuData[i].type == 'VendorReturns' && skuData[i].cogs != undefined) {
          sku.cogs_removal += parseFloat(skuData[i].cogs)
        }
      }
      sku.tcogs = parseFloat(sku.cogs_shipped) + parseFloat(sku.cogs_removal) +
        parseFloat(sku.cogs_return) + parseFloat(sku.cogs_lost) + parseFloat(sku.cogs_adjusted)
    }
  })
  return rs;
}
let GenerateFile = async () => {
  const worksheet = wb_payment.Sheets["Payment T3"];
  const ws3 = workbook.Sheets["Ads Portfolio"]
  const ws4 = workbook.Sheets["Ads T3"]
  const ws5 = wb_storage_fee.Sheets["Storage Fee T3"]
  const ws6 = wb_removal.Sheets["Removal Fee T3"]
  const ws7 = wb_surcharge.Sheets["Surcharge Fee T3"]
  const ws8 = rm.Sheets["Thành"]
  const ws9 = wb_inventory_ledger.Sheets["Inventory Ledger T3"]
  const cogs_sheet = cogs.Sheets["transaction list"]
  // handle ads
  const wb_ads_product = XLSX.readFile('ads_prodduct.xlsx')
  const wb_ads_display = XLSX.readFile('ads_display.xlsx')
  const wb_ads_brand = XLSX.readFile('ads_brand.xlsx')
  let ads_product = []
  ads_product = XLSX.utils.sheet_to_json(wb_ads_product.Sheets['Sponsored Product Advertised Pr']).map(row => {
    return new Ads(
      row['Advertised SKU'],
      row['Spend']
    )
  })
  let ads_display = []
  ads_display = XLSX.utils.sheet_to_json(wb_ads_display.Sheets['Sponsored Display Advertised Pr']).map(r => {
    return new Ads(
      r['Advertised SKU'],
      r['Spend']
    )
  })
  let ads_brand = []
  ads_brand = XLSX.utils.sheet_to_json(wb_ads_brand.Sheets['Sponsored Brands Campaign Repor']).map(r => {
    return new AdsT3(
      r['Portfolio name'],
      r['Spend']
    )
  })

  const cogs_fixed = await axios.get("https://dl.dropboxusercontent.com/scl/fi/rtlsuqthkunvzp2ga8epp/COGS.xlsx?rlkey=peh76yk28t1lmuwiz2f7b5j6n&dl=1", { responseType: "arraybuffer" });
  const cog_wb = XLSX.read(cogs_fixed.data);
  const arr2 = XLSX.utils.sheet_to_json(cog_wb.Sheets["Product List"])
  let costOfGods = []
  costOfGods = arr2.map((row) => {
    return new CostOfGods(
      row['SKU'],
      row['FNSKU'],
      row['Price'],
      row['Portfolio']
    );
  });


  let cogs_data = XLSX.utils.sheet_to_json(cogs_sheet).map(row => {
    return new Transaction(
      row['date'],
      row['sku'],
      row['type'],
      row['quantity'],
      row['disposition'],
      row['cogs']
    )
  })

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
      row['fnsku'],
      row['short-time-range-long-term-storage-fee']
    )
  })

  const arr8 = XLSX.utils.sheet_to_json(ws9)
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

  const arr9 = XLSX.utils.sheet_to_json(ws6);
  let customerReturn = arr9.map((row) => {
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

  let rs = getSKUData(payments, costOfGods, ads_brand, ads_display, ads_product, storageFee, surChargeFee, adjustment,
    customerReturn);
  let final = await findCogs(rs, cogs_data)
  console.log(final[final.length-1]);
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
  const newWorksheet = XLSX.utils.json_to_sheet(final);
  XLSX.utils.book_append_sheet(workbook, newWorksheet, "Thành T3");
  XLSX.writeFile(workbook, 'data.xlsx');
}

GenerateFile()




