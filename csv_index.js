const XLSX = require('xlsx');
const fs = require('fs');
const csv = require('csv-parser');
const { getJsDateFromExcel } = require("excel-date-to-js");
const cogs = XLSX.readFile("inventory.xlsx")
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
  constructor(sku, fnsku, total_amount, portfolio, group) {
    this.sku = sku;
    this.fnsku = fnsku;
    this.total_amount = total_amount;
    this.portfolio = portfolio
    this.group = group;
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
  constructor(date, last_updated_date, sku, fnsku, dispotition, order_type, order_status, shipped_quantity, disposed_quantity, removal_fee) {
    this.date = date;
    this.last_updated_date = last_updated_date
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
  constructor(group, sku, fnsku, sale_quantity, refund_quantity, product_sales, refund_amount, liquidations, gross_sales,
    product_sales_tax, shipping_credits, shipping_credit_tax, gift_wrap_credits, gift_wrap_credits_tax, regulatory_fee,
    regulatory_fee_tax, promotional_rebates, promotional_rebates_tax, marketplace_withheld_tax, referral_fees, fullfillment_fees, refund_commission, other_transaction_fee,
    other_adjustment, gross_profits, ads, storage_fee, disposal_fee, vine_fee, aged_inventory_surcharge, gross_profits_overall, mcf_quantity, lost_quantity_by_aw,
    adjusted_quantity_by_aw, removal_liquidations, removal_return, removal_disposal, customer_return_sellable,
    customer_return_unsellable, sellable_return_percent, missing_received_quantity, shipmentID,
    quantity_found, reimbursed_quantity, not_reimbursed_quantity, reimbursement_for_missing_quantity, cogs_for_missing_quantity, reconcile_cogs, cogs_shipped, cogs_return, cogs_lost, cogs_adjusted, cogs_removal, tcogs, business_expense, net_profit) {
    this.group = group
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
    this.missing_received_quantity = missing_received_quantity
    this.shipmentID = shipmentID
    this.quantity_found = quantity_found;
    this.reimbursed_quantity = reimbursed_quantity;
    this.not_reimbursed_quantity = not_reimbursed_quantity;
    this.reimbursement_for_missing_quantity = reimbursement_for_missing_quantity
    this.cogs_for_missing_quantity = cogs_for_missing_quantity;
    this.reconcile_cogs = reconcile_cogs;
    this.cogs_shipped = cogs_shipped;
    this.cogs_return = cogs_return;
    this.cogs_lost = cogs_lost;
    this.cogs_adjusted = cogs_adjusted;
    this.cogs_removal = cogs_removal;
    this.tcogs = tcogs;
    this.business_expense = business_expense;
    this.net_profit = net_profit
  }
}

function getSKUData(paymentList, costOfGods, ads_brand, ads_display, ads_product, storageFee, surChargeFee, adjustment,
  customerReturn) {
  const skuData = {};
  costOfGods.forEach(skuInfo => {
    const { sku, fnsku, group, portfolio } = skuInfo;
    if (!skuData[sku]) {
      skuData[sku] = {
        group: group,
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
        sellable_return_percent: 0,
        reimbursed_quantity: 0,
        reimbursement_for_missing_quantity: 0
      }
      ads_brand?.forEach(ad => {
        if (ad.portfolio == portfolio?.replace(/\r/g, "")) {
          skuData[sku].ads += parseFloat(ad.spend)
        }
      });
      ads_display?.forEach(ads => {
        if (sku === ads.sku) {
          skuData[sku].ads += ads.spend;
        }
      })
      ads_product?.forEach(ads => {
        if (sku === ads.sku) {
          skuData[sku].ads += ads.spend;
        }
      })
    }
  });
  // sửa date ở ddây
  paymentList.forEach(payment => {
    if (new Date("09/01/2023") <= new Date(payment.date.includes("PDT") ? payment.date.replace(" PDT", "") : payment.date.replace(" PST", "")) &&
      new Date(payment.date.includes("PDT") ? payment.date.replace(" PDT", "") : payment.date.replace(" PST", "")) < new Date("10/01/2023")) {
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
        };
      }
      if (type === 'Order') {
        skuData[sku].sale_quantity += parseInt(quantity);
        skuData[sku].product_sales_order += parseFloat(product_sales);
        skuData[sku].gross_sales += parseFloat(product_sales);
        skuData[sku].referral_fees += parseFloat(selling_fee);
      }
      if (type === 'Refund') {
        skuData[sku].refund_quantity += parseInt(quantity);
        skuData[sku].product_sales_refund += parseFloat(product_sales);
        skuData[sku].gross_sales += parseFloat(product_sales);
        skuData[sku].refund_commission += parseFloat(selling_fee);
      }
      if (type === 'Liquidations' || type === "Liquidations Adjustments") {
        skuData[sku].liquidations += parseFloat(product_sales);
        skuData[sku].gross_sales += parseFloat(product_sales)
      }
      if (market_place?.slice(0, 3) === "sim" || market_place?.slice(1, 4) === "sim") {
        skuData[sku].mcf_quantity += parseInt(quantity)
      }
      skuData[sku].product_sales_tax += parseFloat(product_sales_tax);
      skuData[sku].shipping_credits += parseFloat(shipping_credits);
      skuData[sku].shipping_credit_tax += parseFloat(shipping_credit_tax);
      skuData[sku].gift_wrap_credits += parseFloat(gift_wrap_credits);
      skuData[sku].gift_wrap_credits_tax += parseFloat(gift_wrap_credits_tax);
      skuData[sku].regulatory_fee += parseFloat(regulatory_fee);
      skuData[sku].regulatory_fee_tax += parseFloat(regulatory_fee_tax);
      skuData[sku].promotional_rebates += parseFloat(promotional_rebates);
      skuData[sku].promotional_rebates_tax += parseFloat(promotional_rebates_tax);
      skuData[sku].marketplace_withheld_tax += parseFloat(marketplace_withheld_tax);
      skuData[sku].fullfillment_fees += parseFloat(fba_fee);
      skuData[sku].other_transaction_fee += parseFloat(other_transaction_fee);
      skuData[sku].other += parseFloat(other);
      skuData[sku].gross_profits += parseFloat(total);
    }
  });

  // costOfGods.forEach(ads => {
  //   const { sku, portfolio } = ads;
  //   if (skuData[sku]) {
  //     ads_brand?.forEach(ad => {
  //       if (ad.portfolio == portfolio?.replace(/\r/g, "")) {
  //         skuData[sku].ads += parseFloat(ad.spend)
  //       }
  //     });
  //     ads_display?.forEach(ads => {
  //       if (sku === ads.sku) {
  //         skuData[sku].ads += ads.spend;
  //       }
  //     })
  //     ads_product?.forEach(ads => {
  //       if (sku === ads.sku) {
  //         skuData[sku].ads += ads.spend;
  //       }
  //     })
  //   }
  // });

  surChargeFee?.forEach(s => {
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
        v.aged_inventory_surcharge += parseFloat(s.short_time_range_long_term_storage_fee);
      }
    });
    // const { sku, short_time_range_long_term_storage_fee } = s;
    // if (skuData[sku]) {
    //   console.log("surcharge fee", skuData[sku]);
    //   skuData[sku].aged_inventory_surcharge += short_time_range_long_term_storage_fee;
    // }
  });
  storageFee?.forEach(s => {
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
        v.storage_fee += parseFloat(s.estimated_monthly_storage_fee.replace(/,/g, ''));
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
    var startDate = parseDate("09/01/2023")
    var endDate = parseDate("10/01/2023")
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
        aDate && aDate < endDate) {
        if (a.msku === v.sku && a.event_type === "Adjustments" && a.quantity < 0 && a.disposition === "SELLABLE") {
          v.lost_quantity_by_aw += parseInt(a.quantity);
        }
        if (a.msku === v.sku && a.event_type === "Adjustments" && a.quantity > 0 && a.disposition === "SELLABLE") {
          v.adjusted_quantity_by_aw += parseInt(a.quantity);
        }
        if (a.msku === v.sku && a.event_type === "CustomerReturns") {
          if (a.disposition === "SELLABLE") {
            v.customer_return_sellable += parseInt(a.quantity);
          } else {
            v.customer_return_unsellable += parseInt(a.quantity);
          }
        }
      }
    });
  });
  // sửa date
  customerReturn.forEach(c => {
    var startDate = parseDate("09/01/2023")
    var endDate = parseDate("10/01/2023")
    var aDate = c.last_updated_date.length > 7 ? parseDate(c.last_updated_date.substring(0, 10)) : parseDate(getJsDateFromExcel(c.last_updated_date).getDate() + "/" + Number(getJsDateFromExcel(c.last_updated_date).getMonth() + 1) + "/" + getJsDateFromExcel(c.last_updated_date).getFullYear())
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
    const { order_status, disposition, shipped_quantity, disposed_quantity, sku, fnsku,order_type, date, removal_fee } = c;
    Object.values(skuData).forEach(v => {
      if (fnsku === v.fnsku && startDate <= aDate && aDate < endDate) {
        if (order_status === 'Completed' && disposition === 'Sellable') {
          if (order_type === 'Liquidations') {
            v.removal_liquidations += parseInt(shipped_quantity)
          }
          if (order_type === 'Return') {
            v.removal_return += parseInt(shipped_quantity)
          }
          v.removal_disposal += (disposed_quantity != undefined && disposed_quantity != "") ? parseInt(disposed_quantity) : 0
        }
        console.log(removal_fee);
        v.disposal_fee += (removal_fee != undefined && removal_fee != "") ? parseFloat(removal_fee) : 0;
      }
    });
  })


  return Object.values(skuData).map(sku => new Result(
    sku.group,
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
    sku.ads ? Number(-sku.ads) : null,
    -sku.storage_fee,
    -sku.disposal_fee,
    -sku.vine_fee,
    -sku.aged_inventory_surcharge,
    parseFloat(sku.gross_profits) - (Number(sku.ads ? sku.ads : 0) + parseFloat(sku.storage_fee) + parseFloat(sku.disposal_fee) + parseFloat(sku.vine_fee) + parseFloat(sku.aged_inventory_surcharge)),
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
// cần sửa date ở đây nhé
let findCogs = async (rs, cogs_data) => {
  rs.forEach(sku => {
    let skuData = cogs_data.filter(t => t.sku === sku.sku && (new Date(t.date) >= new Date("09/01/2023")) && (new Date(t.date) < new Date("10/01/2023")) && t.disposition == "SELLABLE")
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
    } else {
      sku.cogs_shipped = 0;
      sku.cogs_removal = 0;
      sku.cogs_adjusted = 0;
      sku.cogs_lost = 0;
      sku.cogs_return = 0;
      sku.tcogs = 0
    }
  })
  return rs;
}
// cần sửa date ở đây
let splitEmptySku = async (data, payments) => {
  const nullSkuElements = data.filter((item) => (item.sku == undefined || item.sku == ''));
  const nonNullSkuElements = data.filter((item) => (item.sku !== undefined && item.sku != ''));
  let newData = [...nonNullSkuElements];
  let filterPayments = payments.filter(p => (p.sku == '' && (new Date(p.date) >= new Date("2023-09-01T00:00:00-07:00"))
    && (new Date(p.date) < new Date("2023-10-01T00:00:00-07:00"))));
  let count = 0;
  let sub_fee = 0, sub_fee_adjustment = 0, early_reviewer_program_fee = 0, vine_Enrollment_Fee = 0, coupom_fee = 0, lighting_deal_fee = 0,
    commisstion_adjustment = 0, fee_adjustment = 0, fba_inventory = 0, other = 0, fba_amazon = 0, fba_international = 0, advertising_payment = 0, previous_storage_fee = 0,
    surcharge_fee = 0, disposal_fee = 0, return_fee = 0, debt = 0, transfer = 0, retrocharge = 0, retrocharge_reversal = 0
  filterPayments.forEach(p => {
    if (p.description == "Subscription") {
      sub_fee += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "Subscription Fee Adjustment") {
      sub_fee_adjustment += parseFloat(p.total.replace(/,/g, ''));
    }
    else if (p.description?.includes("Early Reviewer Program fee")) {
      early_reviewer_program_fee += parseFloat(p.total.replace(/,/g, ''));
    }
    else if (p.description == "Vine Enrollment Fee") {
      vine_Enrollment_Fee += parseFloat(p.total.replace(/,/g, ''));
    }
    else if (p.description?.includes("Coupon Redemption Fee")) {
      coupom_fee += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.type == "Deal Fee") {
      lighting_deal_fee += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "Commission Adjustment") {
      commisstion_adjustment += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "Fee Adjustment - Weight and Dimension Change") {
      fee_adjustment += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "FBA Inventory Reimbursement - Fee Correction") {
      fba_inventory += parseFloat(p.total.replace(/,/g, ''))
    } else if (p.type == "Order_Retrocharge") {
      retrocharge += parseFloat(p.marketplace_withheld_tax.replace(/,/g, ''))
    } else if (p.type == "Refund_Retrocharge") {
      retrocharge_reversal += parseFloat(p.marketplace_withheld_tax.replace(/,/g, ''))
    }
    // if(p.description == "Other" || p.type == "Order_Retrocharge"){
    //   other+= p.total
    // }
    else if (p.description == "FBA Amazon-Partnered Carrier Shipment Fee") {
      fba_amazon += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "FBA international shipping charge") {
      fba_international += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "Cost of Advertising") {
      advertising_payment += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "FBA Inventory Storage Fee") {
      previous_storage_fee += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "FBA Long-Term Storage Fee") {
      surcharge_fee += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "FBA Removal Order: Disposal Fee") {
      disposal_fee += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.description == "FBA Removal Order: Return Fee") {
      return_fee += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.type == "Debt") {
      debt += parseFloat(p.total.replace(/,/g, ''))
    }
    else if (p.type == "Transfer") {
      transfer += parseFloat(p.total.replace(/,/g, ''))
    } else {
      other += parseFloat(p.total.replace(/,/g, ''))
    }
  })
  console.log("jee", count);
  newData.push({})
  newData.push({ group: "Subscription Fee", other_adjustment: sub_fee, gross_profits: sub_fee, gross_profits_overall: sub_fee, net_profit: sub_fee })
  newData.push({ group: "Subscription Fee Adjustment", other_adjustment: sub_fee_adjustment, gross_profits: sub_fee_adjustment, gross_profits_overall: sub_fee_adjustment, net_profit: sub_fee_adjustment })
  newData.push({ group: "Debt", other_adjustment: debt, gross_profits: debt, gross_profits_overall: debt, net_profit: debt })
  newData.push({ group: "Early Reviewer Program Fee", other_transaction_fee: early_reviewer_program_fee, gross_profits: early_reviewer_program_fee, gross_profits_overall: early_reviewer_program_fee, net_profit: early_reviewer_program_fee })
  newData.push({ group: "Vine Enrollment Fee", other_adjustment: vine_Enrollment_Fee, gross_profits: vine_Enrollment_Fee, gross_profits_overall: vine_Enrollment_Fee, net_profit: vine_Enrollment_Fee })
  newData.push({ group: "Coupon Redemption Fee", other_transaction_fee: coupom_fee, gross_profits: coupom_fee, gross_profits_overall: coupom_fee, net_profit: coupom_fee })
  newData.push({ group: "Lightning Deal Fee", other_transaction_fee: lighting_deal_fee, gross_profits: lighting_deal_fee, gross_profits_overall: lighting_deal_fee, net_profit: lighting_deal_fee })
  newData.push({ group: "Commission Adjustment", other_adjustment: commisstion_adjustment, gross_profits: commisstion_adjustment, gross_profits_overall: commisstion_adjustment, net_profit: commisstion_adjustment })
  newData.push({ group: "Fee Adjustment - Weight and Dimension Change", fullfillment_fees: fee_adjustment, gross_profits: fee_adjustment, gross_profits_overall: fee_adjustment, net_profit: fee_adjustment })
  newData.push({ group: "FBA Inventory Reimbursement - Fee Correction", other_adjustment: fba_inventory, gross_profits: fba_inventory, gross_profits_overall: fba_inventory, net_profit: fba_inventory })
  newData.push({ group: "Retrocharge", marketplace_withheld_tax: retrocharge, gross_profits: 0, gross_profits_overall: 0, net_profit: 0 })
  newData.push({ group: "Retrocharge Reversal", marketplace_withheld_tax: retrocharge_reversal, gross_profits: 0, gross_profits_overall: 0, net_profit: 0 })
  newData.push({ group: "Other", other_adjustment: other, gross_profits: other, gross_profits_overall: other, net_profit: other })
  newData.push({ group: "FBA Amazon-Partnered Carrier Shipment Fee", other_adjustment: fba_amazon, gross_profits: fba_amazon, gross_profits_overall: fba_amazon, net_profit: fba_amazon })
  newData.push({ group: "FBA International Shipping Charge", other_adjustment: fba_international, gross_profits: fba_international, gross_profits_overall: fba_international, net_profit: fba_international })
  newData.push({ group: "Advertising Payment", other_transaction_fee: advertising_payment, gross_profits: advertising_payment, gross_profits_overall: advertising_payment, net_profit: advertising_payment })
  newData.push({ group: "Previous Storage Fee", other_adjustment: previous_storage_fee, gross_profits: previous_storage_fee, gross_profits_overall: previous_storage_fee, net_profit: previous_storage_fee })
  newData.push({ group: "Surcharge Fee", other_adjustment: surcharge_fee, gross_profits: surcharge_fee, gross_profits_overall: surcharge_fee, net_profit: surcharge_fee })
  newData.push({ group: "Disposal Fee", other_adjustment: disposal_fee, gross_profits: disposal_fee, gross_profits_overall: disposal_fee, net_profit: disposal_fee })
  newData.push({ group: "Return Fee", other_adjustment: return_fee, gross_profits: return_fee, gross_profits_overall: return_fee, net_profit: return_fee })
  newData.push({ group: "Transfer Payment", other_adjustment: transfer, gross_profits: transfer, gross_profits_overall: transfer, net_profit: transfer })
  newData.push({})
  // tính statistic cho dòng total
  let total_sales_quantity = newData.reduce((acc, obj) => acc + (obj.sale_quantity || 0), 0);
  let refund_quantity_total = newData.reduce((acc, obj) => acc + (obj.refund_quantity || 0), 0);
  let total_product_sales = newData.reduce((acc, obj) => acc + (obj.product_sales || 0), 0);
  let total_refund_amount = newData.reduce((acc, obj) => acc + (obj.refund_amount || 0), 0);
  let liquidations = newData.reduce((acc, obj) => acc + (obj.liquidations || 0), 0);
  let gross_sales = newData.reduce((acc, obj) => acc + (obj.gross_sales || 0), 0);
  let product_sales_tax = newData.reduce((acc, obj) => acc + (obj.product_sales_tax || 0), 0);
  let shipping_credits = newData.reduce((acc, obj) => acc + (obj.shipping_credits || 0), 0);
  let shipping_credit_tax = newData.reduce((acc, obj) => acc + (obj.shipping_credit_tax || 0), 0);
  let gift_wrap_credits = newData.reduce((acc, obj) => acc + (obj.gift_wrap_credits || 0), 0);
  let gift_wrap_credits_tax = newData.reduce((acc, obj) => acc + (obj.gift_wrap_credits_tax || 0), 0);
  let regulatory_fee = newData.reduce((acc, obj) => acc + (obj.regulatory_fee || 0), 0);
  let regulatory_fee_tax = newData.reduce((acc, obj) => acc + (obj.regulatory_fee_tax || 0), 0);
  let promotional_rebates = newData.reduce((acc, obj) => acc + (obj.promotional_rebates || 0), 0);
  let promotional_rebates_tax = newData.reduce((acc, obj) => acc + (obj.promotional_rebates_tax || 0), 0);
  let marketplace_withheld_tax = newData.reduce((acc, obj) => acc + (obj.marketplace_withheld_tax || 0), 0);
  let referral_fees = newData.reduce((acc, obj) => acc + (obj.referral_fees || 0), 0);
  let fullfillment_fees = newData.reduce((acc, obj) => acc + (obj.fullfillment_fees || 0), 0);
  let refund_commission = newData.reduce((acc, obj) => acc + (obj.refund_commission || 0), 0);
  let other_transaction_fee = newData.reduce((acc, obj) => acc + (obj.other_transaction_fee || 0), 0);
  let other_adjustment = newData.reduce((acc, obj) => acc + (obj.other_adjustment || 0), 0);
  let gross_profits = newData.reduce((acc, obj) => acc + (obj.gross_profits || 0), 0);
  let ads = newData.reduce((acc, obj) => acc + (obj.ads || 0), 0);
  let storage_fee = newData.reduce((acc, obj) => acc + (obj.storage_fee || 0), 0);
  let disposal_fee_total = newData.reduce((acc, obj) => acc + (obj.disposal_fee || 0), 0);
  let vine_fee = newData.reduce((acc, obj) => acc + (obj.vine_fee || 0), 0);
  let aged_inventory_surcharge = newData.reduce((acc, obj) => acc + (obj.aged_inventory_surcharge || 0), 0);
  let gross_profits_overall = newData.reduce((acc, obj) => acc + (obj.gross_profits_overall || 0), 0);
  let mcf_quantity = newData.reduce((acc, obj) => acc + (obj.mcf_quantity || 0), 0);
  let lost_quantity_by_aw = newData.reduce((acc, obj) => acc + (obj.lost_quantity_by_aw || 0), 0);
  let adjusted_quantity_by_aw = newData.reduce((acc, obj) => acc + (obj.adjusted_quantity_by_aw || 0), 0);
  let removal_liquidations = newData.reduce((acc, obj) => acc + (obj.removal_liquidations || 0), 0);
  let removal_return = newData.reduce((acc, obj) => acc + (obj.removal_return || 0), 0);
  let removal_disposal = newData.reduce((acc, obj) => acc + (obj.removal_disposal || 0), 0);
  let customer_return_sellable = newData.reduce((acc, obj) => acc + (obj.customer_return_sellable || 0), 0);
  let customer_return_unsellable = newData.reduce((acc, obj) => acc + (obj.customer_return_unsellable || 0), 0);
  let sellable_return_percent = newData.reduce((acc, obj) => acc + (parseFloat(obj.sellable_return_percent) || 0), 0);
  let cogs_shipped = newData.reduce((acc, obj) => acc + (obj.cogs_shipped || 0), 0);
  let cogs_return = newData.reduce((acc, obj) => acc + (obj.cogs_return || 0), 0);
  let cogs_lost = newData.reduce((acc, obj) => acc + (obj.cogs_lost || 0), 0);
  let cogs_adjusted = newData.reduce((acc, obj) => acc + (obj.cogs_adjusted || 0), 0);
  let cogs_removal = newData.reduce((acc, obj) => acc + (obj.cogs_removal || 0), 0);
  let tcogs = newData.reduce((acc, obj) => acc + (obj.tcogs || 0), 0);
  let missing_received_quantity = newData.reduce((acc, obj) => acc + (obj.missing_received_quantity || 0), 0);
  //  shipmentID,
  let quantity_found = newData.reduce((acc, obj) => acc + (obj.quantity_found || 0), 0);
  let reimbursed_quantity = newData.reduce((acc, obj) => acc + (obj.reimbursed_quantity || 0), 0);
  let not_reimbursed_quantity = newData.reduce((acc, obj) => acc + (obj.not_reimbursed_quantity || 0), 0);
  let reimbursement_for_missing_quantity = newData.reduce((acc, obj) => acc + (obj.reimbursement_for_missing_quantity || 0), 0);
  let cogs_for_missing_quantity = newData.reduce((acc, obj) => acc + (obj.cogs_for_missing_quantity || 0), 0);
  let reconcile_cogs = newData.reduce((acc, obj) => acc + (obj.reconcile_cogs || 0), 0);
  let business_expense = newData.reduce((acc, obj) => acc + (obj.business_expense || 0), 0);
  let net_profit = newData.reduce((acc, obj) => acc + (obj.net_profit || 0), 0);
  newData.push({
    group: "Total", sale_quantity: total_sales_quantity, refund_quantity: refund_quantity_total, product_sales: total_product_sales, refund_amount: total_refund_amount,
    liquidations: liquidations, gross_sales: gross_sales, product_sales_tax: product_sales_tax, shipping_credits: shipping_credits, shipping_credit_tax: shipping_credit_tax, gift_wrap_credits: gift_wrap_credits,
    gift_wrap_credits_tax: gift_wrap_credits_tax, regulatory_fee: regulatory_fee, regulatory_fee_tax: regulatory_fee_tax, promotional_rebates: promotional_rebates, promotional_rebates_tax: promotional_rebates_tax,
    marketplace_withheld_tax: marketplace_withheld_tax, referral_fees: referral_fees, fullfillment_fees: fullfillment_fees, refund_commission: refund_commission, other_transaction_fee: other_transaction_fee,
    other_adjustment: other_adjustment, gross_profits: gross_profits, ads: ads, storage_fee: storage_fee, disposal_fee: disposal_fee_total, vine_fee: vine_fee, aged_inventory_surcharge: aged_inventory_surcharge,
    gross_profits_overall: gross_profits_overall, mcf_quantity: mcf_quantity, lost_quantity_by_aw: lost_quantity_by_aw, adjusted_quantity_by_aw: adjusted_quantity_by_aw, removal_liquidations: removal_liquidations,
    removal_return: removal_return, removal_disposal: removal_disposal, customer_return_sellable: customer_return_sellable, customer_return_unsellable: customer_return_unsellable,
    sellable_return_percent: sellable_return_percent, missing_received_quantity: missing_received_quantity, shipmentID: undefined, quantity_found: quantity_found, reimbursed_quantity: reimbursed_quantity,
    not_reimbursed_quantity: not_reimbursed_quantity, reimbursement_for_missing_quantity: reimbursement_for_missing_quantity, cogs_for_missing_quantity: cogs_for_missing_quantity,
    reconcile_cogs: reconcile_cogs, cogs_shipped: cogs_shipped, cogs_return: cogs_return, cogs_lost: cogs_lost, cogs_adjusted: cogs_adjusted, cogs_removal: cogs_removal,
    tcogs: tcogs, business_expense: business_expense, net_profit: net_profit
  })
  return newData;
}
let handleRemoveDuplicated = async (records) => {
  let result = [];
  //  xóa dòng empty
  // let empty = records.filter(r=> r.sku == undefined)[0]
  // if(empty){
  //   console.log(empty);
  //   result.push(empty)
  // }
  let processedFnskus = {};
  for (let i = 0; i < records.length; i++) {
    let currentRecord = records[i];
    // Kiểm tra xem fnsku đã được xử lý chưa
    if (!processedFnskus[currentRecord.fnsku] && currentRecord.sku != undefined) {
      let reversedRecord = records.find(
        (record) =>
          record.sku === currentRecord.fnsku && record.fnsku === currentRecord.sku
      );

      if (reversedRecord) {
        // Nếu tìm thấy, cộng dồn data của bản ghi đảo ngược vào data của bản ghi hiện tại
        currentRecord.sale_quantity += reversedRecord.sale_quantity;
        currentRecord.refund_quantity += reversedRecord.refund_quantity;
        currentRecord.product_sales += reversedRecord.product_sales;
        currentRecord.refund_amount += reversedRecord.refund_amount;
        currentRecord.liquidations += reversedRecord.liquidations;
        currentRecord.gross_sales += reversedRecord.gross_sales;
        currentRecord.product_sales_tax += reversedRecord.product_sales_tax
        currentRecord.shipping_credits += reversedRecord.shipping_credits
        currentRecord.shipping_credit_tax += reversedRecord.shipping_credit_tax
        currentRecord.gift_wrap_credits += reversedRecord.gift_wrap_credits
        currentRecord.gift_wrap_credits_tax += reversedRecord.gift_wrap_credits_tax
        currentRecord.regulatory_fee += reversedRecord.regulatory_fee
        currentRecord.regulatory_fee_tax += reversedRecord.regulatory_fee_tax
        currentRecord.promotional_rebates += reversedRecord.promotional_rebates
        currentRecord.promotional_rebates_tax += reversedRecord.promotional_rebates_tax
        currentRecord.marketplace_withheld_tax += reversedRecord.marketplace_withheld_tax
        currentRecord.referral_fees += reversedRecord.referral_fees
        currentRecord.fullfillment_fees += reversedRecord.fullfillment_fees,
          currentRecord.refund_commission += reversedRecord.refund_commission,
          currentRecord.other_transaction_fee += reversedRecord.other_transaction_fee,
          currentRecord.other_adjustment += reversedRecord.other_adjustment
        currentRecord.gross_profits += reversedRecord.gross_profits
        currentRecord.ads += reversedRecord.ads;
        currentRecord.storage_fee += reversedRecord.storage_fee
        currentRecord.disposal_fee += reversedRecord.disposal_fee
        currentRecord.vine_fee += reversedRecord.vine_fee
        currentRecord.aged_inventory_surcharge += reversedRecord.aged_inventory_surcharge
        currentRecord.gross_profits_overall += reversedRecord.gross_profits_overall
        currentRecord.mcf_quantity += reversedRecord.mcf_quantity;
        currentRecord.lost_quantity_by_aw += reversedRecord.lost_quantity_by_aw
        currentRecord.adjusted_quantity_by_aw += reversedRecord.adjusted_quantity_by_aw;
        currentRecord.removal_liquidations += reversedRecord.removal_liquidations;
        currentRecord.removal_return += reversedRecord.removal_return;
        currentRecord.removal_disposal += reversedRecord.removal_disposal;
        currentRecord.customer_return_sellable += reversedRecord.customer_return_sellable
        currentRecord.customer_return_unsellable += reversedRecord.customer_return_unsellable;
        currentRecord.sellable_return_percent = (currentRecord.customer_return_sellable + currentRecord.customer_return_unsellable) != 0 ?
          currentRecord.customer_return_sellable / (currentRecord.customer_return_sellable + currentRecord.customer_return_unsellable) * 100 + "%" : "0%"
        currentRecord.cogs_shipped += reversedRecord?.cogs_shipped;
        currentRecord.cogs_return += reversedRecord?.cogs_return
        currentRecord.cogs_lost += reversedRecord?.cogs_lost
        currentRecord.cogs_adjusted += reversedRecord?.cogs_adjusted
        currentRecord.cogs_removal += reversedRecord?.cogs_removal
        currentRecord.tcogs += reversedRecord?.tcogs
        // Đánh dấu fnsku đã được xử lý
        processedFnskus[currentRecord.fnsku] = true;
      }
    }

    // Thêm bản ghi hiện tại vào kết quả nếu nó chưa tồn tại trong kết quả
    if (!processedFnskus[currentRecord.sku] && currentRecord.sku != undefined) {
      result.push(currentRecord);
    }
  }
  return result
}
// sửa date ở đây
let findRemainFields = async (splitEmptySkuData, payments) => {
  splitEmptySkuData.forEach(sku => {
    sku.net_profit = sku.gross_profits_overall + sku.tcogs
    // sku.group = undefined;
    let tmp = payments.filter(p => p.sku == sku.sku && p.description == "FBA Inventory Reimbursement - Lost:Inbound" &&
      (new Date(p.date) >= new Date("09/01/2023")) && (new Date(p.date) < new Date("10/01/2023")))
    sku.reimbursed_quantity = 0;
    sku.reimbursement_for_missing_quantity = 0;
    if (tmp) {
      for (var i = 0; i < tmp.length; i++) {
        sku.reimbursed_quantity += parseInt(tmp[i].quantity)
        sku.reimbursement_for_missing_quantity += parseFloat(tmp[i].other);
      }
    }
  })
  return splitEmptySkuData
}
function hasAllZeroPropertiesExcept(obj, excludedProps) {
  for (const key in obj) {
    if (obj.hasOwnProperty(key)) {
      if (!excludedProps.includes(key) && obj[key] !== 0 && obj[key] != undefined && obj[key] != "0%" && obj[key] !== -0
        && obj[key] !== "0") {
        return false; // Nếu có ít nhất một thuộc tính không nằm trong danh sách loại trừ và có giá trị khác 0, trả về false
      }
    }
  }
  return true; // Nếu tất cả các thuộc tính nằm trong danh sách loại trừ hoặc có giá trị bằng 0, trả về true
}
let removeNullRow = async (data) => {
  let rs = []
  data.forEach(d => {
    if (!hasAllZeroPropertiesExcept(d, ['group', 'sku', 'fnsku'])) {
      rs.push(d);
    }
  })
  rs.sort((a, b) => b.sale_quantity - a.sale_quantity);
  return rs;
}
let GenerateFile = async () => {
  const cogs_sheet = cogs.Sheets["transaction list"]

  // another way using csv
  const payment_csv = './input/2023/T8/Payment 01.01.23 - 30.09.23.csv';
  const storage_csv = './input/2023/T8/Storage Fee T8.23.csv';
  const surcharge_csv = './input/2023/T9/Surcharge Fee T9.23.csv';
  const removal_csv = "./input/2023/T8/Removal Fee 01.01.23 - 30.09.23.csv"
  const inventory_ledger_csv = "./input/2023/T8/Inventory Ledger 01.01.23 - 30.09.23.csv"

  // handle ads
  const wb_ads_product = XLSX.readFile('./input/2023/T9/Sponsored Products Advertised product report T9.23.xlsx')
  const wb_ads_display = XLSX.readFile('ads_display.xlsx')
  const wb_ads_brand = XLSX.readFile('./input/2023/T9/Sponsored Brands Campaign report T9.23.xlsx')
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
  const cogs_fixed = await axios.get("https://dl.dropboxusercontent.com/scl/fi/ugm3pgv1d9b6vdqu8dyj1/COGS-Code-Web.xlsx?rlkey=b3ol7hpetuiwav9sxr6m2p0a4&dl=1", { responseType: "arraybuffer" });
  const cog_wb = XLSX.read(cogs_fixed.data);
  const arr2 = XLSX.utils.sheet_to_json(cog_wb.Sheets["Product List"])
  let costOfGods = []
  costOfGods = arr2.map((row) => {
    return new CostOfGods(
      row['SKU'],
      row['FNSKU'],
      row['Price'],
      row['Portfolio'],
      row['Product Group']
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

  // payment
  const payment_data = new Promise((resolve, reject) => {
    const data = [];
    fs.createReadStream(payment_csv)
      .pipe(csv())
      .on('data', (row) => {
        const x = Object.keys(row)[0];
        let date = row[x]
        // Process each row of CSV data and store in an array
        data.push(new Payment(
          date,
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
        ));
      })
      .on('end', () => {
        resolve(data);
      })
      .on('error', (error) => {
        reject(error);
      });
  });

  const storage_fee_data = new Promise((resolve, reject) => {
    const data = [];
    fs.createReadStream(storage_csv)
      .pipe(csv())
      .on('data', (row) => {
        data.push(new StorageFee(
          row['fnsku'],
          row['estimated_monthly_storage_fee']
        ));
      })
      .on('end', () => {
        resolve(data);
      })
      .on('error', (error) => {
        reject(error);
      });
  });

  const surcharge_data = new Promise((resolve, reject) => {
    const data = [];
    fs.createReadStream(surcharge_csv)
      .pipe(csv())
      .on('data', (row) => {
        data.push(new SurchargeFee(
          row['sku'],
          row['fnsku'],
          row['short-time-range-long-term-storage-fee'] || row['amount-charged']
        ));
      })
      .on('end', () => {
        resolve(data);
      })
      .on('error', (error) => {
        reject(error);
      });
  });


  const inventory_ledger_data = new Promise((resolve, reject) => {
    const data = [];
    fs.createReadStream(inventory_ledger_csv)
      .pipe(csv())
      .on('data', (row) => {
        data.push(new InventoryLedger(
          row['Date and Time'],
          row['FNSKU'],
          row['MSKU'],
          row['Quantity'],
          row['Disposition'],
          row['Event Type']
        ));
      })
      .on('end', () => {
        resolve(data);
      })
      .on('error', (error) => {
        reject(error);
      });
  });

  const customer_data = new Promise((resolve, reject) => {
    const data = [];
    fs.createReadStream(removal_csv)
      .pipe(csv())
      .on('data', (row) => {
        data.push(new CustomerReturn(
          row['request-date'],
          row['last-updated-date'],
          row['sku'],
          row['fnsku'],
          row['disposition'],
          row['order-type'],
          row['order-status'],
          row['shipped-quantity'],
          row['disposed-quantity'],
          row['removal-fee']
        ));
      })
      .on('end', () => {
        resolve(data);
      })
      .on('error', (error) => {
        reject(error);
      });
  });
  const [payments, storageFee, surChargeFee, adjustment,
    customerReturn] = await Promise.all([
      payment_data, storage_fee_data, surcharge_data, inventory_ledger_data, customer_data
    ]);
  //   let rs = getSKUData(payments, costOfGods, ads_brand, ads_display, ads_product, storageFee, surChargeFee, adjustment,
  //     customerReturn);
  let rs = getSKUData(payments, costOfGods, ads_brand, null, ads_product, null, surChargeFee, adjustment,
    customerReturn);
  let final = await findCogs(rs, cogs_data)
  let finalData = []
  if (final) {
    finalData = await handleRemoveDuplicated(final)
  }
  let findRemainFieldsData = await findRemainFields(finalData, payments)
  let removeNullRowData = await removeNullRow(findRemainFieldsData)
  let splitEmptySkuData = await splitEmptySku(removeNullRowData, payments);

  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(splitEmptySkuData);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "P&L");
  // rename colummn
  XLSX.utils.sheet_add_aoa(newWorksheet, [["Group", "Sku", "Fnsku", "Sale Quantity", "Refund Quantity", "Product Sales", "Refund Amount", "Liquidations", "Gross Sales",
    "Product Sales Tax", "Shipping Credits", "Shipping Credits Tax", "Gift Wrap Credits", "Giftwrap Credits Tax", "Regulatory Fee", "Regulatory Fee Tax", "Promotional Rebates", "Promotional Rebates Tax",
    "Marketplace Withheld Tax", "Referral Fees", "Fulfillment Fees", "Refund Commission", "Other Transaction Fees", "Other Adjustment", "Gross Profit (by product)", "Ads", "Storage Fee",
    "Removal Fee", "Vine Fee", "Aged-inventory Surcharge", "Gross Profit (overall)", "MCF Quantity", "Lost/Damaged Quantity", "Adjusted Quantity", "Removal Quantity of Sellable Units (Liquidations)"
    , "Removal Quantity of Sellable Units (Return)", "Removal Quantity of Sellable Units (Disposal)", "Customer Return Quantity (Sellable)", "Customer Return Quantity (Unsellable)", "% Sellable Returns",
    "Discrepancy Quantity", "Shipment ID", "Quantity Found", "Reimbursed Quantity", "Quantity is Not Reimbursed",
    "Reimbursement for Missing Quantity", "COGS for Missing Quantity", "COGS Reconcile", "COGS Shipped", "COGS Customer Return", "COGS Lost/Damaged", "COGS Adjusted", "COGS Removal", "TCOGS", "Other Expenses", "Net Profit"]], { origin: "A1" });
  XLSX.writeFile(newWorkbook, 'P&L.xlsx');
}

GenerateFile()




