const findPreviousDate = (skuData, transaction) => {
  let listTransactionOfSku = transaction.filter(t => t.sku === skuData.sku);
  let tmp = [];
  let rs = [];
  let index = skuData.listShipmentID.indexOf(skuData.shipmentID);
  for (let i = index; i < skuData.listShipmentID.length; i++) {
    tmp.push({
      shipmentID: skuData.listShipmentID[i],
      quantityOfShipment: skuData.listQuantityOfShipment[i]
    });
  }
  const startIndex = listTransactionOfSku.findIndex(t => t.shipmentID === skuData.shipmentID);
  const filteredTransactions = listTransactionOfSku.slice(startIndex);

  const processNextTmpElement = (index) => {
    if (index >= tmp.length) {
      // Base case: Reached the end of tmp array
      return;
    }

    let total = 0;
    for (let j = 0; j < filteredTransactions.length; j++) {
      const t = filteredTransactions[j];
      if (t.sku === skuData.sku) {
        total += t.quantity;
        if (-tmp[index].quantityOfShipment >= total) {
          let d = new Date(t.date);
          d.setFullYear(d.getFullYear() + 3);
          t.shipmentID = tmp[index].shipmentID;
          rs.push(new Cog(t.sku, t.fnsku, tmp[index].shipmentID, null, new Date(t.date), new Date(d), total + tmp[index].quantityOfShipment, tmp[index + 1]?.shipmentID, null));
          break;
        }
      }
    }

    // Process the next element in tmp array recursively
    processNextTmpElement(index + 1);
  };

  // Start the recursive processing from the first element of tmp array
  processNextTmpElement(0);

  // Return the result array rs or perform additional operations if needed
  return rs;
};

// Usage
const skuData = {
  sku: 'abc',
  shipmentID: 'xyz',
  listShipmentID: ['xyz', 'abc', 'def'],
  listQuantityOfShipment: [10, 20, 30]
};

const transactionData = [
  { sku: 'abc', shipmentID: 'abc', data: 5, quantity: 15, date: '2022-01-01' },
  { sku: 'abc', shipmentID: 'xyz', data: 10, quantity: 20, date: '2022-02-01' },
  { sku: 'abc', shipmentID: 'def', data: 15, quantity: 25, date: '2022-03-01' }
];

const cogs = [];

const result = findPreviousDate(skuData, transactionData);
console.log(result);
