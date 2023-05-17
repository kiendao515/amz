function findMatchingTransactions(data, listTransaction) {
    const result = [];
  
    data.forEach(dataItem => {
      let { Sku, Data } = dataItem;
      let sum = 0;
  
      const matchingTransactions = listTransaction.filter(transaction => transaction.Sku === Sku);
  
      matchingTransactions.forEach(transaction => {
        console.log(Data);
         Data = Data - transaction.quantity;
         if(Data == 0){
            result.push(transaction);
         }
      });
      
    });
  
    return result;
  }
  
  // Mảng dữ liệu data
  const data = [
    { Sku: 1, fnsku: 'abc', Data: 20 },
    {Sku:3,fnsku:'haohan',Data:10}
  ];
  
  // Mảng transaction
  const listTransaction = [
    { Sku: 1, fnsku: 'abc', quantity: 1, date: '02/02/2020' },
    { Sku: 2, fnsku: 'ehd', quantity: 1, date: '03/02/2020' },
    { Sku: 1, fnsku: 'abc', quantity: 3, date: '10/02/2020' },
    { Sku: 1, fnsku: 'abc', quantity: 16, date: '20/02/2020' },
    { Sku: 1, fnsku: 'abc', quantity: 16, date: '22/02/2020' },
  ];
  
  // Gọi hàm findMatchingTransactions
  const result = findMatchingTransactions(data, listTransaction);
  
  console.log(result);
  