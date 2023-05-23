const data = 
[
    {
      date: '2023-04-17T00:00:00-0700',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA1710GKRGS',
      quantity: -1
    },
    {
      date: '2023-03-28T00:00:00-0700',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA170CCSSLT',
      quantity: 1
    },
    {
      date: '2023-02-28T00:00:00-0800',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA1710GKRGS',
      quantity: 96
    },
    {
      date: '2023-02-27T00:00:00-0800',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA1710GKRGS',
      quantity: -1
    },
    {
      date: '2023-02-25T00:00:00-0800',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA1710GKRGS',
      quantity: 192
    },
    {
      date: '2023-01-18T00:00:00-0800',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA170CCSSLT',
      quantity: -1
    },
    {
      date: '2023-01-09T00:00:00-0800',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA170CCSSLT',
      quantity: 24
    },
    {
      date: '2023-01-06T00:00:00-0800',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA170CCSSLT',
      quantity: 45
    },
    {
      date: '2023-05-03T00:00:00-0700',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA17337P5GZ',
      quantity: 100
    },
    {
      date: '2023-04-03T00:00:00-0700',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA1710GKRGS',
      quantity: 1
    },
    {
      date: '2023-03-03T00:00:00-0800',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA1710GKRGS',
      quantity: 1
    },
    {
      date: '2023-05-02T00:00:00-0700',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA17337P5GZ',
      quantity: 300
    },
    {
      date: '2023-04-02T00:00:00-0700',
      msku: '9K-FBJO-XTOF',
      fnsku: 'X003K6AQYR',
      shipmentID: 'FBA171HSLF02',
      quantity: 300
    }
  ]
  const groupedRecords = data.reduce((groups, record) => {
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

console.log(distinctRecords);

distinctRecords.sort((a, b) => {
    if (a.shipmentID !== b.shipmentID) {
        // Sắp xếp các bản ghi có cùng sku và cùng referenceID theo date tăng dần
        return new Date(b.date) - new Date(a.date);
    }
});
const referenceIDs = distinctRecords.map(item => item.shipmentID);
console.log(referenceIDs);


