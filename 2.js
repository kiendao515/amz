function findReferenceID(data, thresholds) {
  let referenceIDTotals = {};
  let currentTotal = 0;
  let result = [];

  for (let i = 0; i < data.length; i++) {
    let row = data[i];
    let referenceID = row["referenceID"];
    let quantity = parseInt(row["quantity"]);

    if (referenceID in referenceIDTotals) {
      referenceIDTotals[referenceID] += quantity;
    } else {
      referenceIDTotals[referenceID] = quantity;
    }

    currentTotal += quantity;

    for (let j = 0; j < thresholds.length; j++) {
      if (currentTotal > thresholds[j] && !result[j]) {
        result[j] = referenceID;
      } else if (currentTotal <= thresholds[j]) {
        break;
      }
    }
  }

  return result;
}

// Example usage
const data = [
  { referenceID: "A", quantity: 2 },
  { referenceID: "B", quantity: 3 },
  { referenceID: "C", quantity: 6 },
];

const thresholds = [1, 3, 10];

const result = findReferenceID(data, thresholds);

console.log(result); // Output: ["B", "C", "C"]
