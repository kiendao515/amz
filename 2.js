data = [
  {
      "Sku": 1,
      "fnsku": "abc",
      "referenceID": "FBA170L0THG1",
      "date": "2023-03-17T00:00:00-0700"
  },
  {
      "Sku": 1,
      "fnsku": "abc",
      "referenceID": "FBA170L0THG1",
      "date": "2022-01-01T00:00:00-0800"
  },
  {
      "Sku": 1,
      "fnsku": "abc",
      "referenceID": "FBA16V64L910",
      "date": "2023-01-04T00:00:00-0800"
  },
  {
      "Sku": 1,
      "fnsku": "abc",
      "referenceID": "FBA16V64L910",
      "date": "2023-05-01T00:00:00-0800"
  },
  {
    "Sku": 1,
    "fnsku": "def",
    "referenceID": "FBA16V64L910",
    "date": "2022-03-04T00:00:00-0800"
  }
]
data.sort((a, b) => {
    if (a.referenceID === b.referenceID) {
      // Sắp xếp các bản ghi có cùng sku và cùng referenceID theo date tăng dần
      return new Date(a.date) - new Date(b.date);
    } else {
      // Sắp xếp các bản ghi có cùng sku nhưng khác referenceID theo date giảm dần
      return new Date(b.date) - new Date(a.date);
    }
});

console.log(data);

