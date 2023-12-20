// Mảng các bản ghi chứa các thuộc tính sku, fnsku, data
let records = [
  { sku: 'A', fnsku: '111', data: 'some data 1' },
  { sku: 'B', fnsku: '222', data: 'some data 2' },
  { sku: 'C', fnsku: '333', data: 'some data 3' },
  { sku: '111', fnsku: 'A', data: 'reversed data 1' },
  { sku: '222', fnsku: 'B', data: 'reversed data 2' },
  { sku: 'D', fnsku: 'C', data: 'okkk' },
];

// Tạo một bản sao của mảng records để lưu trữ kết quả
let result = [];

// Tạo một đối tượng tạm thời để theo dõi các fnsku đã được xử lý
let processedFnskus = {};

// Duyệt qua mảng records và tìm và cộng dồn các bản ghi bị đảo ngược sku và fnsku
for (let i = 0; i < records.length; i++) {
  let currentRecord = records[i];

  // Kiểm tra xem fnsku đã được xử lý chưa
  if (!processedFnskus[currentRecord.fnsku]) {
    let reversedRecord = records.find(
      (record) =>
        record.sku === currentRecord.fnsku && record.fnsku === currentRecord.sku
    );

    if (reversedRecord) {
      // Nếu tìm thấy, cộng dồn data của bản ghi đảo ngược vào data của bản ghi hiện tại
      currentRecord.data += ` | ${reversedRecord.data}`;

      // Đánh dấu fnsku đã được xử lý
      processedFnskus[currentRecord.fnsku] = true;
    }
  }

  // Thêm bản ghi hiện tại vào kết quả nếu nó chưa tồn tại trong kết quả
  if (!processedFnskus[currentRecord.sku]) {
    result.push(currentRecord);
  }
}

console.log(result);
