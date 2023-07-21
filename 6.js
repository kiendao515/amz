const inputId = "FBA170MXM13R";
const dataStr = "FBA172J7C75H (60), FBA17337P5GZ (190), FBA170MXM13R (76), FBA172J3X26V (231), FBA1702DS2J7 (300), FBA16ZSQL93L (350), FBA16V77Q3T4 (1300), FBA16V76SJS1 (500), FBA167L56FMJ (1), FBA16PFFKY38 (301), FBA16P4CTPN2 (560), FBA16MR87K37 (-4)";

// Tách chuỗi thành mảng các phần tử
const dataArray = dataStr.split(', ');

// Tìm vị trí của inputId trong mảng
const index = dataArray.findIndex((element) => element.includes(inputId));

// Kiểm tra nếu inputId không tồn tại trong mảng, hoặc nếu nó là phần tử đầu tiên
// thì không có phần tử nào được trả về trước nó
if (index === -1 || index === 0) {
  console.log("Không có phần tử nào trước " + inputId);
} else {
  // Lấy các phần tử trước inputId (theo thứ tự xuôi)
  const elementsBeforeId = dataArray.slice(0, index).reverse();
  console.log(elementsBeforeId.join(' >> '));
}
