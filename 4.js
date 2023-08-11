var danhSachObj = [
    { data: 5 },
    { data: 3 },
    { data: 8 },
    { data: 1 }
  ];
  
  danhSachObj.sort((a, b) => b.data - a.data);
  
  console.log(danhSachObj);