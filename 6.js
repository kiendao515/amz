const XLSX = require("xlsx");
const axios = require("axios");

const url = "https://dl.dropboxusercontent.com/scl/fi/rei6i0wpbwocuxyym9xzd/COGS.xlsx?rlkey=hjy92xapyauogrgwplfbobmvq&dl=1";
(async() => {
  const res = await axios.get(url, {responseType: "arraybuffer"});
  /* res.data is a Buffer */
  const workbook = XLSX.read(res.data);
  const inventoryData = XLSX.utils.sheet_to_json(workbook.Sheets["Sheet1"])
  console.log(inventoryData[0].COGS);
  /* DO SOMETHING WITH workbook HERE */
})();