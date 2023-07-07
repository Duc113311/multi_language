const XLSX = require("xlsx");
const fs = require("fs");

// Đường dẫn tới file Excel
const filePath = "./language.xlsx";

// Đọc file Excel
// Đọc file Excel
const workbook = XLSX.readFile(filePath);

// Lấy tên sheet đầu tiên trong file Excel
const sheetName = workbook.SheetNames[0];

// Lấy sheet đầu tiên
const worksheet = workbook.Sheets[sheetName];

// Chuyển đổi sheet thành đối tượng JSON
const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

// Tạo đối tượng JSON kết quả
const result = {};
for (let index = 3; index < 100; index++) {
  const element = index;
  console.log("element", element);
  for (let i = 0; i < jsonData.length; i++) {
    const row = jsonData[i];
    if (row.length >= 5) {
      const code = row[0];
      const value = row[element];
      // Thêm dữ liệu vào đối tượng JSON kết quả
      result[code] = value;
    }
  }

  const nameFile = result.Key_code;
  var batDau = nameFile.indexOf("(") + 1;
  var ketThuc = nameFile.indexOf(")");
  var ketQua = nameFile.substring(batDau, ketThuc);
  console.log(nameFile);
  delete result["Key_code"];
  const jsonContent = JSON.stringify(result, null, 2);

  // Đường dẫn và tên file JSON đầu ra
  const outputFilePath = `./file/${ketQua}.json`;

  // Ghi dữ liệu vào file JSON
  fs.writeFileSync(outputFilePath, jsonContent);

  console.log(`File JSON đã được xuất ra: ${outputFilePath}`);
}
// Duyệt qua từng dòng trong dữ liệu Excel (bắt đầu từ dòng thứ 2)
