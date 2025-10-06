import fs from "fs";
import path from "path";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";

// Hàm fill data
function fillDocx() {
  // 1. Đọc file docx trong src
  const templatePath = path.resolve(__dirname, "template.docx");
  const content = fs.readFileSync(templatePath, "binary");

  // 2. Load bằng pizzip
  const zip = new PizZip(content);

  // 3. Khởi tạo docxtemplater
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    delimiters: {
      start: "[[", // Bắt đầu bằng 2 dấu ngoặc vuông
      end: "]]", // Kết thúc bằng 2 dấu ngoặc vuông
    },
  });

  // 4. Data cần fill
  const data = {
    name: "Nguyễn Văn A",
    address: "123 Đường A, Quận B, Hà Nội",
    loanAmount: "1,000,000 USD",
  };

  // 5. Render
  doc.render(data);

  // 6. Xuất ra file mới
  const buffer = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  const outputPath = path.resolve(__dirname, "output.docx");
  fs.writeFileSync(outputPath, buffer);

  console.log("✅ Đã tạo file:", outputPath);
}

// Chạy thử
fillDocx();
