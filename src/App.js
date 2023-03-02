import * as ExcelJS from "exceljs";
import gambar from "./logo.svg";
import { Buffer } from "buffer";

function App() {
  const createExcelFile = async (data) => {
    // Membuat sebuah workbook baru
    const workbook = new ExcelJS.Workbook();

    // Membuat sebuah worksheet baru
    const worksheet = workbook.addWorksheet("Sheet 1");
    // Menambahkan header ke worksheet
    worksheet.columns = [
      { header: "", key: "no",  },
      { header: "", key: "nama" },
      { header: "", key: "umur" },
      { header: "", key: "email", width: 20 },
    ];

    const rown = worksheet.getRow(1);
    rown.height = 80;

    const headerImage = await fetch("/logo192.png")
      .then((response) => response.arrayBuffer())
      .then((buffer) => Buffer.from(buffer));

    const imageId = workbook.addImage({
      buffer: headerImage,
      extension: "png",
    });

    worksheet.addImage(imageId, "D1:D1");
    // Menambahkan data header dan ke worksheet
    worksheet.addRow({
      no: "no",
      nama: "nama",
      umur: "umur",
      email: "email",
    });

    data.forEach((row) => {
      worksheet.addRow(row);
    });

    // Mengatur lebar kolom

    // Membuat file Excel
    const buffer = await workbook.xlsx.writeBuffer();

    return buffer;
  };

  const downloadExcelFile = async () => {
    // Membuat data yang akan ditampilkan di file Excel
    const data = [
      { no: 1, nama: "John Doe", umur: 30, email: "john.doe@example.com" },
      { no: 2, nama: "Jane Doe", umur: 25, email: "jane.doe@example.com" },
      { no: 3, nama: "Bob Smith", umur: 40, email: "bob.smith@example.com" },
    ];

    // Membuat file Excel
    const buffer = await createExcelFile(data);

    // Membuat file blob dari buffer
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    // Mengatur nama file
    const fileName = "example.xlsx";

    // Membuat link untuk mengunduh file Excel
    const link = document.createElement("a");
    link.href = window.URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
  };
  return (
    <div className="App">
      <button onClick={downloadExcelFile}>Download Excel</button>
    </div>
  );
}

export default App;
