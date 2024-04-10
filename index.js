const express = require("express");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const app = express();
const port = 3000;
app.use(express.static("public"));

const jsonData = {
  structure: ["name", "age", "ville", "code_postal", "adresse", "zip"],
  data: [
    {
      id: 1,
      name: "test1",
      age: 45,
      ville: "agadir",
      code_postal: 80750,
      adresse: "aourir",
      zip: "12345",
    },
    {
      id: 2,
      name: "test2",
      age: 45,
      ville: "agadir",
      code_postal: 80750,
      adresse: "aourir",
      zip: "12345",
    },
    {
      id: 3,
      name: "test3",
      age: 42,
      ville: "rabat",
      code_postal: 80700,
      adresse: "ta9adome",
      zip: "123458",
    },
  ],
};

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/src/index.html");
});

app.get("/save", (req, res) => {
  const groupByField = req.query.groupBy || "id";
  const groupedData = {};

  jsonData.data.forEach((obj) => {
    const groupByKey = obj[groupByField].toString();
    if (!groupedData[groupByKey]) {
      groupedData[groupByKey] = [];
    }
    groupedData[groupByKey].push(obj);
  });

  const srcFolderPath = path.join(__dirname, "src");
  const excelFolderPath = path.join(srcFolderPath, "excel");
  fs.mkdirSync(srcFolderPath, { recursive: true });
  +fs.mkdirSync(excelFolderPath, { recursive: true });

  const fileNames = [];
  for (const key in groupedData) {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data");
    const headers = jsonData.structure;
    worksheet.addRow(headers);

    groupedData[key].forEach((obj) => {
      const row = [];
      headers.forEach((header) => {
        row.push(obj[header]);
      });
      worksheet.addRow(row);
    });

    const fileName = `${groupByField}_${key}.xlsx`;
    const filePath = path.join(excelFolderPath, fileName);
    workbook.xlsx
      .writeFile(filePath)
      .then(() => {
        fileNames.push(filePath);
        console.log(`Excel file ${filePath} saved.`);
        if (Object.keys(groupedData).length === fileNames.length) {
          res.send(`Excel files saved locally: ${fileNames.join(", ")}`);
        }
      })
      .catch((err) => {
        console.error(`Error generating Excel file ${filePath}:`, err);
        res.status(500).send(`Error generating Excel file ${filePath}`);
      });
  }
});

app.get("/files", (req, res) => {
  const excelFolderPath = path.join(__dirname, "src", "excel");
  fs.readdir(excelFolderPath, (err, files) => {
    if (err) {
      console.error("Error reading directory:", err);
      res.status(500).send("Error reading directory");
      return;
    }
    res.json(files);
  });
});

app.get("/download", (req, res) => {
  const fileName = req.query.file;
  const filePath = path.join(__dirname, "src", "excel", fileName);
  res.download(filePath, (err) => {
    if (err) {
      console.error(`Error downloading file ${fileName}:`, err);
      res.status(500).send(`Error downloading file ${fileName}`);
    } else {
      console.log(`File ${fileName} downloaded successfully.`);
    }
  });
});

app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});
