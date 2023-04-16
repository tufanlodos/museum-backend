const express = require("express");
const XLSX = require("xlsx");

const WB_NAME = "TSMuze.xlsx";
const WS_TROPHIES = "Trophies";
let workbook = null;

const app = express();

// parse json request body
app.use(express.json());

// parse urlencoded request body
app.use(express.urlencoded({ extended: true }));

app.get("/trophies", (req, res) => {
  const trophies = XLSX.utils.sheet_to_json(workbook.Sheets[WS_TROPHIES]);

  res.send({ trophies });
});

app.post("/trophy", (req, res) => {
  const { name, description, year } = req.body;
  if (!name || !description || !year) {
    return res
      .status(400)
      .send({ message: "Name, description and year fields are required" });
  }

  let worksheet;
  const trophies = XLSX.utils.sheet_to_json(workbook.Sheets[WS_TROPHIES]);
  const trophy = {
    name,
    description,
    year,
    created_at: new Date().toISOString(),
    updated_at: new Date().toISOString(),
  };
  if (trophies.length === 0) {
    trophy.id = "1";
    worksheet = XLSX.utils.json_to_sheet([trophy]);
    XLSX.utils.book_append_sheet(workbook, worksheet, WS_TROPHIES);
  } else {
    const lastTrophyID = Number(trophies.slice(-1)[0].id);
    trophy.id = (lastTrophyID + 1).toString();
    XLSX.utils.sheet_add_json(workbook.Sheets[WS_TROPHIES], [trophy], {
      skipHeader: true,
      origin: -1,
    });
  }
  XLSX.writeFile(workbook, WB_NAME);

  return res.status(200).send({ success: true, data: trophy });
});

app.put("/trophy/:id", (req, res) => {
  const { id } = req.params;
  if (!id) {
    return res.status(400).send({ message: "Id is required" });
  }

  const trophies = XLSX.utils.sheet_to_json(workbook.Sheets[WS_TROPHIES]);
  const trophy = trophies.find((t) => t.id === id);
  if (!trophy) {
    return res.status(400).send({ message: "Trophy not found" });
  }

  const { name, description, year } = req.body;
  trophy.name = name || trophy.name;
  trophy.description = description || trophy.description;
  trophy.year = year || trophy.year;
  trophy.updated_at = new Date().toISOString();

  workbook = XLSX.utils.book_new();
  const updatedWorksheet = XLSX.utils.json_to_sheet(trophies);
  XLSX.utils.book_append_sheet(workbook, updatedWorksheet, WS_TROPHIES);
  XLSX.writeFile(workbook, WB_NAME);

  res.status(200).send("PUT");
});

app.delete("/trophy/:id", (req, res) => {
  const { id } = req.params;
  if (!id) {
    return res.status(400).send({ message: "Id is required" });
  }

  const trophies = XLSX.utils.sheet_to_json(workbook.Sheets[WS_TROPHIES]);
  const trophyIndex = trophies.findIndex((t) => t.id === id);
  if (trophyIndex === -1) {
    return res.status(400).send({ message: "Trophy not found" });
  }

  workbook = XLSX.utils.book_new();
  trophies.splice(trophyIndex, 1);
  const updatedWorksheet = XLSX.utils.json_to_sheet(trophies);
  XLSX.utils.book_append_sheet(workbook, updatedWorksheet, WS_TROPHIES);
  XLSX.writeFile(workbook, WB_NAME);

  res.status(200).send("DELETE");
});

app.listen(1967, () => {
  try {
    workbook = XLSX.readFile(WB_NAME);
  } catch (error) {}

  if (workbook === null) {
    workbook = XLSX.utils.book_new();
  }
  console.log("started at 1967");
});
