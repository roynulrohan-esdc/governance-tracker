import { readFileSync } from "fs";
import { exit } from "process";
import * as xlsx from "xlsx";
import pptxgen from "pptxgenjs";

let buffer: Buffer;

try {
  buffer = readFileSync("CAD.xlsx");
} catch (e) {
  if (-4058 === e.errno) {
    console.log("\n\nERROR: CAD.xlsx not found. Make sure it is present in the current directory and named correctly.\n\n");
  } else {
    console.log(e);
  }
  exit(1);
}

const workbook = xlsx.read(buffer);

const createPowerpoint = () => {
  const sheets = Object.entries(workbook.Sheets);

  let pres = new pptxgen();
  let slides: pptxgen.Slide[] = [];

  sheets.forEach((sheet, i) => {
    slides[i] = pres.addSlide();

    slides[i].addText(`${sheet[0]}`, { x: 1, y: 1, color: "363636" });

    slides[i].addText(`${sheet[1]["A1"]["w"]}`, { x: 1, y: 2, color: "363636" });
  });

  pres
    .writeFile({ fileName: `Presentation.pptx` })
    .then(() => {
      console.log("\n\nSuccessfully converted to powerpoint => Presentation.pptx\n\n");
    })
    .catch((err) => {
      console.log(err);
      exit(0);
    });
};

createPowerpoint();
