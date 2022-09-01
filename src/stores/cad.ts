import * as xlsx from "xlsx";
import pptxgen from "pptxgenjs";
import { get, writable } from "svelte/store";

const cad = writable(undefined);
const presentation = writable(undefined);

export const loading = writable(false);
export const readyToExport = writable(false);

export const parseExcel = async (file: File) => {
  loading.set(true);
  readyToExport.set(false);

  const data = await file.arrayBuffer();

  const workbook = xlsx.read(data);

  cad.set(workbook);

  createPowerpoint();
};

const createPowerpoint = () => {
  const sheets = get(cad).Sheets;

  let pres = new pptxgen();

  let slide = pres.addSlide();

  let textboxText = `${sheets["Foo"]["A1"]["w"]} ${sheets["Bar"]["A1"]["w"]}`;
  let textboxOpts = { x: 1, y: 1, color: "363636" };
  slide.addText(textboxText, textboxOpts);

  presentation.set(pres);

  loading.set(false);
  readyToExport.set(true);
};

export const exportPowerpoint = (fileName: string) => {
  get(presentation).writeFile({ fileName: `${fileName}.pptx` });
};
