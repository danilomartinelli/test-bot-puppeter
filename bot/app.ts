// Imports
import * as fs from "fs";
import * as Excel from "exceljs";
import * as puppeteer from "puppeteer";
import * as downloadPDF from "download-pdf";
import * as PDFMerge from "pdf-merge";

import { asyncForEach } from "./utils";

// Types
interface IData {
	statusCell: string;
	codes: string[];
}

// Constants
const EXCEL_DIR = "./excel/data.xlsx";
const BASE_URL = "https://repositorio.ufba.br/ri/handle/ri/";
const TEMP_DIR = "./.tmp";
const PDF_DIR = "./dist";

// Entry method to start BOT
init();

async function getUniquePDF(
	codes: string[],
	page: puppeteer.Page,
	row: string
) {
	await asyncForEach(codes, async code => {
		const path = TEMP_DIR + "/";
		await page.goto(BASE_URL + code, {
			waitUntil: "networkidle2"
		});

		const element = await page.$(
			"body > table.centralPane > tbody > tr:nth-child(1) > td.pageContents > table:nth-child(6) > tbody > tr > td > table > tbody > tr:nth-child(2) > td:nth-child(1) > a"
		);
		const url =
			"https://repositorio.ufba.br" +
			(await page.evaluate(obj => {
				return obj.getAttribute("href");
			}, element));

		console.log(`\t- Download do PDF de código: ${code}`);
		await downloadPDF(
			url,
			{ directory: path, filename: code + ".pdf" },
			err => {
				if (err) {
					console.log(
						`\t- Não foi possível fazer o download do PDF de código: ${code}`
					);
				}
			}
		);
	});

	console.log(`Juntando os PDFs da linha ${row}.\n`);
	try {
		await PDFMerge(codes.map(code => `${TEMP_DIR}/${code}.pdf`), {
			output: `${PDF_DIR}/linha_${row}.pdf`
		});
	} catch (err) {
		console.log(
			"Pulando o merge dos PDFs da linha ${row}. Erro: " + err.message + "\n"
		);
	}
}

async function runCrawler(
	page: puppeteer.Page,
	archivesIDs: IData[],
	excelData: Excel.Worksheet
) {
	await asyncForEach(archivesIDs, async (value, idx) => {
		console.log(`Executando ${idx + 1} de ${archivesIDs.length}...`);
		const statusCell = excelData.getCell(value.statusCell);
		if (statusCell.text === "Pendente") {
			await getUniquePDF(value.codes, page, statusCell.row);
		}
	});
}

async function initBrowser() {
	const browser = await puppeteer.launch();
	return browser;
}

// Loading Excel File
async function loadingExcel() {
	const workbook = new Excel.Workbook();
	try {
		const excelFile = await workbook.xlsx.readFile(EXCEL_DIR);
		return excelFile.getWorksheet(1);
	} catch (err) {
		console.error(err);
		process.exit(1);
	}
}

// Get all Codes and Status
function getArchivesIDs(excelData: Excel.Worksheet) {
	const column = excelData.getColumn("A");
	const res: IData[] = [];

	const arr: number[] = [];
	column.eachCell((_, rowNumber) => {
		if (rowNumber !== 1) {
			arr.push(rowNumber);
		}
	});

	arr.forEach(rowID => {
		const row = excelData.getRow(rowID);
		const obj: IData = {
			statusCell: "A" + rowID,
			codes: []
		};
		row.eachCell((cell, idx) => {
			if (idx !== 1) {
				obj.codes.push(cell.text);
			}
		});
		res.push(obj);
	});

	return res;
}

// Init Function
async function init() {
	// Create folders and files
	if (!fs.existsSync(TEMP_DIR)) {
		fs.mkdirSync(TEMP_DIR);
	}
	if (!fs.existsSync(PDF_DIR)) {
		fs.mkdirSync(PDF_DIR);
	}

	// Init services
	const browser = await initBrowser();
	const excelData = await loadingExcel();
	const page = await browser.newPage();

	const archivesIDs = getArchivesIDs(excelData);
	await runCrawler(page, archivesIDs, excelData);

	await browser.close();
}
