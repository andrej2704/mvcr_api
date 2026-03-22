const ExcelJS = require("exceljs");
const axios = require("axios");
const cheerio = require("cheerio");
let workbook = undefined;
let fullFileName;

let appRouter = function(app) {
    app.get("/", function(req, res) {
        if (!req.query.search) {
            res.send(JSON.stringify({ found: "Please specify what to search!" }));
            return;
        }
        const search = req.query.search;
        const page = req.query.page || 0;
        if (!workbook) {
            downloadExcel(search, page).then((result) => {
                res.send(JSON.stringify(result));
            }).catch((err) => {
                console.error(err);
                res.status(500).send(JSON.stringify({ error: "Failed to fetch data from MVCR." }));
            });
        } else {
            searchInWorkBook(workbook, search, page).then((result) => {
                res.send(JSON.stringify(result));
            }).catch((err) => {
                console.error(err);
                res.status(500).send(JSON.stringify({ error: "Failed to search workbook." }));
            });
        }
    });
}

let downloadExcel = function(search, page = 0) {
    return axios
        .get("https://www.mvcr.cz/clanek/informace-o-stavu-rizeni.aspx")
        .then(response => {
            const tmp = response.data;
            const fileName = cheerio
                .load(tmp)(".dark")
                .attr("href");
            fullFileName = fileName;
            const fullUrl = "https://www.mvcr.cz/" + fullFileName;
            return axios.get(fullUrl, { responseType: 'arraybuffer' });
        })
        .then(function(res) {
            const wb = new ExcelJS.Workbook();
            return wb.xlsx.load(Buffer.from(res.data)).then(() => {
                workbook = wb;
                return searchInWorkBook(workbook, search, page);
            });
        });
}

let searchInWorkBook = function(workbook, search, page = 0) {
    const sheetIndex = parseInt(page, 10) || 0;
    if (sheetIndex < 0 || sheetIndex >= workbook.worksheets.length) {
        return Promise.reject(new Error("Invalid page number."));
    }
    const worksheet = workbook.worksheets[sheetIndex];
    let found = false;
    worksheet.eachRow(function(row) {
        if (found) return;
        row.eachCell(function(cell) {
            if (found) return;
            if (typeof cell.value === 'string' && cell.value.includes(search)) {
                found = true;
            }
        });
    });
    const result = found
        ? { found: "Found!!! Pick it Up!", fileName: fullFileName }
        : { found: "NOT Found!!!", fileName: fullFileName };
    return Promise.resolve(result);
}

module.exports = appRouter;