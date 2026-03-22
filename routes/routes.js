const XLSX = require("xlsx");
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
            const data = new Uint8Array(res.data);
            workbook = XLSX.read(data, { type: "array" });
            return searchInWorkBook(workbook, search, page);
        });
}

let searchInWorkBook = function(workbook, search, page = 0) {
    const sheetIndex = parseInt(page, 10) || 0;
    if (sheetIndex < 0 || sheetIndex >= workbook.SheetNames.length) {
        return Promise.reject(new Error("Invalid page number."));
    }
    const sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (var R = range.s.r; R <= range.e.r; ++R) {
        for (var C = range.s.c; C <= range.e.c; ++C) {
            var cellref = XLSX.utils.encode_cell({ c: C, r: R });
            if (!sheet[cellref]) continue;
            var cell = sheet[cellref];
            if (!(cell.t == "s" || cell.t == "str")) continue;
            if (cell.v.includes(search)) {
                return Promise.resolve({ found: "Found!!! Pick it Up!", fileName: fullFileName });
            }
        }
    }
    return Promise.resolve({ found: "NOT Found!!!", fileName: fullFileName });
}

module.exports = appRouter;