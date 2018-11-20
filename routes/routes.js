const XLSX = require("xlsx");
const axios = require("axios");
const cheerio = require("cheerio");
let workbook = undefined;
let result;
let fullFileName;

let appRouter = function(app) {
    app.get("/", function(req, res) {
        if (!req.query.search) {
            result = { found: "Please specify what to search!" };
            res.send(JSON.stringify(result));
            return;
        }
        if (!workbook) {
            downloadExcel(req.query.search, req.query.page).then(()=>{
                res.send(JSON.stringify(result));
            });
        }
        else {
            searchInWorkBook(workbook, req.query.search, req.query.page).then(()=>{
                res.send(JSON.stringify(result));
            });
        }
    });
}

let downloadExcel = function(search, page=0){      
    return axios
        .get("https://www.mvcr.cz/clanek/informace-o-stavu-rizeni.aspx")
        .then(response => {
        const tmp = response.data;
        const fileName = cheerio
            .load(tmp)(".dark")
            .attr("href");
        fullFileName = fileName;
        const fullUrl = "https://www.mvcr.cz/" + fullFileName;
        return axios.get(fullUrl, {
                responseType: 'arraybuffer'
            })
            .then(function(res) {
            /* get the data as a Blob */
            if (res.statusText !== "OK") throw new Error("fetch failed");
            // return res.arrayBuffer();
            return res.data;
            })
            .then(function(ab) {
            /* parse the data when it is received */
            const data = new Uint8Array(ab);
            workbook = XLSX.read(data, { type: "array" });
            return searchInWorkBook(workbook, search, page);
            }).catch(err=>console.log(err));
        }).catch(err=>{
            console.log(err);
        });
}

let searchInWorkBook = function(workbook, search, page=0) {
    const sheet = workbook.Sheets[workbook.SheetNames[page]];
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    for (var R = range.s.r; R <= range.e.r; ++R) {
      for (var C = range.s.c; C <= range.e.c; ++C) {
        /* find the cell object */
        var cellref = XLSX.utils.encode_cell({ c: C, r: R }); // construct A1 reference for cell
        if (!sheet[cellref]) continue; // if cell doesn't exist, move on
        var cell = sheet[cellref];
        /* if the cell is a text cell with the searched string */
        if (!(cell.t == "s" || cell.t == "str")) continue; // skip if cell is not text
        if (cell.v.includes(search)) {
          result = {found: "Found!!! Pick it Up!", fileName: fullFileName};
        } else {
          result = {found: "NOT Found!!!", fileName: fullFileName};
        }
      }
    }
    return Promise.resolve(result);
  }

module.exports = appRouter;