import express from 'express'
import excel from 'exceljs';

const app = express();

app.use(express.json(
    { "limit": "1000mb" }
))


app.get("/api/health", function (req, res) {
    res.end("Service workng!");
});

app.post("/api/get-xlsx-data", express.json({ "limit": "1000mb" }), async function (req, res) {
    if (!req.body || (req.body.base64File == undefined || req.body.base64File.length === 0)) {
        return res.send("base64File is Empty");
    }

    const workbook = new excel.Workbook();

    try {

        const fileDataDecoded = Buffer.from(req.body.base64File, 'base64');
        await workbook.xlsx.load(fileDataDecoded);

        let jsonData: any[] = [];

        workbook.worksheets.forEach(function (sheet: excel.Worksheet) {

            let firstRow = sheet.getRow(1);

            if (!firstRow.cellCount) return;

            let keys: any = firstRow.values;

            sheet.eachRow((row, rowNumber) => {

                if (rowNumber == 1) return;
                let values: any = row.values;
                let obj: any = {};

                for (let i = 1; i < keys.length; i++) {
                    obj[keys[i]] = values[i];
                }

                jsonData.push(obj);
            });

        });

        return res.status(200).send(JSON.stringify(jsonData));
    }
    catch (err) {
        res.status(400).send("Error get-xlsx-data(): " + err);
    }
})

const port = process.env.PORT || 8080

app.listen(port, () => {
    console.log("get-xlsx-data are listening port : " + port);
});
