const express = require('express');
const fs = require('fs');
const ExcelJS = require('exceljs');
require('dotenv').config();


const app = express();
app.use(express.static('/'));
app.use(express.json());

function extractHeaders(data) {
    const headerData = data[0];
    let headers = [];
    headers = Object.keys(headerData).map(element => { return { header: element, key: element, width: 30 } });
    return headers;
}

app.post('/jsontoexcel', async (req, res) => {


    try {
        // const workbook = new ExcelJS.Workbook();
        // console.log(req);
        // console.log(req.params);
        // console.log(req.body.data);
        // const data = req.body.data;
        // data?.forEach((ele, index) => {
        //     let sheet = workbook.addWorksheet(`sheet${index + 1}`)
        //     sheet.columns = extractHeaders(ele);
        //     // ele?.forEach((item) => {
        //     //     sheet.addRow(item);
        //     // })
        // })

          const workbook = new ExcelJS.Workbook();
           const data = req.body.data;

        let sheet = workbook.addWorksheet(`datasheet`)
            sheet.columns = extractHeaders(data);

        data.forEach((ele, index) => {
                sheet.addRow(ele);
        })

        await workbook.xlsx.writeFile('multi_sheet_excel.xlsx');
        return res.json({ success: true, message: "successfully created excel file",url:process.env.BACKEND_URL+'/download'});
    } catch (err) {

        return res.json({ success: false, message: err.message });
    }
})


app.get('/download',async(req,res)=>{
    try{
    const file = fs.readFileSync('multi_sheet_excel.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="data.xlsx"');

     return res.send(file);
    }catch(err){
        return res.json({success:false,message:err.message});
    }
})


app.get('/', (req, res) => {
    return res.json({ message: "Hello Welcome to Serve" });
})

app.listen(80, () => {
    console.log(`server running at ${process.env.BACKEND_URL}`);
})
