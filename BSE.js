
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const csv = require('csv-parser');
const chrome = require('selenium-webdriver/chrome');
const path = require('chromedriver').path;
const fs = require('fs');
const dirpath = require('path');
const XLSX = require('xlsx');
const { Builder, By, Key, until } = require('selenium-webdriver');
const service = new chrome.ServiceBuilder(path).build();
chrome.setDefaultService(service);
let url = 'https://www.bseindia.com/stock-share-price/';
let newjson = [];
const stocklist = [];
let header = [];
// const filePath = dirpath.resolve(__dirname, 'OLD MIS.xlsx');
// const workbook = XLSX.readFile(filePath);
// const sheet_name_list = workbook.SheetNames;
// const stocklist = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
//stocklist = JSON.parse(data)

const filename = "Equity_IT.csv";
var originfilename = __dirname + "/originFiles/" + filename;
fs.createReadStream(originfilename)
    .pipe(csv())
    .on('data', (row) => {
        stocklist.push(row)
    })
    .on('end', () => {
        console.log('CSV file successfully processed');
        Object.keys(stocklist[0]).forEach(element => {
            header.push({ 'id': element, 'title': element });
        });
        header.push({ id: 'price', title: 'price' });
        header.push({ id: 'marketcap', title: 'marketcap' });
        header.push({ id: 'Pros', title: 'Pros' });
        header.push({ id: 'Cons', title: 'Cons' });


        console.log(header);
    });
const csvWriter = createCsvWriter({
    path: 'EQUITY_IT_cunsumer_price.csv',
    header: header
});

(async function example1() {
    let driver = await new Builder().forBrowser('chrome').build();
    try {
        let count = 0;
        for (const key in stocklist) {
            count++;
            if (count > 5)
                break;
            if (stocklist.hasOwnProperty(key)) {
                const element = stocklist[key];
                let sercurityname = (element['Security Name']).replace(/\ /g, "-");
                let geturl = url + sercurityname + '/' + element['Security Id'] + '/' + element['Security Code'] + '/';
                console.log(geturl);
                let comobj;
                let isfond = true;
                try {
                    await driver.get(geturl);
                    await driver.getTitle().then((title) => { console.log(title); });
                    await driver.wait(until.elementLocated(By.id('idcrval')), 10000).then(function (elm) {
                    }).catch(function (ex) {
                        console.log(ex);
                        isfond = false;
                        return;
                    });
                 
                    if (isfond) {
                        await (await driver.findElement({ xpath: '//*[@id="idcrval"]' })).getText().then((ele) => {
                            comobj = { ...element, price: ele };
                        }).catch(function (ex) {
                            console.log(ex);
                            return;
                        });
                        await (await driver.findElement({ xpath: '//*[@id="getquoteheader"]/div[6]/div[3]/div/div[3]/div/table/tbody/tr[5]/td[2]' })).getText().then((ele) => {
                            comobj = { ...comobj, marketcap: ele };
                            newjson.push(comobj);
                        }).catch(function (ex) {
                            console.log(ex);
                            return;
                        });
                        await (await driver.findElement({ xpath: '//*[@id="getquoteheader"]/div[6]/div[3]/div/div[3]/div/table/tbody/tr[5]/td[2]' })).getText().then((ele) => {
                            comobj = { ...comobj, marketcap: ele };
                            newjson.push(comobj);
                        }).catch(function (ex) {
                            console.log(ex);
                            return;
                        });
                       
                    }
                } catch (error) {
                    console.log(error);
                    return;
                }
            }
        }
        const newsheet = XLSX.utils.json_to_sheet(newjson);
        const neworkBook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(neworkBook, newsheet);
       // const newfilename = dirpath.resolve(__dirname, '/newFiles/' + new Date() + filename.split('.')[0]);
        var newfilename = __dirname +  '/newFiles/' + new Date() + filename.split('.')[0];
        console.log(originfilename);
        XLSX.writeFile(neworkBook, newfilename);
        console.log("The excel has been created!!!");
        // csvWriter.writeRecords(newjson).then(() => console.log('The CSV file was written successfully'));
        // var json = JSON.stringify(newjson);
        // fs.writeFile('./stock.json', json, 'utf8', () => { console.log('CSV file successfully created.'); });
    } catch (error) {
        console.log(error);
    } finally {
        await driver.quit();
    }
})();