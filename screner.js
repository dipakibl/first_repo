
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const csv = require('csv-parser');
const chrome = require('selenium-webdriver/chrome');
const path = require('chromedriver').path;
const fs = require('fs');
const Jquery = require("jquery");
const XLSX = require('xlsx');
//let jsonData = {}
var service = new chrome.ServiceBuilder(path).build();
const { Builder, By, Key, until } = require('selenium-webdriver');
chrome.setDefaultService(service);
let url = 'https://www.screener.in/company/';
//https://www.bseindia.com/stock-share-price/infosys-ltd/infy/500209/
let newjson = [];
let stocklist = [];
let header = [];
const filename = "Equity_IT_cunsumer.csv";
fs.createReadStream('./originFiles/Equity_IT_cunsumer.csv')
    .pipe(csv())
    .on('data', (row) => {
        stocklist.push(row)
    })
    .on('end', () => {
        console.log('CSV file successfully processed');
        Object.keys(stocklist[0]).forEach(element => {
            header.push({ 'id': element, 'title': element });
        });
        header.push({ id: 'marketcap', title: 'marketcap' });
        header.push({ id: 'price', title: 'price' });
        header.push({ id: 'Pros', title: 'Pros' });
        header.push({ id: 'Cons', title: 'Cons' });
        //  console.log(header);
    });
const csvWriter = createCsvWriter({
    path: 'EQUITY_IT_cunsumer_price.csv',
    header: header
});

(async function example1() {
    let driver = await new Builder().forBrowser('chrome').build();
    try {
        // let data = await fs.readFileSync('equity.json', 'utf-8');
        // jsonData = JSON.parse(data)
        let count = 0;
        for (const key in stocklist) {
            count++;
            if (count > 5)
                break;
            if (stocklist.hasOwnProperty(key)) {
                const element = stocklist[key];
                //let geturl = url + element.SYMBOL + '/';
                let geturl = url + element['Security Id'] + '/';
                console.log(geturl);
                let comobj;
                let isfond = true;
                try {
                    await driver.get(geturl);
                    await driver.getTitle().then((title) => { console.log(title); });
                    await driver.wait(until.elementLocated(By.id('company-info')), 1000).then(function (elm) {
                    }).catch(function (ex) {
                        //console.log(ex);
                        isfond = false;
                        return;
                    });
                    // function getprocons() {
                    //     try {
                    //         var srt = $('#analysis').textContent.replace(/[\n\r]+|[\s]{2,}/g, ' ').trim().replace('Pros:', '[{"Pros" :"').replace('Cons:', '"},{"Cons":" ') + "\"}]";
                    //         return JSON.parse(srt);
                    //     } catch (error) {
                    //         //console.log(error);
                    //         return JSON.parse('[{"Pros":""},{"Cons":""}]');
                    //     }
                    // }
                   
                    if (isfond) {
                        await (await driver.findElement({ xpath: '//*[@id="content-area"]/section[1]/ul/li[1]/b' })).getText().then((ele) => {
                            comobj = { ...element, marketcap: ele };
                        }).catch(function (ex) {
                            //console.log(ex);
                            return;
                        });
                        await (await driver.findElement({ xpath: '//*[@id="content-area"]/section[1]/ul/li[2]/b' })).getText().then((ele) => {
                            comobj = { ...comobj, price: ele };
                            newjson.push(comobj);
                        }).catch(function (ex) {
                            return;
                        });
                        var myfunction = function () {
                          //  jQuery(function ($) {
                                let srt = Jquery('#analysis').textContent.replace(/[\n\r]+|[\s]{2,}/g, ' ').trim().replace('Pros:', '[{"Pros" :"').replace('Cons:', '"},{"Cons":" ') + "\"}]";
                                return new Promise(JSON.parse(srt));
                           // });
                        }

                        await driver.executeScript(myfunction).then(function (ele) {
                            comobj = { ...comobj, Pros: ele.Pros };
                            comobj = { ...comobj, Cons: ele.Cons };
                            newjson.push(comobj);
                            console.log('Return Value by myfunction -> ' + ele);
                        }).catch(function (ex) {
                            console.log(ex);
                            return;
                        });
                        

                        // await (await driver.executeScript(getprocons)).then((ele) => {
                        //     comobj = { ...comobj, Pros: ele.Pros };
                        //     comobj = { ...comobj, Cons: ele.Cons };
                        //     newjson.push(comobj);
                        // }).catch(function (ex) {
                        //     console.log(ex);
                        //     return;
                        // });
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

        csvWriter.writeRecords(newjson).then(() => console.log('The CSV file was written successfully'));
        var json = JSON.stringify(newjson);
        fs.writeFile('./newFiles/stock.json', json, 'utf8', () => { console.log('CSV file successfully created.'); });
    } catch (error) {
        console.log(error);
    } finally {
        await driver.quit();
    }
})();