
const createCsvWriter = require('csv-writer').createObjectCsvWriter;
const csv = require('csv-parser');
const chrome = require('selenium-webdriver/chrome');
const path = require('chromedriver').path;
const fs = require('fs')
//let jsonData = {}
var service = new chrome.ServiceBuilder(path).build();
const { Builder, By, Key, until } = require('selenium-webdriver');
chrome.setDefaultService(service);
let url = 'https://mortgageloan.indiainfoline.com/mortgage/LoginForm.aspx';
let newjson = [];

let header = [];
const filePath = dirpath.resolve(__dirname, 'OLD MIS.xlsx');
const workbook = XLSX.readFile(filePath);
const sheet_name_list = workbook.SheetNames;
const xlData = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
const stocklist = JSON.parse(xlData);

// const stocklist = [];
// fs.createReadStream('./MC.csv')
//     .pipe(csv())
//     .on('data', (row) => {
//         stocklist.push(row)
//     })
//     .on('end', () => {
//         console.log('CSV file successfully processed');
//         Object.keys(stocklist[0]).forEach(element => {
//             header.push({ 'id': element, 'title': element });
//         });
//         header.push({ id: 'mobile', title: 'mobile' });
//         console.log(header);
//         // console.log(stocklist);
//     });
// const csvWriter = createCsvWriter({
//     path: 'MCwithNumber.csv',
//     header: header
// });

let isfond = true;
(async function example1() {
    let driver = await new Builder().forBrowser('chrome').build();
    await driver.get(url);
    try {
        await driver.wait(until.elementLocated(By.id('txtUserName')), 10000).then(function (elm) {
        }).catch(function (ex) {
            console.log(ex);
            isfond = false;
            return;
        });
        if (isfond) {
            await (await driver.findElement({ xpath: '//*[@id="txtUserName"]' })).sendKeys('C170000').then((ele) => {
            }).catch(function (ex) {
                console.log(ex);
                return;
            });
            await (await driver.findElement({ xpath: '//*[@id="txtPassword"]' })).sendKeys('Mar#2020').then((ele) => {
            }).catch(function (ex) {
                console.log(ex);
                return;
            });
            await (await driver.findElement({ xpath: '//*[@id="btnLogin"]' })).click().then((ele) => {
            }).catch(function (ex) {
                console.log(ex);
                return;
            });
            await driver.wait(until.elementLocated({ xpath: '/html/body/form/div[3]/div[2]/div/div[2]/table[2]/tbody/tr/td[2]/div' }), 10000).then(function (elm) {
            }).catch(function (ex) {
                console.log(ex);
                isfond = false;
                return;
            });
            if (isfond) {
                await (await driver.findElement({ xpath: '//*[@id="select-radio-2"]' })).click().then((ele) => {
                }).catch(function (ex) {
                    console.log(ex);
                    return;
                });
            }
         

            await driver.get('https://mortgageloan.indiainfoline.com/mortgage/ClientExpressEntry.aspx');
        }
    } catch (error) {
        console.log(error);
        return;
    }
    try {
        for (const key in stocklist) {
            if (stocklist.hasOwnProperty(key)) {
                const element = stocklist[key];
                let sercurityname = (element['PROSPECT NO']);
                let comobj;
                let isfond = true;
                try {
                    await driver.getTitle().then((title) => { console.log(title); });
                    await driver.wait(until.elementLocated(By.id('ContentPlaceHolder1_txtProspectNo')), 10000).then(function (elm) {
                    }).catch(function (ex) {
                        console.log(ex);
                        isfond = false;
                        return;
                    });
                    if (isfond) {
                        await (await driver.findElement(By.id('ContentPlaceHolder1_txtProspectNo'))).sendKeys(sercurityname).then(() => {
                        }).catch(function (ex) {
                            console.log(ex);
                            return;
                        });
                        await (await driver.findElement({  xpath: '//*[@id="ContentPlaceHolder1_btnSearchProspect"]' })).click().then(() => {
                        }).catch(function (ex) {
                            console.log(ex);
                            return;
                        });

                        console.log('find click 1');//*[@id="1"]
                        await driver.wait(until.elementLocated(By.id('1')), 10000).then(function (elm) {
                        }).catch(function (ex) {
                            console.log(ex);
                            isfond = false;
                            return;
                        });
                        if (isfond) {
                            console.log('Onclick 1');
                            await driver.findElement(By.id('1')).click().then(function (elm) {
                            }).catch(function (ex) {
                                console.log(ex);
                                isfond = false;
                                return;
                            });
                            console.log('find click 2');
                            await driver.wait(until.elementLocated({ xpath: '/html/body/form/div[5]/div[4]/div[1]/ul/li[2]/a/span' }), 10000).then(function (elm) {
                            }).catch(function (ex) {
                                console.log(ex);
                                isfond = false;
                                return;
                            });
                            console.log('Onclick 1');
                            await driver.findElement(({ xpath: '/html/body/form/div[5]/div[4]/div[1]/ul/li[2]/a/span' })).click().then(function (elm) {
                            }).catch(function (ex) {
                                console.log(ex);
                                isfond = false;
                                return;
                            });
                            console.log('copy mobile');
                            await (await driver.findElement({ xpath: '/html/body/form/div[3]/div[1]/div[1]/fieldset/div[1]/div[5]/div[2]/div[2]/input[1]' })).getText().then((elem) => {
                                console.log(ele);
                                comobj = { ...comobj, mobile: ele };
                                newjson.push(comobj);
                            }).catch(function (ex) {
                                console.log(ex);
                                return;
                            });
                        }

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
        XLSX.writeFile(neworkBook, "BESNEW.xlsx");
        console.log("The excel has been created!!!");
        // csvWriter.writeRecords(newjson).then(() => console.log('The CSV file was written successfully'));
        // var json = JSON.stringify(newjson);
        // fs.writeFile('./MC.json', json, 'utf8', () => { console.log('CSV file successfully created.'); });
    } catch (error) {
        console.log(error);
    } finally {
        await driver.quit();
    }
})();