const path = require('path');
const puppeteer = require('puppeteer');
const LibraryFunctions = require('./main/libraryFunctions');
const lib = new LibraryFunctions();
const filePath = path.join(__dirname,'ResourceRequirements.xlsx');
let linkArr = [];

const fileData = lib.getRowsBySheetName('Main',filePath);
for (let i = 0; i< fileData.length; i++){
    const totObj = Object.values(fileData)[i];
    if(totObj.Link){
        linkArr.push(totObj.Link)
    }
}

(async () => {
    console.log("Execution Has Started");
    let valueArr= [];
    const browser = await puppeteer.launch();
    const page = await browser.newPage();
    for (link of linkArr){
        await page.goto(link,{waitUntil: 'load'});
        let [element] = await page.$x("//span[@id='last_last']");
        let text = await page.evaluate(element => element.textContent, element);
        valueArr.push(text)
    }
    await lib.updateSheetWithCurrentValue(filePath,valueArr);
    await browser.close();
    console.log('Data Updated Successfully');
})();
