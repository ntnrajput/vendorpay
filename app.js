const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const numberToWords = require('number-to-words');
const { Console } = require('console');

const app = express();
const port = 3000;

app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});

app.get('/directory.html', (req, res) => {
    res.sendFile(__dirname + '/directory.html');
});




app.post('/launch', async (req, res) => {
    
    const websiteURL = 'https://vendinvtracking.rites.com:81/(S(rmlveb55xstdnj55ungzxv45))/index.aspx'
    const {
        trackingID, diary, diarydate, dummy, amount
    } = req.body;
    
    const txt_dummy = "dummy-" + dummy
    const format_date = diarydate.split('-').reverse().join('/');

    // data received from index.html

    try {
        const browser = await puppeteer.launch({
            headless: false,
            defaultViewport: false,
            executablePath:"c:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
        });
        
        const [page] = await browser.pages();
        await page.goto(websiteURL);
        // await page.waitForTimeout(1000);
        await page.waitForSelector('#txtUsername', { timeout: 10000 });

        await page.type('#txtUsername', '12191');
        await page.type('#txtPwd', '12191@123');
        await page.keyboard.press('Enter');        
        //-----------Login Done---------------        
        await page.waitForNavigation({ waitUntil: 'networkidle2' });

        await page.goto ('https://vendinvtracking.rites.com:81/(S(rmlveb55xstdnj55ungzxv45))/Anjani/InvoiceStatus.aspx')

        await page.screenshot({ path: 'screenshot.png' })
 

        await page.waitForSelector('#ctl00_ContentPlaceHolder2_txtTrackingId')
        await page.type('#ctl00_ContentPlaceHolder2_txtTrackingId',trackingID)
        await page.keyboard.press('Tab')
        await page.waitForTimeout(1000)

        const invoicevalue = await page.$eval('#ctl00_ContentPlaceHolder2_txtInvAmt', input => input.value);

        
        const error = Math.abs(((invoicevalue-amount)/invoicevalue) * 100 )

        if(error>0.1){
            divide(10,0)
        }else{
            await page.select('select#ctl00_ContentPlaceHolder2_ddInvStatus', '0');
            await page.click('#ctl00_ContentPlaceHolder2_BtnSubmit')
            await page.waitForTimeout(2000)
            await page.waitForSelector('#ctl00_ContentPlaceHolder2_txtTrackingId')
            await page.$eval('#ctl00_ContentPlaceHolder2_txtTrackingId', input => (input.value = ''));
            await page.type('#ctl00_ContentPlaceHolder2_txtTrackingId',trackingID)
            await page.keyboard.press('Tab')
            await page.waitForTimeout(1000)
            await page.select('select#ctl00_ContentPlaceHolder2_ddInvStatus', '10');
            await page.waitForTimeout(1000)
            await page.click('#ctl00_ContentPlaceHolder2_BtnSubmit')
            await page.waitForTimeout(2000)
            await page.waitForSelector('#ctl00_ContentPlaceHolder2_txtTrackingId')
            await page.$eval('#ctl00_ContentPlaceHolder2_txtTrackingId', input => (input.value = ''));
            await page.type('#ctl00_ContentPlaceHolder2_txtTrackingId',trackingID)
            await page.keyboard.press('Tab')
           
            await page.waitForTimeout(1000)
            await page.type('#ctl00_ContentPlaceHolder2_txtDiaryEntryNo',diary)
            
            await page.waitForTimeout(1000)

            console.log(format_date)
            await page.type('#ctl00_ContentPlaceHolder2_txtSESDate',format_date)
            await page.waitForTimeout(1000)
            
            await page.type('#ctl00_ContentPlaceHolder2_txtSESNo',txt_dummy)
            
            await page.waitForTimeout(1000)
            await page.select('select#ctl00_ContentPlaceHolder2_ddInvStatus', '3');

            await page.waitForTimeout(1000)

            await page.click('#ctl00_ContentPlaceHolder2_BtnSubmit')
            





        }



            
        
        await page.waitForTimeout(2000);
        await browser.close();
        res.send(` Step  Completed `);        
    
    } catch (error) {
        res.status(500).send(`Error completing step `);
    }

});



async function ie_login(CaseNumber,format_call_date,Consignee_Code, Book, Set, Off_Qty,Rem_Qty,txt_qty) {  
    
    const call_serial_num = '2'
    const call_serial_num_txt = "SNO="+ call_serial_num

    console.log(CaseNumber,"hello case")

    const browser = await puppeteer.launch({
        headless: false,
        defaultViewport: false,
        executablePath:"c:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    });
    const [page] = await browser.pages();
    await page.goto("https://www.ritesinsp.com/RBS/IE_Login_Form.aspx");

    await page.waitForTimeout(2000);
    await page.type('#txtUserId','88888');
    await page.type('#txtPwd','88888');
    await page.keyboard.press('Enter');
    await page.waitForTimeout(2000);

    const allPages = await browser.pages();
    const pass_page1 = allPages.find((page) => page.url() === "https://www.ritesinsp.com/RBS/IE_Instructions.aspx?pIECD=908&pIENAME=CR/BHILAI");
    pass_page1.click('#btnNext')

    await page.waitForTimeout(2000)

    const allPages1 = await browser.pages();
    const ie_login_menu = allPages1.find((page) => page.url() === "https://www.ritesinsp.com/RBS/IE_Menu.aspx");

    await page.waitForTimeout(1000)   
    

    await hoverAndClick(ie_login_menu, 'Calls Status', 'Calls Status', 'Sorted on Call Date');

    await page.waitForTimeout(3000)
    

    const allPages2 = await browser.pages();
    const call_pending = allPages2.find((page) => page.url() === "https://www.ritesinsp.com/RBS/Calls_Marked_To_IE.aspx?ACTION=C");
    
    await call_pending.screenshot({ path: 'screenshot.png' });

    await page.waitForTimeout(3000)
    
    const call_table = await call_pending.$('body > table:nth-child(2)');
    
    const rows = await call_table.$$('tr');
    
    let shouldBreak = false; // Flag variable to control the outer loop
    let row_matching = null;

    for (let rowIndex = 1; rowIndex < rows.length && !shouldBreak; rowIndex++) {
        const row = rows[rowIndex];
        const columns = await row.$$('td'); // Get all columns in the row

        const row_data = [];

        for (const column of columns) {
            const text = await column.evaluate((cell) => cell.textContent); // Get the text content of the column
            row_data.push(text);
        }
        

        if (row_data[12].includes(CaseNumber) && row_data[14].includes("Pending")) {
            shouldBreak = true; // Set the flag to true to break out of both loops 
            row_matching = rowIndex+1; 
        }
    }
    
    
    const rowSelector = `body > table:nth-child(2) > tbody > tr:nth-child(${row_matching})`;   
    const rowHandle = await call_pending.$(rowSelector);
    const linkSelector = `a:first-child`; // Assuming the link is in the first column
    const linkHandle = await rowHandle.$(linkSelector);
    await linkHandle.click();

    
    await page.waitForTimeout(2000);
    let call_status_edit;
    for (const page of allPages) {
        const pageUrl = page.url();            
        if(pageUrl.includes(CaseNumber)){
            call_status_edit = page  
            break;                
        }                    
    } 

    
    await call_status_edit.select('select#lstCallStatus', 'A');
    await call_status_edit.waitForSelector(`#HyperLink2AR`)
    await call_status_edit.click(`#HyperLink2AR`);

    await page.waitForTimeout(3000)
    let call_status_edit_form;
    for (const page of allPages) {
        const pageUrl = page.url();            
        if(pageUrl.includes(CaseNumber)){
            call_status_edit_form= page  
            break;                
        }                    
    } 
    
    await call_status_edit_form.waitForSelector(`#ddlCondignee`)
    await call_status_edit_form.select('select#ddlCondignee', Consignee_Code)

    await page.waitForTimeout(3000)
    
    // ---------skipped for testing----------------

    await call_status_edit_form.type('#txtBKNO1',Book)
    await call_status_edit_form.type('#txtSetNo1',Set)


    const fileInput = await call_status_edit_form.$('#File4'); 
    await fileInput.uploadFile('D:\\IC PHOTO\\1.jpg');
    const fileInput1 = await call_status_edit_form.$('#File5'); 
    await fileInput1.uploadFile('D:\IC PHOTO\\2.jpg');
    const fileInput2 = await call_status_edit_form.$('#File6'); 
    await fileInput2.uploadFile('D:\IC PHOTO\\3.jpg');
    const fileInput3 = await call_status_edit_form.$('#File7'); 
    await fileInput3.uploadFile('D:\IC PHOTO\\4.jpg');
    const fileInput4 = await call_status_edit_form.$('#File8'); 
    await fileInput4.uploadFile('D:\IC PHOTO\\5.jpg');     

    await call_status_edit_form.click(`#btnSaveFiles`);

    // ---------skipped for testing----------------

    await page.waitForTimeout(3000)

    await call_status_edit_form.waitForSelector(`#btnViewIC`);
    await call_status_edit_form.click(`#btnViewIC`);

    await page.waitForTimeout(2000)

    let ic_report;
    for (const page of allPages) {
        const pageUrl = page.url();            
        if(pageUrl.includes(CaseNumber)){
            ic_report= page  
            break;                
        }             
               
    } 
    await ic_report.screenshot({ path: 'screenshot.png' });

    await page.waitForTimeout(3000)

    await ic_report.waitForSelector(`#TxtItemRemarkeh`)
    await ic_report.type(`#TxtItemRemarkeh`,txt_qty)

    await ic_report.waitForTimeout(3000)

    await ic_report.type(`#TxtQTY_PASSED`,Off_Qty)
    await ic_report.type(`#TxtQTY_REJECTED`,'0')
    await ic_report.type(`#TxtQTY_DUE`,Rem_Qty)
    await ic_report.type(`#txtHologram`,"HARD STEEL PUNCH ON ONE END FACE OF EACH RAIL")
    await ic_report.type(`#txtHologram`,"HARD STEEL PUNCH ON ONE END FACE OF EACH RAIL")
    
    await ic_report.type(`#btnSaveICData`)


    await browser.close();
    return;
}
  
//allied functions

async function hoverAndClick(page, hoverText1, hoverText2, clickText) {
    await page.waitForTimeout(500)
    const hoverDiv1 = await page.$x(`//div[contains(text(), '${hoverText1}')]`);
    await hoverDiv1[0].hover();
    await page.waitForTimeout(500)
    const hoverDiv2 = await page.$x(`//div[contains(text(), '${hoverText2}')]`);
    await hoverDiv2[0].hover();
    await page.waitForTimeout(500)
    const clickDiv = await page.$x(`//div[contains(text(), '${clickText}')]`);
    await clickDiv[0].click();
}

app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});



function IC_Description(section, grade, Raillen, railclass, rake) {
    
    let description = null; // Use 'let' instead of 'const' to allow reassignment
    let list_desc_num = null; 
    let PL_No = null; 
    
    if(section==="60E1"){
        PL_No='1'
    }else{PL_No='2'}

    if (Raillen === "260m" && section === '60E1') {        
        description = "(PRIORITY PROGRAMME -01 , RAKE NO. " + rake + "   )   1)  60 E1 R-260 GRADE RAILS (260M) WITH 100% ULTRASONICALLY TESTED SATISFYING THE REQUIREMENTS OF IRS SPECIFICATION NO. IRS-T-12-2009 CL-A PRIME QUALITY RAILS WITH LATEST AMENDMENTS 2) ALL FLASH BUTT WELDED RAIL JOINTS AND THEIR USFD TESTING ARE SATISFYING THE REQUIREMENTS OF IRFBWM 2012 WITH LATEST AMENDMENTS";
        list_desc_num = '8'
        
    }else if(Raillen === "260m" && section === 'IRS52'){
        description = "(PRIORITY PROGRAMME -02 , RAKE NO. " + rake + "   )   1)  IRS-52  R-260 GRADE RAILS (260M) WITH 100% ULTRASONICALLY TESTED SATISFYING THE REQUIREMENTS OF IRS SPECIFICATION NO. IRS-T-12-2009 CL-A PRIME QUALITY RAILS WITH LATEST AMENDMENTS 2) ALL FLASH BUTT WELDED RAIL JOINTS AND THEIR USFD TESTING ARE SATISFYING THE REQUIREMENTS OF IRFBWM 2012 WITH LATEST AMENDMENTS";
        list_desc_num = '7'
    }else if(section === "60E1" && grade === "R260" && Raillen == "26m" && railclass =="A"){
        description = " 60 E1 R-260 GRADE RAILS (26M) WITH 100% ULTRASONICALLY TESTED SATISFYING THE REQUIREMENTS OF IRS SPECIFICATION NO. IRS-T-12-2009 CL-A PRIME QUALITY RAILS WITH LATEST AMENDMENTS";
        list_desc_num = '4'
    } else if(section === "60E1" && grade === "R260" && Raillen == "26m" && railclass =="B"){
        description = " 60 E1 R-260 GRADE RAILS (26M) WITH 100% ULTRASONICALLY TESTED SATISFYING THE REQUIREMENTS OF IRS SPECIFICATION NO. IRS-T-12-2009 CL-B PRIME QUALITY RAILS WITH LATEST AMENDMENTS";
        list_desc_num = '4'
    }else if(section === "60E1" && grade === "R260" && Raillen == "13m" && railclass =="A"){
        description = "60 E1 R-260 GRADE RAILS (13M) WITH 100% ULTRASONICALLY TESTED SATISFYING THE REQUIREMENTS OF IRS SPECIFICATION NO. IRS-T-12-2009 CL-A PRIME QUALITY RAILS WITH LATEST AMENDMENTS";
        list_desc_num = '2'
    }else if(section === "60E1" && grade === "R260" && Raillen == "13m" && railclass =="B"){
        description = "60 E1 R-260 GRADE RAILS (13M) WITH 100% ULTRASONICALLY TESTED SATISFYING THE REQUIREMENTS OF IRS SPECIFICATION NO. IRS-T-12-2009 CL-B PRIME QUALITY RAILS WITH LATEST AMENDMENTS";
        list_desc_num = '2'
    }else if(section === "60E1" && grade === "880" && Raillen == "26m" && railclass =="A"){
        description = "60e1 880 26m cl A";
        list_desc_num = '4'
    } else if(section === "60E1" && grade === "880" && Raillen == "26m" && railclass =="B"){
        description = "60e1 880 26m cl B";
        list_desc_num = '4'
    }else if(section === "60E1" && grade === "880" && Raillen == "13m" && railclass =="A"){
        description = "60e1 880 13m cl A";
        list_desc_num = '2'
    }else if(section === "60E1" && grade === "880" && Raillen == "13m" && railclass =="B"){
        description = " 60e1 880 13m cl B";
        list_desc_num = '2'
    }else if(section === "IRS52" && grade === "R260" && Raillen == "26m" && railclass =="A"){
        description = "irs52 r260 26m cl A";
        list_desc_num = '3'
    } else if(section === "IRS52" && grade === "R260" && Raillen == "26m" && railclass =="B"){
        description = " irs5 r260 26m cl B";
        list_desc_num = '3'
    }else if(section === "IRS52" && grade === "R260" && Raillen == "13m" && railclass =="A"){
        description = "irs52 r260 13m cl A";
        list_desc_num = '1'
    }else if(section === "IRS52" && grade === "R260" && Raillen == "13m" && railclass =="B"){
        description = " irs52 r260 13m cl B";
        list_desc_num = '1'
    }else if(section === "IRS52" && grade === "880" && Raillen == "26m" && railclass =="A"){
        description = " 52KG (13M) GR 880 RAILS TO IRS SPECIFICATION NO. IRS T-12-2009 CL A PRIME QUALITY RAILS 100% ULTRASONICALLY TESTED & FOUND SATISFACTORY";
        list_desc_num = '3'
    } else if(section === "IRS52" && grade === "880" && Raillen == "26m" && railclass =="B"){
        description = "52KG (13M) GR 880 RAILS TO IRS SPECIFICATION NO. IRS T-12-2009 CL B PRIME QUALITY RAILS 100% ULTRASONICALLY TESTED & FOUND SATISFACTORY";
        list_desc_num = '3'
    }else if(section === "6IRS52" && grade === "880" && Raillen == "13m" && railclass =="A"){
        description = " 52KG (13M) GR 880 RAILS TO IRS SPECIFICATION NO. IRS T-12-2009 CL B PRIME QUALITY RAILS 100% ULTRASONICALLY TESTED & FOUND SATISFACTORY";
        list_desc_num = '1'
    }else if(section === "IRS52" && grade === "880" && Raillen == "13m" && railclass =="B"){
        description = "irs52 880 13m cl B";
        list_desc_num = '1'
    }

    console.log(description)
    return [description,list_desc_num,PL_No]; // Return the description at the end
}

function capitalizeFirstLetter(word) {
    return word.charAt(0).toUpperCase() + word.slice(1);
}
  
function convertToText(number) {
    if (isNaN(number)) {
      return "Invalid input";
    }
  
    const wholePart = Math.floor(number);
    const fractionalPart = number - wholePart;
  
    const wholeText = numberToWords.toWords(wholePart).split(' ').map(capitalizeFirstLetter).join(' ');
  
    // Convert the fractional part to text with exactly four decimal places
    const fractionalText = (Math.round(fractionalPart * 10000)).toString().padStart(4, '0').slice(0, 4);
  
    if (fractionalText === "0000") {
      return `${wholeText} point zero zero zero zero`;
    } else {
      const fractionalWords = Array.from(fractionalText).map(digit => numberToWords.toWords(parseInt(digit))).join(' ');
      // Use a specific regular expression to remove the dash between "Eighty" and "Nine"
      const formattedFractionalWords = fractionalWords.replace(/\bEighty-\b/g, 'Eighty Nine').replace(/ and /g, ' ').replace(/\b\w/g, l => l.toUpperCase());
      return `${wholeText} point ${formattedFractionalWords}`;
    }
}
  
 
  
  
