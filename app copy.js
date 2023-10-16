const express = require('express');
const puppeteer = require('puppeteer');
const path = require('path');
const numberToWords = require('number-to-words');

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
    const websiteURL = 'https://www.ritesinsp.com/rbs/Login_Form.aspx'
    const {
        CaseNumber, PODate, PONumber, calldate, section,
        grade, Raillen, railclass, rake, PO_Qty,
        Rate, Consignee_Code, BPO_Code, f_s, irfc, Cumm_Pass_Qty, Off_Qty, Book, Set
    } = req.body;
    const Rem_Qty = PO_Qty -Cumm_Pass_Qty-Off_Qty;
    const format_call_date = calldate.split('-').reverse().join('-');
    const txt_qty_half = await convertToText(Off_Qty,4)
    const txt_qty = "QUANTITY NOW PASSED & DISPATCHED - " + txt_qty_half + " MT ONLY."
    console.log(txt_qty)

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
        await page.waitForSelector('#txtUname', { timeout: 10000 });

        await page.type('#txtUname', 'CRTECH');
        await page.type('#txtPwd', 'BSPINSCR');
        await page.keyboard.press('Enter');

        //-----------Login Done---------------
        
        await page.waitForNavigation({ waitUntil: 'networkidle2' });

        const allPages = await browser.pages();    
        let Main_Page = null;
        for (const page of allPages) {  
            const pageUrl = page.url();
            if(pageUrl === "https://www.ritesinsp.com/rbs/MainForm.aspx?Role=1"){
                Main_Page = page                
            }
        }     

        await hoverAndClick(Main_Page, 'TRANSACTIONS', 'Inspection & Billing', 'Purchase Order Form');
       
        await page.waitForTimeout(3000);
        let PO_Page = null
        for (const page of allPages) {
            const pageUrl = page.url();
            // console.log(pageUrl)
            if(pageUrl === "https://www.ritesinsp.com/rbs/PurchesOrder1_Form.aspx"){
                PO_Page = page
            }
        }

        await PO_Page.type('#txtCsNo',CaseNumber);
        await PO_Page.click('#btnSearchPO');
        
        await PO_Page.waitForSelector('a'); // Wait for at least one anchor element
        
        // Use a selector to locate the anchor tag link you want to click
        
        const linkSelector = `a[href*="PurchesOrder_Form"][href*="${CaseNumber}"]`;

        // const linkSelector = 'a[href="PurchesOrder_Form.aspx?CASE_NO=${variableValue}"]'; // Replace with the href of the specific link you want to click
        
        await PO_Page.waitForSelector(linkSelector);
        await PO_Page.click(linkSelector);       
        
        await page.waitForTimeout(2000);

        //-----------Purchase Order Form Opened

        let PO_Page_Case = null;
        const PO_Page_Case_url = `https://www.ritesinsp.com/rbs/PurchesOrder_Form.aspx?CASE_NO=${CaseNumber}`;
       
        for (const page of allPages) {
            const pageUrl = page.url();
            // console.log(pageUrl)
            if(pageUrl === PO_Page_Case_url){
                PO_Page_Case = page
            }
        }

        const PO_Date_Box = await PO_Page_Case.$('#txtPOdate');
        const PO_Date = await PO_Page_Case.evaluate(input => input.value, PO_Date_Box);
        const tableId = 'grdCB'
        const table = await PO_Page_Case.$(`#${tableId}`);
        
        const rowCount = await PO_Page_Case.evaluate(table => {
            const rows = table.querySelectorAll('tr');
            return rows.length;
        }, table);

        //-----------Going Through all Consingee already added, if cosignee present, opening PO Details

        for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
            const columns = await PO_Page_Case.evaluate((table, rowIndex) => {
              const row = table.querySelectorAll('tr')[rowIndex];
              const columns = Array.from(row.querySelectorAll('td')).map(column => column.textContent.trim());
              return columns;
            }, table, rowIndex);
            const regex = new RegExp(Consignee_Code);
            if (regex.test(columns[1])) {                
                await PO_Page_Case.click('#btnPODetails'); 
                break;
            }
        }

        await PO_Page_Case.waitForTimeout(3000);
        
        // console.log(PO_Details_url)
        for (const page of allPages) {
            const pageUrl = page.url();
            
            if(pageUrl.includes(CaseNumber)){
                PO_Details = page  
                break;                
            }                    
        }

        await page.waitForTimeout(3000); // time for pressing okay on page

        //-----------Getting Description from the function as per details Entered---------------------
        const [description,list_desc_num,PL_No] = IC_Description(section, grade, Raillen, railclass);        

        await PO_Details.select('select#lstItemDesc', list_desc_num);

        //-----------Giving Time to press OK-------------
        
        await page.waitForTimeout(5000);
        await PO_Details.type(`#txtItemDescpt`, description);
        await PO_Details.type(`#txtPLNO`,PL_No);
        await PO_Details.select('select#ddlConsigneeCD', Consignee_Code);
        await PO_Details.type(`#txtQty`,PO_Qty);
        await PO_Details.keyboard.press('Tab');
        await PO_Details.type(`#txtRate`,Rate);
        for (let i = 0; i < 6; i++) {
            await PO_Details.keyboard.press('Tab');
        }
        await PO_Details.type (`#txtSaleTaxPre`,'18');
        for (let i = 0; i < 6; i++) {
            await PO_Details.keyboard.press('Tab');
        }
        await PO_Details.type (`#txtExtDelvDate`,'31-03-2024'); 

        const outerTable = await PO_Details.$('table#Table1');
        const innerTable = await outerTable.$('table#DgPO');

       
        const rows = await innerTable.$$('tr');
        const ic_made = (rows.length);
        
        
        
        await PO_Details.click('#btnSave');
        await page.waitForTimeout(5000);


        //-----------Part Created-------------

        {
        // //Accessing last Part and deleting it.. till it goes real

        // const last_part = `a[id*="DgPO_ctl"][id*="${ic_made + 1}"][id*="Hyperlink2"]`;   
        
        // await PO_Details.click(last_part);    
        // let PO_Part_Page = null;
        // await page.waitForTimeout(2000);

        // for (const page of allPages) {
        //     const pageUrl = page.url();            
        //     if(pageUrl.includes(CaseNumber)){
        //         PO_Part_Page = page  
        //         break;                
        //     }                    
        // } 
        // await page.waitForTimeout(2000);
        // await PO_Part_Page.click(`#btnDelete`);

        // //Accessing last Part and deleting it.. till it goes real

        // await page.waitForTimeout(4000);

        // await PO_Part_Page.click(`#WebUserControl11_HyperLink1`) // Change PO_Part_page to PO_details when deleting part is commented
        }

        //-----------Going Back to Main Page

        await PO_Details.click(`#WebUserControl11_HyperLink1`)

        await page.waitForTimeout(2000);

        

        await hoverAndClick(Main_Page, 'TRANSACTIONS', 'Inspection & Billing', 'Call Registration/Cancellation');

        const call_page_url = 'https://www.ritesinsp.com/rbs/Call_Register_Edit.aspx'
        let call_page = null;
        await page.waitForTimeout(2000);

        for (const page of allPages) {
            const pageUrl = page.url();            
            if(pageUrl === call_page_url){
                call_page = page  
                break;                
            }                    
        } 

        
        await call_page.type(`#txtCaseNo`,CaseNumber)
        await call_page.type(`#txtDtOfReciept`,format_call_date)

        
        await call_page.click(`#btnAdd`);

        let new_call_page = null;

        await page.waitForTimeout(2000);

        for (const page of allPages) {
            const pageUrl = page.url();            
            if(pageUrl.includes(CaseNumber)){
                new_call_page = page  
                break;                
            }                    
        } 
        
        
        await new_call_page.select('select#lstIE', '557');
        await new_call_page.select('select#ddlDept', 'C');
        if (f_s === 'f') {
            await new_call_page.click('#rdbFinal');
          } else if (f_s === 's') {
            await new_call_page.click('#rdbStage');
        }
        await page.waitForTimeout(500);

        const irfcSelect = await new_call_page.$('select#ddlIRFC');
        if (irfcSelect) {
        
            if(irfc === 'funded'){
                await new_call_page.select('select#ddlIRFC', 'Y');
            }else{
                await new_call_page.select('select#ddlIRFC', 'N');
            }
        }
        await new_call_page.type('#txtMName','45338') ;  
        await new_call_page.click('#btnFCList');
        await page.waitForTimeout(1000)
        await new_call_page.click (`#btnSave`); 
        await page.waitForTimeout(3000)
        await new_call_page.click (`#btnCDetails`);

        // -------------------------------------------------------------------------
        // -------------------------------------------------------------------------
        

        // // // -------------This code only testing Purpose-------------

        // // await call_page.click(`#btnSearch`);
        // // await call_page.waitForSelector('#grdCNO_ctl02_Hyperlink2');
        // // await call_page.click('#grdCNO_ctl02_Hyperlink2');
        // // await call_page.waitForSelector('#btnMod');
        // // await call_page.click(`#btnMod`);
        // // await page.waitForTimeout(1000)
        // // const allPages11 = await browser.pages();
        // // const call_mod = allPages11.find((page) => page.url() === "https://www.ritesinsp.com/rbs/Call_Register_Form.aspx?Action=M&Case_No=C19040066&DT_RECIEPT=27/09/2023&CALL_SNO=2");
        // // await page.waitForTimeout(1000)
        // // await call_mod.waitForSelector('#btnCDetails')
        // // await call_mod.click(`#btnCDetails`)
        

        // // // --------------------This code onyl for Testing Purpose---------------

        let call_details = null;

        await page.waitForTimeout(2000);

        for (const page of allPages) {
            const pageUrl = page.url();            
            if(pageUrl.includes(CaseNumber)){
                call_details = page  
                break;                
            }                    
        } 
        

        // Assuming the outer table has an id 'outerTableId' and the inner table has an id 'innerTableId'
        const outerTableSelector = 'table#Table1';
        const innerTableSelector = 'table#grdCDetails';

        // Wait for the outer table to appear
        await call_details.waitForSelector(outerTableSelector);

        // Get the inner table within the outer table
        await call_details.waitForSelector(innerTableSelector, { visible: true, timeout: 0 });

        // Get the last row in the inner table
        const lastRowSelector = `${innerTableSelector} tr:last-child`;

        // Get the link in the first column of the last row of the inner table
        const linkSelector1 = `${lastRowSelector} td:first-child a`;

        // Click the link
        await call_details.click(linkSelector1);


        await call_details.waitForSelector(`#txtQuanInsp`);
        await call_details.type(`#txtCQty`,Cumm_Pass_Qty);
        await call_details.type(`#txtQPrePassed`,Cumm_Pass_Qty);
        await call_details.type(`#txtQuanInsp`,Off_Qty);
        await call_details.click(`#btnSave`);

        page.waitForTimeout(3000)
        


        await call_details.click(`#WebUserControl11_HyperLink2`)

        await page.waitForTimeout(3000)
        

      
        await ie_login(CaseNumber,format_call_date,Consignee_Code, Book, Set, Off_Qty,Rem_Qty,txt_qty)


        await page.waitForTimeout(2000);
        await browser.close();
        res.send(`Launched ${CaseNumber} `);
    } catch (error) {
        res.status(500).send(`Error launching ${websiteURL}`);
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



function IC_Description(section, grade, Raillen, railclass) {
    let description = null; // Use 'let' instead of 'const' to allow reassignment
    let list_desc_num = null; 
    let PL_No = null; 
    
    if(section==="60E1"){
        PL_No='1'
    }else{PL_No='2'}

    if (Raillen === "260m" && section === '60E1') {        
        description = "(PRIORITY PROGRAMME- , RAKE NO. )   1)  60 E1 R-260 GRADE RAILS (260M) WITH 100% ULTRASONICALLY TESTED SATISFYING THE REQUIREMENTS OF IRS SPECIFICATION NO. IRS-T-12-2009 CL-A PRIME QUALITY RAILS WITH LATEST AMENDMENTS 2) ALL FLASH BUTT WELDED RAIL JOINTS AND THEIR USFD TESTING ARE SATISFYING THE REQUIREMENTS OF IRFBWM 2012 WITH LATEST AMENDMENTS";
        list_desc_num = '8'
        
    }else if(Raillen === "260m" && section === 'IRS52'){
        description = " 52 kg rake";
        list_desc_num = '7'
    }else if(section === "60E1" && grade === "R260" && Raillen == "26m" && railclass =="A"){
        description = " 60e1 r260 26m cl A";
        list_desc_num = '4'
    } else if(section === "60E1" && grade === "R260" && Raillen == "26m" && railclass =="B"){
        description = " 60e1 r260 26m cl B";
        list_desc_num = '4'
    }else if(section === "60E1" && grade === "R260" && Raillen == "13m" && railclass =="A"){
        description = "60e1 r260 13m cl A";
        list_desc_num = '2'
    }else if(section === "60E1" && grade === "R260" && Raillen == "13m" && railclass =="B"){
        description = "60e1 r260 13m cl B";
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
        description = " irs52 880 26m cl A";
        list_desc_num = '3'
    } else if(section === "IRS52" && grade === "880" && Raillen == "26m" && railclass =="B"){
        description = "irs52 880 26m cl B";
        list_desc_num = '3'
    }else if(section === "6IRS52" && grade === "880" && Raillen == "13m" && railclass =="A"){
        description = " 52KG (13M) GR 880 RAILS TO IRS SPECIFICATION NO. IRS T-12-2009 CL B PRIME QUALITY RAILS 100% ULTRASONICALLY TESTED & FOUND SATISFACTORY";
        list_desc_num = '1'
    }else if(section === "IRS52" && grade === "880" && Raillen == "13m" && railclass =="B"){
        description = "irs52 880 13m cl B";
        list_desc_num = '1'
    }
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
  
  console.log(convertToText(989.39)); // Output: "Nine Hundred Eighty Nine point Three Nine Two Zero"
  
  
