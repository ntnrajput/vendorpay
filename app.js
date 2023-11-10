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
    
    const websiteURL = 'https://vptp.rites.com:81/(S(tbq3tgmhyrh5xs2dbtj31yz0))/index.aspx'
    const {
        trackingID, diary, diarydate, dummy, amount,url,
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
        await page.waitForTimeout(3000)

        await page.goto (url)

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




app.listen(port, () => {
    console.log(`Server is running on http://localhost:${port}`);
});



 
  
  
