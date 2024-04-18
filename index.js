const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

const workbook = xlsx.readFile('data.xlsx');
const worksheet = workbook.Sheets['Sheet1'];

const f = 'F';
const s = 'S';
const y = 'Y';
const as = 'AS';



const columnData = [];
const technical_service = []
const instance_name = []
const instance_account = []
const expiry_date = []

let rowIndex = 1;

const intervalset = setInterval(() => {
    if (worksheet[`${y}${rowIndex}`] !== undefined) {
        instance_account.push(worksheet[`${y}${rowIndex}`].v);
        technical_service.push(worksheet[`${f}${rowIndex}`].v);
        instance_name.push(worksheet[`${s}${rowIndex}`].v);
        expiry_date.push(worksheet[`${as}${rowIndex}`].w);
        const lastIndex = expiry_date.length - 1
        const daysLeft = 90
        const expiryDate = new Date(expiry_date[lastIndex]);
        const currentDate = new Date();
        const differenceInDays = Math.floor((currentDate - expiryDate) / (1000 * 60 * 60 * 24));
        if (differenceInDays <= daysLeft) {
            console.log(`${instance_name[lastIndex]}, your password is due for ${daysLeft}. Other details, technical service: ${technical_service[lastIndex]}, instance account: ${instance_account[lastIndex]}
            `);
        } 
        // sendEmail(columnData[columnData.length-1])
        rowIndex++;
    }else{
        clearInterval(intervalset)
    }    

}, 2000);

// const sender_id = 'abc@xyz.com';

function sendEmail(receiver) {

    const transporter = nodemailer.createTransport({
        service: 'Outlook',
        auth: {
            user: sender_id,
            pass: 'XXXXXXXX'
        }
    });

    const missedIds = []

    fs.readFile('emailContent.txt', 'utf8', function(err, data) {
        if (err) {
            console.error('Error reading file:', err);
            return;
        }
    
        // Assign the content of the text file to the 'text' property of 'mailOptions'
        const mailOptions = {
            from: sender_id,
            to: receiver,
            subject: 'Password expired',
            text: data 
        };
    
        transporter.sendMail(mailOptions, function (error, info) {
            if (error) {
                console.log(error);
                missedIds.push(receiver);
                console.log("Missed Ids: " + missedIds);
            } else {
                console.log('Email sent: ' + info.response);
                console.log(receiver);
            }
    
            transporter.close();
        });
    });

    
}
