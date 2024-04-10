const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

const workbook = xlsx.readFile('data.xlsx');
const worksheet = workbook.Sheets['Sheet1'];

const columnToExtract = 'A';

const columnData = [];
let rowIndex = 1;

const intervalset = setInterval(() => {
    if (worksheet[`${columnToExtract}${rowIndex}`] !== undefined) {
        columnData.push(worksheet[`${columnToExtract}${rowIndex}`].v);
        sendEmail(columnData[columnData.length-1])
        rowIndex++;
    }else{
        clearInterval(intervalset)
    }    

}, 5000);

const sender_id = 'abc@xyz.com';

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




