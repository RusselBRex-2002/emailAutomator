const fs = require('fs');
const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

const workbook = xlsx.readFile('data.xlsx');
const worksheet = workbook.Sheets['Export'];

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
        rowIndex++;
    }else{
        clearInterval(intervalset)
    }    

}, 2000);
