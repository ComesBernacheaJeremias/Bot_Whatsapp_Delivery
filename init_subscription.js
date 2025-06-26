const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ALGORITHM = 'aes-256-cbc';
const SECRET_KEY = "jeremiascomesbernachea1234567890"; 
const IV_LENGTH = 16;

function encryptDate(date) {
    const iv = crypto.randomBytes(IV_LENGTH);
    const cipher = crypto.createCipheriv(ALGORITHM, Buffer.from(SECRET_KEY), iv);
    let encrypted = cipher.update(date, 'utf8', 'hex');
    encrypted += cipher.final('hex');
    return { iv: iv.toString('hex'), encryptedData: encrypted };
}

const date = '2025-07-25T23:59:59.999Z';
const encrypted = encryptDate(date);
fs.writeFileSync('subscription.json', JSON.stringify({ expiration: encrypted }, null, 2));
console.log('subscription.json creado.');