const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ALGORITHM = 'aes-256-cbc';
const SECRET_KEY = 'jeremiascomesbernachea1234567890'; // Usa la misma clave que en init_subscription.js
const IV_LENGTH = 16;

// Validar SECRET_KEY
if (!SECRET_KEY || SECRET_KEY.length !== 32) {
    console.error('❌ SECRET_KEY debe tener exactamente 32 caracteres.');
    process.exit(1);
}

function decryptDate(encryptedObj) {
    try {
        const iv = Buffer.from(encryptedObj.iv, 'hex');
        const encryptedText = Buffer.from(encryptedObj.encryptedData, 'hex');
        const decipher = crypto.createDecipheriv(ALGORITHM, Buffer.from(SECRET_KEY), iv);
        let decrypted = decipher.update(encryptedText, 'hex', 'utf8');
        decrypted += decipher.final('utf8');
        return decrypted;
    } catch (error) {
        console.error('❌ Error al desencriptar la fecha:', error);
        return null;
    }
}

// Leer subscription.json
try {
    const subscription = JSON.parse(fs.readFileSync(path.join(__dirname, 'subscription.json'), 'utf8'));
    const decryptedDate = decryptDate(subscription.expiration);

    if (decryptedDate) {
        console.log('Fecha de vencimiento:', decryptedDate);
        console.log('Fecha legible:', new Date(decryptedDate).toLocaleString());
    } else {
        console.log('No se pudo desencriptar la fecha. Verifica que SECRET_KEY sea correcta.');
    }
} catch (error) {
    console.error('❌ Error al leer subscription.json:', error);
}