const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');

const crypto = require('crypto');
require('dotenv').config();


// Manejador para cerrar el bot limpiamente
process.on('SIGINT', async () => {
    console.log('🛑 Cerrando el bot...');
    await client.destroy();
    process.exit(0);
});

const client = new Client({
  authStrategy: new LocalAuth({
    dataPath: './session_data',
    clientId: "bot-esquina-food"
  }),
  puppeteer: { 
    headless: true,
    args: ['--no-sandbox']
  }
});
let botStartTime = Math.floor(Date.now() / 1000);
const usuarios = {};

const sinonimos = {
    "coca": "coca cola",
    "coca-cola": "coca cola",
    "empanada": "empanada",
    "empanadas de jamon queso": "empanada de jamón y queso",
    "empanadas de jyq": "empanada de jamón y queso",
    "jyq": "jamón y queso",
    "empanadas de carne": "empanada de carne",
    "empanada de jamon y queso": "empanada de jamón y queso",
    "hamburguesa doble": "hamburguesa doble carne",
    "papas fritas": "papas",
    "pizza muza": "pizza de muzzarella",
    "muzarella": "muzzarella",
    "misiles": "misil",
    "miciles": "misil",
    "dosena": "docena",
    "dosenas": "docenas",
    "empanadas": "empanada",

};



// Configuración de la encriptación
require('dotenv').config();
const ALGORITHM = 'aes-256-cbc';
const SECRET_KEY = process.env.SECRET_KEY; // Cambia esto por una clave secreta de 32 caracteres
const IV_LENGTH = 16; // Longitud del vector de inicialización para AES

// Función para encriptar la fecha
function encryptDate(date) {
    const iv = crypto.randomBytes(IV_LENGTH);
    const cipher = crypto.createCipheriv(ALGORITHM, Buffer.from(SECRET_KEY), iv);
    let encrypted = cipher.update(date, 'utf8', 'hex');
    encrypted += cipher.final('hex');
    return {
        iv: iv.toString('hex'),
        encryptedData: encrypted
    };
}

// Función para desencriptar la fecha
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

// Función para leer la suscripción
const SUBSCRIPTION_FILE = path.join(__dirname, 'subscription.json');

function readSubscription() {
    try {
        if (!fs.existsSync(SUBSCRIPTION_FILE)) {
            console.error('❌ Archivo subscription.json no encontrado. Creando uno con suscripción vencida.');
            const defaultDate = new Date(0).toISOString();
            const encryptedDefault = encryptDate(defaultDate);
            const defaultSubscription = { expiration: encryptedDefault };
            fs.writeFileSync(SUBSCRIPTION_FILE, JSON.stringify(defaultSubscription, null, 2));
            return defaultDate;
        }
        const data = fs.readFileSync(SUBSCRIPTION_FILE, 'utf8');
        const subscription = JSON.parse(data);
        const decryptedDate = decryptDate(subscription.expiration);
        return decryptedDate || new Date(0).toISOString();
    } catch (error) {
        console.error('❌ Error al leer subscription.json:', error);
        return new Date(0).toISOString();
    }
}

// Función para verificar si la suscripción está activa
function isSubscriptionActive() {
    const expirationDateStr = readSubscription();
    const expirationDate = new Date(expirationDateStr);
    const now = new Date();
    return expirationDate > now;
}

// Función para actualizar la fecha de vencimiento (solo administrador)
function updateSubscription(newExpirationDate, adminNumber) {
    try {
        const dateStr = new Date(newExpirationDate).toISOString();
        const encryptedDate = encryptDate(dateStr);
        const newSubscription = { expiration: encryptedDate };
        fs.writeFileSync(SUBSCRIPTION_FILE, JSON.stringify(newSubscription, null, 2));
        console.log(`✅ Suscripción actualizada hasta ${newExpirationDate}`);
        client.sendMessage(adminNumber, `✅ Suscripción renovada hasta ${newExpirationDate}`);
    } catch (error) {
        console.error('❌ Error al actualizar subscription.json:', error);
        client.sendMessage(adminNumber, '❌ Error al renovar la suscripción. Contactá al soporte.');
    }
}

// Verificación periódica de la suscripción (cada hora)
function startSubscriptionCheck() {
    setInterval(() => {
        if (!isSubscriptionActive()) {
            console.error('❌ La suscripción ha vencido. Deteniendo el bot.');
            client.sendMessage(process.env.ADMIN_NUMBER, '❌ La suscripción ha vencido. Por favor, renovala con !renovar YYYY-MM-DD');
            process.exit(1);
        }
    }, 3600000);
}



function normalizarTexto(texto) {
    let limpio = texto.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().trim();
    for (let clave in sinonimos) {
        const regex = new RegExp(`\\b${clave}\\b`, 'g');
        limpio = limpio.replace(regex, sinonimos[clave]);
    }
    return limpio;
}

function cargarMenuDesdeExcel() {
    const rutaMenu = path.join(__dirname, 'menu.xlsx');
    if (!fs.existsSync(rutaMenu)) {
        console.warn('⚠️ El archivo menu.xlsx no fue encontrado. Se usará un menú vacío.');
        return { productos: {}, categorias: {} };
    }

    try {
        const workbook = xlsx.readFile(rutaMenu);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const data = xlsx.utils.sheet_to_json(sheet);

        const productos = {};
        const categorias = {};

        data.forEach(row => {
            const nombre = row.producto?.toLowerCase().trim();
            const precio = parseInt(row.precio);
            if (nombre && !isNaN(precio)) {
                productos[nombre] = precio;
                const categoria = nombre.split(' de ')[0];
                if (!categorias[categoria]) categorias[categoria] = [];
                categorias[categoria].push({ nombre, precio });
            }
        });

        return { productos, categorias };
    } catch (error) {
        console.error('❌ Error al leer menu.xlsx:', error);
        return { productos: {}, categorias: {} };
    }
}

let menu = cargarMenuDesdeExcel();

const numerosTexto = {
    'una': 1, 'un': 1, 'uno': 1,
    'dos': 2, 'tres': 3, 'cuatro': 4, 'cinco': 5,
    'seis': 6, 'siete': 7, 'ocho': 8, 'nueve': 9, 'diez': 10, 'once': 11, 'doce': 12,
    'trece': 13, 'catorce': 14, 'quince': 15, 'dieciseis': 16, 'diecisiete': 17, 'dieciocho': 18, 'diecinueve': 19, 'veinte': 20,
    'una docena': 12, 'una docena de': 12, 'una docena y media': 18, 'una docena y media de': 18, 'media docena': 6, 'media docena de': 6,
    'dos docenas': 24, 'dos docenas de': 24, 'dos docenas y media': 30, 'dos docenas y media de': 30,
    'tres docenas': 36, 'tres docenas de': 36, 'cuatro docenas': 48, 'cuatro docenas de': 48,
    '2 docenas': 24, '2 docenas de': 24, '2 docenas y media': 30, '2 docenas y media de': 30,
    '3 docenas': 36, '3 docenas de': 36, '4 docenas': 48, '4 docenas de': 48
};

function convertirTextoACantidad(texto) {
    if (!texto) return 1;
    texto = texto.toLowerCase().trim();

    // Ordenar frases largas primero para evitar coincidencias parciales
    const frasesOrdenadas = Object.keys(numerosTexto).sort((a, b) => b.length - a.length);

    for (let frase of frasesOrdenadas) {
        if (texto.includes(frase)) {
            return numerosTexto[frase];
        }
    }

    // Intentar parsear número directo si no se encontró ninguna coincidencia
    const numero = parseInt(texto);
    return isNaN(numero) ? 1 : numero;
}


client.on('qr', (qr) => {
    console.log('📲 Escaneá este QR con tu WhatsApp:');
    qrcode.generate(qr, { small: true });
});

client.on('ready', () => {
    console.log('✅ Bot conectado correctamente a WhatsApp');
});

client.on('disconnected', async (reason) => {
    console.log('⚠️ Sesión desconectada:', reason);
    setTimeout(async () => {
        try {
            await client.initialize();
            console.log('🔄 Reconexión exitosa');
        } catch (error) {
            console.error('❌ Error al intentar reconectar:', error);
        }
    }, 5000); // Retraso de 5 segundos
});

function leerConfig() {
    try {
        const data = fs.readFileSync('config.txt', 'utf8');
        const config = {};
        data.split('\n').forEach(line => {
            const [key, value] = line.split('=');
            if (key && value) config[key.trim()] = value.trim();
        });
        console.log('Configuración cargada:', config);
        return config;
    } catch (err) {
        console.error('Error al cargar config.txt:', err);
        return { alias: 'LaEsquinaFood' };
    }
}

client.on('message', async (message) => {
    const chat = await message.getChat();
    if (chat.isGroup) return;
    if (message.timestamp < botStartTime) return;

    const from = message.from;
    const texto = normalizarTexto(message.body);
    const numeroLimpio = from.replace('@c.us', '');
    console.log(`Mensaje recibido de ${from}: "${texto}", esperandoConfirmacion: ${usuarios[from]?.esperandoConfirmacion || false}, esperandoOpcionEntrega: ${usuarios[from]?.esperandoOpcionEntrega || false}`);

    // Verificar si la suscripción está activa
    if (!isSubscriptionActive()) {
        console.log(`❌ Suscripción vencida. Modo restringido activado. Mensaje recibido de ${from}: ${texto}`);
        if (from === process.env.ADMIN_NUMBER && texto.startsWith('!renovar')) {
            console.log(`Comando !renovar recibido. Procesando fecha: ${texto.split(' ')[1]}`);
            const newDate = texto.split(' ')[1];
            if (!newDate || isNaN(Date.parse(newDate))) {
                console.log(`Fecha inválida: ${newDate}`);
                client.sendMessage(from, '⚠️ Formato inválido. Usa: !renovar YYYY-MM-DD');
                return;
            }
            updateSubscription(newDate, from);
        } else {
            console.log(`Acceso denegado: ${from} no es administrador o mensaje inválido`);
            client.sendMessage(from, '❌ La suscripción ha vencido. Solo el administrador puede renovarla con !renovar YYYY-MM-DD.');
        }
        return;
    }

    console.log('✅ Suscripción activa. Procesando mensaje:', texto);
    menu = cargarMenuDesdeExcel();
    const config = leerConfig();

    console.log('Texto normalizado:', texto);
    console.log('Productos en menú:', Object.keys(menu.productos));

    if (texto.startsWith('!renovar')) {
        console.log(`Comando !renovar recibido. from: ${from}, ADMIN_NUMBER: ${process.env.ADMIN_NUMBER}`);
        if (from === process.env.ADMIN_NUMBER) {
            console.log(`Comando válido. Procesando fecha: ${texto.split(' ')[1]}`);
            const newDate = texto.split(' ')[1];
            if (!newDate || isNaN(Date.parse(newDate))) {
                console.log(`Fecha inválida: ${newDate}`);
                client.sendMessage(from, '⚠️ Formato inválido. Usa: !renovar YYYY-MM-DD');
                return;
            }
            updateSubscription(newDate, from);
        } else {
            console.log(`Acceso denegado: ${from} no es administrador`);
            client.sendMessage(from, '❌ Solo el administrador puede usar este comando.');
        }
        return;
    }

    if (!usuarios[from]) {
        usuarios[from] = {
            saludoEnviado: false,
            pedido: [],
            total: 0,
            esperandoConfirmacion: false,
            esperandoOpcionEntrega: false,
            esperandoNombre: false,
            esperandoDireccion: false,
            esperandoMetodoPago: false,
            esperandoComprobante: false,
            esperandoHorario: false, // Nuevo estado
            nombre: '',
            direccion: '',
            metodoPago: '',
            tipoEntrega: '',
            horarioEntrega: '', // Nuevo campo para almacenar el horario
            ultimoProductoConsultado: null,
            mostrandoResumen: false
        };
    }

    const user = usuarios[from];
    const { productos, categorias } = menu;

    // Manejo de la opción de entrega (retiro o envío)
    if (user.esperandoOpcionEntrega) {
        console.log(`Opción de entrega esperada. Texto recibido: "${texto}"`);
        const retiroPhrases = ["retirar", "retiro", "en el local", "voy a buscar"];
        const envioPhrases = ["envío", "envio", "a domicilio", "delivery"];
        if (retiroPhrases.some(phrase => texto.includes(phrase))) {
            console.log(`Opción detectada: retiro`);
            user.tipoEntrega = 'retiro';
            user.esperandoOpcionEntrega = false;
            user.esperandoNombre = true;
            client.sendMessage(from, '🙋‍♂️ ¿Podés decirme tu *nombre o apellido* para el pedido?');
            return;
        } else if (envioPhrases.some(phrase => texto.includes(phrase))) {
            console.log(`Opción detectada: envío`);
            user.tipoEntrega = 'envio';
            user.esperandoOpcionEntrega = false;
            user.esperandoNombre = true;
            client.sendMessage(from, '🙋‍♂️ ¿Podés decirme tu *nombre o apellido* para el pedido?');
            return;
        } else {
            console.log(`Respuesta no válida para opción de entrega: "${texto}"`);
            client.sendMessage(from, '🧐 No entendí. Por favor, responde con "retirar" o "envío".');
            return;
        }
    }

    // Manejo de confirmación
    if (user.esperandoConfirmacion) {
        console.log(`Confirmación esperada. Texto recibido: "${texto}"`);
        const confirmPhrases = ["quiero confirmar", "nada mas", "eso es todo", "confirmar", "esta bien"];
        if (confirmPhrases.some(phrase => texto === phrase || texto.includes(phrase))) {
            console.log(`Confirmación detectada: "${texto}" coincide con alguna frase válida`);
            user.esperandoConfirmacion = false;
            user.esperandoOpcionEntrega = true;
            client.sendMessage(from, '📍 ¿Querés retirar el pedido en el local o preferís que lo enviemos a tu dirección? Responde con "retirar" o "envío".');
            return;
        } else if (["no", "cancelar", "cancelo", "cancelame"].some(phrase => texto === phrase || texto.includes(phrase))) {
            console.log(`Cancelación detectada: "${texto}"`);
            client.sendMessage(from, '❌ Pedido cancelado. Si querés volver a pedir, escribí el producto.');
            delete usuarios[from];
            return;
        } else {
            console.log(`Respuesta no válida para confirmación: "${texto}"`);
            client.sendMessage(from, '🧐 No entendí. ¿Querés confirmar el pedido? Responde con "quiero confirmar", "nada más", "eso es todo", "confirmar" o "está bien".');
            return;
        }
    }

    if (user.esperandoNombre) {
        user.nombre = message.body.trim();
        if (user.tipoEntrega === 'envio') {
            user.esperandoNombre = false;
            user.esperandoDireccion = true;
            client.sendMessage(from, '📍 ¿Cuál es tu *dirección* para el envío?');
        } else {
            user.esperandoNombre = false;
            user.esperandoMetodoPago = true;
            client.sendMessage(from, '💳 ¿Cómo vas a pagar? Escribe *efectivo* o *transferencia*.');
        }
        return;
    }

    if (user.esperandoDireccion) {
        user.direccion = message.body.trim();
        user.esperandoDireccion = false;
        user.esperandoMetodoPago = true;
        client.sendMessage(from, '💳 ¿Cómo vas a pagar? Escribe *efectivo* o *transferencia*.');
        return;
    }

    if (user.esperandoMetodoPago) {
        const metodo = texto.trim().toLowerCase();
        if (['efectivo', 'cash', 'en efectivo'].includes(metodo)) {
            user.metodoPago = 'efectivo';
            user.esperandoMetodoPago = false;
            user.esperandoHorario = true; // Activar estado para esperar horario
            client.sendMessage(from, '⏰ Tu pedido estará listo en aproximadamente *45 minutos*. ¿Querés que lo preparemos para este tiempo o preferís un *horario específico*?');
            return;

        } else if (['transferencia', 'transf', 'por transferencia', 'te transfiero', 'ahi te transfiero', 'ahí te transfiero', 'x transferencia', 'transferir', 'te voy a transferir', 'x transf'].includes(metodo)) {
            user.metodoPago = 'transferencia';
            user.esperandoMetodoPago = false;
            user.esperandoComprobante = true;
            client.sendMessage(from, `💸 Por favor, realizá la transferencia al alias: *${config.alias}*\nLuego, enviá el comprobante de pago (imagen o archivo).`);
            return;
        } else {
            client.sendMessage(from, '🧐 No entendí. Por favor, escribe *efectivo* o *transferencia*.');
            return;
        }
    }

    if (user.esperandoComprobante) {
        if (message.hasMedia) {
            console.log('Comprobante recibido para:', from);
            const media = await message.downloadMedia();
            if (!media) {
                client.sendMessage(from, '❌ Error al descargar el comprobante. Por favor, enviálo nuevamente.');
                return;
            }

            user.esperandoComprobante = false;
            user.esperandoHorario = true; // Activar estado para esperar horario
            client.sendMessage(from, '⏰ Tu pedido estará listo en aproximadamente *45 minutos*. ¿Querés que lo preparemos para este tiempo o preferís un *horario específico*?');
            return;
        } else {
            client.sendMessage(from, '📸 Por favor, enviá el comprobante de pago (imagen o archivo).');
            return;
        }
    }


    if (user.esperandoHorario) {
        const respuestaHorario = texto.trim().toLowerCase();
        console.log(`Horario recibido: "${respuestaHorario}"`);

        // Lista de frases que confirman los 45 minutos
        const confirmHorarioPhrases = [
            "me parece bien", "está bien", "esta bien", "no hay problema", "perfecto",
            "ok", "sí", "si", "dale", "todo bien"
        ];

        // Expresión regular para detectar horarios específicos y frases naturales
        const horarioRegex = /^(?:.*?(?:para las|a las|me pod[ée]s preparar para las|quiero que est[ée] para las|preparalo para las|preparame para las|puede ser para las|me pueden hacer para las|me preparan para las)\s*)?(\d{1,2})(?::(\d{2}))?(?:\s*(?:horas|hs|h)?(?:\s*por favor)?)?(?:\s*[\?!]?)?$/i;

        if (confirmHorarioPhrases.some(phrase => respuestaHorario.includes(phrase))) {
            user.horarioEntrega = '45 minutos';
        } else {
            const matchHorario = respuestaHorario.match(horarioRegex);
            if (matchHorario || ['en 45 minutos', '45 minutos', 'ahora', 'listo'].some(h => respuestaHorario.includes(h))) {
                if (['en 45 minutos', '45 minutos', 'ahora', 'listo'].some(h => respuestaHorario.includes(h))) {
                    user.horarioEntrega = '45 minutos';
                } else {
                    const hora = parseInt(matchHorario[1]);
                    const minutos = matchHorario[2] ? matchHorario[2] : '00'; // Si no hay minutos, usar "00"
                    if (hora >= 0 && hora <= 23) {
                        user.horarioEntrega = `${hora.toString().padStart(2, '0')}:${minutos}`;
                    } else {
                        client.sendMessage(from, '🧐 El horario ingresado no es válido. Por favor, especificá un horario como "22", "22:00", "puede ser para las 22" o "en 45 minutos", o confirmá con "está bien" o "perfecto".');
                        return;
                    }
                }
            } else {
                client.sendMessage(from, '🧐 No entendí. Por favor, especificá un horario como "22", "22:00", "puede ser para las 22" o "en 45 minutos", o confirmá con "está bien" o "perfecto".');
                return;
            }
        }

        user.esperandoHorario = false;

        // Generar el resumen del pedido
        const resumen = `🛎️ *Nuevo pedido confirmado (${user.tipoEntrega === 'retiro' ? 'retiro en local' : 'envío'})*\n` +
                        `👤 Nombre: ${user.nombre}\n` +
                        `${user.tipoEntrega === 'envio' ? `📍 Dirección: ${user.direccion}\n` : ''}` +
                        `📱 Número: ${numeroLimpio}\n` +
                        `💳 Método de pago: ${user.metodoPago.charAt(0).toUpperCase() + user.metodoPago.slice(1)}\n` +
                        `⏰ Horario: ${user.horarioEntrega === '45 minutos' ? 'En 45 minutos' : `Para las ${user.horarioEntrega}`}\n` +
                        `${user.pedido.join('\n')}\n` +
                        `💵 Total: $${user.total}`;

        // Mensaje al cliente
        const mensajeCliente = user.tipoEntrega === 'retiro'
            ? `✅ ¡Pedido confirmado para retirar en La Esquina Food! Estará listo ${user.horarioEntrega === '45 minutos' ? 'en aproximadamente 45 minutos' : `para las ${user.horarioEntrega}`}. Te esperamos en Av. Ejemplo 123. 🍽️`
            : `✅ ¡Pedido confirmado para envío a ${user.direccion}! Estará listo ${user.horarioEntrega === '45 minutos' ? 'en aproximadamente 45 minutos' : `para las ${user.horarioEntrega}`}. 🍽️`;

        client.sendMessage(from, mensajeCliente);

        // Guardar el pedido
        fs.appendFileSync('pedidos.txt', `[${new Date().toLocaleString()}] ${resumen.replace(/\n/g, ' | ')}\n`);

        // Enviar al grupo
        const chats = await client.getChats();
        const grupo = chats.find(c => c.isGroup && c.name === 'Bot_chat');
        if (grupo) {
            await grupo.sendMessage(resumen);
            if (user.metodoPago === 'transferencia') {
                const media = await message.downloadMedia();
                if (media) {
                    await grupo.sendMessage(new MessageMedia(media.mimetype, media.data, `comprobante_${numeroLimpio}_${Date.now()}.png`));
                }
            }
        }

        delete usuarios[from];
        return;
    }

    const palabrasMenu = [
        'menu', 'menú', 'lista de precios', 'carta', 'que tenes', 'qué tenés', 'quiero ver el menu', 'quiero ver el menú',
        'quiero ver la carta', 'quiero ver la lista de precios', 'quiero ver la lista de productos', 'pasame el menu',
        'pasame el menú', 'mandame el menu', 'mandame el menú', 'me pasas el menu', 'me pasas el menú', 'me mandas el menu',
        'me mandas el menú', 'me podes pasar el menu', 'me podes pasar el menú', 'me puedes pasar el menu', 'me puedes pasar el menú',
        'me podés pasar el menu', 'me podés pasar el menú', 'me puedes pasar el menu', 'me puedes pasar el menú',
        'que hay', 'qué hay', 'lista de productos', 'ver menu', 'ver menú', 'mostrame el menu', 'mostrame el menú',
        'me puede pasar el menu', 'me puede pasar el menú', 'pasame la lista deprecios', 'pasame la lista de precios',
        'pasame la lista de presio', 'pasame la lista de precio'
    ];

    if (palabrasMenu.some(p => texto.includes(p))) {
        const rutaPDF = path.join(__dirname, 'menu.pdf');
        if (fs.existsSync(rutaPDF)) {
            try {
                const media = MessageMedia.fromFilePath(rutaPDF);
                await client.sendMessage(from, media, { caption: '📋 Aquí tenés el menú de *La Esquina Food*. ¡Echale un vistazo! 😋' });
            } catch (error) {
                console.error('❌ Error al enviar el PDF:', error);
                client.sendMessage(from, '❌ Hubo un error al enviar el menú. Por favor, intentá de nuevo más tarde.');
            }
        } else {
            console.warn('⚠️ El archivo menu.pdf no fue encontrado en la carpeta assets.');
            client.sendMessage(from, '❌ No encontramos el menú en PDF. Podés pedir detalles de los productos escribiendo, por ejemplo, "cuánto sale la hamburguesa".');
        }
        return;
    }

    const palabrasPreguntas = [
        'cuanto', 'cuánto', 'precio', 'sale', 'cuesta', 'vale', 'valor', 'cuanto esta', 'cuanto está', 'cuánto esta', 'cuánto está',
        'cuanto sale', 'cuanto cuesta', 'cuanto vale', 'cuánto sale', 'cuánto cuesta', 'cuánto vale'
    ];

    if (palabrasPreguntas.some(p => texto.includes(p))) {
        let productosEncontrados = [];
        for (let nombre in productos) {
            if (texto.includes(nombre)) {
                productosEncontrados.push(`💡 *${nombre}*: $${productos[nombre]}`);
                user.ultimoProductoConsultado = nombre;
                console.log('Producto consultado guardado:', nombre);
            }
        }
        if (productosEncontrados.length > 0) {
            client.sendMessage(from, productosEncontrados.join('\n'));
            return;
        } else {
            client.sendMessage(from, '🤔 No encontré el producto que mencionaste. ¿Podés especificar mejor? Por ejemplo: "cuánto sale la hamburguesa"');
            return;
        }
    }

    const palabrasTotal = [
        'total', 'cuanto es', 'cuánto es', 'resumen', 'cuanto debo', 'cuánto debo', 'cuanto te debo', 'cuánto te debo',
        'total pedido', 'cuánto es el total', 'cuanto seria', 'cuanto sería', 'cual es el total', 'cuanto sale todo', 'cuánto sale todo'
    ];

    if (palabrasTotal.some(p => texto.includes(p))) {
        console.log(`Procesando total. Pedido: ${user.pedido.length > 0 ? user.pedido.join(', ') : 'vacío'}, Estado esperandoConfirmacion: ${user.esperandoConfirmacion}`);
        if (user.pedido.length > 0) {
            client.sendMessage(from, `🧾 Tu pedido actual es:\n${user.pedido.join('\n')}\n\n💵 Total: $${user.total}\n¿Querés confirmar o agregar algo más?`);
            user.esperandoConfirmacion = true;
            console.log(`Estado actualizado: esperandoConfirmacion = ${user.esperandoConfirmacion}`);
            return;
        } else {
            client.sendMessage(from, '🛒 No tenés ningún pedido en curso. ¿Qué querés pedir?');
            return;
        }
    }

    if (["no", "nono", "eso es todo", "ya está", "ya esta", "nada más", "nada mas"].some(f => texto.includes(f))) {
        if (user.pedido.length > 0 && !user.esperandoConfirmacion && !user.esperandoOpcionEntrega && !user.esperandoNombre && !user.esperandoDireccion && !user.esperandoMetodoPago && !user.esperandoComprobante) {
            client.sendMessage(from, `🧾 Perfecto. Tu pedido es:\n${user.pedido.join('\n')}\n\n💵 Total: $${user.total}\n¿Querés confirmar o agregar algo más?`);
            user.esperandoConfirmacion = true;
            console.log(`Estado actualizado: esperandoConfirmacion = ${user.esperandoConfirmacion}`);
            return;
        }
    }

    const nombresProductosRegex = Object.keys(menu.productos)
        .map(p => p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'))
        .sort((a, b) => b.length - a.length)
        .join('|');

    const cantidadSinProductoRegex = new RegExp(`(?:quiero|dame|pedime|traeme|puede ser)\\s*(?:(\\d+|una|un|uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|once|doce|trece|catorce|quince|dieciseis|diecisiete|dieciocho|diecinueve|veinte|una docena|una docena de|una docena y media|una docena y media de|media docena|media docena de|dos docenas|dos docenas de|dos docenas y media|dos docenas y media de|tres docenas|tres docenas de|cuatro docenas|cuatro docenas de|2 docenas|2 docenas de|2 docenas y media|2 docenas y media de|3 docenas|3 docenas de|4 docenas|4 docenas de))?\\s*$`, 'i');

    let agregado = false;
    let nuevosPedidos = [];

    let matchSinProducto = texto.match(cantidadSinProductoRegex);
    if (matchSinProducto && user.ultimoProductoConsultado && productos[user.ultimoProductoConsultado]) {
        console.log('Procesando pedido sin producto explícito. Usando último consultado:', user.ultimoProductoConsultado);
        let cantidad = 1;
        if (matchSinProducto[1]) {
            cantidad = convertirTextoACantidad(matchSinProducto[1]);
            if (isNaN(cantidad)) {
                console.log('Error: Cantidad no válida:', matchSinProducto[1]);
                cantidad = 1;
            }
        }
        const productoBase = user.ultimoProductoConsultado;
        const precio = productos[productoBase] * cantidad;
        user.total += precio;
        const pedidoLinea = `🛍️ ${cantidad} ${productoBase}(s) - $${precio}`;
        user.pedido.push(pedidoLinea);
        nuevosPedidos.push(pedidoLinea);
        console.log('Pedido agregado:', pedidoLinea);
        console.log('Estado actual del pedido:', user.pedido, 'Total:', user.total);
        agregado = true;
    }

    if (!agregado) {
        const cantidadRegex = new RegExp(
            `(?:quiero|dame|pedime|traeme|puede ser)?\\s*` +
            `(?:(una docena y media|dos docenas y media|una docena|media docena|\\d+ docenas|\\d+|una|un|uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|once|doce|trece|catorce|quince|dieciseis|diecisiete|dieciocho|diecinueve|veinte)\\s*(?:de)?)?\\s*` +
            `(${nombresProductosRegex})` +
            `(?:\\s|$|s)`, 'ig'
        );

        console.log('Procesando mensaje completo:', texto);

        let matches = [...texto.matchAll(cantidadRegex)];
        console.log('Coincidencias encontradas:', matches);

        for (let match of matches) {
            console.log('Match encontrado:', match);
            console.log('Cantidad capturada:', match[1], 'Producto:', match[2]);
            let cantidad = 1;
            if (match[1]) {
                cantidad = convertirTextoACantidad(match[1]);
                if (isNaN(cantidad)) {
                    console.log('Error: Cantidad no válida:', match[1]);
                    cantidad = 1;
                }
            }
            let productoBase = match[2].toLowerCase();
            console.log('Buscando producto en menú:', productoBase);

            if (productos[productoBase]) {
                const precio = productos[productoBase] * cantidad;
                user.total += precio;
                const pedidoLinea = `🛍️ ${cantidad} ${productoBase}(s) - $${precio}`;
                user.pedido.push(pedidoLinea);
                nuevosPedidos.push(pedidoLinea);
                console.log('Pedido agregado:', pedidoLinea);
                console.log('Estado actual del pedido:', user.pedido, 'Total:', user.total);
                agregado = true;
            } else {
                const categoria = productoBase.split(' de ')[0];
                if (categorias[categoria]) {
                    let respuesta = `📦 Variantes de *${categoria}*:\n`;
                    categorias[categoria].forEach(item => {
                        respuesta += `💡 *${item.nombre}*: $${item.precio}\n`;
                    });
                    respuesta += `\nPor favor, especificá qué tipo de ${categoria} querés (por ejemplo, "${categoria} de carne").`;
                    client.sendMessage(from, respuesta);
                    agregado = true;
                    return;
                }
            }
        }
    }

    if (agregado) {
        client.sendMessage(from, `✅ Producto(s) agregado(s):\n${nuevosPedidos.join('\n')}\n¿Querés algo más?`);
        return;
    }

    let productosEncontrados = [];
    for (let categoria in categorias) {
        if (texto.includes(categoria) && !texto.includes('docena') && !texto.includes('docenas')) {
            productosEncontrados.push(`📦 Variantes de *${categoria}*:`);
            categorias[categoria].forEach(item => {
                productosEncontrados.push(`💡 *${item.nombre}*: $${item.precio}`);
            });
            client.sendMessage(from, productosEncontrados.join('\n'));
            return;
        }
    }

    if (!user.saludoEnviado) {
        client.sendMessage(from, '👋 ¡Hola! Bienvenido al bot de *La Esquina Food*. ¿Qué querés pedir hoy?');
        user.saludoEnviado = true;
        return;
    }

    client.sendMessage(from, '🤔 No entendí. Podés escribir "quiero una hamburguesa y dos pizzas", "cuánto sale la hamburguesa" o "total" para ver tu pedido.');
});


// Verificar suscripción antes de inicializar
if (!isSubscriptionActive()) {
    console.warn('⚠️ La suscripción ha vencido. El bot está en modo restringido. Solo el administrador puede usar !renovar.');
    // No hacemos process.exit(1) para permitir el modo restringido
}


// Iniciar verificación periódica
startSubscriptionCheck();

client.initialize();

