const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');

const client = new Client({ authStrategy: new LocalAuth() });
let botStartTime = Math.floor(Date.now() / 1000);
const usuarios = {};

const sinonimos = {
    "coca": "coca cola",
    "coca-cola": "coca cola",
    "empanadas jamon queso": "empanada de jamón y queso",
    "empanadas de jyq": "empanada de jamón y queso",
    "jyq": "jamón y queso",
    "empanadas carne": "empanada de carne",
    "empanada de jamon y queso": "empanada de jamón y queso",
    "hamburguesa doble": "hamburguesa doble carne",
    "papas fritas": "papas",
    "pizza muza": "pizza de muzzarella",
    "muzarella": "muzzarella",
    "misiles": "misil",
    "miciles": "misil",
    "dosena": "docena",
    "dosenas": "docenas"
};

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
    let texto = normalizarTexto(message.body);
    const numeroLimpio = from.replace('@c.us', '');
    menu = cargarMenuDesdeExcel();
    const config = leerConfig();

    console.log('Texto normalizado:', texto);
    console.log('Productos en menú:', Object.keys(menu.productos));

    if (!usuarios[from]) {
        usuarios[from] = {
            saludoEnviado: false,
            pedido: [],
            total: 0,
            esperandoConfirmacion: false,
            esperandoNombre: false,
            esperandoDireccion: false,
            esperandoMetodoPago: false,
            esperandoComprobante: false,
            nombre: '',
            direccion: '',
            metodoPago: '',
            ultimoProductoConsultado: null,
            mostrandoResumen: false
        };
    }

    const user = usuarios[from];
    const { productos, categorias } = menu;

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

    const palabrasTotal = [
        'total', 'cuanto es', 'cuánto es', 'resumen', 'cuanto debo', 'cuánto debo', 'cuanto te debo', 'cuánto te debo',
        'total pedido', 'cuánto es el total', 'cuanto seria', 'cuanto sería', 'cual es el total', 'cuanto sale todo', 'cuánto sale todo'
    ];

    if (["no", "nono", "eso es todo", "ya está", "ya esta", "nada más", "nada mas"].some(f => texto.includes(f))) {
        if (user.pedido.length > 0 && !user.esperandoConfirmacion && !user.esperandoNombre && !user.esperandoDireccion && !user.esperandoMetodoPago && !user.esperandoComprobante) {
            client.sendMessage(from, `🧾 Perfecto. Tu pedido es:\n${user.pedido.join('\n')}\n\n💵 Total: $${user.total}\n¿Confirmás el pedido? (sí/no)`);
            user.esperandoConfirmacion = true;
            return;
        }
    }

    if (user.esperandoConfirmacion) {
        if (["si", "sí", "confirmar", "confirmo"].includes(texto)) {
            client.sendMessage(from, '🙋‍♂️ ¿Podés decirme tu *nombre o apellido*?');
            user.esperandoConfirmacion = false;
            user.esperandoNombre = true;
            return;
        } else if (["no", "cancelar", "cancelo", "cancelame"].includes(texto)) {
            client.sendMessage(from, '❌ Pedido cancelado. Si querés volver a pedir, escribí el producto.');
            delete usuarios[from];
            return;
        } else {
            client.sendMessage(from, '🧐 No entendí. ¿Confirmás el pedido? (sí/no)');
            return;
        }
    }

    if (user.esperandoNombre) {
        user.nombre = message.body.trim();
        client.sendMessage(from, '📍 ¿Cuál es tu *dirección* para el envío?');
        user.esperandoNombre = false;
        user.esperandoDireccion = true;
        return;
    }

    if (user.esperandoDireccion) {
        user.direccion = message.body.trim();
        client.sendMessage(from, '💳 ¿Cómo vas a pagar? Escribe *efectivo* o *transferencia*.');
        user.esperandoDireccion = false;
        user.esperandoMetodoPago = true;
        return;
    }

    if (user.esperandoMetodoPago) {
        const metodo = texto.trim().toLowerCase();
        if (['efectivo', 'cash', 'en efectivo'].includes(metodo)) {
            user.metodoPago = 'efectivo';
            user.esperandoMetodoPago = false;

            const resumen = `🛎️ *Nuevo pedido confirmado*\n👤 Nombre: ${user.nombre}\n📍 Dirección: ${user.direccion}\n📱 Número: ${numeroLimpio}\n💳 Método de pago: Efectivo\n${user.pedido.join('\n')}\n💵 Total: $${user.total}`;

            client.sendMessage(from, '✅ ¡Pedido confirmado! Lo estamos preparando 🍽️');

            fs.appendFileSync('pedidos.txt', `[${new Date().toLocaleString()}] ${resumen.replace(/\n/g, ' | ')}\n`);

            const chats = await client.getChats();
            const grupo = chats.find(c => c.isGroup && c.name === 'Bot_chat');
            if (grupo) {
                await grupo.sendMessage(resumen);
            }

            delete usuarios[from];
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

            const resumen = `🛎️ *Nuevo pedido confirmado*\n👤 Nombre: ${user.nombre}\n📍 Dirección: ${user.direccion}\n📱 Número: ${numeroLimpio}\n💳 Método de pago: Transferencia\n${user.pedido.join('\n')}\n💵 Total: $${user.total}`;

            client.sendMessage(from, '✅ ¡Comprobante recibido! Pedido confirmado, lo estamos preparando 🍽️');

            fs.appendFileSync('pedidos.txt', `[${new Date().toLocaleString()}] ${resumen.replace(/\n/g, ' | ')}\n`);

            const chats = await client.getChats();
            const grupo = chats.find(c => c.isGroup && c.name === 'Bot_chat');
            if (grupo) {
                await grupo.sendMessage(resumen);
                await grupo.sendMessage(new MessageMedia(media.mimetype, media.data, `comprobante_${numeroLimpio}_${Date.now()}.png`));
            }

            delete usuarios[from];
            return;
        } else {
            client.sendMessage(from, '📸 Por favor, enviá el comprobante de pago (imagen o archivo).');
            return;
        }
    }

    if (palabrasTotal.some(p => texto.includes(p))) {
        if (user.pedido.length > 0) {
            client.sendMessage(from, `🧾 Tu pedido actual es:\n${user.pedido.join('\n')}\n\n💵 Total: $${user.total}\n¿Querés confirmar o agregar algo más?`);
            return;
        } else {
            client.sendMessage(from, '🛒 No tenés ningún pedido en curso. ¿Qué querés pedir?');
            return;
        }
    }

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

    // Reemplazar la parte del código donde procesas los pedidos
if (!agregado) {
        // Expresión regular mejorada para capturar múltiples productos en el mensaje
        const cantidadRegex = new RegExp(
            `(?:quiero|dame|pedime|traeme|puede ser)?\\s*` +
            `(?:(una docena y media|dos docenas y media|una docena|media docena|\\d+ docenas|\\d+|una|un|uno|dos|tres|cuatro|cinco|seis|siete|ocho|nueve|diez|once|doce|trece|catorce|quince|dieciseis|diecisiete|dieciocho|diecinueve|veinte)\\s*(?:de)?)?\\s*` +
            `(${nombresProductosRegex})` +
            `(?:\\s|$|s)`, 'ig' // 'i' para ignorar mayúsculas, 'g' para capturar todas las coincidencias
        );

        console.log('Procesando mensaje completo:', texto);

        // Buscar todas las coincidencias en el mensaje
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
                console.log('Producto no encontrado en menú:', productoBase);
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



// 🔍 Pruebas manuales
console.log(convertirTextoACantidad("una docena y media de empanadas de jyq")); // debería dar 18
console.log(convertirTextoACantidad("3 docenas de empanadas de carne"));       // debería dar 36
console.log(convertirTextoACantidad("media docena de empanadas"));             // debería dar 6
console.log(convertirTextoACantidad("dos"));                                    // debería dar 2
console.log(convertirTextoACantidad("7"));                                      // debería dar 7
console.log(convertirTextoACantidad("cuatro docenas de empanadas"));           // debería dar 48

client.initialize();