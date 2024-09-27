const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const mysql = require('mysql2');
const path = require('path');
const bodyParser = require('body-parser');

const app = express();

// Configuración de Multer para subir archivos
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, file.originalname);
    }
});
const upload = multer({ storage: storage });

// Configuración de EJS y Body-Parser
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: false }));
app.use(express.static(path.join(__dirname, 'public')));
app.use((err, req, res, next) => {
    console.error(err.stack);
    res.status(500).send('Ocurrió un error inesperado');
});

// Conexión a la base de datos MySQL
const connection = mysql.createConnection({
    host: '192.168.4.6',
    user: 'desarrolloti',
    password: 'd3cr3t05',
    database: 'bd_inmobiliario'
});

connection.connect((err) => {
    if (err) {
        console.error('Error al conectar a la base de datos:', err);
        process.exit(1); // Finaliza el proceso si no se puede conectar
    } else {
        console.log('Conectado a la base de datos MySQL.');
    }
});

// Función para convertir fechas de Excel a formato MySQL (YYYY-MM-DD)
function excelDateToJSDate(excelDate) {
    // Verificar si la fecha es un número (formato de serie de Excel)
    if (typeof excelDate === 'number') {
        const daysSinceEpoch = excelDate - 25569; // El sistema de fechas de Excel empieza el 1 de enero de 1970
        const date = new Date(daysSinceEpoch * 86400 * 1000); // Convertir días en milisegundos
        return date.toISOString().split('T')[0]; // Devolver en formato YYYY-MM-DD
    }
    // Si la fecha ya está en formato de texto, devolverla directamente
    return excelDate;
}

// Rutas
app.get('/', (req, res) => {
    res.render('index');
});

app.get('/upload', (req, res) => {
    res.render('upload');
});

app.post('/upload', upload.single('excelFile'), (req, res) => {
    if (!req.file) {
        return res.status(400).send('No se ha subido ningún archivo.');
    }

    const filePath = req.file.path;
    let workbook;
    try {
        workbook = xlsx.readFile(filePath);
    } catch (error) {
        return res.status(400).send('Error al leer el archivo Excel.');
    }

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(worksheet);
    
    if (data.length === 0) {
        return res.status(400).send('El archivo Excel está vacío.');
    }

    connection.query('TRUNCATE TABLE gestion_comercial_dfvivienda', (err) => {
        if (err) throw err;

        const query = 'INSERT INTO gestion_comercial_dfvivienda (fecha_asignacion, numero_caso, identificacion, nombre_cliente, tramite, matricula, regional, estado, actividad, responsable,trazabilidad) VALUES ?';
        const values = data.map(row => [
            excelDateToJSDate(row.fecha_asignacion), // Conviertes la fecha aquí
            row.numero_caso,
            row.identificacion,
            row.nombre_cliente,
            row.tramite,
           
            row.matricula,
            row.regional,
            row.estado,
            row.actividad,
            row.responsable,
            row.trazabilidad
        ]);

        connection.query(query, [values], (err) => {
            console.log(err)
            if (err) return res.status(500).send('Error al insertar los datos en la base de datos.');
            res.send('Datos importados exitosamente.');
            
        });
    });
});

// Iniciar el servidor
app.listen(3000, () => {
    console.log('Servidor iniciado en el puerto 3000.');
 
});
