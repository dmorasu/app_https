const https=require('https');
const fs=require('fs');
const express = require('express');
const mysql = require('mysql2');
const bodyParser = require('body-parser');
const excelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 5000;

app.use('/resources',express.static('public'));
app.use('/resources',express.static(__dirname + 'public'));

const options = {
  key: fs.readFileSync('/etc/letsencrypt/live/gestionactivosgpa.com/privkey.pem'),
  cert: fs.readFileSync('/etc/letsencrypt/live/gestionactivosgpa.com/fullchain.pem')
  
};


// Configuración de la base de datos
const db = mysql.createConnection({
    host: '192.168.4.6',
    user: 'desarrolloti',
    password: 'd3cr3t05',
    database: 'bd_inmobiliario',
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0,
    connectTimeout: 10000,
    keepAliveInitialDelay: 10000
});

db.connect((err) => {
    if (err) throw err;
    console.log('Conectado a la base de datos MySQL');
});

// Habilitar keep-alive en el pool
db.on('connection', function (connection) {
    connection.ping();  // Envía un ping para mantener la conexión activa
  });

// Pool de conexiones para la base de datos
db.on('error', function(err) {
    console.log('Error en la base de datos:', err);
    if (err.code === 'PROTOCOL_CONNECTION_LOST') {
      // En caso de pérdida de conexión, intenta reconectar
      console.log('Conexión perdida. Intentando reconectar...');
      pool.getConnection((error, connection) => {
        if (error) {
          console.error('Error al reconectar:', error);
        } else {
          console.log('Reconexión exitosa.');
          connection.release();
        }
      });
    } else {
      throw err;  // Lanza otros errores
    }
  });

// Middleware
app.use(bodyParser.urlencoded({ extended: false }));
app.use(express.json());
app.set('view engine', 'ejs');
app.use(express.static(path.join(__dirname, 'public')));

// Ruta principal
app.get('/', (req, res) => {
    res.render('index');
});

// Ruta para buscar clientes por cédula
app.post('/buscar', (req, res) => {
    const numero_caso = req.body.numero_caso;
    const sql = "SELECT *, DATE_FORMAT(fecha_asignacion,'%d-%b-%y') AS fecha FROM  gestion_comercial_dfvivienda WHERE numero_caso = ?";
    
    db.query(sql, [numero_caso], (err, results) => {
        if (err) throw err;
        res.render('resultados', { clientes: results });
    });
});

// Ruta para exportar los datos a Excel
// app.get('/exportar/:cedula', (req, res) => {
//     const cedula = req.params.cedula;
//     const sql = 'SELECT * FROM gestion_comercial WHERE usuario_comercial = ?';

//     db.query(sql, [cedula], async (err, results) => {
//         if (err) throw err;

//         const workbook = new excelJS.Workbook();
//         const worksheet = workbook.addWorksheet('Clientes');

//         worksheet.columns = [
//             { header: 'Cédula Cliente', key: 'cedula_cliente', width: 20 },
//             { header: 'Nombre Cliente', key: 'nombre_cliente', width: 30 },
//             { header: 'Fecha Asignación', key: 'fecha_asignacion', width: 20 },
//             { header: 'Número Caso', key: 'numero_caso', width: 20 },
//             { header: 'Valor Financiado', key: 'valor_financiado', width: 15 },
//             { header: 'Matrícula', key: 'matricula', width: 20 },
//             { header: 'Subclasificación', key: 'subclasificacion', width: 20 },
//             { header: 'Regional', key: 'regional', width: 20 },
//             { header: 'Responsable', key: 'responsable', width: 30 },
//             { header: 'Estado', key: 'estado', width: 20 },
//             { header: 'Comercial', key: 'comercial', width: 30 },
//             { header: 'Resumen Gestión', key: 'resumen_gestion', width: 40 }
//         ];

//         worksheet.addRows(results);

//         res.setHeader(
//             'Content-Type',
//             'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
//         );
//         res.setHeader(
//             'Content-Disposition',
//             'attachment; filename=clientes.xlsx'
//         );

//         await workbook.xlsx.write(res);
//         res.end();
//     });
// });

https.createServer(options, app).listen(5000, () => {
  console.log('Servidor HTTPS corriendo en puerto 5000, https://localhost:5000');
});

// app.listen(port, () => {
//     console.log(`Servidor iniciado en http://localhost:${port}`);
// });