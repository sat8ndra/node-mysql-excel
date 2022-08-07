const express = require('express')
const fs = require('fs')
const path = require('path')
const bodyparser = require('body-parser')
const readXlsxFile = require('read-excel-file/node')
const mysql = require('mysql')
const multer = require('multer')
const { ifError } = require('assert')
const app = express()
app.use(express.static('./public'))
app.use(bodyparser.json())
app.use(
  bodyparser.urlencoded({
    extended: true,
  }),
)


const db = mysql.createConnection({
  host: 'localhost',
  user: 'root',
  password: 'root',
  database: 'test',
})
db.connect(function (err) {
  if (err) {
    return console.error('error: ' + err.message)
  }
  console.log('Database connected.')
})
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, '/uploads/'))
  },
  filename: (req, file, cb) => {
    cb(null, file.fieldname + '-' + Date.now() + '-' + file.originalname)
  },
})
const uploadFile = multer({ storage: storage })
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html')
})
app.post('/import-excel', uploadFile.single('import-excel'), (req, res) => {
  var fileextname = path.extname(__dirname + '/uploads/' + req.file.filename);
  console.log(fileextname);
  if (fileextname == '.xlsx') {
  importFileToDb(__dirname + '/uploads/' + req.file.filename)
  console.log(res)
  }
  else {
    console.log("file is not uploaded to db upload only xlsx file");
  }
})
function importFileToDb(exFile) {
    readXlsxFile(exFile).then((rows) => {
      rows.shift()
          let query = 'INSERT INTO user (id, name, email) VALUES ?'
          db.query(query, [rows], (error, response) => {
            console.log(error || response)
          })
        })
}
  
  

let nodeServer = app.listen(4000, function () {
  let port = nodeServer.address().port
  let host = nodeServer.address().address
  console.log('App working on: ', host, port)
})