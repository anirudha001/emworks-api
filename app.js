const express = require('express')
const path = require('path')
const logger = require('morgan')
const cookieParser = require('cookie-parser')
const bodyParser = require('body-parser')
const fileUpload = require('express-fileupload')
xlsxj = require("xlsx-to-json-lc");
var sql = require('mssql');
const cors = require('cors')

const app = express()

const sqlConfig = {
  user: 'saa',
  password: '123456',
  server: 'ANIRUDHA-PC\\SQLEXPRESS',
  database: 'emworks',
  beforeConnect: conn => {
    conn.once('connect', err => { err ? console.error(err) : console.log('mssql connected')})
    conn.once('end', err => { err ? console.error(err) : console.log('mssql disconnected')})
  }
}
 
app.use(logger('dev'))
app.use(cors())
app.use(bodyParser.json())
app.use(
  bodyParser.urlencoded({
    extended: false,
  }),
)
app.use(cookieParser())
app.use(fileUpload())
app.use('/public', express.static(__dirname + '/public'))

app.post('/upload', async (req, res, next) => {
  let uploadFile = req.files.file
  const fileName = req.files.file.name  
  let resultJsonData;                                           
  try {
     resultJsonData = await new Promise(function(resolve, reject) {
      uploadFile.mv(
        `${__dirname}/public/files/${fileName}`,
        function (err) {
          if (err) {
            reject(err);
          }
          xlsxj({
            input: `${__dirname}/public/files/${fileName}`, 
            output: `${__dirname}/public/files/employee data.json`,
            lowerCaseHeaders:true
          }, function(err, result) {
            if(err) {
              console.error(err);
              reject(err);
            }else {
              console.log(result);
              resolve(result);
            }
          });
        },
      )
    });
  } catch(err) {
      return res.status(500).send(err);
  }

  try {
    await sql.connect(sqlConfig);

    const table = new sql.Table('employees')
    table.create = false
    table.columns.add('firstname', sql.VarChar(255), {nullable: true})
    table.columns.add('lastname', sql.VarChar(255), {nullable: true})
    table.columns.add('currentcompany', sql.VarChar(255), {nullable: true})
    table.columns.add('currentsalary', sql.VarChar(255), {nullable: true})
    table.columns.add('expectedsalary', sql.VarChar(255), {nullable: true})
    table.columns.add('reasonforchange', sql.VarChar(500), {nullable: true})
  
    resultJsonData.forEach(emp => {
      table.rows.add(emp['first name'], emp['last name'], emp['current company'], emp['current salary'], emp['expected salary'], emp['reason for change']);
    });
   
    const request = new sql.Request();
    const result = await request.bulk(table);

  } catch(err) {
    console.log(err);
    return res.status(500).send(err);
  } finally {
    sql.close();
  }
        
  res.json({
    file: `public/${req.files.file.name}`,
  })
})

// catch 404 and forward to error handler
app.use(function (req, res, next) {
  const err = new Error('Not Found')
  err.status = 404
  next(err)
})

// error handler
app.use(function (err, req, res, next) {
  // set locals, only providing error in development
  res.locals.message = err.message
  res.locals.error = req.app.get('env') === 'development' ? err : {}

  // render the error page
  res.status(err.status || 500)
  res.render('error')
})

module.exports = app