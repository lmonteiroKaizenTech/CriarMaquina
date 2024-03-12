const sql = require('mssql')
const config = require('./connection.js')

console.log(config);
async function runSelectQuery(sqlQuery) {
  try {
    
    await sql.connect(config);

    const result = await sql.query(sqlQuery);
    //console.log(result.recordset);
    
    await sql.close();

    return result; 

  } catch (err) {
    console.error('Error:', err);
  }
}

module.exports = {runSelectQuery};