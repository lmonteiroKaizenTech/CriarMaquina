const express = require('express');
const app = express();
const cors = require('cors');
const PORT = process.env.PORT || 5000;
app.use(cors());

// Servir o arquivo HTML
app.get('/', (req, res) => {
  res.sendFile(__dirname + '/index.html');
});

// Servir o arquivo JSON
app.get('/test-results', (req, res) => {
  const json = require('../test-results.json');
  res.json(json);
});

// Inicia o servidor
app.listen(PORT, () => {
  console.log(`Servidor rodando em http://localhost:${PORT}`);
});