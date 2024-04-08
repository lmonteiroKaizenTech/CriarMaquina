const { exec } = require('child_process');
// Executar um comando PowerShell e capturar a saída
exec('whoami', (error, stdout, stderr) => {
    if (error) {
        console.error(`Erro ao executar o comando: ${error.message}`);
        return;
    }
    if (stderr) {
        console.error(`Erro do PowerShell: ${stderr}`);
        return;
    }
});