import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';

// Configurações de conexão
const sql = require('mssql');
const config = require('../../../../CRIARMÁQUINA/tests/dbConnection/connection.js');

test('CriarAreaPai', async ({ page }) => {

    // Definindo o tipo de uma linha do Excel
    type LinhaExcel = Record<string, unknown>;

    // Função para ler o arquivo Excel
    function lerArquivoExcel(nomeArquivo: string) {
        // Carrega o arquivo
        const workbook = XLSX.readFile(nomeArquivo);
    
        // Pega a primeira planilha do arquivo
        const primeiraPlanilha = workbook.Sheets[workbook.SheetNames[0]];
    
        // Converte os dados da planilha em um objeto JSON
        const dados = XLSX.utils.sheet_to_json(primeiraPlanilha);
    
        // Retorna os dados
        return dados;
    }
    
    // Exemplo de uso
    const dadosExcel = lerArquivoExcel('C:\\Users\\LeandroMonteiro\\Desktop\\CriarAreaMaquina.xlsx');
    console.log(dadosExcel);

    // ------------------------------Recolher dados------------------------------

});