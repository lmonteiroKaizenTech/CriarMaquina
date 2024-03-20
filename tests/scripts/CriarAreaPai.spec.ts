import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';

// Configurações de conexão
const sql = require('mssql');
const config = require('../../../CRIARMAQUINA/tests/dbConnection/connection.js');

// -----------Ambientes-----------

let ambientes_nome: any[] = ['AC_PRD','AC_QLD','AC_TST','AFL_PRD','AFL_QLD','AFL_TST','ACF_PRD','ACF_QLD','ACF_TST','ACC_PRD','ACC_QLD','ACC_TST','DEV','AQS_PRD','AQS_TST','ARC_PRD','ARC_TST','ACO_PRD','ACO_TST','CLP_PRD','CLP_TST','DISNEYLAND','MCS_TST'];
let ambientes_links: any[] = ['AMR-MES15','AMRMMES89','ktmesapp04','AMR-MES16','AMRMMES88','KTMESAPP03','AMRMMES28','AMRMMES87','KTMESAPP05','AMRMMES30','AMRMMES84','ktmesapp02','ktmesapp01','KTARCMESAPP01','KTMESAPP11','KTARCMESAPP01','KTMESAPP10','KTACOMESAPP01','KTMESAPP08','KTCLPMESAPP01','KTMESAPP07','ktdisneyland01','ktmesapp06'];

test('CriarAreaPai', async ({ page }) => {

    // Definindo o tipo de uma linha do Excel
    type LinhaExcel = Record<string, string | null>;

    // Função para ler o arquivo Excel
    function lerArquivoExcel(nomeArquivo: string): LinhaExcel[] {
        // Carrega o arquivo
        const workbook = XLSX.readFile(nomeArquivo);

        // Pega a primeira planilha do arquivo
        const primeiraPlanilha = workbook.Sheets[workbook.SheetNames[0]];

        // Converte os dados da planilha em um objeto JSON
        const dados = XLSX.utils.sheet_to_json(primeiraPlanilha, { header: 1 }) as string[][];

        // Extrai os cabeçalhos da primeira linha
        const colunas = dados[0];

        // Inicializa um array para armazenar os dados
        const dadosFormatados: LinhaExcel[] = [];

        // Itera sobre as linhas de dados, começando da segunda linha
        for (let i = 1; i < dados.length; i++) {
            const linha: LinhaExcel = {};
            // Itera sobre as colunas
            for (let j = 0; j < colunas.length; j++) {
                const valor = dados[i][j];
                linha[colunas[j]] = valor !== undefined ? valor.toString() : null;
            }
            dadosFormatados.push(linha);
        }

        // Retorna os dados formatados
        return dadosFormatados;
    }
    
    // Exemplo de uso
    const dadosExcel = lerArquivoExcel('C:\\Users\\LeandroMonteiro\\Desktop\\CriarAreaMaquina.xlsx');
    //console.log(dadosExcel);

    // ------------------------------Recolher dados------------------------------

    //------------Variáveis------------

    let ambiente;
    let site;

    if (dadosExcel) {
        for (var i = 0; i < 4; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;

            if (i == 1) ambiente = segundaLinha['Site'] as string;

            if (i == 3) site = segundaLinha['Site'] as string;

        }
    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log(ambiente);
    console.log(site);

    var position = 0;
    for (var i = 0; i < ambientes_nome.length; i++)
    {
        if (ambiente == ambientes_nome[i]) position = i;
    }

    let ambiente_final;
    for (var i = 0; i < ambientes_links.length; i++)
    {
        if (i == position) ambiente_final = ambientes_links[i];
    }

    let General;
    let Notificacoes;
    let Notes;

    // Verifica se os dados foram lidos corretamente
    if (dadosExcel) {
        // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
        const segundaLinha: LinhaExcel = dadosExcel[1] as LinhaExcel;

        // Por exemplo, para acessar um valor específico de uma coluna, você pode usar a chave correspondente ao cabeçalho
        General = segundaLinha['General'] as string;
        Notificacoes = segundaLinha['Notificações'] as string;
        Notes = segundaLinha['Notes'] as string;
        console.log(General);
        console.log(Notificacoes);
        console.log(Notes);

    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }

    //------------Vetores------------
    var i = 0, idep = 1;
    let CaminhoArea: any[] = [], continuar: boolean = true;

    // Verifica se os dados foram lidos corretamente
    if (dadosExcel) {

        while (continuar) {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;
            const segundaLinhadep: LinhaExcel = dadosExcel[idep] as LinhaExcel;

            CaminhoArea.push(segundaLinha['Caminho da Area'] as string);

            if (segundaLinhadep['Caminho da Area'] == null) break;

            i++;
            idep++;
        }

    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log(CaminhoArea);

    let RegraLote: any[] = [];

    if (dadosExcel) {
        for (var i = 0; i < dadosExcel.length; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;

            RegraLote[i] = segundaLinha['Regra Lote'] as string;
            //console.log(segundaLinha['Regra Lote']);

        }
    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log(RegraLote);

    let SAP: any[] = [];

    if (dadosExcel) {
        for (var i = 0; i < 10; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;

            SAP[i] = segundaLinha['SAP'] as string;

        }
    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log(SAP);

    let LuzAzul: any[] = [];

    if (dadosExcel) {
        for (var i = 0; i < 10; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;

            LuzAzul[i] = segundaLinha['Luz Azul'] as string;
        }
    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log(LuzAzul);

    // ---------------Gerar Key---------------

    await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');

    await page.getByLabel('Login').fill('kt0032'); //utilizador kt 
    await page.getByLabel('Password').click();
    await page.getByLabel('Password').fill('12345'); // password
    await page.getByRole('button', { name: 'Sign In' }).click();

    await page.click('#contentPage_ctl05');
    await page.click('.btn-item-key-btn_GerarKey');
    await page.waitForTimeout(3000);
    const key2 = await page.locator('#contentPage_ctl04').textContent();
    let final_key2
    if (key2) final_key2 = key2.trim();

    await page.waitForTimeout(3000);

    // ---------------Login Site Principal---------------
    
    await page.goto('http://' + ambiente_final + '/TS/');
    await page.waitForTimeout(2000);
    //Verificação de Login
    const currentURL = page.url();
    await page.waitForTimeout(2000);
    if (currentURL.includes('http://' + ambiente_final + '/TS/Account/LogOn.aspx'))
    {
        await page.getByLabel('Login').fill('kt0032'); //utilizador kt 
        await page.getByLabel('Password').click();
        await page.getByLabel('Password').fill('12345'); // password
        await page.getByRole('button', { name: 'Sign In' }).click();
    }
    await page.waitForTimeout(3000);

    // ---------------Criar Área---------------

    await page.goto('http://' + ambiente_final + '/TS/pages/' + site + '/config/systems/');

    await page.waitForTimeout(3000);
    for (var i = 0; i < CaminhoArea.length; i++) await page.click(`li:has-text("${CaminhoArea[i]}")`);
    await page.waitForTimeout(3000);
    await page.click(`a:has-text("New Child")`);
    await page.waitForTimeout(3000);

    // ----------Parametrizações da Área----------

    await page.fill('#tseditName', General);
    await page.waitForTimeout(2000);
    await page.fill('#tseditKey', final_key2);
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Regra Lote")`);
    await page.waitForTimeout(3000);

    for (var i = 0; i < RegraLote.length; i++)
    {
        switch (i) {
            case 0:
                if (RegraLote[i+1] != null) await page.selectOption('#tseditcp_CPS0000000017_CP0000000051', RegraLote[i+1]);
                break;
            case 2:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000048', RegraLote[i+1]);
                break;
            case 4:
                if (RegraLote[i+1] == "Sim") await page.click('#tseditcp_CPS0000000017_CP0000000092');
                break;
            case 6:
                if (RegraLote[i+1] == "Sim") await page.click('#tseditcp_CPS0000000017_CP0000000184');
                break;
            case 8:
                if (RegraLote[i+1] == "Sim") await page.click('#tseditcp_CPS0000000017_CP0000000234');
                break;
            case 10:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000052', RegraLote[i+1]);
                break;
            case 12:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000053', RegraLote[i+1]);
                break;
            case 14:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000063', RegraLote[i+1]);
                break;
            case 16:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000082', RegraLote[i+1]);
                break;
            case 18:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000084', RegraLote[i+1]);
                break;
            case 20:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000098');
                break;
            case 22:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000099');
                break;
            case 24:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000124');
                break;
            case 26:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000130');
                break;
            case 28:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000140', RegraLote[i+1]);
                break;
            case 30:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000153', RegraLote[i+1]);
                break;
            case 32:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000182');
                break;
            case 34:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000154');
                break;
            case 36:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000166');
                break;
            case 38:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000168', RegraLote[i+1]);
                break;
            case 40:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000170', RegraLote[i+1]);
                break;
            case 42:
                if (RegraLote[i+1] != null) await page.selectOption('#tseditcp_CPS0000000017_CP0000000177', RegraLote[i+1]);
                break;
            case 44:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000197', RegraLote[i+1]);
                break;
            case 46:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000220', RegraLote[i+1]);
                break;
            case 48:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000222', RegraLote[i+1]);
                break;
            case 50:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000225', RegraLote[i+1]);
                break;
            case 52:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000232', RegraLote[i+1]);
                break;
            case 54:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000244', RegraLote[i+1]);
                break;
            case 56:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000253', RegraLote[i+1]);
                break;
            case 58:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000254', RegraLote[i+1]);
                break;
            case 60:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000256', RegraLote[i+1]);
                break;
            case 62:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000257', RegraLote[i+1]);
                break;
            case 64:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000261', RegraLote[i+1]);
                break;
            case 66:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000263');
                break;
            case 68:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000264');
                break;
            case 70:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000265', RegraLote[i+1]);
                break;
            case 72:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000268', RegraLote[i+1]);
                break;
            case 74:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000271', RegraLote[i+1]);
                break;
            case 76:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000274');
                break;
            case 78:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000299', RegraLote[i+1]);
                break;
            case 80:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000300', RegraLote[i+1]);
                break;
            case 82:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000301', RegraLote[i+1]);
                break;
            case 84:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000310');
                break;
            case 86:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000312');
                break;
            case 88:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000333');
                break;
            case 90:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000342');
                break;
            case 92:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000353', RegraLote[i+1]);
                break;
            case 94:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000357', RegraLote[i+1]);
                break;
            case 96:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000365');
                break;
            case 98:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000178', RegraLote[i+1]);
                break;
            case 100:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000179', RegraLote[i+1]);
                break;
            case 102:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000180', RegraLote[i+1]);
                break;
            case 104:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000192', RegraLote[i+1]);
                break;
            case 106:
                if (RegraLote[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000017_CP0000000302');
                break;
            case 108:
                if (RegraLote[i+1] != null) await page.fill('#tseditcp_CPS0000000017_CP0000000303', RegraLote[i+1]);
                break;
        }
    }

    await page.waitForTimeout(3000);
    await page.click(`li:has-text("SAP")`);
    await page.waitForTimeout(3000);

    for (var i = 0; i < SAP.length; i++)
    {
        switch (i) {
            case 0:
                if (SAP[i+1] != null) await page.selectOption('#tseditcp_CPS0000000009_CP0000000078', SAP[i+1]);
                break;
            case 2:
                if (SAP[i+1] != null) await page.selectOption('#tseditcp_CPS0000000009_CP0000000020', SAP[i+1]);
                break;
            case 4:
                if (SAP[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000009_CP0000000047', SAP[i+1]);
                break;
            case 6:
                if (SAP[i+1] == 'Sim') await page.click('#tseditcp_CPS0000000009_CP0000000106', SAP[i+1]);
                break;
            case 8:
                if (SAP[i+1] != null) await page.fill('#tseditcp_CPS0000000009_CP0000000135', SAP[i+1]);
                break;
        }
    }

    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Luz Azul")`);
    await page.waitForTimeout(3000);

    for (var i = 0; i < LuzAzul.length; i++)
    {
        switch (i) {
            case 0:
                if (LuzAzul[i+1] != null) await page.fill('#tseditcp_CPS0000000021_CP0000000066', LuzAzul[i+1]);
                break;
            case 2:
                if (LuzAzul[i+1] != null) await page.fill('#tseditcp_CPS0000000021_CP0000000067', LuzAzul[i+1]);
                break;
            case 4:
                if (LuzAzul[i+1] != null) await page.fill('#tseditcp_CPS0000000021_CP0000000068', LuzAzul[i+1]);
                break;
            case 6:
                if (LuzAzul[i+1] != null) await page.fill('#tseditcp_CPS0000000021_CP0000000069', LuzAzul[i+1]);
                break;
            case 8:
                if (LuzAzul[i+1] != null) await page.fill('#tseditcp_CPS0000000021_CP0000000071', LuzAzul[i+1]);
                break;
        }
    }

    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Notificações")`);
    await page.waitForTimeout(3000);

    await page.fill('#tseditcp_CPS0000000032_CP0000000233', Notificacoes);
    await page.waitForTimeout(3000);

    await page.click(`li:has-text("Notes")`);
    await page.waitForTimeout(3000);

    await page.fill('#tseditAltName', Notes);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(3000);

    // ----------------------Criar Máquina Pai ou individual----------------------

    // --------Recolha de dados--------

// Função para ler o arquivo Excel
function lerArquivoExcel2(nomeArquivo: string): LinhaExcel[] {
    // Carrega o arquivo
    const workbook = XLSX.readFile(nomeArquivo);

    // Pega a primeira planilha do arquivo
    const primeiraPlanilha = workbook.Sheets[workbook.SheetNames[0]];

    // Converte os dados da planilha em um objeto JSON
    const dados = XLSX.utils.sheet_to_json(primeiraPlanilha, { header: 1 }) as string[][];

    // Extrai os cabeçalhos da primeira linha
    const colunas = dados[0];

    // Inicializa um array para armazenar os dados
    const dadosFormatados: LinhaExcel[] = [];

    // Itera sobre as linhas de dados, começando da segunda linha
    for (let i = 1; i < dados.length; i++) {
        const linha: LinhaExcel = {};
        // Itera sobre as colunas
        for (let j = 0; j < colunas.length; j++) {
            const valor = dados[i][j];
            linha[colunas[j]] = valor !== undefined ? valor.toString() : null;
        }
        dadosFormatados.push(linha);
    }

    // Retorna os dados formatados
    return dadosFormatados;
}
    
    // Exemplo de uso
    const dadosExcel2 = lerArquivoExcel2('C:\\Users\\LeandroMonteiro\\Desktop\\CriarMaquinaPai.xlsx');
    console.log(dadosExcel2);

    let tipomaquina, templatetags, locationname, area;

    // Verifica se os dados foram lidos corretamente
    if (dadosExcel2) {
        // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
        const segundaLinha: LinhaExcel = dadosExcel2[0] as LinhaExcel;

        // Por exemplo, para acessar um valor específico de uma coluna, você pode usar a chave correspondente ao cabeçalho
        tipomaquina = segundaLinha['Tipo'] as string;
        templatetags = segundaLinha['Template Tags'] as string;
        locationname = segundaLinha['Nome da Location'] as string;
        area = segundaLinha['Area da Maquina'] as string;
        console.log(tipomaquina);
        console.log(templatetags);
        console.log(locationname);
        console.log(area);

    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }

    var linha3 = 0;
    let location: any[] = [];

    while (1 < 2)
    {
        const segundaLinha: LinhaExcel = dadosExcel2[linha3] as LinhaExcel; 
        const prov = segundaLinha['Location'];
        if (prov) location.push(segundaLinha['Location'] as string);
        else break;
        linha3++;
    }

    console.log('-------------------------------');
    console.log(location);

    let name, schedule, script, numero_maquina, protocolo_automacao, alternate_name, KPI;

    if (dadosExcel2) {
        for (var i = 0; i < 4; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel2[i] as LinhaExcel;

            switch (i) {
                case 1:
                    name = segundaLinha['General'] as string;
                    script = segundaLinha['Advanced'] as string;
                    numero_maquina = segundaLinha['Maquina'] as string;
                    alternate_name = segundaLinha['Notes'] as string;
                    KPI = segundaLinha['KPI'] as string;
                    break;
                case 3:
                    schedule = segundaLinha['General'] as string;
                    protocolo_automacao = segundaLinha['Maquina'] as string;
                    break;
                default:
                    break;
            }
        }
    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log('-----------------' + name + '-----------------');

    // --------Criar Máquina--------

    // ---------------Gerar Key Location---------------

    await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');

    await page.click('#contentPage_ctl25');
    await page.click('.btn-item-key-btn_GerarKey');
    await page.waitForTimeout(3000);
    const keylacation = await page.locator('#contentPage_ctl04').textContent();
    let final_keylocation
    if (keylacation) final_keylocation = keylacation.trim();

    // ---------------Criar Location---------------

    await page.goto('http://ktmesapp04/TS/pages/home/config/locations/');

    await page.waitForTimeout(3000);
    for (var i = 0; i < location.length; i++)
    {
        await page.getByText(new RegExp("^" + location[i] + "$", "i")).click();
        await page.waitForTimeout(3000);
    }

    await page.waitForTimeout(3000);
    await page.click(`li:has-text("New Child")`);
    await page.waitForTimeout(3000);
    await page.fill('#tseditName', locationname);
    await page.waitForTimeout(2000);
    if (final_keylocation) await page.fill('#tseditUniqueID', final_keylocation);
    await page.waitForTimeout(2000);
    await page.selectOption('#tseditLocationTypeID','LT_Maquinas');
    await page.waitForTimeout(2000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(5000);

    await page.goto('http://ktmesapp04/TS/pages/home/config/tags/');
    await page.waitForTimeout(3000);

    await page.click(`li:has-text("Template")`);
    await page.click(`li:has-text("Multi")`);
    await page.click(`li:has-text("MultiXX")`);
    await page.waitForTimeout(3000);
    await page.click(`a:has-text("Duplicate")`);

    if (site == "home") await page.fill('#contentPage_slice2_FromPrefixInput', 'CHK.CHK_MULTI.MultiXX');
    else await page.fill('#contentPage_slice2_FromPrefixInput', site + '.' + site + '_MULTI.MultiXX');

    await page.fill('#contentPage_slice2_ToPrefixInput', templatetags);
    await page.waitForTimeout(3000);

    await page.click('#contentPage_slice2_DuplicateButton');
    
    await page.waitForTimeout(3000);

    const va1 = await page.locator(`li:has-text("MultiXX")`).nth(1);
    const vatextoHandle1 = await va1.first();
    await vatextoHandle1.click();

    await page.waitForTimeout(3000);
    await page.click('.fa-edit');
    await page.waitForTimeout(3000);
    await page.fill('#tseditName', templatetags);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(3000);

    await page.click(`a:has-text("Move")`);
    await page.waitForTimeout(3000);

    await page.click(`span:has-text("Expand All")`);
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("${area}")`);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_slice2_Move');

    await page.waitForTimeout(3000);

    await page.click(`li:has-text("Systems")`);
    await page.waitForTimeout(3000);

    for (var i = 0; i < CaminhoArea.length; i++) await page.click(`li:has-text("${CaminhoArea[i]}")`);

    await page.waitForTimeout(3000);

    await page.click(`li:text("${General}")`);
    await page.waitForTimeout(3000);
    const va2 = await page.locator(`a:has-text("New")`).nth(2);
    const vatextoHandle2 = await va2.first();
    await vatextoHandle2.click();
    //await page.getByTitle(tipomaquina).click();
    await page.waitForTimeout(3000);
    await page.getByText(tipomaquina).click();
    await page.waitForTimeout(3000);
    await page.fill('#tseditName', name);
    await page.waitForTimeout(3000);
    const clicar = await page.locator('#contentPage_tseditScheduleID_Picker').first();
    if (clicar) clicar.click();
    await page.waitForTimeout(3000);
    
    await page.click(`a:has-text("Expand All")`);
    await page.click(`li:text("${schedule}")`);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Picker_ScheduleID_AssignButton');
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Event Splits")`);
    await page.waitForTimeout(3000);
    await page.click('#tseditSplitEventOnDayChange');
    await page.click('#tseditSplitEventOnShiftChange');
    await page.click('#tseditSplitEventOnProductChange');
    await page.click('#tseditSplitEventOnJobChange');
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Job")`);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_tseditJobTagID_Picker');
    await page.fill('#contentPage_Picker_JobTagID_Name_TextBox',templatetags + '.Ord.Ordem');
    await page.click('#contentPage_Picker_JobTagID_Find_Button');
    await page.waitForTimeout(2000);
    const clicarbut = await page.locator(`button:has-text("Assign")`).first();
    await page.waitForTimeout(2000);
    if (clicarbut) clicarbut.click();
    await page.waitForTimeout(3000);
    const segundo = await page.locator(`li:has-text("Product")`).nth(6);
    const vatextoHandle3 = await segundo.first();
    await vatextoHandle3.click();
    await page.waitForTimeout(5000);
    await page.click('#contentPage_tseditProductTagID_Picker');
    await page.fill('#contentPage_Picker_ProductTagID_Name_TextBox', templatetags + '.Prod.CodigoProduto');
    await page.click('#contentPage_Picker_ProductTagID_Find_Button');
    await page.waitForTimeout(3000);
    const clicarbut2 = await page.locator(`button:has-text("Assign")`).first();
    await page.waitForTimeout(2000);
    if (clicarbut2) clicarbut2.click();
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Advanced")`);

    if (script) await page.fill('#tseditScriptClassName',script);

    await page.fill('#tseditTemplateTagPrefix', templatetags);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_tseditLocationID_Picker');
    await page.waitForTimeout(3000);
    await page.click(`a:has-text("Expand All")`);
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("${locationname}")`);
    await page.waitForTimeout(3000);
    await page.click("#contentPage_Picker_LocationID_AssignButton");
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Maquina")`);
    await page.waitForTimeout(3000);
    if (numero_maquina) await page.fill('#tseditcp_CPS0000000013_CP0000000083', numero_maquina);
    await page.locator(`li:has-text("Maquina")`);
    await page.selectOption('#tseditcp_CPS0000000013_CP0000000045', protocolo_automacao);
    await page.waitForTimeout(3000);
    if (alternate_name) await page.fill('#tseditAltName', alternate_name);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(5000);

    // -------------------KPI's Máquina Pai-------------------

    // await page.getByTitle(name).click();
    // await page.waitForTimeout(3000);
    // await page.click(`li:has-text("KPI Calculations")`);
    // await page.waitForTimeout(3000);
    // await page.click(`a:has-text("New")`);
    // await page.waitForTimeout(3000);
    // await page.fill('#tseditName','OEE');
    // await page.selectOption('#tseditOeeCalculationTypeID','KPI_Producao');
    // await page.waitForTimeout(3000);
    // await page.click(`li:has-text("Rates")`);
    // await page.waitForTimeout(3000);
    // const primeiro = await page.getByTitle('Constant').first();
    // if (primeiro) primeiro.click();
    // await page.waitForTimeout(3000);
    // const primeiro_segundo = await page.locator('.glyphicon-tag').first();
    // if (primeiro_segundo) primeiro_segundo.click();
    // await page.waitForTimeout(3000),
    // await page.fill('#contentPage_Picker_TheoreticalCalculationUnitsPerMinuteTagID_Name_TextBox', templatetags + '.Prod.TaxaProducaoTeorica');
    // await page.waitForTimeout(3000);
    // await page.click('#contentPage_Picker_TheoreticalCalculationUnitsPerMinuteTagID_Find_Button');
    // await page.waitForTimeout(3000);
    // await page.click('button:has-text("Assign")');
    // await page.waitForTimeout(3000);
    // await page.selectOption('#tseditTargetRateUnitType','Units per Minute');
    // await page.waitForTimeout(3000);
    // await page.fill('#tseditScriptClassName','OeeCalculationScriptKPI2ITEM');
    // await page.waitForTimeout(3000);
    // await page.click('#contentPage_Save_Button');
    // await page.waitForTimeout(3000);
    // await page.getByTitle('OEE').click();
    // await page.waitForTimeout(3000);
    // await page.click(`li:has-text("Good")`);
    // await page.waitForTimeout(3000);
    // await page.click(`a:has-text("New")`);
    // await page.waitForTimeout(3000);
    // await page.fill('#tseditName','ProdutoXX');
    // await page.waitForTimeout(3000);

    await page.close();

});