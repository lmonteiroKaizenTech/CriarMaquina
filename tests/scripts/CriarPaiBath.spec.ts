import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';
const { exec } = require('child_process');

// Configurações de conexão
const sql = require('mssql');
const config = require('../../../CRIARMAQUINA/tests/dbConnection/connection.js');

// -----------Ambientes-----------

let ambientes_nome: any[] = ['AC_PRD','AC_QLD','AC_TST','AFL_PRD','AFL_QLD','AFL_TST','ACF_PRD','ACF_QLD','ACF_TST','ACC_PRD','ACC_QLD','ACC_TST','DEV','AQS_PRD','AQS_TST','ARC_PRD','ARC_TST','ACO_PRD','ACO_TST','CLP_PRD','CLP_TST','DISNEYLAND','MCS_TST'];
let ambientes_links: any[] = ['AMR-MES15','AMRMMES89','ktmesapp04','AMR-MES16','AMRMMES88','KTMESAPP03','AMRMMES28','AMRMMES87','KTMESAPP05','AMRMMES30','AMRMMES84','ktmesapp02','ktmesapp01','172.16.5.2','172.16.1.15','172.16.3.1','172.16.1.9','172.16.4.2','172.16.1.13','10.60.101.20','ktmesapp07','ktdisneyland01','172.16.1.10'];

let output, user = '';
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

    // Faça o que quiser com a saída, como armazená-la em uma variável ou vetor
    output = stdout.trim(); // Remove espaços em branco extras

    if (output)
    {
        for (var i = 0; i < output.length; i++)
        {
            if (output[i] == '\\')
            {
                for (var j = 1; j < output.length - i; j++) user += output[i + j];
            }
        }
    }
});
test('CriarMaquinaBatch', async ({ page }) => {

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
    const dadosExcel = lerArquivoExcel('C:\\Users\\' + user + '\\Desktop\\CriarPaiBath.xlsx');
    console.log(dadosExcel);

    // ------------------------------Recolher dados------------------------------

    // --------Recolha de dados--------

    let protocolo;
    let CP_MAQUINA_PERMITE_PARAR_AUTOMACAO;
    let CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO;
    let CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO;
    let CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO;
    let CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE;
    var soma = 0, linha4 = 0;
    let aut: boolean = true;

    let CP_MAQUINA_PERMITE_PARAR_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE_preenchido: boolean = true;

    if (dadosExcel) {
        for (var i = 0; i < 13; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[linha4] as LinhaExcel;

            switch (i) {
                case 1:
                    protocolo = segundaLinha['Campos Automação'] as string;
                    break;
                case 3:
                    CP_MAQUINA_PERMITE_PARAR_AUTOMACAO = segundaLinha['Campos Automação'] as string;
                    if (segundaLinha['Campos Automação'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_PARAR_AUTOMACAO_preenchido = false;
                    break;
                case 5:
                    CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO = segundaLinha['Campos Automação'] as string;
                    if (segundaLinha['Campos Automação'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO_preenchido = false;
                    break;
                case 7:
                    CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO = segundaLinha['Campos Automação'] as string;
                    if (segundaLinha['Campos Automação'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO_preenchido = false;
                    break;
                case 9:
                    CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO = segundaLinha['Campos Automação'] as string;
                    if (segundaLinha['Campos Automação'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO_preenchido = false;
                    break;
                case 11:
                    CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE = segundaLinha['Campos Automação'] as string;
                    if (segundaLinha['Campos Automação'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE_preenchido = false;
                    break;
            
                default:
                    break;
            }
            linha4++;
        }
        if (soma >= 1) aut = true;

    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log('-------------------------------');
    console.log(protocolo);
    console.log(CP_MAQUINA_PERMITE_PARAR_AUTOMACAO);
    console.log(CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO);
    console.log(CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO);
    console.log(CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO);
    console.log(CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE);

    let tipo, templatetags, locationname, area;

    // Verifica se os dados foram lidos corretamente
    if (dadosExcel) {
        // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
        const segundaLinha: LinhaExcel = dadosExcel[0] as LinhaExcel;

        // Por exemplo, para acessar um valor específico de uma coluna, você pode usar a chave correspondente ao cabeçalho
        tipo = segundaLinha['Tipo Máquina'] as string;
        templatetags = segundaLinha['Template Tags'] as string;
        locationname = segundaLinha['Nome da Location'] as string;
        area = segundaLinha['Area da Maquina'] as string;
        console.log(tipo);
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
        const segundaLinha: LinhaExcel = dadosExcel[linha3] as LinhaExcel;
        const prov = segundaLinha['Location'];
        if (prov) location.push(segundaLinha['Location'] as string);
        else break;
        linha3++;
    }

    console.log('-------------------------------');
    console.log(location);

    var linha2 = 0;
    let tags: any[] = [];

    while (1 < 2)
    {
        const segundaLinha: LinhaExcel = dadosExcel[linha2] as LinhaExcel; 
        const prov = segundaLinha['Caminho Tags'];
        if (prov) tags.push(segundaLinha['Caminho Tags'] as string);
        else break;
        linha2++;
    }
    
    console.log('-------------------------------');
    console.log(tags);

    let name, schedule, script, numero_maquina, protocolo_automacao, alternate_name, rejeitados, consumos_automaticos, capture_scheme, ProductSet;

    if (dadosExcel) {
        for (var i = 0; i < 6; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;

            switch (i) {
                case 1:
                    name = segundaLinha['General'] as string;
                    script = segundaLinha['Advanced'] as string;
                    numero_maquina = segundaLinha['Maquina'] as string;
                    alternate_name = segundaLinha['Notes'] as string;
                    rejeitados = segundaLinha['KPI'] as string;
                    break;
                case 3:
                    schedule = segundaLinha['General'] as string;
                    protocolo_automacao = segundaLinha['Maquina'] as string;
                    consumos_automaticos = segundaLinha['KPI'] as string;
                    break;
                case 5:
                    capture_scheme = segundaLinha['KPI'] as string;
                    ProductSet = segundaLinha['General'] as string;
                default:
                    break;
            }
        }
    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log('-----------------' + name + '-----------------');

    var linha5 = 1, soma = 0;
    let GroupName: any[] = [];

    while (1 < 2)
    {
        const segundaLinha: LinhaExcel = dadosExcel[linha5] as LinhaExcel; 
        const prov = segundaLinha['Event Definition (Group)'];
        soma++;
        if (prov) GroupName.push(segundaLinha['Event Definition (Group)'] as string);
        else break;
        linha5++;
    }
    console.log(GroupName);

    let EventName: any[] = [], EventDefinitionType: any[] = [], Priority: any[] = [], TriggerwhenEquals: any[] = [], OEEEventType: any[] = [], ReEvaluateSystemEventonStart: any[] = [], ReEvaluateSystemEventonEnd: any[] = [], ShowForAcknowledge: any[] = [], MTBFType: any[] = [], Duration: any[] = [], IsolationType: any[] = [], CP_EventDefinitionKey_ForMTBFTypeFailure: any[] = [], CP_EventDefinitionKey_ForMTBFTypeNONFailure: any[] = [], CP_EventDefinitionKey_ForMTBFTypeExcluded: any[] = [], CP_EventDefinitionIDLigada: any[] = [], CP_TagEventoCodigoAutomacao: any[] = [], CP_LoockupSetKeyCategoriaEventosAuto: any[] = [];

    for (var i = 1; i < soma; i++)
    {
        const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;
        EventName.push(segundaLinha['Event Definition (Event)'] as string);
        EventDefinitionType.push(segundaLinha['1'] as string);
        Priority.push(segundaLinha['2'] as string);
        TriggerwhenEquals.push(segundaLinha['3'] as string);
        OEEEventType.push(segundaLinha['4'] as string);
        ReEvaluateSystemEventonStart.push(segundaLinha['5'] as string);
        ReEvaluateSystemEventonEnd.push(segundaLinha['6'] as string);
        ShowForAcknowledge.push(segundaLinha['7'] as string);
        MTBFType.push(segundaLinha['8'] as string);
        Duration.push(segundaLinha['9'] as string);
        IsolationType.push(segundaLinha['10'] as string);
        CP_EventDefinitionKey_ForMTBFTypeFailure.push(segundaLinha['11'] as string);
        CP_EventDefinitionKey_ForMTBFTypeNONFailure.push(segundaLinha['12'] as string);
        CP_EventDefinitionKey_ForMTBFTypeExcluded.push(segundaLinha['13'] as string);
        CP_EventDefinitionIDLigada.push(segundaLinha['14'] as string);
        CP_TagEventoCodigoAutomacao.push(segundaLinha['15'] as string);
        CP_LoockupSetKeyCategoriaEventosAuto.push(segundaLinha['16'] as string);
    }
    console.log(EventName);

    console.log('erroooooooo');
    console.log('"C:\\Users\\' + user + '\\Desktop\\disneyland.xlsx"');
    console.log('erroooooooo');

    //------------Ir buscar site------------

    let ambiente;
    let site;

    if (dadosExcel) {
        for (var i = 0; i < 4; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[i] as LinhaExcel;

            if (i == 1) ambiente = segundaLinha['Site'] as string;

            if (i == 3)
            {
                site = segundaLinha['Site'] as string;
            }

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

    //------------Fim do Ir buscar site------------

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

    await page.goto('http://' + ambiente_final + '/TS/pages/' + site +'/config/tags/import');

    await page.waitForTimeout(3000);

    //-----------------Criar Tags------------------

    //Importar ficheiro excel

    // Localize o input de arquivo e insira o caminho do arquivo Excel
    const inputFile = await page.$('input[type="file"]');
    if (inputFile) await inputFile.setInputFiles('"C:\\Users\\' + user + '\\Desktop\\disneyland.xlsx"');
    await page.waitForTimeout(3000);
    await page.click('#Buttons_Import');

    await page.waitForTimeout(5000);

    for (var i = 0; i < tags.length; i++)
    {

        await page.locator('#contentPage_slice1_TreeList_Tree_TreeView').getByText(tags[i]).click();
    
    }

    await page.waitForTimeout(3000);

    await page.click(`li:has-text("Evento")`);

    await page.waitForTimeout(3000);

    await page.click('#tsslice-index-2 .fa-plus');
    await page.waitForTimeout(3000);
    const ScriptTag = await page.getByTitle('Script Tag').first();
    if (ScriptTag) ScriptTag.click();
    await page.waitForTimeout(3000);
    await page.fill('#tseditName', templatetags + '.Evento.FalhaComunicacoes');
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditDataType','Discrete');
    await page.waitForTimeout(2000);
    await page.selectOption('#tseditScriptType','Advanced (Multi-Line C#.NET Function)');
    await page.waitForTimeout(2000);
    await page.getByText(/^Evaluation$/i).click();
    await page.waitForTimeout(2000);
    await page.click('#tseditForceEvaluation');
    await page.waitForTimeout(2000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(5000);
    const tag = await page.getByTitle(templatetags + '.Evento.FalhaComunicacoes').first();
    if (tag) tag.click();
    await page.waitForTimeout(3000);
    await page.click('.fa-code');
    await page.waitForTimeout(3000);
    await page.click('.fa-cog');
    await page.waitForTimeout(5000);
    await page.selectOption('#InputEditorType','Text');
    await page.waitForTimeout(3000);
    await page.click('#Btns_Save');
    await page.waitForTimeout(3000);
    await page.fill('#contentPage_Editor_Code', 'if (Tags["' + templatetags + '.Evento.EstadoMaquina"].Quality != 0) return 0; else return 1;');

    await page.waitForTimeout(5000);

    await page.click('.tsoperation-toolbar-saveandclose');
    await page.waitForTimeout(5000);

    await page.goto('http://' + ambiente_final + '/TS/pages/' + site + '/config/tags/');

    await page.waitForTimeout(3000);

    for (var i = 0; i < tags.length; i++)
    {

        await page.locator('#contentPage_slice1_TreeList_Tree_TreeView').getByText(tags[i]).click();

    }

    await page.waitForTimeout(3000);

    await page.click(`li:has-text("Evento")`);

    await page.waitForTimeout(3000);

    await page.click('#tsslice-index-2 .fa-plus');
    await page.waitForTimeout(3000);
    const compare = await page.getByTitle('Compare Tag').first();
    if (compare) compare.click();
    await page.waitForTimeout(3000);
    await page.fill('#tseditName',templatetags + '.Evento.HeartBeatUpdate');
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditDataType','Integer');
    await page.waitForTimeout(3000);
    const primeiro = await page.getByTitle('Constant').first();
    if (primeiro) primeiro.click();
    await page.waitForTimeout(3000);
    const primeiro_segundo = await page.locator('.bi-tag-fill').first();
    if (primeiro_segundo) primeiro_segundo.click();
    await page.waitForTimeout(3000);
    await page.fill('#contentPage_Picker_LeftTagID_Name_TextBox','Global.HeartBeat');
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Picker_LeftTagID_Find_Button');
    await page.waitForTimeout(3000);
    await page.click('button:has-text("Assign")');
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditOperation','<>');
    await page.waitForTimeout(3000);
    const va7 = await page.getByTitle('Constant').nth(1);
    const vatextoHandle5 = await va7.first();
    await vatextoHandle5.click();
    await page.waitForTimeout(3000);
    const va8 = await page.locator('.bi-tag-fill').nth(1);
    const vatextoHandle6 = await va8.first();
    await vatextoHandle6.click();
    await page.fill('#contentPage_Picker_RightTagID_Name_TextBox', templatetags + '.Evento.HeartBeatMaquina');
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Picker_RightTagID_Find_Button');
    await page.waitForTimeout(3000);
    await page.click('button:has-text("Assign")');
    await page.waitForTimeout(3000);
    await page.getByText(/^Assign$/i).click();
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditAssignOnTrueOnly','While True');
    await page.waitForTimeout(3000);
    await page.getByText(/^Evaluation$/i).click();
    await page.waitForTimeout(3000);
    await page.click('#tseditForceEvaluation');
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(3000);

    await page.waitForTimeout(3000);
    await page.goto('http://' + ambiente_final + '/TS/pages/' + site + '/config/systems/');
    await page.waitForTimeout(3000);
    
    await page.click(`li:has-text("  Expand All")`);
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("${area}")`);
    await page.waitForTimeout(3000);
    const va2 = await page.locator(`a:has-text("New")`).nth(2);
    const vatextoHandle2 = await va2.first();
    await vatextoHandle2.click();
    await page.waitForTimeout(3000);
    await page.getByRole('link', { name: 'Batch System', exact: true }).click();
    await page.waitForTimeout(3000);

    // --------------Criar Máquina--------------

    await page.fill('#tseditName', name);
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditSystemTypeID', tipo);
    await page.waitForTimeout(3000);
    const clicar = await page.locator('#contentPage_tseditScheduleID_Picker').first();
    if (clicar) clicar.click();
    await page.waitForTimeout(3000);
    
    await page.click(`a:has-text("Expand All")`);
    await page.waitForTimeout(3000);
    await page.click(`li:text("${schedule}")`);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Picker_ScheduleID_AssignButton');
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditProductSetID', ProductSet);
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Event Splits")`);
    await page.waitForTimeout(3000);
    await page.click('#tseditSplitEventOnDayChange');
    await page.click('#tseditSplitEventOnShiftChange');
    await page.click('#tseditSplitEventOnProductChange');
    await page.click('#tseditSplitEventOnJobChange');
    await page.click('#tseditSplitEventOnBatchChange');
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Advanced")`);

    if (script) await page.fill('#tseditTemplateTagPrefix', templatetags);

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
    await page.click(`div:text("${name}")`);
    await page.waitForTimeout(3000);
    await page.click(`a:text("  Sub-Systems")`);

    await page.waitForTimeout(3000);

    var record1;
    var quantidade;
    try {
        await sql.connect(config)
        record1 = await sql.query`select count(id) as Quantidade from tTag where [Name] like '%${templatetags}%.Evento.EstadoMaquina%'` // select distinct
        quantidade = record1.recordset[0].Quantidade;
    
    } catch (e) {
        console.log(e);
    }

    await page.waitForTimeout(3000);

    

    await page.close();

});