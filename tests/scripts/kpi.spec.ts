import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';

// Configurações de conexão
const sql = require('mssql');
const config = require('../../../CRIARMAQUINA/tests/dbConnection/connection.js');

// -----------Ambientes-----------

let ambientes_nome: any[] = ['AC_PRD','AC_QLD','AC_TST','AFL_PRD','AFL_QLD','AFL_TST','ACF_PRD','ACF_QLD','ACF_TST','ACC_PRD','ACC_QLD','ACC_TST','DEV','AQS_PRD','AQS_TST','ARC_PRD','ARC_TST','ACO_PRD','ACO_TST','CLP_PRD','CLP_TST','DISNEYLAND'];
let ambientes_links: any[] = ['AMR-MES15','AMRMMES89','ktmesapp04','AMR-MES16','AMRMMES88','KTMESAPP03','AMRMMES28','AMRMMES87','KTMESAPP05','AMRMMES30','AMRMMES84','ktmesapp02','ktmesapp01','KTMESAPP11','KTARCMESAPP01','KTMESAPP10','KTACOMESAPP01','KTMESAPP08','KTCLPMESAPP01','KTMESAPP07','ktdisneyland01'];

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

    let name, schedule, script, numero_maquina, protocolo_automacao, alternate_name, rejeitados, consumos_automaticos, capture_scheme;

    if (dadosExcel2) {
        for (var i = 0; i < 6; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel2[i] as LinhaExcel;

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
        const segundaLinha: LinhaExcel = dadosExcel2[linha5] as LinhaExcel; 
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
        const segundaLinha: LinhaExcel = dadosExcel2[i] as LinhaExcel;
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

    // ---------------Gerar Key Event Definition---------------

    await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');

    await page.getByLabel('Login').fill('kt0032'); //utilizador kt 
    await page.getByLabel('Password').click();
    await page.getByLabel('Password').fill('12345'); // password
    await page.getByRole('button', { name: 'Sign In' }).click();

    await page.click('#contentPage_ctl15');
    await page.waitForTimeout(3000);
    await page.click('.btn-item-key-btn_GerarKey');
    await page.waitForTimeout(3000);
    const keyEventDefinition = await page.locator('#contentPage_ctl04').textContent();
    let final_keyEventDefinition
    if (keyEventDefinition) final_keyEventDefinition = keyEventDefinition.trim();
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

    await page.goto('http://ktmesapp04/TS/pages/home/config/systems/');
    for (var i = 0; i < CaminhoArea.length; i++) await page.click(`li:has-text("${CaminhoArea[i]}")`);
    await page.waitForTimeout(3000);
    await page.click(`li:text("${General}")`);
    

    await page.waitForTimeout(3000);

    // -------------------KPI's Máquina Pai-------------------

    await page.click(`div:text("${name}")`);
    await page.waitForTimeout(3000);
    await page.click(`a:text("  KPI Calculations")`);
    await page.waitForTimeout(3000);
    await page.click(`a:has-text("New")`);
    await page.waitForTimeout(3000);
    await page.fill('#tseditName','OEE');
    await page.selectOption('#tseditOeeCalculationTypeID','KPI_Producao');
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Rates")`);
    await page.waitForTimeout(3000);
    const primeiro = await page.getByTitle('Constant').first();
    if (primeiro) primeiro.click();
    await page.waitForTimeout(3000);
    const primeiro_segundo = await page.locator('.bi-tag-fill').first();
    if (primeiro_segundo) primeiro_segundo.click();
    await page.waitForTimeout(3000),
    await page.fill('#contentPage_Picker_TheoreticalCalculationUnitsPerMinuteTagID_Name_TextBox', templatetags + '.Prod.TaxaProducaoTeorica');
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Picker_TheoreticalCalculationUnitsPerMinuteTagID_Find_Button');
    await page.waitForTimeout(3000);
    await page.click('button:has-text("Assign")');
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditTargetRateUnitType','Units per Minute');
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("Advanced")`);
    await page.waitForTimeout(3000);
    await page.fill('#tseditScriptClassName','OeeCalculationScriptKPI2ITEM');
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(3000);
    const clicarOEE = await page.getByTitle('OEE').first();
    if (clicarOEE) clicarOEE.click();
    await page.waitForTimeout(3000);
    await page.click(`a:text("  Good")`);
    await page.waitForTimeout(3000);

    var record1;
    var contagem;
    try {
        await sql.connect(config)
        record1 = await sql.query`select count (id) - 2 as Contar from tTag where [Name] like '%' + ${templatetags.toString()} + '.Prod.ContadorProduto%'` // select distinct
        contagem = record1.recordset[0].Contar;
    
    } catch (e) {
        console.log(e);
    }

    await page.waitForTimeout(3000);

    for (var i = 1; i < contagem + 1; i++)
    {
        if (i == 2 && rejeitados == 'Sim')
        {
            await page.click(`a:text("  Bad")`);
            await page.waitForTimeout(3000);
            await page.click(`a:has-text("New")`);
            await page.waitForTimeout(3000);
            await page.fill('#tseditName', 'Produto0' + i);
            await page.waitForTimeout(3000);
            const primeiro = await page.getByTitle('Constant').first();
            if (primeiro) primeiro.click();
            await page.waitForTimeout(3000);
            const primeiro_segundo = await page.locator('.bi-tag-fill').first();
            if (primeiro_segundo) primeiro_segundo.click();
            await page.waitForTimeout(3000);
            await page.fill('#contentPage_Picker_CounterTagID_Name_TextBox', templatetags + '.Prod.ContadorProduto0' + i);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Picker_CounterTagID_Find_Button');
            await page.waitForTimeout(3000);
            const clicarbutkpi = await page.locator(`button:has-text("Assign")`).first();
            if (clicarbutkpi) clicarbutkpi.click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditMaxPlusTagConstant_Constant','999999');
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Advanced")`);
            await page.waitForTimeout(3000);
            await page.fill('#tseditRolloverTagConstant_Constant','999999');
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
            await page.click(`a:text("  Good")`);
            await page.waitForTimeout(3000);
        }
        else
        {
            await page.click(`a:has-text("New")`);
            await page.waitForTimeout(3000);
            if (i <= 10) await page.fill('#tseditName', 'Produto0' + i);
            else await page.fill('#tseditName', 'Produto' + i);
            await page.waitForTimeout(3000);
            const primeiro = await page.getByTitle('Constant').first();
            if (primeiro) primeiro.click();
            await page.waitForTimeout(3000);
            const primeiro_segundo = await page.locator('.bi-tag-fill').first();
            if (primeiro_segundo) primeiro_segundo.click();
            await page.waitForTimeout(3000);
            if (i <= 10) await page.fill('#contentPage_Picker_CounterTagID_Name_TextBox', templatetags + '.Prod.ContadorProduto0' + i);
            else await page.fill('#contentPage_Picker_CounterTagID_Name_TextBox', templatetags + '.Prod.ContadorProduto' + i);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Picker_CounterTagID_Find_Button');
            await page.waitForTimeout(3000);
            const clicarbutkpi = await page.locator(`button:has-text("Assign")`).first();
            if (clicarbutkpi) clicarbutkpi.click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditMaxPlusTagConstant_Constant','999999');
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Advanced")`);
            await page.waitForTimeout(3000);
            await page.fill('#tseditRolloverTagConstant_Constant','999999');
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
        }
    }

    await page.waitForTimeout(3000);

    // ---------Automáticos---------

    if (consumos_automaticos == 'Sim')
    {
        var record2;
        var contagem2;
        try {
            await sql.connect(config)
            record2 = await sql.query`select count (id) - 2 as Contar from tTag where [Name] like '%' + ${templatetags.toString()} + '.Consumo.ContadorProduto%'` // select distinct
            contagem2 = record2.recordset[0].Contar;
        
        } catch (e) {
            console.log(e);
        }

        await page.waitForTimeout(3000);

        for (var i = 1; i < contagem + 1; i++)
        {
            await page.click(`a:has-text("New")`);
            await page.waitForTimeout(3000);
            if (i <= 10) await page.fill('#tseditName', 'Produto0' + i);
            else await page.fill('#tseditName', 'Produto' + i);
            await page.waitForTimeout(3000);
            const primeiro = await page.getByTitle('Constant').first();
            if (primeiro) primeiro.click();
            await page.waitForTimeout(3000);
            const primeiro_segundo = await page.locator('.bi-tag-fill').first();
            if (primeiro_segundo) primeiro_segundo.click();
            await page.waitForTimeout(3000);
            if (i <= 10) await page.fill('#contentPage_Picker_CounterTagID_Name_TextBox', templatetags + '.Cons.ContadorProduto0' + i);
            else await page.fill('#contentPage_Picker_CounterTagID_Name_TextBox', templatetags + '.Cons.ContadorProduto' + i);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Picker_CounterTagID_Find_Button');
            await page.waitForTimeout(3000);
            const clicarbutkpi = await page.locator(`button:has-text("Assign")`).first();
            if (clicarbutkpi) clicarbutkpi.click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditMaxPlusTagConstant_Constant','999999');
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Advanced")`);
            await page.waitForTimeout(3000);
            await page.fill('#tseditRolloverTagConstant_Constant','999999');
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
        }
    }

    await page.waitForTimeout(3000);

    await page.click(`a:text("  Set KPI Capture Scheme")`);
    await page.waitForTimeout(3000);
    await page.selectOption('#contentPage_slice2_SchemeDropDownList', capture_scheme);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_slice2_AssignButton');
    await page.waitForTimeout(3000);
    await page.click(`a:text("  KPI Capture Tags")`);
    await page.waitForTimeout(3000);
    const split = await page.getByTitle('Split').first();
    if (split) await split.click();

    await page.waitForTimeout(3000);

    const va5 = await page.locator('.fa-edit').nth(1);
    const vatextoHandle5 = await va5.first();
    await vatextoHandle5.click();

    await page.waitForTimeout(3000);

    await page.click('#contentPage_tseditCaptureTagID_Picker');
    await page.waitForTimeout(3000);
    await page.fill('#contentPage_Picker_CaptureTagID_Name_TextBox',templatetags + '.Prod.Split');
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Picker_CaptureTagID_Find_Button');
    await page.waitForTimeout(3000);
    const clicarbutkpi = await page.locator(`button:has-text("Assign")`).first();
    if (clicarbutkpi) clicarbutkpi.click();
    await page.waitForTimeout(3000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(3000);
    await page.click(`a:text("${name}")`);
    await page.waitForTimeout(3000);
    if (GroupName)
    {
        for (var i = 0; i < GroupName.length; i++)
        {
            await page.click(`a:text("  Event Definitions")`);
            await page.waitForTimeout(3000);
            await page.click(`a:text("  New Group")`);
            await page.waitForTimeout(3000);
            await page.fill('#tseditName', GroupName[i]);
            await page.click('#contentPage_Save_Button');

            await page.waitForTimeout(3000);

            const va11 = await page.getByTitle(GroupName[i]).first();
            const EventHandler11 = await va11.first();
            await EventHandler11.click();
            await page.waitForTimeout(5000);
            const va10 = await page.locator(`a:text("  New")`).nth(1);
            const EventHandler12 = await va10.first();
            await EventHandler12.click();
            await page.waitForTimeout(5000);
            await page.fill('#tseditName',EventName[i]);
            await page.waitForTimeout(3000);
            if (EventDefinitionType[i]) await page.selectOption('#tseditEventDefinitionTypeID', EventDefinitionType[i]);
            await page.waitForTimeout(3000);
            await page.fill('#tseditKey', final_keyEventDefinition);
            if (Priority[i]) await page.fill('#tseditPriority', Priority[i]);
            if (TriggerwhenEquals[i]) await page.fill('#tseditTriggerWhenEquals', TriggerwhenEquals[i]);
            await page.waitForTimeout(3000);
            await page.click(`a:text("OEE")`);
            await page.waitForTimeout(3000);
            if (OEEEventType[i]) await page.selectOption('#tseditOeeEventType', OEEEventType[i]);
            await page.waitForTimeout(3000);
            await page.click(`a:text("Split")`);
            await page.waitForTimeout(3000);
            if (ReEvaluateSystemEventonStart[i]) await page.click('#tseditReEvaluateSystemEventOnStart');
            if (ReEvaluateSystemEventonEnd[i]) await page.click('#tseditReEvaluateSystemEventOnEnd');
            await page.waitForTimeout(3000);
            await page.click(`a:text("Event")`);
            await page.waitForTimeout(3000);
            if (ShowForAcknowledge[i]) await page.selectOption('#tseditShowForAcknowledge', ShowForAcknowledge[i]);
            await page.waitForTimeout(3000);
            await page.click(`a:text("Advanced")`);
            await page.waitForTimeout(3000);
            if (MTBFType[i]) await page.selectOption('#tseditMtbfType', MTBFType[i]);
            await page.waitForTimeout(3000);
            if (Duration[i]) await page.fill('#tseditDurationSeconds', Duration[i]);
            await page.waitForTimeout(3000);
            if (IsolationType[i]) await page.selectOption('#tseditEventIsolationType', IsolationType[i]);
            await page.waitForTimeout(3000);
            await page.click(`a:text("Definições")`);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeFailure[i]) await page.fill('#tseditcp_CPS0000000039_CP0000000321', CP_EventDefinitionKey_ForMTBFTypeFailure[i]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeNONFailure[i]) await page.fill('#tseditcp_CPS0000000039_CP0000000322', CP_EventDefinitionKey_ForMTBFTypeNONFailure[i]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeExcluded[i]) await page.fill('#tseditcp_CPS0000000039_CP0000000323', CP_EventDefinitionKey_ForMTBFTypeExcluded[i]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionIDLigada[i]) await page.fill('#tseditcp_CPS0000000039_CP0000000345', CP_EventDefinitionIDLigada[i]);
            await page.waitForTimeout(3000);
            if (CP_TagEventoCodigoAutomacao[i]) await page.fill('#tseditcp_CPS0000000039_CP0000000367', CP_TagEventoCodigoAutomacao[i]);
            await page.waitForTimeout(3000);
            if (CP_LoockupSetKeyCategoriaEventosAuto[i]) await page.fill('#tseditcp_CPS0000000039_CP0000000368', CP_LoockupSetKeyCategoriaEventosAuto[i]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
        }
    }

    await page.close();

});