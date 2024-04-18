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
    //console.log(dadosExcel);

    // ------------------------------Recolher dados------------------------------

    // --------Recolha de dados--------

    let protocolo;
    let templatetag;
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
                case 0:
                    protocolo = segundaLinha['Campos Automação'] as string;
                    templatetag = segundaLinha['Template Tag'] as string;
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
        console.log('TEMPLATE DA TAGGGGG:' + templatetag);
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

    let tipo, locationname, area;

    // Verifica se os dados foram lidos corretamente
    if (dadosExcel) {
        // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
        const segundaLinha: LinhaExcel = dadosExcel[0] as LinhaExcel;

        // Por exemplo, para acessar um valor específico de uma coluna, você pode usar a chave correspondente ao cabeçalho
        tipo = segundaLinha['Tipo Máquina'] as string;
        locationname = segundaLinha['Nome da Location'] as string;
        area = segundaLinha['Area da Maquina'] as string;
        console.log(tipo);
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

    console.log('erroooooooo');
    console.log('C:\\Users\\' + user + '\\Desktop\\Tags (1).xlsx');
    console.log('erroooooooo');


    await page.waitForTimeout(3000);


    // Definindo o tipo de uma linha do Excel
    type LinhaExcel2 = Record<string, string | null>;

    // Função para ler o arquivo Excel
    function lerArquivoExcel2(nomeArquivo: string): LinhaExcel2[] {
        // Carrega o arquivo
        const workbook = XLSX.readFile(nomeArquivo);

        // Pega a primeira planilha do arquivo
        const primeiraPlanilha = workbook.Sheets[workbook.SheetNames[0]];

        // Converte os dados da planilha em um objeto JSON
        const dados = XLSX.utils.sheet_to_json(primeiraPlanilha, { header: 1 }) as string[][];

        // Extrai os cabeçalhos da primeira linha
        const colunas = dados[0];

        // Inicializa um array para armazenar os dados
        const dadosFormatados: LinhaExcel2[] = [];

        // Itera sobre as linhas de dados, começando da segunda linha
        for (let i = 1; i < dados.length; i++) {
            const linha: LinhaExcel2 = {};
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
    const dadosExcel2 = lerArquivoExcel2('C:\\Users\\' + user + '\\Desktop\\CriarSubMaquina.xlsx');
    //console.log(dadosExcel2);

    // ------------------------------------------

        // ------------------------------------------

        let templatetags: any[] = [], taggroup: any[] = [], functiondefinition_name: any[] = [], functiondefinition_type: any[] = [], functiondefinition_script: any[] = [], name_vetor: any[] = [], schedule_vetor: any[] = [], script_vetor: any[] = [], numero_maquina_vetor: any[] = [], protocolo_automacao_vetor: any[] = [], alternate_name_vetor: any[] = [], rejeitados_vetor: any[] = [], consumos_automaticos_vetor: any[] = [], capture_scheme_vetor: any[] = [], ProductSet_vetor: any[] = [];
        let permanecer: boolean = true;
        let temp;
        var i = -1, j = 0, soma2;
    
        if (dadosExcel2) {
            while (permanecer)
            {
                // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
                const linhatemp: LinhaExcel2 = dadosExcel2[j] as LinhaExcel2;
    
                temp = linhatemp['General'] as string;
    
                soma2 = j + i;
                const segundaLinha2: LinhaExcel2 = dadosExcel2[soma2] as LinhaExcel2;
    
                if (temp != null)
                {
                    i++;
                    switch (soma2) {
                        case j:
                            templatetags.push(segundaLinha2['Template Tag'] as string);
                            taggroup.push(segundaLinha2['Tag Group'] as string);
                            break;
                        case j + 1:
                            name_vetor.push(segundaLinha2['General'] as string);
                            script_vetor.push(segundaLinha2['Advanced'] as string);
                            numero_maquina_vetor.push(segundaLinha2['Maquina'] as string);
                            alternate_name_vetor.push(segundaLinha2['Notes'] as string);
                            rejeitados_vetor.push(segundaLinha2['KPI'] as string);
                            functiondefinition_name.push(segundaLinha2['Function Definition'] as string);
                            break;
                        case j + 3:
                            schedule_vetor.push(segundaLinha2['General'] as string);
                            protocolo_automacao_vetor.push(segundaLinha2['Maquina'] as string);
                            consumos_automaticos_vetor.push(segundaLinha2['KPI'] as string);
                            consumos_automaticos_vetor.push(segundaLinha2['KPI'] as string);
                            functiondefinition_type.push(segundaLinha2['Function Definition'] as string);
                            break;
                        case j + 5:
                            capture_scheme_vetor.push(segundaLinha2['KPI'] as string);
                            ProductSet_vetor.push(segundaLinha2['General'] as string);
                            consumos_automaticos_vetor.push(segundaLinha2['KPI'] as string);
                            functiondefinition_script.push(segundaLinha2['Function Definition'] as string);
                            j += 8;
                            i = -1;
                        default:
                            break;
                    }
                }
                else
                {
                    permanecer = false;
                }
            }
        } else {
            console.log("Não foi possível ler os dados do arquivo Excel.");
        }
        console.log('-----------------' + name_vetor + '-----------------');
        console.log('-----------------' + schedule_vetor + '-----------------');
        console.log('-----------------' + numero_maquina_vetor + '-----------------');
        console.log('-----------------' + script_vetor + '-----------------');
        console.log('-----------------12345 ' + templatetags + '-----------------');
    
        // var linha5 = 1, soma3 = 0;
        // let GroupName: any[] = [];
    
        // while (1 < 2)
        // {
        //     const segundaLinha: LinhaExcel2 = dadosExcel2[linha5] as LinhaExcel2;
        //     const segundaLinha2: LinhaExcel2 = dadosExcel2[linha5 + 1] as LinhaExcel2;
        //     const segundaLinha3: LinhaExcel2 = dadosExcel2[linha5 + 2] as LinhaExcel2;
        //     const segundaLinha4: LinhaExcel2 = dadosExcel2[linha5 + 5] as LinhaExcel2;
        //     const prov = segundaLinha['Event Definition (Group)'];
        //     const prov2 = segundaLinha2['Event Definition (Group)'];
        //     const prov3 = segundaLinha3['Event Definition (Group)'];
        //     const prov4 = segundaLinha4['Event Definition (Group)'];
        //     soma3++;
        //     if (prov == null && prov2 == null && prov3 == null && prov4 == null) break;
        //     else if (segundaLinha['Event Definition (Group)'] != 'Name') GroupName.push(segundaLinha['Event Definition (Group)'] as string);
        //     linha5++;
        // }
        console.log('soma2: ' + soma2);
        //console.log('soma3: ' + soma3);
        
        var linha5 = 1, soma3 = 0, EventDefinition_acumuladora = 0;
    
        while (1 < 2)
        {
            const segundaLinha: LinhaExcel2 = dadosExcel2[linha5] as LinhaExcel2;
            const segundaLinha2: LinhaExcel2 = dadosExcel2[linha5 + 1] as LinhaExcel2;

            const prov = segundaLinha['Event Definition (Event)'];
            const prov2 = segundaLinha2['Event Definition (Event)'];

            soma3++;
            if (prov == null && prov2 == null) break;
            else if (segundaLinha['Event Definition (Event)'] != 'Name') EventDefinition_acumuladora++;
            linha5++;
        }
    
        let EventName: any[] = [], EventDefinitionType: any[] = [], Priority: any[] = [], TriggerwhenEquals: any[] = [], OEEEventType: any[] = [], ReEvaluateSystemEventonStart: any[] = [], ReEvaluateSystemEventonEnd: any[] = [], ShowForAcknowledge: any[] = [], MTBFType: any[] = [], Duration: any[] = [], IsolationType: any[] = [], CP_EventDefinitionKey_ForMTBFTypeFailure: any[] = [], CP_EventDefinitionKey_ForMTBFTypeNONFailure: any[] = [], CP_EventDefinitionKey_ForMTBFTypeExcluded: any[] = [], CP_EventDefinitionIDLigada: any[] = [], CP_TagEventoCodigoAutomacao: any[] = [], CP_LoockupSetKeyCategoriaEventosAuto: any[] = [];
        var forms = 0;

        for (var i = 1; i < soma2; i++)
        {
            const segundaLinha: LinhaExcel2 = dadosExcel2[i] as LinhaExcel2;
            if (i%8 != 0)
            {
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
            else forms++;
        }
        console.log('forms: ' + forms);
        console.log('ReEvaluateSystemEventonStart: ' + ReEvaluateSystemEventonStart);
        console.log(EventName);
        console.log(Priority);
        console.log(EventDefinitionType);
        console.log(MTBFType);
        console.log(CP_TagEventoCodigoAutomacao);
    
        // ------------------------------------------

    // ------------------------------------------

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

    await page.waitForTimeout(3000);

    //------------Fim do Ir buscar site------------

    // ------------------------------Gerar Key's------------------------------

    let keys: any[] = [];

    await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');

    await page.getByLabel('Login').fill('kt0032'); //utilizador kt
    await page.getByLabel('Password').click();
    await page.getByLabel('Password').fill('12345'); // password
    await page.getByRole('button', { name: 'Sign In' }).click();

    await page.waitForTimeout(5000);

    await page.click('#contentPage_ctl25');
    await page.click('.btn-item-key-btn_GerarKey');
    await page.waitForTimeout(3000);
    const key2 = await page.locator('#contentPage_ctl04').textContent();
    let final_key2;
    if (key2) final_key2 = key2.trim();
    await page.waitForTimeout(5000);

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

    await page.goto('http://ktmesapp04/TS/pages/home/config/systems/?S2ID=2829&S3Key=Item.SystemBatchSubComposite&S3ID=2829&c=ETS.Configuration.Slices.Nav&S1Key=Item.System&S1ID=2827');
    var jota = 0;

        await page.click(`a:text("  Event Definitions")`);
        await page.waitForTimeout(3000);

        for (var j = 0; j < EventDefinition_acumuladora; j++)
        {
            const novo = await page.locator(`a:has-text("New")`).nth(1);
            const eventhandler = novo.first();
            eventhandler.click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditName', EventName[jota]);
            await page.waitForTimeout(3000);
            if (EventDefinitionType[jota] != null) await page.selectOption('#tseditEventDefinitionTypeID', EventDefinitionType[jota]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_tseditTriggerTagID_Picker');
            await page.waitForTimeout(3000);
            await page.fill('#contentPage_Picker_TriggerTagID_Name_TextBox', templatetags[i] + '.Evento.EstadoMaquina');
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Picker_TriggerTagID_Find_Button');
            await page.waitForTimeout(3000);
            await page.click('button:has-text("Assign")');
            await page.waitForTimeout(3000);
            if (Priority[jota] != null) await page.fill('#tseditPriority', Priority[jota]);
            await page.waitForTimeout(3000);
            if (TriggerwhenEquals[jota] != '') await page.fill('#tseditTriggerWhenEquals', TriggerwhenEquals[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("OEE")`);
            await page.waitForTimeout(3000);
            if (OEEEventType[jota] != null) await page.selectOption('#tseditOeeEventType', OEEEventType[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Split")`);
            await page.waitForTimeout(3000);
            if (ReEvaluateSystemEventonStart[jota] == "Sim") await page.click('#tseditReEvaluateSystemEventOnStart');
            await page.waitForTimeout(3000);
            if (ReEvaluateSystemEventonEnd[jota] == "Sim") await page.click('#tseditReEvaluateSystemEventOnEnd');
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Event")`);
            await page.waitForTimeout(3000);
            if (ShowForAcknowledge[jota] != null) await page.selectOption('#tseditShowForAcknowledge', ShowForAcknowledge[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Advanced")`);
            await page.waitForTimeout(3000);
            if (MTBFType[jota] != null) await page.selectOption('#tseditMtbfType', MTBFType[jota]);
            await page.waitForTimeout(3000);
            if (Duration[jota] != null) await page.fill('#tseditDurationSeconds', Duration[jota]);
            await page.waitForTimeout(3000);
            if (IsolationType[jota] != null) await page.selectOption('#tseditEventIsolationType', IsolationType[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Definições")`);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeFailure[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000321', CP_EventDefinitionKey_ForMTBFTypeFailure[jota]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeNONFailure[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000322', CP_EventDefinitionKey_ForMTBFTypeNONFailure[jota]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeExcluded[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000323', CP_EventDefinitionKey_ForMTBFTypeExcluded[jota]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionIDLigada[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000345', CP_EventDefinitionIDLigada[jota]);
            await page.waitForTimeout(3000);
            if (CP_TagEventoCodigoAutomacao[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000367', CP_TagEventoCodigoAutomacao[jota]);
            await page.waitForTimeout(3000);
            if (CP_LoockupSetKeyCategoriaEventosAuto[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000368', CP_LoockupSetKeyCategoriaEventosAuto[jota]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
            jota++;
        }
        jota += 6;
        await page.click(`a:text("${name}")`);
        await page.waitForTimeout(3000);

    await page.close();

});