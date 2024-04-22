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
        console.log('TEMPLATE DA TAG:' + templatetag);
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
                            j += 13;
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
    
        let EventName: any[] = [], EventDefinitionType: any[] = [], triggertag: any[] = [], Priority: any[] = [], TriggerwhenEquals: any[] = [], OEEEventType: any[] = [], ReEvaluateSystemEventonStart: any[] = [], ReEvaluateSystemEventonEnd: any[] = [], ShowForAcknowledge: any[] = [], MTBFType: any[] = [], Duration: any[] = [], IsolationType: any[] = [], CP_EventDefinitionKey_ForMTBFTypeFailure: any[] = [], CP_EventDefinitionKey_ForMTBFTypeNONFailure: any[] = [], CP_EventDefinitionKey_ForMTBFTypeExcluded: any[] = [], CP_EventDefinitionIDLigada: any[] = [], CP_TagEventoCodigoAutomacao: any[] = [], CP_LoockupSetKeyCategoriaEventosAuto: any[] = [];
        var forms = 0;

        for (var i = 1; i < soma2; i++)
        {
            const segundaLinha: LinhaExcel2 = dadosExcel2[i] as LinhaExcel2;
            if (i%13 != 0)
            {
                EventName.push(segundaLinha['Event Definition (Event)'] as string);
                EventDefinitionType.push(segundaLinha['1'] as string);
                triggertag.push(segundaLinha['17'] as string);
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
        }EventName
        console.log('forms: ' + forms);
        console.log('ReEvaluateSystemEventonStart: ' + ReEvaluateSystemEventonStart);
        console.log(EventName);
        console.log(Priority);
        console.log(EventDefinitionType);
        console.log(MTBFType);
        console.log(CP_TagEventoCodigoAutomacao);

        var linha5 = 1, soma3 = 0, EventDefinition_acumuladora = 0, max = 0;
        var tent: any[] = [], dif: any[] = [];
        let isTrue: boolean = true;
    
        while (1 < 2)
        {
            const segundaLinha: LinhaExcel2 = dadosExcel2[linha5] as LinhaExcel2;
            const segundaLinha2: LinhaExcel2 = dadosExcel2[linha5 + 1] as LinhaExcel2;

            const prov = segundaLinha['Event Definition (Event)'];
            const prov2 = segundaLinha2['Event Definition (Event)'];

            soma3++;
            if (prov == null && prov2 == null && isTrue)
            {
                tent.push(EventDefinition_acumuladora);
                EventDefinition_acumuladora = 0;
                max++;
                isTrue = false;
            }
            else if (segundaLinha['Event Definition (Event)'] != 'Name' && segundaLinha['Event Definition (Event)'] != null)
            {
                EventDefinition_acumuladora++;
                isTrue = true;
            }
            if (max == forms + 1) break;
            linha5++;
            console.log('------------------------------------------------------------------------------------------------------');
            console.log('linha5: ' + linha5);
            console.log('------------------------------------------------------------------------------------------------------');
        }
        console.log('Tent: ' + tent);
        // -----------------Inicio Filragem-----------------

        let EventName_filtrado = EventName.filter(elemento => elemento !== null);
        let EventDefinitionType_filtrado = EventDefinitionType.filter(elemento => elemento !== null);
        let triggertag_filtrado = triggertag.filter(elemento => elemento !== null);
        let Priority_filtrado = Priority.filter(elemento => elemento !== null);
        let TriggerwhenEquals_filtrado = TriggerwhenEquals.filter(elemento => elemento !== null);
        let OEEEventType_filtrado = OEEEventType.filter(elemento => elemento !== null);
        let ReEvaluateSystemEventonStart_filtrado = ReEvaluateSystemEventonStart.filter(elemento => elemento !== null);
        let ReEvaluateSystemEventonEnd_filtrado = ReEvaluateSystemEventonEnd.filter(elemento => elemento !== null);
        let ShowForAcknowledge_filtrado = ShowForAcknowledge.filter(elemento => elemento !== null);
        let MTBFType_filtrado = MTBFType.filter(elemento => elemento !== null);
        let Duration_filtrado = Duration.filter(elemento => elemento !== null);
        let IsolationType_filtrado = IsolationType.filter(elemento => elemento !== null);
        let CP_EventDefinitionKey_ForMTBFTypeFailure_filtrado = CP_EventDefinitionKey_ForMTBFTypeFailure.filter(elemento => elemento !== null);
        let CP_EventDefinitionKey_ForMTBFTypeNONFailure_filtrado = CP_EventDefinitionKey_ForMTBFTypeNONFailure.filter(elemento => elemento !== null);
        let CP_EventDefinitionKey_ForMTBFTypeExcluded_filtrado = CP_EventDefinitionKey_ForMTBFTypeExcluded.filter(elemento => elemento !== null);
        let CP_EventDefinitionIDLigada_filtrado = CP_EventDefinitionIDLigada.filter(elemento => elemento !== null);
        let CP_TagEventoCodigoAutomacao_filtrado = CP_TagEventoCodigoAutomacao.filter(elemento => elemento !== null);
        let CP_LoockupSetKeyCategoriaEventosAuto_filtrado = CP_LoockupSetKeyCategoriaEventosAuto.filter(elemento => elemento !== null);

        // -----------------Fim Filragem-----------------
    
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

    await page.goto('http://' + ambiente_final + '/TS/pages/' + site +'/config/tags/import');

    await page.waitForTimeout(3000);

    //-----------------Criar Tags------------------

    //Importar ficheiro excel

    // Localize o input de arquivo e insira o caminho do arquivo Excel
    const inputFile = await page.$('input[type="file"]');
    if (inputFile) await inputFile.setInputFiles('C:\\Users\\' + user + '\\Desktop\\Tags (1).xlsx');
    await page.waitForTimeout(3000);
    await page.click('#Buttons_Import');

    await page.waitForTimeout(5000);

    // ------------Inicio Quantidade------------

    var record1;
    var quantidade;
    try {
        await sql.connect(config)
        record1 = await sql.query`select count(id) as Quantidade from ttag where [Name] like '%' + ${templatetags[0].toString()} + '.BatchReal.' + ${functiondefinition_name[0].toString()} + '%.Ativo%'` // select distinct
        quantidade = record1.recordset[0].Quantidade;
    
    } catch (e){
        console.log(e);
    }

    console.log('Quantidade: ' + quantidade);

    for (var i = 0; i < (quantidade) * (forms + 1); i++)
    {
        await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');
        await page.waitForTimeout(3000);
        await page.click('#contentPage_ctl17');
        await page.click('.btn-item-key-btn_GerarKey');
        await page.waitForTimeout(3000);
        const key3 = await page.locator('#contentPage_ctl04').textContent();
        let final_key3;
        if (key3) final_key3 = key3.trim();
        await page.waitForTimeout(3000);
        keys.push(final_key3 as string);
        await page.waitForTimeout(3000);
    }
    console.log('Keys: ' + keys);

    // ------------Fim Quantidade------------

    await page.goto('http://' + ambiente_final + '/TS/pages/' + site +'/config/tags/');

    for (var j = 0; j < tags.length; j++)
    {

        await page.locator('#contentPage_slice1_TreeList_Tree_TreeView').getByText(tags[j]).click();

    }

    await page.locator('#contentPage_slice1_TreeList_Tree_TreeView').getByText('Geral').click();

    await page.waitForTimeout(3000);

    await page.click(`li:has-text("Evento")`);

    await page.waitForTimeout(3000);

    await page.click('#tsslice-index-2 .fa-plus');
    await page.waitForTimeout(3000);
    const ScriptTag = await page.getByTitle('Script Tag').first();
    if (ScriptTag) ScriptTag.click();
    await page.waitForTimeout(3000);
    await page.fill('#tseditName', templatetag + '.Geral.Evento.FalhaComunicacoes');
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
    const tag = await page.getByTitle(templatetag + '.Geral.Evento.FalhaComunicacoes').first();
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

    // if (Tags["CHK_Bath.Teste1.Evento.EstadoMaquina"].Quality != 0"] && Tags["CHK_Bath.Teste2.Evento.EstadoMaquina"].Quality != 0"] && Tags["CHK_Bath.Teste3.Evento.EstadoMaquina"].Quality != 0) return 0; else return 1;

    let texto1 = 'if (';
    for (var i = 0; i < forms + 1; i++) {
        if (i != forms) texto1 += 'Tags["' + templatetags[i] + '.Evento.EstadoMaquina"].Quality != 0 && ';
        else texto1 += 'Tags["' + templatetags[i] + '.Evento.EstadoMaquina"].Quality != 0';
    }
    let texto2 = ') return 0; else return 1;';
    let textofinal = texto1 + texto2;
    await page.fill('#contentPage_Editor_Code', textofinal);

    await page.waitForTimeout(5000);

    await page.click('.tsoperation-toolbar-saveandclose');
    await page.waitForTimeout(5000);

    await page.goto('http://' + ambiente_final + '/TS/pages/' + site + '/config/tags/');

    await page.waitForTimeout(3000);

    for (var j = 0; j < tags.length; j++)
    {

        await page.locator('#contentPage_slice1_TreeList_Tree_TreeView').getByText(tags[j]).click();

    }

    await page.waitForTimeout(3000);
    
    await page.locator('#contentPage_slice1_TreeList_Tree_TreeView').getByText('Geral').click();

    await page.waitForTimeout(3000);

    await page.click(`li:has-text("Evento")`);

    await page.waitForTimeout(3000);

    await page.click('#tsslice-index-2 .fa-plus');
    await page.waitForTimeout(3000);
    const compare = await page.getByTitle('Compare Tag').first();
    if (compare) compare.click();
    await page.waitForTimeout(3000);
    await page.fill('#tseditName', templatetag + '.Geral.Evento.HeartBeatUpdate');
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
    await page.fill('#contentPage_Picker_RightTagID_Name_TextBox', templatetag + '.Geral.Evento.HeartBeatMaquina');
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

    //-----------------Criar Location------------------

    await page.click(`li:has-text("Locations")`);
    await page.waitForTimeout(3000);
    // for (var i = 0; i < tags.length; i++)
    // {
    //     await page.getByText(new RegExp("^" + location[i] + "$", "i")).click();
    //     await page.waitForTimeout(3000);
    // }
    for (var i = 0; i < location.length; i++)
    {

        await page.locator('#contentPage_slice1_TreeList_Tree_TreeView').getByText(location[i]).click();

    }
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("New Child")`);
    await page.waitForTimeout(3000);
    await page.fill('#tseditName',locationname);
    await page.waitForTimeout(2000);
    if (key2) await page.fill('#tseditUniqueID', final_key2);
    await page.waitForTimeout(2000);
    await page.selectOption('#tseditLocationTypeID','LT_Maquinas');
    await page.waitForTimeout(2000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(5000);

    // --------------Criar Máquina--------------

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

    if (script) await page.fill('#tseditTemplateTagPrefix', templatetag);

    await page.waitForTimeout(3000);
    await page.click('#contentPage_tseditLocationID_Picker');
    await page.waitForTimeout(5000);
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

    // ----------------Criar SubMáquina----------------

    var jota = 0, capa = 0, capatemp = quantidade, keys_acumuladaora = 0;

    console.log(forms);
    for (var i = 0; i < forms + 1; i++)
    {
        var functiondefinition_acumuladora = 1;
        console.log('Formulário: ' + forms);
        await page.click(`a:has-text("New")`);

        await page.waitForTimeout(3000);
        await page.fill('#tseditName', name_vetor[i]);
        await page.waitForTimeout(3000);

        await page.click('#contentPage_tseditJobTagID_Picker');
        await page.fill('#contentPage_Picker_JobTagID_Name_TextBox',templatetags[i] + '.Ord.Ordem');
        await page.click('#contentPage_Picker_JobTagID_Find_Button');
        await page.waitForTimeout(2000);
        const clicarbut = await page.locator(`button:has-text("Assign")`).first();
        await page.waitForTimeout(2000);
        if (clicarbut) clicarbut.click();
        await page.waitForTimeout(3000);

        await page.click('#contentPage_tseditBatchTagID_Picker');
        await page.fill('#contentPage_Picker_BatchTagID_Name_TextBox', templatetags[i] + '.Batch.BatchTag');
        await page.click('#contentPage_Picker_BatchTagID_Find_Button');
        await page.waitForTimeout(2000);
        const clicarbut2 = await page.locator(`button:has-text("Assign")`).first();
        await page.waitForTimeout(2000);
        if (clicarbut2) clicarbut2.click();
        await page.waitForTimeout(3000);
        const segundo = await page.locator(`li:has-text("Product")`).nth(6);
        const vatextoHandle3 = await segundo.first();
        await vatextoHandle3.click();
        await page.waitForTimeout(3000);

        await page.click('#contentPage_tseditProductTagID_Picker');
        await page.fill('#contentPage_Picker_ProductTagID_Name_TextBox', templatetags[i] + '.Batch.ReceitaProduto');
        await page.click('#contentPage_Picker_ProductTagID_Name_TextBox');
        await page.waitForTimeout(2000);
        const clicarbut3 = await page.locator(`button:has-text("Assign")`).first();
        await page.waitForTimeout(2000);
        if (clicarbut3) clicarbut3.click();
        await page.waitForTimeout(3000);

        await page.click(`li:has-text("Advanced")`);
        await page.waitForTimeout(3000);
        await page.fill('#tseditScriptClassName', script_vetor[i]);
        await page.waitForTimeout(3000);
        await page.click(`li:has-text("Maquina")`);
        await page.waitForTimeout(3000);
        await page.fill('#tseditcp_CPS0000000013_CP0000000083', numero_maquina_vetor[i]);
        await page.waitForTimeout(3000);
        await page.click('#contentPage_Save_Button');
        await page.waitForTimeout(3000);
        await page.click(`div:text("${name_vetor[i]}")`);
        await page.waitForTimeout(3000);
        await page.click(`a:text("  Function Definitions")`);
        await page.waitForTimeout(3000);

        console.log(EventDefinition_acumuladora);

        // -----------Function Parameters-----------

        for (var k = capa; k < capatemp; k++)
        {
            const novo = await page.locator(`a:has-text("New")`).nth(1);
            const eventhandler = novo.first();
            eventhandler.click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditName', functiondefinition_name[i] + functiondefinition_acumuladora.toString());
            await page.waitForTimeout(3000);
            await page.selectOption('#tseditFunctionDefinitionTypeID', functiondefinition_type[i]);
            await page.waitForTimeout(3000);
            await page.fill('#tseditKey', keys[keys_acumuladaora]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_tseditTriggerTagID_Picker');
            await page.waitForTimeout(3000);
            await page.fill('#contentPage_Picker_TriggerTagID_Name_TextBox', templatetags[i] + '.BatchReal.' + functiondefinition_name[i] + functiondefinition_acumuladora.toString() + '.Ativo');
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Picker_TriggerTagID_Find_Button');
            await page.waitForTimeout(3000);
            await page.click('button:has-text("Assign")');
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Advanced")`);
            await page.waitForTimeout(3000);
            await page.fill('#tseditScriptClassName', functiondefinition_script[i]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
            functiondefinition_acumuladora++;
            keys_acumuladaora++;
        }

        capa = k + 1;
        capatemp += quantidade;

        await page.click(`a:text("  Event Definitions")`);
        await page.waitForTimeout(3000);

        for (var j = 0; j < tent[i]; j++)
        {
            const novo = await page.locator(`a:has-text("New")`).nth(1);
            const eventhandler = novo.first();
            eventhandler.click();
            await page.waitForTimeout(3000);
            await page.fill('#tseditName', EventName_filtrado[jota]);
            await page.waitForTimeout(3000);
            if (EventDefinitionType_filtrado[jota] != null) await page.selectOption('#tseditEventDefinitionTypeID', EventDefinitionType_filtrado[jota]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_tseditTriggerTagID_Picker');
            await page.waitForTimeout(3000);
            await page.fill('#contentPage_Picker_TriggerTagID_Name_TextBox', templatetags[i] + '.' + triggertag_filtrado[i]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Picker_TriggerTagID_Find_Button');
            await page.waitForTimeout(3000);
            await page.click('button:has-text("Assign")');
            await page.waitForTimeout(3000);
            if (Priority_filtrado[jota] != null) await page.fill('#tseditPriority', Priority_filtrado[jota]);
            await page.waitForTimeout(3000);
            if (TriggerwhenEquals_filtrado[jota] != '') await page.fill('#tseditTriggerWhenEquals', TriggerwhenEquals_filtrado[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("OEE")`);
            await page.waitForTimeout(3000);
            if (OEEEventType_filtrado[jota] != null) await page.selectOption('#tseditOeeEventType', OEEEventType_filtrado[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Split")`);
            await page.waitForTimeout(3000);
            if (ReEvaluateSystemEventonStart_filtrado[jota] == "Sim") await page.click('#tseditReEvaluateSystemEventOnStart');
            await page.waitForTimeout(3000);
            if (ReEvaluateSystemEventonEnd_filtrado[jota] == "Sim") await page.click('#tseditReEvaluateSystemEventOnEnd');
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Event")`);
            await page.waitForTimeout(3000);
            if (ShowForAcknowledge_filtrado[jota] != null) await page.selectOption('#tseditShowForAcknowledge', ShowForAcknowledge_filtrado[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Advanced")`);
            await page.waitForTimeout(3000);
            if (MTBFType_filtrado[jota] != null) await page.selectOption('#tseditMtbfType', MTBFType_filtrado[jota]);
            await page.waitForTimeout(3000);
            if (Duration_filtrado[jota] != null) await page.fill('#tseditDurationSeconds', Duration_filtrado[jota]);
            await page.waitForTimeout(3000);
            if (IsolationType_filtrado[jota] != null) await page.selectOption('#tseditEventIsolationType', IsolationType_filtrado[jota]);
            await page.waitForTimeout(3000);
            await page.click(`li:has-text("Definições")`);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeFailure_filtrado[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000321', CP_EventDefinitionKey_ForMTBFTypeFailure[jota]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeNONFailure_filtrado[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000322', CP_EventDefinitionKey_ForMTBFTypeNONFailure[jota]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionKey_ForMTBFTypeExcluded_filtrado[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000323', CP_EventDefinitionKey_ForMTBFTypeExcluded_filtrado[jota]);
            await page.waitForTimeout(3000);
            if (CP_EventDefinitionIDLigada_filtrado[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000345', CP_EventDefinitionIDLigada_filtrado[jota]);
            await page.waitForTimeout(3000);
            if (CP_TagEventoCodigoAutomacao_filtrado[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000367', CP_TagEventoCodigoAutomacao_filtrado[jota]);
            await page.waitForTimeout(3000);
            if (CP_LoockupSetKeyCategoriaEventosAuto_filtrado[jota] != null) await page.fill('#tseditcp_CPS0000000039_CP0000000368', CP_LoockupSetKeyCategoriaEventosAuto_filtrado[jota]);
            await page.waitForTimeout(3000);
            await page.click('#contentPage_Save_Button');
            await page.waitForTimeout(3000);
            jota++;
        }
        //vetorFiltrado

        // while (1 < 2)
        // {
        //     const segundaLinha: LinhaExcel2 = dadosExcel2[linha5] as LinhaExcel2;
        //     const segundaLinha2: LinhaExcel2 = dadosExcel2[linha5 + 1] as LinhaExcel2;

        //     const prov = segundaLinha['Event Definition (Event)'];
        //     const prov2 = segundaLinha2['Event Definition (Event)'];

        //     soma3++;
        //     if (prov == null && prov2 == null)
        //     {
        //         linha5 += Math.abs(linha5 - 13);
        //         tent.push(EventDefinition_acumuladora);
        //         EventDefinition_acumuladora = 0;

        //     }
        //     else if (segundaLinha['Event Definition (Event)'] != 'Name') EventDefinition_acumuladora++;
        //     linha5++;
        // }

        await page.click(`a:text("${name}")`);
        await page.waitForTimeout(3000);
    }

    await page.close();

});