import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';

// Configurações de conexão
const sql = require('mssql');
const config = require('../../../../CRIARMÁQUINA/tests/dbConnection/connection.js');
test('CriarMinhaMáquina', async ({ page }) => {
    
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
    const dadosExcel = lerArquivoExcel('C:\\Users\\LeandroMonteiro\\Desktop\\CriarMáquina.xlsx');
    console.log(dadosExcel);

    // ------------------------------Recolher dados (Não tags)------------------------------

    let site;
    let excel_AUT;
    let maquina;
    let nome_maquina;
    let numero_maquina;
    let tipo;
    let template;
    let nome_location;
    let ProduzItemsDefinitionID;
    let TaxaProducaoTeorica;
    let EstadoMaquina;

    // Verifica se os dados foram lidos corretamente
    if (dadosExcel) {
        // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
        const segundaLinha: LinhaExcel = dadosExcel[0] as LinhaExcel;

        // Por exemplo, para acessar um valor específico de uma coluna, você pode usar a chave correspondente ao cabeçalho
        site = segundaLinha['Site'] as string;
        excel_AUT = segundaLinha['Excel AUT'] as string;
        maquina = segundaLinha['Máquina'] as string;
        nome_maquina = segundaLinha['Nome Máquina'] as string;
        numero_maquina = segundaLinha['Número Máquina'] as string;
        tipo = segundaLinha['Tipo'] as string;
        template = segundaLinha['Template'] as string;
        nome_location = segundaLinha['Nome Location'] as string;
        ProduzItemsDefinitionID = segundaLinha['Ord.ProduzItemsDefinitionID'] as string;
        TaxaProducaoTeorica = segundaLinha['Prod.TaxaProducaoTeorica'] as string;
        EstadoMaquina = segundaLinha['Evento.EstadoMaquina'] as string;
        console.log(site);
        console.log(excel_AUT);
        console.log(maquina);
        console.log(nome_maquina);
        console.log(numero_maquina);
        console.log(tipo);
        console.log(template);
        console.log(nome_location);
        console.log(ProduzItemsDefinitionID);
        console.log(TaxaProducaoTeorica);
        console.log(EstadoMaquina);

    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }

    // ------------------------------Recolher dados (tags)------------------------------

    var linha = 0;
    let ArmazemDestinoProduto: any[] = [];
    let ContentorTipoDestinoProduto: any[] = [];
    let ArmazemOrigemProduto: any[] = [];
    let ContentorOrigemProduto: any[] = [];

    // Verifica se os dados foram lidos corretamente
    if (dadosExcel) {
        for (var i = 0; i < dadosExcel.length; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[linha] as LinhaExcel; 

            // Por exemplo, para acessar um valor específico de uma coluna, você pode usar a chave correspondente ao cabeçalho
            ArmazemDestinoProduto.push(segundaLinha['Prod.ArmazemDestinoProduto'] as string);
            ContentorTipoDestinoProduto.push(segundaLinha['Prod.ContentorTipoDestinoProduto'] as string);
            ArmazemOrigemProduto.push(segundaLinha['Cons.ArmazemOrigemProduto'] as string);
            ContentorOrigemProduto.push(segundaLinha['Cons.ContentorOrigemProduto'] as string);

            linha++;
        }

    } else {
        console.log("Não foi possível ler os dados do arquivo Excel.");
    }
    console.log('-------------------------------');
    console.log(ArmazemDestinoProduto);
    console.log(ContentorTipoDestinoProduto);
    console.log(ArmazemOrigemProduto);
    console.log(ContentorOrigemProduto);

    //Tags restantes (Script e Compare)

    var linha2 = 0;
    let tags: any[] = [];

    // Verifica se os dados foram lidos corretamente
    // if (dadosExcel) {
    //     for (var i = 0; i < dadosExcel.length; i++)
    //     {
    //         // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
    //         const segundaLinha: LinhaExcel = dadosExcel[linha2] as LinhaExcel; 

    //         // Por exemplo, para acessar um valor específico de uma coluna, você pode usar a chave correspondente ao cabeçalho
    //         tags.push(segundaLinha['Caminho Tags'] as string);

    //         linha2++;
    //     }

    // } else {
    //     console.log("Não foi possível ler os dados do arquivo Excel.");
    // }

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

    // --------------------------------

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

    let protocolo: any[] = [];
    let CP_MAQUINA_PERMITE_PARAR_AUTOMACAO: any[] = [];
    let CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO: any[] = [];
    let CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO: any[] = [];
    let CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO: any[] = [];
    let CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE: any[] = [];
    var soma = 0;
    let aut: boolean = true;

    let protocolo_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_PARAR_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO_preenchido: boolean = true;
    let CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE_preenchido: boolean = true;

    if (dadosExcel) {
        for (var i = 0; i < 13; i++)
        {
            // Por exemplo, para armazenar os valores da segunda linha do Excel (índice 1)
            const segundaLinha: LinhaExcel = dadosExcel[linha] as LinhaExcel;

            switch (i) {
                case 3:
                    protocolo.push(segundaLinha['ProtocoloAutomacao'] as string);
                    if (segundaLinha['ProtocoloAutomacao'] == 'Sim') soma++;
                    else protocolo_preenchido = false;
                    break;
                case 5:
                    CP_MAQUINA_PERMITE_PARAR_AUTOMACAO.push(segundaLinha['CP_MAQUINA_PERMITE_PARAR_AUTOMACAO'] as string);
                    if (segundaLinha['CP_MAQUINA_PERMITE_PARAR_AUTOMACAO'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_PARAR_AUTOMACAO_preenchido = false;
                    break;
                case 7:
                    CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO.push(segundaLinha['CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO'] as string);
                    if (segundaLinha['CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO_preenchido = false;
                    break;
                case 9:
                    CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO.push(segundaLinha['CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO'] as string);
                    if (segundaLinha['CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO_preenchido = false;
                    break;
                case 11:
                    CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO.push(segundaLinha['CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO'] as string);
                    if (segundaLinha['CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO_preenchido = false;
                    break;
                case 13:
                    CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE.push(segundaLinha['CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE'] as string);
                    if (segundaLinha['CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE'] == 'Sim') soma++;
                    else CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE_preenchido = false;
                    break;
            
                default:
                    break;
            }
            linha++;
        }
        if (soma == 6) aut = true;

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

    // ------------------------------Gerar Key's------------------------------

    await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');

    await page.getByLabel('Login').fill('kt0032'); //utilizador kt 
    await page.getByLabel('Password').click();
    await page.getByLabel('Password').fill('12345'); // password
    await page.getByRole('button', { name: 'Sign In' }).click();

    await page.click('#contentPage_ctl43');
    await page.click('.btn-item-key-btn_GerarKey');
    await page.waitForTimeout(3000);
    const key = await page.locator('#contentPage_ctl04').textContent();

    await page.waitForTimeout(5000);
    
    await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');

    await page.click('#contentPage_ctl25');
    await page.click('.btn-item-key-btn_GerarKey');
    await page.waitForTimeout(3000);
    const key2 = await page.locator('#contentPage_ctl04').textContent();
    let final_key2
    if (key2) final_key2 = key2.trim();

    // await page.waitForTimeout(3000);

    // await page.goto('http://ktmesapp01/TS/pages/root/dev/osi_teste/pd0000002170/');

    // await page.click('#contentPage_ctl25');
    // await page.click('.btn-item-key-btn_GerarKey');
    // await page.waitForTimeout(3000);
    // const location = await page.locator('#contentPage_ctl25').textContent();

    // ------------------------------Começar Criação de Máquina de Forma Automática------------------------------

    await page.goto(site + 'config/tags/import');

    await page.waitForTimeout(3000);

    //-----------------Criar Tags------------------

    //Importar ficheiro excel

    // Localize o input de arquivo e insira o caminho do arquivo Excel
    const inputFile = await page.$('input[type="file"]');
    if (inputFile) await inputFile.setInputFiles(excel_AUT);
    await page.waitForTimeout(3000);
    await page.click('#Buttons_Import');

    await page.waitForTimeout(5000);

    //Criar Tags restantes(Script e Compare)

    for (var i = 0; i < tags.length; i++)
    {
        await page.getByText(new RegExp("^" + tags[i] + "$", "i")).click();
        await page.waitForTimeout(3000);
    }

    await page.waitForTimeout(3000);

    const elemento2 = await page.$(`[data-nodeid='218']`);
    if (elemento2) elemento2.click();

    await page.waitForTimeout(3000);

    await page.click('#tsslice-index-2 .fa-plus');
    await page.waitForTimeout(3000);
    const ScriptTag = await page.getByTitle('Script Tag').first();
    if (ScriptTag) ScriptTag.click();
    await page.waitForTimeout(3000);
    await page.fill('#tseditName',template + '.Evento.FalhaComunicacoes');
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
    // const va6 = await page.getByTitle(template + '.Evento.FalhaComunicacoes').first();
    // if (va6) va6.click();
    // await page.waitForTimeout(3000);
    // const va5 = await page.locator('.fa-code').nth(1);
    // const vatextoHandle3 = await va5.first();
    // await vatextoHandle3.click();
    // await page.waitForTimeout(3000);
    // await page.evaluate(() => {
    //     const teste = document.querySelector('.view-line');
    //     if (teste) teste.innerHTML = '<span><span class="mtk6">if</span><span class="mtk1">&nbsp;(Tags[</span><span class="mtk20">"CHK.CGL.Madeiras.Multiserra.Multiserra01.Evento.E</span><span class="mtk20">stadoMaquina"</span><span class="mtk1">].Quality&nbsp;!=&nbsp;</span><span class="mtk7">0</span><span class="mtk1">)&nbsp;</span><span class="mtk6">return</span><span class="mtk1">&nbsp;</span><span class="mtk7">0</span><span class="mtk1">;&nbsp;</span><span class="mtk6">else</span><span class="mtk1">&nbsp;</span><span class="mtk6">return</span><span class="mtk1">&nbsp;</span><span class="mtk7">1</span><span class="mtk1">;</span></span>';
    // });
    // await page.waitForTimeout(3000);
    // await page.click('.tsoperation-toolbar-saveandclose');
    // await page.waitForTimeout(10000);
    // await page.waitForTimeout(3000);

// -----------------------------------------------

    // var record19;
    // let texto19;
    // try {
    //     await sql.connect(config)
    //     record19 = await sql.query`select id from tTag where [Name] = ${template.toString()} + '.Evento.FalhaComunicacoes'` // select distinct
    //     texto19 = record19.recordset[0].id;
    // } catch (e) {
    //     console.log(e);
    // }

    // var record20;
    // try {
    //     await sql.connect(config)
    //     record20 = await sql.query`update tTagScript set Script = 'if (Tags["AQS.CGL.Madeiras.Multiserra.Multiserra01.Evento.EstadoMaquina"].Quality != 0)
    //     return 0;
    //   else
    //     return 1;' where TagID = ${texto19}` // select distinct
    
    // } catch (e) {
    //     console.log(e);
    // }

// -----------------------------------------------

    //await page.click('ul .active');
    //await page.waitForTimeout(3000);

    for (var i = 0; i < tags.length; i++)
    {
        await page.getByText(new RegExp("^" + tags[i] + "$", "i")).click();
        await page.waitForTimeout(3000);
    }

    await page.waitForTimeout(3000);

    const elemento3 = await page.$(`[data-nodeid='218']`);
    if (elemento3) elemento3.click();

    await page.waitForTimeout(3000);

    await page.click('#tsslice-index-2 .fa-plus');
    await page.waitForTimeout(3000);
    const compare = await page.getByTitle('Compare Tag').first();
    if (compare) compare.click();
    await page.waitForTimeout(3000);
    await page.fill('#tseditName',template + '.Evento.HeartBeatUpdate');
    await page.waitForTimeout(3000);
    await page.selectOption('#tseditDataType','Integer');
    await page.waitForTimeout(3000);
    const primeiro = await page.getByTitle('Constant').first();
    if (primeiro) primeiro.click();
    await page.waitForTimeout(3000);
    const primeiro_segundo = await page.locator('.glyphicon-tag').first();
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
    const va8 = await page.locator('.glyphicon-tag').nth(1);
    const vatextoHandle6 = await va8.first();
    await vatextoHandle6.click();
    await page.fill('#contentPage_Picker_RightTagID_Name_TextBox',template + '.Evento.HeartBeatMaquina');
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
    for (var i = 0; i < tags.length; i++)
    {
        await page.getByText(new RegExp("^" + location[i] + "$", "i")).click();
        await page.waitForTimeout(3000);
    }
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("New Child")`);
    await page.waitForTimeout(3000);
    await page.fill('#tseditName',nome_location);
    await page.waitForTimeout(2000);
    if (key2) await page.fill('#tseditUniqueID',final_key2);
    await page.waitForTimeout(2000);
    await page.selectOption('#tseditLocationTypeID','LT_Maquinas');
    await page.waitForTimeout(2000);
    await page.click('#contentPage_Save_Button');
    await page.waitForTimeout(5000);

    //-----------------Criar Máquina------------------

    // await page.getByText(/^Systems$/i).click();
    // await page.click('#contentPage_slice1_TreeList_Tree_TreeView');
    // await page.waitForTimeout(3000);
    // await page.click('li:has-text("CHK_A_Colagem")');
    // await page.waitForTimeout(3000);
    // await page.click('#tsslice-index-2 ul li a');
    // await page.waitForTimeout(3000);
    // await page.getByRole('link', { name: 'Discrete System', exact: true }).click();
    // await page.waitForTimeout(3000);

    // await page.fill('#tseditName',nome_maquina);
    // if (key) await page.fill('#tseditKey',key);
    // await page.click('li:has-text("Maquina")');
    // await page.waitForTimeout(3000);
    // await page.fill('#tseditcp_CPS0000000013_CP0000000083', numero_maquina);
    // await page.waitForTimeout(1000);
    // await page.click('#contentPage_Save_Button');


    // ------------------------------------------------

    await page.getByText(/^Systems$/i).click();
    await page.click('#contentPage_slice1_TreeList_Tree_TreeView');
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("${maquina}")`);
    await page.waitForTimeout(3000);

    await page.click(`li:has-text("XXX")`);
    
    const va2 = await page.locator('.fa-share').nth(2);
    const vatextoHandle2 = await va2.first();
    await vatextoHandle2.click();
    await page.click('#contentPage_slice2_CreateButton');

    await page.waitForTimeout(3000);

    await page.click(`li:has-text("New System")`);

    await page.waitForTimeout(3000);
    await page.click('.fa-edit');
    await page.waitForTimeout(3000);

    await page.fill('#tseditName',nome_maquina);
    if (key) await page.fill('#tseditKey',key);
    await page.click('li:has-text("Advanced")');
    await page.fill('#tseditTemplateTagPrefix',template);
    await page.waitForTimeout(3000);
    await page.click('#contentPage_tseditLocationID_Picker');
    await page.waitForTimeout(3000);
    await page.click(`a:has-text("Expand All")`);
    await page.waitForTimeout(3000);
    await page.click(`li:has-text("${nome_location}")`);
    await page.waitForTimeout(3000);
    await page.click("#contentPage_Picker_LocationID_AssignButton");
    await page.waitForTimeout(3000);
    await page.click('li:has-text("Maquina")');
    await page.waitForTimeout(3000);
    await page.fill('#tseditcp_CPS0000000013_CP0000000083', numero_maquina);
    await page.waitForTimeout(3000);
    if (aut)
    {
        await page.selectOption('#tseditcp_CPS0000000013_CP0000000045',protocolo);
        await page.waitForTimeout(3000);
        if (CP_MAQUINA_PERMITE_PARAR_AUTOMACAO_preenchido)
        {
            await page.click('#tseditcp_CPS0000000013_CP0000000090');
        }
        else if (CP_MAQUINA_PERMITE_TROCAR_AUTOMACAO_preenchido)
        {
            await page.click('#tseditcp_CPS0000000013_CP0000000090');
        }
        else if (CP_MAQUINA_PERMITE_CONTENTOR_SEGUINTE_AUTOMACAO_preenchido)
        {
            await page.click('#tseditcp_CPS0000000013_CP0000000090');
        }
        else if (CP_MAQUINA_PERMITE_TRANSPORTE_AUTOMACAO_preenchido)
        {
            await page.click('#tseditcp_CPS0000000013_CP0000000090');
        }
        else if (CP_MAQUINA_PERMITE_ARRANQUE_AUTORIZA_ETIQUETA_SEGUINTE_preenchido)
        {
            await page.click('#tseditcp_CPS0000000013_CP0000000090');
        }
    }
    else await page.selectOption('#tseditcp_CPS0000000013_CP0000000045','Sem Protocolo');

    await page.waitForTimeout(3000);

    await page.click('#contentPage_Save_Button');

    await page.waitForTimeout(3000);

    const va4 = await page.locator('.fa-share').first();
    await va4.click();
    await page.waitForTimeout(3000);
    await page.click('#contentPage_SaveButton');

    //Tags (Parametrização)

    // Colocar Value e DefaultValue a 0 de todas as Tags da máquina

    var record10;
    try {
        await sql.connect(config)
        record10 = await sql.query`update tTag set [Value] = 0 where [Name] like '%' + ${template.toString()} + '%'`; // select distinct
    
    } catch (e) {
        console.log(e);
    }

    var record14;
    try {
        await sql.connect(config)
        record14 = await sql.query`update tTag set [DefaultValue] = 0 where [Name] like '%' + ${template.toString()} + '%'`; // select distinct
    
    } catch (e) {
        console.log(e);
    }

    var record1;
    try {
        await sql.connect(config)
        record1 = await sql.query`update tTag set [Value] = 1 where [Name] = '' + ${template.toString()} + '.Ord.ConsomeItems'` // select distinct
    } catch (e) {
        console.log(e);
    }

    var record2;
    try {
        await sql.connect(config)
        record2 = await sql.query`update tTag set [Value] = 1 where [Name] = '' + ${template.toString()} + '.Ord.ProduzItems'` // select distinct
    
    } catch (e) {
        console.log(e);
    }

    var record3;
    try {
        await sql.connect(config)
        record3 = await sql.query`update tTag set [Value] = ${ProduzItemsDefinitionID} where [Name] = '' + ${template.toString()} + '.Ord.ProduzItemsDefinitionID'` // select distinct
    
    } catch (e) {
        console.log(e);
    }

    var record4;
    try {
        await sql.connect(config)
        record4 = await sql.query`update tTag set [Value] = ${TaxaProducaoTeorica} where [Name] = '' + ${template.toString()} + '.Prod.TaxaProducaoTeorica'` // select distinct
    
    } catch (e) {
        console.log(e);
    }

    var record5;
    try {
        await sql.connect(config)
        record5 = await sql.query`update tTag set [Value] = ${EstadoMaquina} where [Name] = '' + ${template.toString()} + '.Evento.EstadoMaquina'` // select distinct
    
    } catch (e) {
        console.log(e);
    }

    for (var i = 1; i < dadosExcel.length+1; i++)
    {

        var columnName = '';
        if (i < 10) {
            columnName = `.Cons.ArmazemOrigemProduto0${i}`;
        } else {
            columnName = `.Cons.ArmazemOrigemProduto${i}`;
        }
        var record6;
        try {
            await sql.connect(config)
            record6 = await sql.query`update tTag set [Value] = ${ArmazemOrigemProduto[i-1].toString()} where [Name] = ${template.toString()} + ${columnName}`;
        
        } catch (e) {
            console.log(e);
        }

        var columnName2 = '';
        if (i < 10) {
            columnName2 = `.Prod.ArmazemDestinoProduto0${i}`;
        } else {
            columnName2 = `.Prod.ArmazemDestinoProduto${i}`;
        }
        var record13;
        try {
            await sql.connect(config)
            record13 = await sql.query`update tTag set [Value] = ${ArmazemDestinoProduto[i-1].toString()} where [Name] = ${template.toString()} + ${columnName2}`;
        
        } catch (e) {
            console.log(e);
        }

        var columnName3 = '';
        if (i < 10) {
            columnName3 = `.Prod.ContentorTipoDestinoProduto0${i}`;
        } else {
            columnName3 = `.Prod.ContentorTipoDestinoProduto${i}`;
        }
        var record14;
        try {
            await sql.connect(config)
            record14 = await sql.query`update tTag set [Value] = ${ContentorTipoDestinoProduto[i-1].toString()} where [Name] = ${template.toString()} + ${columnName3}`;
        
        } catch (e) {
            console.log(e);
        }

        var columnName4 = '';
        if (i < 10) {
            columnName4 = `.Cons.ContentorOrigemProduto0${i}`;
        } else {
            columnName4 = `.Cons.ContentorOrigemProduto${i}`;
        }
        var record15;
        try {
            await sql.connect(config)
            record15 = await sql.query`update tTag set [Value] = ${ArmazemOrigemProduto[i-1].toString()} where [Name] = ${template.toString()} + ${columnName4}`;
        
        } catch (e) {
            console.log(e);
        }
    }

    await page.close();

});