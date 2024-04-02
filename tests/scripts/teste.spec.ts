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

    await page.goto('http://ktmesapp04/TS/pages/home/config/tags/?c=ETS.Configuration.Scripting.ScriptEditor&Entity=TagScriptComposite&ID=340748');

    await page.click('.fa-cog');
    await page.waitForTimeout(5000);
    await page.selectOption('#InputEditorType','Text');
    await page.waitForTimeout(3000);
    await page.click('#Btns_Save');
    await page.waitForTimeout(3000);
    await page.fill('#contentPage_Editor_Code', 'if (Tags[".Evento.EstadoMaquina"].Quality != 0) return 0; else return 1;');

    await page.waitForTimeout(5000);

    await page.click('.tsoperation-toolbar-saveandclose');

});