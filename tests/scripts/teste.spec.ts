import { test, expect } from '@playwright/test';
import { fail } from 'assert';
import * as XLSX from 'xlsx';

test('CriarMinhaMáquina', async ({ page }) => {

    await page.goto('http://ktmesapp04/TS/pages/home/config/tags/?c=ETS.Configuration.Scripting.ScriptEditor&Entity=TagScriptComposite&ID=333534');

    await page.waitForTimeout(3000);

    // Obtenha o valor do input
    const value = await page.locator('#contentPage_Editor_FieldCode').inputValue();
    console.log(value);

    await page.waitForTimeout(3000);
    //await page.fill('#contentPage_Editor_FieldCode','teste');
    await page.waitForTimeout(3000);

    // Obtenha o valor do input
    const value2 = await page.locator('#contentPage_Editor_FieldCode').inputValue();
    console.log(value2);

    await page.waitForTimeout(5000);

    await page.close();

});