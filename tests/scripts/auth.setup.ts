// import { test as setup } from '@playwright/test';

// const authFile = 'playwright/.auth/user.json';

// setup('authenticate', async ({ page }) => {

//   await page.goto('http://ktmesapp04/TS/Account/LogOn.aspx?ts_deny=true&ts_rurl=%2fTS%2fdefault.aspx');
//   await page.getByLabel('Login').fill('kt0032'); //utilizador kt 
//   await page.getByLabel('Password').click();
//   await page.getByLabel('Password').fill('12345'); // password
//   await page.getByRole('button', { name: 'Sign In' }).click();

//   await page.waitForURL('http://ktmesapp04/TS/pages/root/config/products/materials/');

//   await page.context().storageState({ path: authFile });
// });

/*
setup('logicON', async ({ page }) => {

  await page.goto('http://ktmesapp02/TS/pages/acc/admin/services/logic$1/');
  await page.getByText('Start', {exact :true}).click();
  await page.getByRole('button', { name: 'OK'}).click();
  await page.waitForURL('http://ktmesapp02/TS/pages/root/admin/services/module$1/?c=ETS.Application.Wait.StatusActivityComplete&guid=5041653d-c63e-4706-a102-d31375233aa4');
});

setup('DMSon', async ({ page }) => {

  await page.goto('http://ktmesapp02/TS/pages/acc/admin/services/module$1/');
  await page.getByText('Start', {exact :true}).click();
  await page.getByRole('button', { name: 'OK'}).click();
  await page.waitForURL('http://ktmesapp02/TS/pages/acc/admin/services/module$1/?c=ETS.Application.Wait.StatusActivityComplete&guid=8d269f3a-4c61-432c-a12e-89a4be823889');
});*/