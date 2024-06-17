const { chromium } = require('playwright');
const fs = require('fs').promises;
const path = require('path');

(async () => {
  const date = new Date();

  const suspensa = 'Está com comportamento: Suspenso.';
  const baixada = 'Está com comportamento: Baixado.';
  const notFound = 'Contribuinte não encontrado.';

  const empresas = JSON.parse(
    await fs.readFile(path.join(__dirname, 'empresas.json'), 'utf-8')
  );

  const { init, end } = getPreviousMonthDates();

  function getPreviousMonthDates() {
    const now = new Date();
    const firstDayOfCurrentMonth = new Date(
      now.getFullYear(),
      now.getMonth(),
      1
    );

    // Obter o último dia do mês passado
    const lastDayOfPreviousMonth = new Date(firstDayOfCurrentMonth - 1);
    const end = formatDate(lastDayOfPreviousMonth);

    // Obter o primeiro dia do mês passado
    const firstDayOfPreviousMonth = new Date(
      lastDayOfPreviousMonth.getFullYear(),
      lastDayOfPreviousMonth.getMonth(),
      1
    );
    const init = formatDate(firstDayOfPreviousMonth);

    return { init, end };
  }

  function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Os meses em JavaScript são baseados em zero
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
  }

  async function saveReport(content) {
    const report = path.join(__dirname, 'report.txt');
    await fs.appendFile(report, `${content}\n`);
  }

  const folder = path.join('C:/NFSe');

  async function apagarPasta() {
    try {
      await fs.rm(folder, { recursive: true, force: true });
      console.log('Pasta deletada com sucesso!');
    } catch (err) {
      console.error('Erro ao deletar a pasta: ', err);
    }
  }

  async function acessarEmpresa(page, empresa, cnpj) {
    await page.fill('#TxtCPF', cnpj);
    await page.click('#imbLocalizar');
    await page.waitForLoadState();

    if (await page.isVisible('text=OK')) {
      const statusErr = await page.$eval(
        '.bootbox-body',
        (element) => element.textContent
      );
      if (
        statusErr.includes(baixada) ||
        statusErr.includes('Está com status: Baixado')
      ) {
        await saveReport(`${empresa}: ${baixada}`);
      } else if (statusErr.includes(suspensa)) {
        await saveReport(`${empresa}: ${suspensa}`);
      } else if (statusErr.includes(notFound)) {
        await saveReport(`${empresa}: ${notFound}`);
      } else {
        await saveReport(`${empresa}: status do erro não identificado.`);
      }
      await page.click('text=OK');
      await page.goto(
        'https://iss.fazenda.df.gov.br/online/default/empresas.aspx'
      );
    } else if (await page.isVisible('#dgEmpresas__ctl3_imbSelecione')) {
      const rows = await page.$$('tbody .ItemStyleNovo');
      for (const row of rows) {
        const statusText = await row.$eval('td:nth-child(4)', (element) =>
          element.textContent.trim()
        );
        if (statusText === 'Ativo') {
          const selectLink = await row.$('td:nth-child(7) a');
          if (selectLink) {
            await selectLink.click();
            break;
          }
        }
      }
      await saveReport(`${empresa}: Apareceu mais de um, mas deu certo!`);
      await page.goto(
        'https://iss.fazenda.df.gov.br/online/NotaDigital/consulta_nota.aspx'
      );
    } else {
      await saveReport(`${empresa}: Apareceu só um e deu tudo certo!`);
      await page.goto(
        'https://iss.fazenda.df.gov.br/online/NotaDigital/consulta_nota.aspx'
      );
    }
  }

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext();
  const page = await context.newPage();

  await saveReport(
    `---------------------------------\nInicio da execução: ${date}`
  );

  // Deletar a pasta de forma assíncrona antes de continuar
  await apagarPasta();

  await page.goto('https://iss.fazenda.df.gov.br/');
  await page.click('#btnAcionaCertificado');
  await page.waitForSelector('#imbLocalizar');

  for (const comp of empresas) {
    const empresa = comp.empresa;
    const cnpj = comp.cnpj;
    const caminho = comp.caminho;

    await acessarEmpresa(page, empresa, cnpj);
    if (
      page.url() ===
      'https://iss.fazenda.df.gov.br/online/default/empresas.aspx'
    ) {
      continue;
    }

    await page.selectOption('#ddSerie', {
      label: 'Nota Fiscal de Serviço Eletrônica - NFS-e',
    });
    await page.click('#filtrosUpDown');
    await page.fill('#txtDtEmissaoIni', init);
    await page.fill('#txtDtEmissaoFim', end);
    await page.click('#btnLocalizar2');
    if (await page.isVisible('#dgDocumentos__ctl3_btnExpEste')) {
      const downloadPromise = page.waitForEvent('download');
      await page.click('#dgDocumentos__ctl2_btnExpTodos');
      const download = await downloadPromise;

      await download.saveAs(caminho + `prestados.zip`);
    }
    await page.goto(
      'https://iss.fazenda.df.gov.br/online/default/Empresas.aspx'
    );
  }
  await page.close();
  const endDate = new Date();
  const timeOfExecution = (endDate - date) / 1000;
  await saveReport(
    `Tempo de execução: ${timeOfExecution} segundos ou ${
      timeOfExecution / 60
    } minutos`
  );
  await browser.close();
})();
