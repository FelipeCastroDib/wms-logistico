/*************************************************
 * CONFIGURAÇÕES
 *************************************************/

const CONFIG = {
  ABA_WMS: 'WMS',
  ABA_ESTOQUE: 'Estoque',

  CELULA_SKU: 'I2',
  CELULA_RUA: 'D2',
  CELULA_PAINEL: 'M1',

  RANGE_AF: 'AF3:AF200',

  LINHAS_ENDERECOS: [8, 13, 18, 23, 28],

  COLUNA_INICIAL: 2,
  QTD_COLUNAS: 25,

  TEMPO_RECALCULO: 1200
};

const VERSAO_WMS = '1.3';

/*************************************************
 * UTILIDADES
 *************************************************/

function aguardarRecalculo_() {
  SpreadsheetApp.flush();
  Utilities.sleep(CONFIG.TEMPO_RECALCULO);
  SpreadsheetApp.flush();
}

function obterWMS_() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.ABA_WMS);
}

function obterEstoque_() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.ABA_ESTOQUE);
}

/*************************************************
 * ENDEREÇOS ENCONTRADOS
 *************************************************/

function obterListaEnderecosAF_(sh) {

  aguardarRecalculo_();

  return sh
    .getRange(CONFIG.RANGE_AF)
    .getDisplayValues()
    .flat()
    .map(v =>
      String(v)
        .trim()
        .toUpperCase()
    )
    .filter(v =>
      v &&
      v !== 'NÃO ENCONTRADO' &&
      v.includes('-')
    );
}

function mostrarToastEnderecos_(ss, sh) {

  const lista = obterListaEnderecosAF_(sh);

  if (lista.length > 0) {

    ss.toast(
      'Endereços encontrados:\n\n' +
      lista.join('\n'),
      'WMS',
      15
    );

  } else {

    ss.toast(
      'Nenhum endereço encontrado',
      'WMS',
      5
    );

  }

  return lista;
}

/*************************************************
 * EVENTO AO EDITAR
 *************************************************/

function onEdit(e) {

  const sh = e.range.getSheet();

  if (sh.getName() !== CONFIG.ABA_WMS) return;

  if (e.range.getA1Notation() !== CONFIG.CELULA_SKU) return;

  mostrarToastEnderecos_(e.source, sh);

  atualizarPickingInteligente_();
}

/*************************************************
 * PAINEL DE INFORMAÇÕES
 *************************************************/

function obterLinhaEndereco_(linhaSelecionada) {

  for (const linha of CONFIG.LINHAS_ENDERECOS) {

    if (
      linhaSelecionada >= linha &&
      linhaSelecionada <= linha + 2
    ) {
      return linha;
    }

  }

  return null;
}

function atualizarPainelEndereco_(ss, sh, endereco) {

  const painel = sh.getRange(CONFIG.CELULA_PAINEL);

  const estoque = obterEstoque_();
  const ultimaLinha = estoque.getLastRow();

  if (ultimaLinha < 2) {

    painel.setValue(
      `Drive-in ${endereco}\n\nSem itens cadastrados`
    );

    return;
  }

  const dados = estoque
    .getRange(2, 1, ultimaLinha - 1, 6)
    .getDisplayValues();

  const itens = dados
    .filter(linha =>
      String(linha[2]).trim().toUpperCase() === endereco.toUpperCase()
    )
    .map(linha =>
      `Produto ${linha[3]} — ${linha[5]} paletes`
    );

  if (itens.length === 0) {

    painel.setValue(
      `Drive-in ${endereco}\n\nSem itens cadastrados`
    );

    return;
  }

  painel.setValue(
    `Drive-in ${endereco}\n\n${itens.join('\n')}`
  );
}

function onSelectionChange(e) {

  const sh = e.range.getSheet();

  if (sh.getName() !== CONFIG.ABA_WMS) return;

  const linha = e.range.getRow();
  const coluna = e.range.getColumn();

  const linhaEndereco = obterLinhaEndereco_(linha);

  const painel = sh.getRange(CONFIG.CELULA_PAINEL);

  if (!linhaEndereco) {

    painel.setValue('Clique em um endereço');

    return;
  }

  const endereco = String(
    sh.getRange(linhaEndereco, coluna).getDisplayValue()
  ).trim();

  if (
    !endereco ||
    !endereco.includes('-')
  ) {

    painel.setValue('Clique em uma célula válida');

    return;
  }

  atualizarPainelEndereco_(
    e.source,
    sh,
    endereco
  );
}

/*************************************************
 * DESTAQUES
 *************************************************/

function limparDestaquesSKU_() {

  const wms = obterWMS_();

  CONFIG.LINHAS_ENDERECOS.forEach(linha => {

    wms
      .getRange(
        linha,
        CONFIG.COLUNA_INICIAL,
        1,
        CONFIG.QTD_COLUNAS
      )
      .setBackground('#FFFFFF')
      .setFontColor('#000000')
      .setFontWeight('normal')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');

  });

}

function navegarRuaSKU_(rua) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const wms = obterWMS_();

  wms
    .getRange(CONFIG.CELULA_RUA)
    .setValue(rua);

  aguardarRecalculo_();

  limparDestaquesSKU_();

  const lista = obterListaEnderecosAF_(wms);

  const paraPintar = [];

  CONFIG.LINHAS_ENDERECOS.forEach(linha => {

    const valores = wms
      .getRange(
        linha,
        CONFIG.COLUNA_INICIAL,
        1,
        CONFIG.QTD_COLUNAS
      )
      .getDisplayValues()[0];

    valores.forEach((valor, idx) => {

      const endereco = String(valor)
        .trim()
        .toUpperCase();

      if (lista.includes(endereco)) {

        paraPintar.push(
          wms
            .getRange(
              linha,
              idx + CONFIG.COLUNA_INICIAL
            )
            .getA1Notation()
        );

      }

    });

  });

  if (paraPintar.length > 0) {

    wms
      .getRangeList(paraPintar)
  .setBackground('#B3E5FC')
  .setFontColor('#01579B')
  .setFontWeight('bold')
  .setHorizontalAlignment('center')
  .setVerticalAlignment('middle');

    ss.toast(
      'Endereços encontrados:\n' +
      lista.join(' | '),
      'WMS - Endereços',
      8
    );

  } else {

    ss.toast(
      'Nenhum endereço encontrado',
      'WMS',
      5
    );

  }
}

/*************************************************
 * PICKING INTELIGENTE
 *************************************************/

function atualizarPickingInteligente_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const wms = ss.getSheetByName('WMS');
  const estoque = ss.getSheetByName('Estoque');

  const sku = String(
    wms.getRange('I2').getDisplayValue()
  ).trim();

  // limpa painel
  wms.getRange('S1:Z5').clearContent();

  if (!sku) return;

  const ultimaLinha = estoque.getLastRow();

  if (ultimaLinha < 2) return;

  const dados = estoque
    .getRange(2, 1, ultimaLinha - 1, 6)
    .getDisplayValues();

  const encontrados = dados
    .filter(linha =>
      String(linha[3]).trim() === sku
    )
    .map(linha => ({
      endereco: linha[2],
      descricao: linha[4],
      quantidade: Number(linha[5])
    }));

  if (encontrados.length === 0) {

    wms.getRange('S1').setValue(
      'PICKING INTELIGENTE'
    );

    wms.getRange('S2').setValue(
      'SKU NÃO ENCONTRADO'
    );

    return;
  }

  // ordena pela menor quantidade
  encontrados.sort(
    (a, b) => a.quantidade - b.quantidade
  );

  // título
  wms.getRange('S1').setValue(
    'PICKING INTELIGENTE'
  );

  // mostra até 3 posições
  encontrados
    .slice(0, 3)
    .forEach((item, idx) => {

      wms
        .getRange(2 + idx, 19)
        .setValue(
          `${idx + 1}º ${item.endereco} (${item.quantidade} paletes)`
        );

    });

  // prioridade
  wms.getRange('S5').setValue(
    `Prioridade: ${encontrados[0].endereco}`
  );
}

/*************************************************
 * DRIVE-INS LIVRES
 *************************************************/

function atualizarDriveInsLivres_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const wms = ss.getSheetByName('WMS');
  const resumo = ss.getSheetByName('Resumo_Enderecos');

  const ultimaLinha = resumo.getLastRow();

  if (ultimaLinha < 2) return;

  const dados = resumo
    .getRange(2, 1, ultimaLinha - 1, 10)
    .getDisplayValues();

  const livres = dados.filter(
    linha =>
      String(linha[8]).trim().toUpperCase() === 'LIVRE'
  );

  // limpa painel
  wms.getRange('W1:Z5').clearContent();

  wms.getRange('W1').setValue(
    'DRIVE-INS LIVRES'
  );

  livres
    .slice(0, 3)
    .forEach((linha, idx) => {

      const endereco = linha[2];

      wms
        .getRange(2 + idx, 23)
        .setValue(
          `${idx + 1}º ${endereco}`
        );

    });

  wms.getRange('W5').setValue(
    `Total livres: ${livres.length}`
  );

}

function atualizarDashboard() {

  atualizarDriveInsLivres_();

  SpreadsheetApp
    .getActiveSpreadsheet()
    .toast(
      'Drive-ins livres atualizados',
      'WMS',
      3
    );

}

/*************************************************
 * BOTÕES
 *************************************************/

function destacarRuaA() { navegarRuaSKU_('A'); }
function destacarRuaB() { navegarRuaSKU_('B'); }
function destacarRuaC() { navegarRuaSKU_('C'); }
function destacarRuaD() { navegarRuaSKU_('D'); }
function destacarRuaE() { navegarRuaSKU_('E'); }
function destacarRuaF() { navegarRuaSKU_('F'); }
function destacarRuaG() { navegarRuaSKU_('G'); }
function destacarRuaH() { navegarRuaSKU_('H'); }
function destacarRuaI() { navegarRuaSKU_('I'); }


/*************************************************
 * CHANGELOG
 *************************************************/

function onOpen() {

  const props = PropertiesService.getUserProperties();

  const ultimaVersaoVista =
    props.getProperty('VERSAO_WMS_VISTA');

  if (ultimaVersaoVista === VERSAO_WMS) {
    return;
  }

  SpreadsheetApp.getUi().alert(
    '🚀 WMS - Versão ' + VERSAO_WMS + '\n\n' +
    'Novidades desta versão:\n\n' +
    '✓ Picking Inteligente\n' +
    '✓ Dashboard Drive-ins Livres\n' +
    '✓ Melhorias de desempenho\n\n' +
    'Consulte a aba CHANGELOG para mais detalhes.'
  );

  props.setProperty(
    'VERSAO_WMS_VISTA',
    VERSAO_WMS
  );
}