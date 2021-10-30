const refRegex = /(?:(\w+|'.*?')!)?([A-Z\$]+[:\d]+(?:[A-Z\$\d]+)?)/gm

function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Formula Peek')
      .addItem('Peek formula for current cell', 'showDialog')
      .addToUi();
}

function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('DynamicModal')
      .setWidth(1200)
      .setHeight(800);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Formula Peek');
}

function buildCurrentHtml() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const currentSheet = sheet.getSheetName()
  const range = sheet.getCurrentCell()
  const globalRef = `${currentSheet}!${range.getA1Notation()}`
  return buildRefHtml(globalRef)
}

function getDataFromGlobalRef(globalRef) {
  Logger.log(`Getting data for global ref: ${globalRef}`)
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  let range
  try {
    range = ss.getRange(globalRef)
  }
  catch (error) {
    Logger.log(`Range error on ${globalRef}, ${error}`)
    return {}
  }
  const sheet = range.getSheet()
  const column = range.getColumn()
  const headerRange = findHeaderRange(sheet, column)
  return {
    'ref': globalRef,
    'sheet': sheet.getName(),
    'row': range.getRow(),
    'column': column,
    'value': range.getValue(),
    'formula': range.getFormula(),
    'header': headerRange.getValue()
  }
}

function prettifyFormula(formula){
  var pretty = ''
  var tabNum = 0
  var newLineAdded = false
  formula.split('').forEach(function(c,i){
    if (newLineAdded) {
      c = c.trim()
      newLineAdded = false
    }
    if(/[\{\(]/.test(c)){
      tabNum++
      pretty += c + '\n' + '\t'.repeat(tabNum)
      newLineAdded = true
    } else if(/[\}\)]/.test(c)){
      tabNum--
      pretty += c + '\n' + '\t'.repeat(tabNum)
      newLineAdded = true
    //} else if (/[\+\-\*\/\^,;&]/.test(c)) { 
    } else if (/[,]/.test(c)) { 
      pretty += c + '\n' + '\t'.repeat(tabNum)
      newLineAdded = true
    } else {
      pretty += c
    }
  })
  return pretty
}

function findHeaderRange(sheet, column) {
  for (let i = 1; i < 3; i++) {
    let headerRange = sheet.getRange(i, column)
    if (headerRange.getFontWeight() === 'bold' || headerRange.getBackground() !== '#ffffff') {
      return headerRange
    }
  }
  return sheet.getRange(1, column)
}

function getGlobalRefFromRef(ref, currentSheet) {
  refRegex.lastIndex = 0
  const matches = refRegex.exec(ref).slice(1)
  if (matches !== null) {
    Logger.log(`Matches for ref (${ref}) and sheet (${currentSheet}) are ${matches}`)
    let [sheetRef, cellRef] = matches
    if (sheetRef === undefined) {
      // If sheet name not found in ref, it's local to currentSheet
      sheetRef = currentSheet
    }
    const globalRef = `${sheetRef.trim()}!${cellRef.trim()}`
    Logger.log(`Global ref for ref (${ref}) and sheet (${currentSheet}) is (${globalRef})`)
    return globalRef
  }
  return null
}

function buildRefHtml(globalRef) {
  Logger.log(`Building html for global ref: ${globalRef}`)
  const data = getDataFromGlobalRef(globalRef) || ''
  
  let html = '<li>'
  const headerTitle = data.header.toString().replace(/"/g, '&quot;')
  const valueTitle = data.value.toString().replace(/"/g, '&quot;')
  const sharedHtml = `
    <span class="ref">${data.ref}</span> | 
    <b>Header:</b> <span class="header" title="${headerTitle}">${data.header}</span> | 
    <b>Value:</b> <span class="value" title="${valueTitle}">${data.value}</span>
  `
  if (data.formula.search(refRegex) === -1) {
    // Leaf, just the ref data
    html += sharedHtml
  } else {
    // Node with children, include formula
    let formulaTitle
    try {
      formulaTitle = prettifyFormula(data.formula).replace(/"/g, '&quot;')
    } catch (e) {
      Logger.log(`Could not pretiffy formula (${e}): ${data.formula}`)
      formulaTitle = data.formula.replace(/"/g, '&quot;')
    }
    html += `
      <span class="caret">
        ${sharedHtml} | 
        <b>Formula:</b> <span class="formula" title="${formulaTitle}">${data.formula}</span>
      </span>
    `
  }
  html += '</li>'
  return html
}

function buildChildrenHtml(globalRef) {
  const data = getDataFromGlobalRef(globalRef)
  if (!data) {
    return ''
  }
  let html = '<ul class="nested active">'
  refRegex.lastIndex = 0
  const childRefs = [...new Set(data.formula.match(refRegex))]
  for (let childRef of childRefs) {
    let childGlobalRef = getGlobalRefFromRef(childRef, data.sheet)
    html += buildRefHtml(childGlobalRef)
  }
  html += '</ul>'
  return html
}