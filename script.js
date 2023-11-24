function autoFillCompanyAndContainer() {
  const key = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const l = key.toString().split('-')
  const company = l[0]
  const _key = l[1]
  const date = Utilities.formatDate(new Date(), 'America/Los_Angeles', "MM/dd/yyyy")
  SpreadsheetApp.getActiveSheet().getRange('B2').setValue(company + ' - ' + date)
  SpreadsheetApp.getActiveSheet().getRange('H2').setValue("柜号：" + _key)
  return [company, _key]
}
// autoCompute()
function autoCompute() {
  const isLocked = addLock()
  if (isLocked) { 
    throw new Error('Locked: 前置表被其他人占用，为防止冲突，请你稍后(10秒-1分钟后)再试')
    return 
  }
  const sheet = SpreadsheetApp.getActiveSheet()
  // lock head
  // sheet.setFrozenRows(3)

  let lastRow = parseInt(sheet.getLastRow())
  console.log('lr', lastRow)
  const lastRowValue = sheet.getRange('B' + parseInt(lastRow - 1)).getValue()
  if(!lastRowValue.toString().includes('托盘总数:\nTOTAL PALLETS:')) {
    // insert last rows 
    sheet.insertRowAfter(lastRow)
    sheet.getRange('B'+parseInt(lastRow + 1)).setValue('托盘总数:\nTOTAL PALLETS:')
    sheet.getRange('D'+parseInt(lastRow + 1)).setValue('拆柜人员签字：\nEMPLOYEE SIGNATURE:')
    sheet.getRange(parseInt(lastRow + 1), 2, 1, 2).merge() // total pallets
    sheet.getRange(parseInt(lastRow + 1), 4, 1, 5).merge() // signature
    sheet.getRange(parseInt(lastRow + 1), 9, 1, 3).merge() // total count

    // insert instuction information
    sheet.insertRowAfter(lastRow + 1)
    sheet.getRange(parseInt(lastRow + 2), 2, 1, 10).merge()
    const richTextA1 = SpreadsheetApp.newRichTextValue()
    .setText("Notice: if you found THIS LIST and LABELS were mismatched, or had a lacking of LABELS \nPlease contact label printing person immediately").setTextStyle(0,5,SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#b45f06').build())
    .build()
    const richTextA2 = SpreadsheetApp.newRichTextValue()
    .setText("Aviso: si encontraste ESTA LISTA y ETIQUETAS que no coincidían o faltaban ETIQUETAS \nPor favor, comunícate inmediatamente con el tipo de impresión de etiquetas").setTextStyle(0,5,SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor('#b45f06').build())
    .build()
    sheet.getRange('B'+parseInt(lastRow + 2)).setRichTextValue(richTextA2)
    //.setFontColor('#b45f06').setFontWeight('bold').setFontSize(12).setFontFamily('Arial')
    

    lastRow += 2
  }

  sheet.getRange('I'+parseInt(lastRow - 1)).setValue(`=CONCATENATE(
    "总数\n(TOTAL  BOXES): ", sum(I4:I${lastRow - 2}), 
    "\nSKU: " ,COUNTA(B4:B${lastRow - 2}), "")`
  )
  // style
  const lastRange = sheet.getRange(lastRow - 1, 2, 2, 10)
  lastRange.setBorder(true,true,true,true,true,true)
  lastRange.setHorizontalAlignment('left')
  lastRange.setBackground('#fff2cc')
  lastRange.setFontWeight('bold')
  lastRange.setFontSize('12')
  lastRange.setFontFamily('Arial').setFontColor('black')
  sheet.autoResizeRows(lastRow - 2, 2)
  // sheet.autoResizeRows(lastRow, 1)

  sheet.getRange('B'+parseInt(lastRow)).setFontColor('#b45f06')

  // (start row number, start column number, how many rows, how many columns)
  const selectedRangeSecondCol = sheet.getRange(4, 2, lastRow - 5, 8)
  // console.log(selectedRangeSecondCol.getValues())
  const list = selectedRangeSecondCol.getValues()
  // console.log(list)
  // [[ 'GYR2', 'FBA17D1C2N0W', '4MTXUHNV', 3380.45, 8.4, 9.6, '', 233 ]... ]

  // calculate
  const data = []
  let _i = 4
  list.forEach((_v)=>{
      if (!_v[4]) _v[4] = 0
      if (!_v[7]) _v[7] = 0

      if(_v[0]) {
        data.push({
          to: _v[0],
          fba: _v[1],
          volumn: _v[4],
          boxes: _v[7],
          rowStart: _i,
          rowEnds: _i,
        })
      }else {
        if (_v[1]) data[data.length - 1]['fba'] = _v[1]
        data[data.length - 1]['volumn'] += _v[4]
        data[data.length - 1]['boxes'] += _v[7]
        data[data.length - 1]['rowEnds'] += 1
      }

      _i++
  })
  // [ { to: 'GYR2', volumn: 9.6, boxes: 244, rowStart: 4, rowEnds: 5 }... ]
  // console.log(data)
  const formattedDate = Utilities.formatDate(new Date(), 'America/Los_Angeles', "yyyy-MM-dd HH:mm")
  const date = Utilities.formatDate(new Date(), 'America/Los_Angeles', "yyyy-MM-dd")
  const [compnay, container] = autoFillCompanyAndContainer()

  let hasPA, hasPickUp, hasCL, hasStop, hasUPS

  // wirte to sheet
  data.forEach((v)=> {
    // give a padding to the row
    //const cell_to = sheet.getRange('B'+v['rowStart'])
    //const _value = cell_to.getValue().toString().trim()
    //cell_to.setValue(_value + '\n')
    
    const key1 = 'G'+v['rowStart']
    const key2 = 'J'+v['rowStart']
    const key3 = 'C'+v['rowStart']
    const cell1 = sheet.getRange(key1)
    const cell2 = sheet.getRange(key2)
    // add padding to row
    cell1.setValue(v['volumn'].toFixed(2))
    cell2.setValue(v['boxes']) // 2023-11-10 横跳
    // cell2.setValue('') // 2023-11-01 总数留空
    //var value = (cell1.isPartOfMerge() ? cell.getMergedRanges()[0].getCell(1,1) : cell).getValue();

    // merge cells
    // const ranges = sheet.getRangeList(['G'+v['rowStart'], 'G'+v['rowEnds']]).getRanges()
    // ranges[0].merge()
    if (!cell1.isPartOfMerge()) {
      sheet.getRange(v['rowStart'], 7, v['rowEnds'] - v['rowStart'] + 1, 1).merge() // for volumn
      sheet.getRange(v['rowStart'], 8, v['rowEnds'] - v['rowStart'] + 1, 1).merge() // for pallet
    }
    if (!cell2.isPartOfMerge()) {
      sheet.getRange(v['rowStart'], 10, v['rowEnds'] - v['rowStart'] + 1, 1).merge() // for boxes
      sheet.getRange(v['rowStart'], 11, v['rowEnds'] - v['rowStart'] + 1, 1).merge() // for remark
    }

    // transload marks
    const to = v['to'].toUpperCase()
    if (to.includes('PICK UP') || to.includes('PICKING UP') || to.includes('自提')) {
      hasPickUp = true
      v['to'] = 'PICK UP'
    }
    else if (to.includes('P.A') || to.includes('PERSONAL ADDRESS') || to.includes('私人') ) {
      hasPA = true
      v['to'] = 'P.A'
    }
    else if (to === 'HOLD' || to === 'STOP' || to.includes('拦截') ) {
      hasStop = true
      v['to'] = 'STOP'
    }
    else if (to === 'C.L' || to === 'CHANGE LABEL' || to === '换标') {
      hasCL = true
      v['to'] = 'C.L'
    }
    else if (to === 'UPS' || to === 'FEDEX' || to.includes('联邦') ) {
      if (to.includes('联邦')) v['to'] = 'FEDEX'
      hasUPS = true
    }

    v['fba'] = generateContentWithAnumber(container, v['to'], v['fba'], v['boxes'])
    const cell3 = sheet.getRange(key3)
    cell3.setValue(v['fba'])
  })

  let s = '☐ STOP     ☐ Pick Up     ☐ Personal Address     ☐ Change Label     ☐ UPS/FEDEX' // size 15
  s = hasStop? '✅ STOP' : '☐ STOP'
  s += '     '
  s += hasPickUp? '✅ Pick Up' : '☐ Pick Up'
  s += '     '
  s += hasPA? '✅ Personal Address' : '☐ Personal Address'
  s += '     '
  s += hasCL? '✅ Change Label' : '☐ Change Label'
  s += '     '
  s += hasUPS? '✅ UPS/FEDEX' : '☐ UPS/FEDEX'
  sheet.getRange('D1').setValue(s)


  // format whole sheets
  const range = sheet.getRange(4, 2, lastRow - 5, 10)
  range.setFontFamily('Arial')
  range.setFontSize(12)
  range.setFontWeight('bold')
  range.setVerticalAlignment('middle')
  range.setHorizontalAlignment('center')
  range.setBorder(true, true ,true, true, true, true)

  // auto resize all rows
  sheet.autoResizeRows(4, lastRow - 4)
  sheet.setRowHeights(4, lastRow - 4, 32)

  // Utilities.sleep(10 * 1000)
  exportSheetAsExcel()

  //const next = ScriptApp.newTrigger("exportSheetAsExcel").timeBased()
  //next.after(1).create()

  saveAsJSON({
    data, compnay, container, date
  }, `${formattedDate} ${compnay} - ${container}`)
  
  // must run this at last
  releaseLocks()
}

function getFileAsBlob(exportUrl) {

  const key = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  const l = key.toString().split('-')
  const company = l[0]
  const _key = l[1]

  const formattedDate = Utilities.formatDate(new Date(), "CET", "yyyy-MM-dd")
  const name = `[${formattedDate}] ${company} - ${_key}` 
  const dir = DriveApp.getFolderById("1VKLj9JVy3YyeKh4q8Hk2bPsqEURY89KC")  // Container-Listing

  let response = UrlFetchApp.fetch(exportUrl, {
     muteHttpExceptions: true,
     headers: {
       Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
     },
   })
  const blob = response.getBlob()
  blob.setName(name + ".pdf")
  const file = dir.createFile(blob)
  return blob
}

function exportSheetAsExcel() {
  const id = SpreadsheetApp.getActiveSpreadsheet().getId()
  const blob = getFileAsBlob(`https://docs.google.com/spreadsheets/u/0/d/${id}/export?format=pdf`)
  Logger.log("Content type: " + blob.getContentType())
  Logger.log("File size in MB: " + blob.getBytes().length / 1000000)
}

function saveAsJSON(data, fileName) {
  // const dir = DriveApp.getFolderById("1-C4Ch4Qxoa-e0ohKcK6bPczQQuLLbjVd")
  const dir = DriveApp.getFolderById('1-Dw82L8iepItrevCAa7Kf3GQFnmhrpPj')
  dir.createFile(fileName + '.json', JSON.stringify(data))
}

function splitNumber(string) {
  let out = string.replace(/\'/g,'')
  out = out.split(/(\d+)/)
  out = out.filter(Boolean) 
  // ['A', '032']
  return out
}

// operation: [STOP, P.A, C.L, PICK UP]
// content 
function generateContentWithAnumber(containerId, operation, content, boxes) {
  if (content.includes('\n---')) return content

  if (operation === 'STOP') operation = 'HOLD'
  else if (operation === 'P.A') operation = 'LOCAL'
  else if (operation === 'C.L') operation = 'Change Label'
  else if (operation == 'PICK UP') operation = 'Pick Up'
  else return content
  
  const sheet = getNicksShit()
  const serialRange = sheet.getRange('B2')
  let [prefix, num] = splitNumber(serialRange.getValue())
  let lastRow = parseInt(sheet.getLastRow())
  //generate anumber
  num = parseInt(num)
  if (num === 999) {
    num = 0
    prefix = getNextLetter(prefix)
  }
  let middle = ''
  if (num < 100) {
    middle = '0'
  }
  if (num < 10) {
    middle = '00'
  }
  const anumber = prefix + middle + parseInt(num + 1)
  serialRange.setValue(anumber)
  // gen new row
  sheet.getRange('A'+parseInt(lastRow + 1)).setValue(containerId)
  sheet.getRange('C'+parseInt(lastRow + 1)).setValue(containerId + anumber)
  sheet.getRange('E'+parseInt(lastRow + 1)).setValue(content)
  sheet.getRange('F'+parseInt(lastRow + 1)).setValue(operation)
  sheet.getRange('G'+parseInt(lastRow + 1)).setValue(boxes)
  content += ('\n---' + anumber)
  return content
}

// true if LOCKED
function addLock() {
  let sheet = getNicksShit()
  const lockCell = sheet.getRange('D2')
  const lock = lockCell.getValue()
  if (lock === 'LOCKED') {
    return true
  }
}

function releaseLocks() {
  let sheet = getNicksShit()
  const lockCell = sheet.getRange('D2')
  lockCell.setValue('Vacant') 
}

let nickSheet = undefined

function getNicksShit() {
  if (nickSheet) return nickSheet
  let sheet = SpreadsheetApp.openById("1oJ3GE7hABgbfKUVjMTTw3lMRNW6QeOtRH7Hb56UGJ7o")
  nickSheet = sheet.getSheetByName('明细')
  return nickSheet
}

function getNextLetter (current) {
  if (current === 'Z') return 'A'
  return current.substring(0, current.length - 1)
      + String.fromCharCode(current.charCodeAt(current.length - 1) + 1)
}
