_autoFillCompanyAndContainer()
function _autoFillCompanyAndContainer() {
  const key = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName()
  const l = key.toString().split('-')
  const company = l[0]
  const _key = l[1]
  const date = Utilities.formatDate(new Date(), 'America/Los_Angeles', "MM/dd/yyyy")
  SpreadsheetApp.getActiveSheet().getRange('B2').setValue(company + ' - ' + date)
  SpreadsheetApp.getActiveSheet().getRange('H2').setValue("柜号：" + _key)
  return [company, _key]
}
