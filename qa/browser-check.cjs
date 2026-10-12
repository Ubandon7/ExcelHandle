const { chromium } = require('C:\\Users\\LAAA\\.cache\\codex-runtimes\\codex-primary-runtime\\dependencies\\node\\node_modules\\playwright')

let browser
let uploadedFileId

async function main() {
  browser = await chromium.launch({
    headless: true,
    executablePath: 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
  })
  const page = await browser.newPage({ viewport: { width: 1440, height: 960 }, deviceScaleFactor: 1 })
  page.setDefaultTimeout(15000)
  const errors = []
  page.on('pageerror', (error) => errors.push(`page: ${error.message}`))
  page.on('console', (message) => { if (message.type() === 'error') errors.push(`console: ${message.text()}`) })

  await page.goto('http://127.0.0.1:8000', { waitUntil: 'networkidle' })
  console.log('opened home')
  await page.screenshot({ path: 'prototype/项目完成预览-首页.png', fullPage: true })

  const initialFiles = (await (await page.request.get('http://127.0.0.1:8000/api/files')).json()).files
  const initialFileCount = await page.locator('.sidebar-file').count()
  const chooserPromise = page.waitForEvent('filechooser')
  await page.getByRole('button', { name: '选择文件' }).first().click()
  const chooser = await chooserPromise
  await chooser.setFiles('样例/原始全部数据.xlsx')
  await page.getByRole('heading', { name: '文件概览' }).waitFor({ timeout: 30000 })
  await page.getByText('6,991').first().waitFor({ timeout: 30000 })
  if (await page.locator('.sidebar-file').count() !== initialFileCount + 1) throw new Error('new workbook was not added to the file library')
  const currentFiles = (await (await page.request.get('http://127.0.0.1:8000/api/files')).json()).files
  uploadedFileId = currentFiles.find((item) => !initialFiles.some((initial) => initial.id === item.id))?.id
  console.log('uploaded sample')
  await page.screenshot({ path: 'prototype/项目完成预览-文件概览.png', fullPage: true })

  await page.getByRole('button', { name: '筛选处理' }).click()
  await page.getByTitle('新增结果组').click()
  await page.locator('input[value="Sheet1"]').waitFor({ timeout: 10000 })
  console.log('created group')
  const condition = page.locator('.condition-row').first()
  await condition.locator('select').nth(0).selectOption({ label: '录入备注' })
  await condition.locator('select').nth(1).selectOption('no_prefix')
  if (await condition.getByLabel('前后缀标识符').inputValue() !== '-') throw new Error('no-prefix delimiter did not default to hyphen')
  await condition.locator('select').nth(1).selectOption('starts_with')
  await condition.getByLabel('筛选值').click()
  const prefixOption = page.locator('.suggestion-menu button').first()
  await prefixOption.waitFor({ timeout: 30000 })
  const prefix = await prefixOption.locator('span').textContent()
  if (!prefix || !prefix.endsWith('-')) throw new Error(`invalid detected prefix: ${prefix}`)
  await prefixOption.click()
  console.log(`verified prefix suggestions: ${prefix}`)
  await page.getByTitle('新增结果组').click()
  await page.locator('input[value="Sheet2"]').waitFor()
  await page.locator('.group-card').nth(0).click()
  await page.getByLabel('筛选值').waitFor()
  if (await page.getByLabel('筛选值').inputValue() !== prefix) throw new Error('previous group was not auto-saved')
  await page.getByRole('button', { name: '保存条件' }).click()
  await page.locator('.saved-group-card').filter({ hasText: 'Sheet1' }).waitFor()
  console.log('verified persistent filter preset')
  await page.locator('.group-card').nth(1).click()
  await page.getByLabel('筛选值').fill('__no_match__')
  console.log('verified group auto-save')
  await page.getByRole('button', { name: '开始筛选' }).click()
  console.log('started filter')
  await page.getByRole('button', { name: '保存处理记录' }).waitFor({ timeout: 60000 })
  await page.getByRole('button', { name: /下载 Excel/ }).waitFor({ timeout: 60000 })
  console.log('filter completed')
  await page.screenshot({ path: 'prototype/项目完成预览-筛选结果.png', fullPage: true })

  await page.getByRole('button', { name: '跨 Sheet 求和' }).click()
  const temporarySource = page.locator('.sum-source-card').filter({ hasText: '临时筛选结果' }).first()
  await temporarySource.waitFor({ timeout: 30000 })
  await temporarySource.click()
  await page.locator('.sum-form select').selectOption({ label: '消费金额 ($) · 覆盖 3/3 Sheet' })
  const sumResponsePromise = page.waitForResponse((response) => response.url().endsWith('/api/sum-runs') && response.request().method() === 'POST')
  await page.getByRole('button', { name: '开始求和' }).click()
  const sumRunId = (await (await sumResponsePromise).json()).run.id
  console.log('started sum from temporary filter result')
  let sumState
  for (let attempt = 0; attempt < 120; attempt += 1) {
    sumState = (await (await page.request.get(`http://127.0.0.1:8000/api/sum-runs/${sumRunId}`)).json()).run
    if (sumState.status === 'succeeded' || sumState.status === 'failed') break
    await page.waitForTimeout(500)
  }
  if (sumState?.status !== 'succeeded') throw new Error(`temporary result sum failed: ${sumState?.error || sumState?.status}`)
  await page.locator('.grand-total').waitFor({ timeout: 10000 })
  await page.getByText('完成', { exact: true }).first().waitFor()
  await page.screenshot({ path: 'prototype/项目完成预览-跨表求和.png', fullPage: true })
  console.log('verified temporary filter result sum')

  await page.getByRole('button', { name: '筛选处理' }).click()
  await page.getByRole('heading', { name: '已保存的处理记录' }).waitFor()
  await page.getByRole('button', { name: /下载 Excel/ }).waitFor()
  await page.getByRole('button', { name: '保存处理记录' }).click()
  await page.locator('.run-record').first().waitFor()
  console.log('verified manual save')
  await page.reload({ waitUntil: 'networkidle' })
  await page.getByRole('heading', { name: '文件概览' }).waitFor()
  await page.getByRole('button', { name: '筛选处理' }).click()
  await page.getByRole('button', { name: /下载 Excel/ }).waitFor({ timeout: 30000 })
  console.log('verified result preservation after navigation and reload')

  await page.getByRole('button', { name: '跨 Sheet 求和' }).click()
  await page.locator('.sum-source-card').filter({ hasText: '已保存筛选结果' }).first().waitFor({ timeout: 30000 })
  console.log('verified saved filter result source')

  console.log(JSON.stringify({ title: await page.title(), errors }, null, 2))
  await browser.close()
}

main().catch((error) => {
  console.error(error)
  process.exitCode = 1
}).finally(async () => {
  if (browser) await browser.close()
  if (uploadedFileId) await fetch(`http://127.0.0.1:8000/api/files/${uploadedFileId}`, { method: 'DELETE' }).catch(() => {})
})
