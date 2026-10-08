import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

test.describe('home screen', () => {
  test('shows hero, quick-create cards and tab bar', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'home-basics' })
    const { page } = launched
    try {
      await expect(page.locator('.home-hero')).toBeVisible()
      // six quick-create cards plus the "Open file" browse card
      const quick = page.locator('.quick-card')
      await expect(quick).toHaveCount(7)
      await expect(quick.first()).toContainText('Docs')
      await expect(quick.nth(1)).toContainText('Sheets')
      await expect(quick.nth(2)).toContainText('Slides')
      await expect(quick.nth(3)).toContainText('Markdown')
      await expect(quick.nth(4)).toContainText('HTML')
      await expect(quick.nth(5)).toContainText('PDF')
      // and the starter templates below them
      await expect(page.locator('.template-card')).toHaveCount(5)
      await expect(page.locator('.tab-bar .tab-item.tab-home')).toBeVisible()
      await page.screenshot({ path: screenshotPath('home-overview') })
    } finally {
      await closeAndSaveVideo(launched, 'home-basics')
    }
  })

  test('renders localized UI when GENOFFICE_LANG=zh-CN', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-CN',
      videoDir: 'home-zh-cn',
    })
    const { page } = launched
    try {
      await expect(page.locator('.nav-item .nav-label').first()).toHaveText('最近')
      await page.screenshot({ path: screenshotPath('home-zh-cn') })
    } finally {
      await closeAndSaveVideo(launched, 'home-zh-cn')
    }
  })
})
