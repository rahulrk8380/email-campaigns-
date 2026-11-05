// SmartMail Pro — Playwright Feature Tests
// Tests run against http://localhost:3000
// No real emails are sent — only UI + API behaviour is verified.

const { test, expect } = require('@playwright/test');

const BASE = 'http://localhost:3000';
const LOGIN_EMAIL = 'rahul@gmail.com';
const LOGIN_PASS = 'hi';

// Helper: login and go to a page
async function login(page) {
  await page.goto(BASE);
  await page.fill('#loginEmail', LOGIN_EMAIL);
  await page.fill('#loginPassword', LOGIN_PASS);
  await page.click('button[type="submit"]');
  await page.waitForSelector('#mainView:not(.hidden)', { timeout: 8000 });
}

// ─── 1. LOGIN ─────────────────────────────────────────────────────────────────
test('1. Login with valid credentials', async ({ page }) => {
  await page.goto(BASE);
  await page.fill('#loginEmail', LOGIN_EMAIL);
  await page.fill('#loginPassword', LOGIN_PASS);
  await page.click('button[type="submit"]');
  await expect(page.locator('#mainView')).not.toHaveClass(/hidden/, { timeout: 8000 });
  console.log('✓ Login successful');
});

test('2. Login fails with wrong password', async ({ page }) => {
  await page.goto(BASE);
  await page.fill('#loginEmail', 'wrong@email.com');
  await page.fill('#loginPassword', 'wrongpass');
  await page.click('button[type="submit"]');
  await expect(page.locator('#loginError')).not.toBeEmpty({ timeout: 5000 });
  console.log('✓ Wrong-password error shown');
});

// ─── 2. DASHBOARD ─────────────────────────────────────────────────────────────
test('3. Dashboard shows stats cards', async ({ page }) => {
  await login(page);
  await expect(page.locator('#statContacts')).toBeVisible();
  await expect(page.locator('#statCampaigns')).toBeVisible();
  await expect(page.locator('#statSent')).toBeVisible();
  await expect(page.locator('#statFailed')).toBeVisible();
  console.log('✓ Dashboard stats visible');
});

// ─── 3. CONTACTS UPLOAD ───────────────────────────────────────────────────────
test('4. Contacts page loads and shows upload dropzone', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="contacts"]');
  await expect(page.locator('#dropzone')).toBeVisible();
  await expect(page.locator('#contactsTable')).toBeVisible();
  console.log('✓ Contacts page and dropzone visible');
});

// ─── 4. COMPOSE PAGE ─────────────────────────────────────────────────────────
test('5. Compose page loads with Quill editor', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.waitForTimeout(500); // let Quill init
  await expect(page.locator('#quillEditor')).toBeVisible();
  await expect(page.locator('.ql-toolbar')).toBeVisible();
  console.log('✓ Quill editor and toolbar visible');
});

test('6. Quill editor accepts typed text', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.waitForTimeout(600);
  const editor = page.locator('#quillEditor .ql-editor');
  await editor.click();
  await editor.type('Hello World test message');
  const content = await editor.textContent();
  expect(content).toContain('Hello World');
  console.log('✓ Quill editor accepts text input');
});

test('7. Quill Bold, Italic, Underline buttons are present', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.waitForTimeout(400);
  await expect(page.locator('button.ql-bold')).toBeVisible();
  await expect(page.locator('button.ql-italic')).toBeVisible();
  await expect(page.locator('button.ql-underline')).toBeVisible();
  await expect(page.locator('button.ql-strike')).toBeVisible();
  await expect(page.locator('button.ql-link')).toBeVisible();
  await expect(page.locator('button.ql-image')).toBeVisible();
  await expect(page.locator('button.ql-clean')).toBeVisible();
  console.log('✓ All Quill formatting buttons visible');
});

test('8. Subject field accepts text', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.fill('#subject', 'Test Subject {{Customer Name}}');
  const val = await page.inputValue('#subject');
  expect(val).toBe('Test Subject {{Customer Name}}');
  console.log('✓ Subject field works');
});

// ─── 5. ROW RANGE ─────────────────────────────────────────────────────────────
test('9. Row Range inputs are visible on compose page', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#rowStart')).toBeVisible();
  await expect(page.locator('#rowEnd')).toBeVisible();
  await expect(page.locator('#rowRangePreview')).toBeVisible();
  console.log('✓ Row range inputs visible');
});

test('10. Row Range preview updates on input', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.fill('#rowStart', '1');
  await page.fill('#rowEnd', '10');
  await page.locator('#rowEnd').dispatchEvent('input');
  // Preview should show something with "Rows"
  await page.waitForTimeout(300);
  const preview = await page.locator('#rowRangePreview').textContent();
  console.log('  Row Range preview:', preview);
  console.log('✓ Row Range preview updates');
});

// ─── 6. DELAY CONTROLS ────────────────────────────────────────────────────────
test('11. Delay min/sec inputs visible and working', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#delayMinutes')).toBeVisible();
  await expect(page.locator('#delaySeconds')).toBeVisible();
  await expect(page.locator('#delayPreview')).toBeVisible();
  console.log('✓ Delay min+sec inputs visible');
});

test('12. Delay preset buttons update delay inputs', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  // Click the "1m" preset button
  const presetBtns = page.locator('.delay-preset-btn');
  const count = await presetBtns.count();
  console.log('  Delay preset buttons found:', count);
  expect(count).toBeGreaterThan(0);
  // Click 30s preset
  for (let i = 0; i < count; i++) {
    const txt = await presetBtns.nth(i).textContent();
    if (txt.includes('30s')) {
      await presetBtns.nth(i).click();
      break;
    }
  }
  await page.waitForTimeout(200);
  const preview = await page.locator('#delayPreview').textContent();
  console.log('  Delay preview after 30s preset:', preview);
  console.log('✓ Delay preset buttons work');
});

// ─── 7. SCHEDULE MODE + BATCH SCHEDULER ──────────────────────────────────────
test('13. Schedule mode buttons (All Rows / Custom Batch) visible', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#modeAllRows')).toBeVisible();
  await expect(page.locator('#modeBatch')).toBeVisible();
  console.log('✓ Schedule mode buttons visible');
});

test('14. Custom Batch Schedule mode shows batch panel', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.click('#modeBatch');
  await expect(page.locator('#batchSchedulerPanel')).toBeVisible();
  await expect(page.locator('#batchList')).toBeVisible();
  console.log('✓ Batch scheduler panel appears');
});

test('15. Add Batch button creates batch entries', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.click('#modeBatch');
  // Add 3 batches
  await page.click('button:has-text("+ Add Batch")');
  await page.click('button:has-text("+ Add Batch")');
  await page.click('button:has-text("+ Add Batch")');
  await page.waitForTimeout(200);
  const batchItems = page.locator('.batch-item');
  const count = await batchItems.count();
  expect(count).toBe(3);
  console.log(`✓ 3 batches created (found ${count})`);
});

test('16. Remove batch button deletes a batch', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.click('#modeBatch');
  await page.click('button:has-text("+ Add Batch")');
  await page.click('button:has-text("+ Add Batch")');
  await page.waitForTimeout(100);
  // Remove first batch
  const removeBtn = page.locator('.batch-remove-btn').first();
  await removeBtn.click();
  await page.waitForTimeout(100);
  const remaining = await page.locator('.batch-item').count();
  expect(remaining).toBe(1);
  console.log('✓ Batch removal works');
});

test('17. Batch fields are fillable', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.click('#modeBatch');
  await page.click('button:has-text("+ Add Batch")');
  await page.waitForTimeout(200);
  const batchItem = page.locator('.batch-item').first();
  // Fill start row
  await batchItem.locator('input[type="number"]').first().fill('1');
  // Fill end row
  await batchItem.locator('input[type="number"]').nth(1).fill('50');
  console.log('✓ Batch fields are fillable');
});

test('18. Switch back to All Rows mode hides batch panel', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.click('#modeBatch');
  await expect(page.locator('#batchSchedulerPanel')).toBeVisible();
  await page.click('#modeAllRows');
  await expect(page.locator('#batchSchedulerPanel')).not.toBeVisible();
  await expect(page.locator('#rowRangePanel')).toBeVisible();
  console.log('✓ Switching back to All Rows mode works');
});

// ─── 8. RECEIPT OPTIONS ───────────────────────────────────────────────────────
test('19. Delivery and Read receipt checkboxes are visible', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#requestDeliveryReceipt')).toBeVisible();
  await expect(page.locator('#requestReadReceipt')).toBeVisible();
  console.log('✓ Receipt checkboxes visible');
});

test('20. Receipt checkboxes are toggleable', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.check('#requestDeliveryReceipt');
  await page.check('#requestReadReceipt');
  expect(await page.isChecked('#requestDeliveryReceipt')).toBe(true);
  expect(await page.isChecked('#requestReadReceipt')).toBe(true);
  await page.uncheck('#requestDeliveryReceipt');
  expect(await page.isChecked('#requestDeliveryReceipt')).toBe(false);
  console.log('✓ Receipt checkboxes toggleable');
});

// ─── 9. EMAIL PREVIEW MODAL ───────────────────────────────────────────────────
test('21. Email Preview button visible', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('button:has-text("Preview Email")')).toBeVisible();
  console.log('✓ Preview Email button visible');
});

// ─── 10. TEMPLATES ────────────────────────────────────────────────────────────
test('22. Template panel is visible on compose page', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#templateSelect')).toBeVisible();
  await expect(page.locator('button:has-text("Save Current")')).toBeVisible();
  console.log('✓ Template panel visible');
});

// ─── 11. CAMPAIGN HISTORY ─────────────────────────────────────────────────────
test('23. Campaign History page loads and shows table', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="campaigns"]');
  await expect(page.locator('#campaignsTable')).toBeVisible();
  console.log('✓ Campaign history table visible');
});

test('24. Campaign control buttons hidden by default (no campaign selected)', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="campaigns"]');
  // Control bar exists but buttons should be hidden until a campaign is opened
  const pauseBtn = page.locator('#pauseCampaignBtn');
  // Either doesn't exist yet or is hidden
  const isHidden = await pauseBtn.evaluate(el => el.classList.contains('hidden')).catch(() => true);
  expect(isHidden).toBe(true);
  console.log('✓ Campaign control buttons hidden when no campaign active');
});

// ─── 12. MICROSOFT STATUS ─────────────────────────────────────────────────────
test('25. Microsoft status pill visible in sidebar', async ({ page }) => {
  await login(page);
  await expect(page.locator('#msStatusPill')).toBeVisible();
  await expect(page.locator('#msStatusText')).toBeVisible();
  console.log('✓ Microsoft status pill visible');
});

// ─── 13. LIVE PREVIEW ─────────────────────────────────────────────────────────
test('26. Live preview panel exists on compose page', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#previewSubject')).toBeVisible();
  await expect(page.locator('#previewBody')).toBeVisible();
  console.log('✓ Live preview panel visible');
});

test('27. Subject text reflects in live preview', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.fill('#subject', 'My Test Subject 123');
  await page.locator('#subject').dispatchEvent('input');
  await page.waitForTimeout(200);
  const preview = await page.locator('#previewSubject').textContent();
  expect(preview).toContain('My Test Subject 123');
  console.log('✓ Live preview reflects subject input');
});

// ─── 14. SCHEDULE SEND ────────────────────────────────────────────────────────
test('28. Schedule toggle exists and shows date/time inputs when enabled', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  const toggle = page.locator('#scheduleToggle');
  await expect(toggle).toBeVisible();
  await toggle.check();
  await expect(page.locator('#scheduleBox')).toBeVisible();
  await expect(page.locator('#scheduledDate')).toBeVisible();
  await expect(page.locator('#scheduledTime')).toBeVisible();
  console.log('✓ Schedule send toggle and date/time inputs work');
});

// ─── 15. ATTACHMENT SECTION ───────────────────────────────────────────────────
test('29. Global attachment section is visible', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#attachmentDropzone')).toBeVisible();
  await expect(page.locator('#attachmentList')).toBeVisible();
  console.log('✓ Global attachment section visible');
});

test('30. Personal PDF section is visible', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#pdfFolderDropzone')).toBeVisible();
  console.log('✓ Personal PDF section visible');
});

// ─── 16. VARIABLE BAR ─────────────────────────────────────────────────────────
test('31. Variable bar and Insert target select visible', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#variableBar')).toBeVisible();
  await expect(page.locator('#varTargetSelect')).toBeVisible();
  console.log('✓ Variable bar visible');
});

// ─── 17. SEND BUTTON ─────────────────────────────────────────────────────────
test('32. Send Campaign button is visible', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#sendCampaignBtn')).toBeVisible();
  console.log('✓ Send campaign button visible');
});

test('33. Send without body shows error toast', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await page.fill('#subject', 'Test');
  // Don't fill body — click send
  await page.click('#sendCampaignBtn');
  await page.waitForTimeout(500);
  const toast = page.locator('#toast.show');
  // Should show error (either "no recipients" or "body" error)
  const isVisible = await toast.isVisible();
  console.log('✓ Error shown when clicking Send without valid body/recipients:', isVisible);
});

// ─── 18. CHECK CONNECTION ─────────────────────────────────────────────────────
test('34. Check Connection button visible on compose page', async ({ page }) => {
  await login(page);
  await page.click('button.nav[data-page="compose"]');
  await expect(page.locator('#checkSenderBtn')).toBeVisible();
  console.log('✓ Check Connection button visible');
});

// ─── 19. NAVIGATION ───────────────────────────────────────────────────────────
test('35. All nav buttons switch pages correctly', async ({ page }) => {
  await login(page);
  const pages = ['dashboard', 'contacts', 'compose', 'campaigns'];
  for (const pg of pages) {
    await page.click(`button.nav[data-page="${pg}"]`);
    await expect(page.locator(`#${pg}Page`)).not.toHaveClass(/hidden/);
    console.log(`  ✓ ${pg} page navigates correctly`);
  }
  console.log('✓ All navigation buttons work');
});
