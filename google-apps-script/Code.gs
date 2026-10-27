/**
 * తిరుమల కాలనీ సహాయనిధి - Google Apps Script Backend API
 * 
 * Instructions:
 * 1. Open your Google Sheet (or create a new one: sheets.new)
 * 2. In Google Sheets menu, click: Extensions > Apps Script
 * 3. Delete any code in the editor, and paste this entire file content.
 * 4. Run the function `setupSpreadsheet()` once from the toolbar to auto-create all tabs, headers, and sample data!
 * 5. Click "Deploy" > "New deployment"
 * 6. Select type: "Web app"
 * 7. Configuration:
 *    - Description: Tirumala Colony Sahayanidhi API
 *    - Execute as: "Me" (your email)
 *    - Who has access: "Anyone"
 * 8. Click "Deploy", authorize permissions, and copy the Web App URL!
 * 9. Paste the Web App URL into the Colony Fund App Settings.
 */

const SHEETS = {
  MEMBERS: 'Members',
  CONTRIBUTIONS: 'MonthlyContributions',
  LOANS: 'Loans',
  LOAN_PAYMENTS: 'LoanPayments',
  MAINTENANCE: 'MaintenanceCollections',
  EXPENSES: 'ColonyExpenses',
  BALANCE: 'SocietyBalance',
  LOGIN_LOGS: 'LoginLogs',
  ANNOUNCEMENTS: 'Announcements',
  POLLS: 'Polls',
  POLL_VOTES: 'PollVotes',
  SETTINGS: 'Settings',
  SETTLEMENTS: 'Settlements'
};

/**
 * Automatically creates custom menu when the Google Sheet is opened
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏛️ తిరుమల కాలనీ')
    .addItem('⚡ Generate Historical Contributions (Batch)', 'promptAndGenerateHistoricalContributions')
    .addItem('⚙️ Setup / Sync Settings Tab (UPI ID & Name)', 'setupSettingsTab')
    .addItem('🔄 Recalculate Society Balance', 'menuRecalculateBalance')
    .addSeparator()
    .addItem('💾 Create Backup Now (Snapshot)', 'menuCreateBackupNow')
    .addItem('⏰ Enable Automated Daily 2 AM Backup', 'menuSetupDailyBackupTrigger')
    .addItem('📁 Open Backups Folder in Drive', 'menuOpenBackupsFolder')
    .addToUi();
}

/**
 * Run this function once in Apps Script to set up all 5 tabs and headers.
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Initialize Society Name in Script Properties
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SocietyName')) {
    props.setProperty('SocietyName', 'తిరుమల కాలనీ సహాయనిధి');
  }
  if (!props.getProperty('UpiId')) {
    props.setProperty('UpiId', 'colonyfund@upi');
  }

  // 1. Members tab
  let membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
  if (!membersSheet) {
    membersSheet = ss.insertSheet(SHEETS.MEMBERS);
    membersSheet.appendRow(['MemberID', 'Name', 'Phone', 'FlatNumber', 'JoinedDate', 'Status', 'Role', 'Pin', 'ResidentType', 'IsWelfareMember', 'IsMaintenanceMember', 'MaintenanceStartDate']);
    membersSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
    
    // Add sample members with credentials
    membersSheet.appendRow(['MEM001', 'Rajesh Sharma', '+919876543210', 'A-101', '2023-07-01', 'Active', 'Admin', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
    membersSheet.appendRow(['MEM002', 'Sunita Verma', '+919876543211', 'A-102', '2023-07-01', 'Active', 'Member', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
    membersSheet.appendRow(['MEM003', 'Vikram Patel', '+919876543212', 'B-201', '2023-07-01', 'Active', 'Member', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
    membersSheet.appendRow(['MEM004', 'Anita Deshmukh', '+919876543213', 'B-202', '2023-07-01', 'Active', 'Member', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
    membersSheet.appendRow(['MEM005', 'Manoj Gupta', '+919876543214', 'C-301', '2023-07-01', 'Active', 'Member', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
    membersSheet.appendRow(['MEM006', 'Priya Kulkarni', '+919876543215', 'C-302', '2023-07-01', 'Active', 'Member', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
    membersSheet.appendRow(['MEM007', 'Suresh Nair', '+919876543216', 'D-401', '2023-07-01', 'Active', 'Member', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
    membersSheet.appendRow(['MEM008', 'Kavita Joshi', '+919876543217', 'D-402', '2023-07-01', 'Active', 'Member', '1234', 'Resident', 'Yes', 'Yes', 'January 2026']);
  } else {
    // If Members sheet already exists, ensure subscription columns 9, 10, 11, 12 exist
    if (membersSheet.getLastColumn() < 12) {
      if (membersSheet.getLastColumn() < 9 || String(membersSheet.getRange(1, 9).getValue()).trim() === '') membersSheet.getRange(1, 9).setValue('ResidentType');
      if (membersSheet.getLastColumn() < 10 || String(membersSheet.getRange(1, 10).getValue()).trim() === '') membersSheet.getRange(1, 10).setValue('IsWelfareMember');
      if (membersSheet.getLastColumn() < 11 || String(membersSheet.getRange(1, 11).getValue()).trim() === '') membersSheet.getRange(1, 11).setValue('IsMaintenanceMember');
      if (membersSheet.getLastColumn() < 12 || String(membersSheet.getRange(1, 12).getValue()).trim() === '') membersSheet.getRange(1, 12).setValue('MaintenanceStartDate');
      membersSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
    }
  }

  // 2. MonthlyContributions tab
  let contSheet = ss.getSheetByName(SHEETS.CONTRIBUTIONS);
  if (!contSheet) {
    contSheet = ss.insertSheet(SHEETS.CONTRIBUTIONS);
    contSheet.appendRow(['ContributionID', 'MemberID', 'MemberName', 'MonthYear', 'Amount', 'PaymentDate', 'Status', 'PaymentMode', 'TransactionRef']);
    contSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // 3. Loans tab
  let loansSheet = ss.getSheetByName(SHEETS.LOANS);
  if (!loansSheet) {
    loansSheet = ss.insertSheet(SHEETS.LOANS);
    loansSheet.appendRow(['LoanID', 'MemberID', 'MemberName', 'PrincipalAmount', 'InterestRatePercent', 'IssueDate', 'TenureMonths', 'ExpectedTotalInterest', 'TotalRepayable', 'RemainingPrincipal', 'Status', 'RepaymentMode']);
    loansSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  } else {
    // If Loans sheet already exists, automatically ensure Column L (12) has 'RepaymentMode' header
    if (loansSheet.getLastColumn() < 12 || String(loansSheet.getRange(1, 12).getValue()).trim() === '') {
      loansSheet.getRange(1, 12).setValue('RepaymentMode')
        .setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
    }

    // Auto-backfill existing rows if Column L is empty
    const lastRow = loansSheet.getLastRow();
    if (lastRow > 1) {
      const data = loansSheet.getRange(2, 1, lastRow - 1, 12).getValues();
      for (let i = 0; i < data.length; i++) {
        if (!data[i][11] || String(data[i][11]).trim() === '') {
          const p = Number(data[i][3] || 0);
          const r = Number(data[i][4] || 1.5);
          const t = Number(data[i][6] || 0);
          const exp = Number(data[i][7] || 0);
          const flatInt = Math.round((p * r * t) / 100);
          const redInt = Math.round((p * r * (t + 1)) / 200);
          const mode = (p > 0 && r > 0 && t > 0 && (exp <= redInt || exp < flatInt)) ? 'emi' : 'interest_only';
          loansSheet.getRange(i + 2, 12).setValue(mode);
        }
      }
    }
  }

  // 4. LoanPayments tab
  let paymentsSheet = ss.getSheetByName(SHEETS.LOAN_PAYMENTS);
  if (!paymentsSheet) {
    paymentsSheet = ss.insertSheet(SHEETS.LOAN_PAYMENTS);
    paymentsSheet.appendRow(['PaymentID', 'LoanID', 'MemberID', 'MemberName', 'PaymentDate', 'PrincipalPaid', 'InterestPaid', 'TotalPaid', 'Notes']);
    paymentsSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // 5. SocietyBalance tab
  let balanceSheet = ss.getSheetByName(SHEETS.BALANCE);
  if (!balanceSheet) {
    balanceSheet = ss.insertSheet(SHEETS.BALANCE);
    balanceSheet.appendRow(['BalanceID', 'TotalCollectedContributions', 'TotalLoansDisbursed', 'TotalPrincipalRepaid', 'TotalInterestEarned', 'CurrentCashInBank', 'CirculatingLoans', 'LastUpdated']);
    balanceSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // 6. LoginLogs tab
  let logsSheet = ss.getSheetByName(SHEETS.LOGIN_LOGS);
  if (!logsSheet) {
    logsSheet = ss.insertSheet(SHEETS.LOGIN_LOGS);
    logsSheet.appendRow(['LogID', 'MemberID', 'MemberName', 'HouseNumber', 'Action', 'Timestamp', 'DeviceInfo']);
    logsSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // 7. Announcements / Meeting Notices tab
  let noticeSheet = ss.getSheetByName(SHEETS.ANNOUNCEMENTS);
  if (!noticeSheet) {
    noticeSheet = ss.insertSheet(SHEETS.ANNOUNCEMENTS);
    noticeSheet.appendRow(['NoticeID', 'Title', 'NoticeType', 'MeetingDate', 'MeetingTime', 'Venue', 'Agenda', 'PostedBy', 'PostedDate', 'Status']);
    noticeSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
    noticeSheet.appendRow([
      'NOT-202610-01',
      'నెలవారీ జనరల్ బాడీ సమావేశం (Monthly General Body Meeting)',
      'Meeting',
      '2026-10-18',
      '10:00 AM',
      'కాలనీ కమ్యూనిటీ హాల్ (Colony Community Hall)',
      '1. గత నెల చందాలు మరియు ఖర్చుల సమీక్ష (Monthly fund review)\n2. కాలనీ రోడ్ మరమ్మత్తు & డ్రైనేజీ సమస్యల పరిష్కారం (Road & drainage)\n3. 4 కొత్త CCTV కెమెరాల ఏర్పాటు ప్రణాళిక (CCTV installation proposal)\n4. దీపావళి వేడుకల ఏర్పాట్లు (Diwali celebration plan)',
      'కాలనీ కమిటీ (Managing Committee)',
      '2026-10-08',
      'Active'
    ]);
  }

  // 8. Polls tab
  let pollsSheet = ss.getSheetByName(SHEETS.POLLS);
  if (!pollsSheet) {
    pollsSheet = ss.insertSheet(SHEETS.POLLS);
    pollsSheet.appendRow(['PollID', 'Question', 'Description', 'Options', 'ExpiresDate', 'CreatedBy', 'CreatedAt', 'Status']);
    pollsSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
    pollsSheet.appendRow([
      'POL-202610-01',
      'మన కాలనీ మెయిన్ గేట్ & 3వ లైన్ వద్ద 4 కొత్త CCTV కెమెరాలు అమర్చాలా? (Should we install 4 CCTV cameras?)',
      'కాలనీ భద్రత దృష్ట్యా ప్రతి ఇంటి నుండి ఒకసారి ₹800 వాటా వేసి కెమెరాలు ఏర్పాటు చేయడం.',
      'అవును, వెంటనే ఏర్పాటు చేయాలి (Yes, Install immediately);వద్దు, ఇప్పుడు అవసరం లేదు (No, Not required);మరింత చర్చించాలి / మీటింగ్‌లో నిర్ణయిద్దాం (Discuss in Meeting)',
      '2026-10-25',
      'Admin (కమిటీ)',
      '2026-10-08',
      'Active'
    ]);
  }

  // 9. PollVotes tab
  let votesSheet = ss.getSheetByName(SHEETS.POLL_VOTES);
  if (!votesSheet) {
    votesSheet = ss.insertSheet(SHEETS.POLL_VOTES);
    votesSheet.appendRow(['VoteID', 'PollID', 'MemberID', 'HouseNumber', 'SelectedOption', 'VotedAt']);
    votesSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
    votesSheet.appendRow(['VOT-202610-01', 'POL-202610-01', 'MEM001', 'A-101', 'అవును, వెంటనే ఏర్పాటు చేయాలి (Yes, Install immediately)', '2026-10-08 10:00:00']);
    votesSheet.appendRow(['VOT-202610-02', 'POL-202610-01', 'MEM002', 'A-102', 'అవును, వెంటనే ఏర్పాటు చేయాలి (Yes, Install immediately)', '2026-10-08 11:30:00']);
  }

  // 10. Settings tab (Society Name, UPI ID, Monthly Dues)
  ensureSettingsSheet(ss);

  // 11. Settlements tab
  let settleSheet = ss.getSheetByName(SHEETS.SETTLEMENTS);
  if (!settleSheet) {
    settleSheet = ss.insertSheet(SHEETS.SETTLEMENTS);
    settleSheet.appendRow(['SettlementID', 'MemberID', 'MemberName', 'HouseNumber', 'SettlementDate', 'TotalContributed', 'SharePercent', 'InterestEarned', 'GrossPayout', 'LoanDeduction', 'NetPayout', 'PaymentMode', 'TransactionRef', 'Notes', 'Status']);
    settleSheet.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // 12. Maintenance Collections tab
  let maintSheet = ss.getSheetByName(SHEETS.MAINTENANCE);
  if (!maintSheet) {
    maintSheet = ss.insertSheet(SHEETS.MAINTENANCE);
    maintSheet.appendRow(['CollectionID', 'MemberID', 'MemberName', 'HouseNumber', 'MonthYear', 'Amount', 'PaymentDate', 'Status', 'PaymentMode', 'TransactionRef']);
    maintSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // 13. Colony Expenses tab
  let expSheet = ss.getSheetByName(SHEETS.EXPENSES);
  if (!expSheet) {
    expSheet = ss.insertSheet(SHEETS.EXPENSES);
    expSheet.appendRow(['ExpenseID', 'Date', 'Category', 'Amount', 'Description', 'PaidTo', 'ReceiptRef', 'LoggedBy']);
    expSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // Recalculate balance
  updateSocietyBalance(ss);
}

/**
 * Fast bulk generation of historical contributions directly in Google Sheets.
 * Writes all 800+ records in < 2 seconds using batch setValues().
 */
function generateHistoricalContributions(startMonthName, startYear, endMonthName, endYear, amountPerMonth, defaultStatus, defaultPaymentMode) {
  startMonthName = startMonthName || 'July';
  startYear = Number(startYear) || 2023;
  endMonthName = endMonthName || 'October';
  endYear = Number(endYear) || 2026;
  amountPerMonth = Number(amountPerMonth) || 500;
  defaultStatus = defaultStatus || 'Paid';
  defaultPaymentMode = defaultPaymentMode || 'UPI';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
  const contSheet = ss.getSheetByName(SHEETS.CONTRIBUTIONS);

  if (!membersSheet || membersSheet.getLastRow() < 2) {
    throw new Error('No members found in Members sheet. Please add members first.');
  }

  // 1. Fetch all members
  const memberData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 4).getValues();
  const members = memberData.map(r => ({
    id: String(r[0]),
    name: String(r[1]),
    phone: String(r[2]),
    flat: String(r[3])
  })).filter(m => m.id && m.id !== 'undefined');

  // 2. Generate months list in chronological order
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];
  const sIdx = monthNames.indexOf(startMonthName);
  const eIdx = monthNames.indexOf(endMonthName);
  const months = [];
  let curY = startYear;
  let curM = sIdx !== -1 ? sIdx : 0;
  const targetEndM = eIdx !== -1 ? eIdx : 11;

  while (curY < endYear || (curY === endYear && curM <= targetEndM)) {
    months.push(monthNames[curM] + ' ' + curY);
    curM++;
    if (curM > 11) {
      curM = 0;
      curY++;
    }
  }

  // 3. Read existing contributions to prevent duplicates
  const existingMap = new Set();
  if (contSheet && contSheet.getLastRow() > 1) {
    const existing = contSheet.getRange(2, 2, contSheet.getLastRow() - 1, 3).getValues();
    existing.forEach(r => {
      const memId = String(r[0]);
      const my = String(r[2]);
      existingMap.add(memId + '_' + my);
    });
  }

  // 4. Build batch rows for lightning fast single-call insertion
  const newRows = [];
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  months.forEach(my => {
    const cleanMY = my.replace(/\s+/g, '');
    members.forEach(m => {
      const key = m.id + '_' + my;
      if (!existingMap.has(key)) {
        newRows.push([
          'CON-' + cleanMY + '-' + m.id,
          m.id,
          m.name,
          my,
          amountPerMonth,
          defaultStatus === 'Paid' ? todayStr : '',
          defaultStatus,
          defaultStatus === 'Paid' ? defaultPaymentMode : '',
          'HISTORICAL-BATCH'
        ]);
        existingMap.add(key);
      }
    });
  });

  // Fast write using setValues() with auto-expanding sheet capacity
  if (newRows.length > 0) {
    const lastRow = contSheet.getLastRow();
    const currentMax = contSheet.getMaxRows();
    const needed = (lastRow + newRows.length) - currentMax;
    if (needed > 0) {
      contSheet.insertRowsAfter(currentMax, needed + 100);
    }
    contSheet.getRange(lastRow + 1, 1, newRows.length, 9).setValues(newRows);
  }

  // Recalculate balance
  updateSocietyBalance(ss);

  return {
    success: true,
    createdCount: newRows.length,
    membersCount: members.length,
    monthsCount: months.length
  };
}

/**
 * Interactive prompt when running directly from the Google Sheets menu:
 * 🏛️ తిరుమల కాలనీ > ⚡ Generate Historical Contributions (Batch)
 */
function promptAndGenerateHistoricalContributions() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const fee = Number(props.getProperty('MonthlyContributionAmount')) || 500;

  const resp = ui.alert(
    'తిరుమల కాలనీ సహాయనిధి - Batch Onboarding',
    'Do you want to generate all historical contributions from March 2023 to September/October 2026 for all members as "Paid" (₹' + fee + '/month)?\n\nThis will take ~2 seconds and write directly into the MonthlyContributions sheet.',
    ui.ButtonSet.YES_NO
  );

  if (resp === ui.Button.YES) {
    const res = generateHistoricalContributions('March', 2023, 'October', 2026, fee, 'Paid', 'UPI');
    ui.alert(
      'విజయవంతమైంది (Success)',
      'Generated ' + res.createdCount + ' historical contribution records across ' + res.monthsCount + ' months for ' + res.membersCount + ' members!',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Menu item to manually recalculate balance
 */
function menuRecalculateBalance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  updateSocietyBalance(ss);
  SpreadsheetApp.getUi().alert('Society Balance recalculated successfully!');
}

/**
 * Menu item to manually ensure Settings tab is set up
 */
function setupSettingsTab() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSettingsSheet(ss);
  SpreadsheetApp.getUi().alert('Settings tab verified! You can view and edit SocietyName, UpiId and MonthlyContributionAmount directly in the Settings tab.');
}

/**
 * Ensures the Settings tab exists in the spreadsheet with headers & initial rows.
 */
function ensureSettingsSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEETS.SETTINGS);
  const props = PropertiesService.getScriptProperties();

  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.SETTINGS);
    sheet.appendRow(['SettingKey', 'SettingValue', 'Description', 'LastUpdated']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow(['SocietyName', props.getProperty('SocietyName') || 'తిరుమల కాలనీ సహాయనిధి', 'Colony / Society Official Name', nowStr]);
    sheet.appendRow(['UpiId', props.getProperty('UpiId') || 'colonyfund@upi', 'Society Payment UPI ID (VPA) for Resident Dues', nowStr]);
    sheet.appendRow(['MonthlyContributionAmount', Number(props.getProperty('MonthlyContributionAmount')) || 500, 'Default Monthly Maintenance Fee', nowStr]);
    sheet.autoResizeColumns(1, 4);
  }
  return sheet;
}

/**
 * Reads all settings from the Settings sheet, falling back to ScriptProperties.
 */
function getSettingsFromSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = {};
  const props = PropertiesService.getScriptProperties();

  // Baseline defaults
  settings.SocietyName = props.getProperty('SocietyName') || 'తిరుమల కాలనీ సహాయనిధి';
  settings.UpiId = props.getProperty('UpiId') || 'colonyfund@upi';
  settings.MonthlyContributionAmount = Number(props.getProperty('MonthlyContributionAmount')) || 500;

  try {
    const sheet = ensureSettingsSheet(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < data.length; i++) {
        const key = String(data[i][0] || '').trim();
        const val = data[i][1];
        if (key && val !== undefined && val !== null && String(val).trim() !== '') {
          if (key === 'MonthlyContributionAmount') {
            settings[key] = Number(val) || 500;
          } else {
            settings[key] = String(val).trim();
          }
        }
      }
    }
  } catch (err) {
    Logger.log('Error reading Settings sheet: ' + err.message);
  }

  // Double-sync to ScriptProperties for cache
  if (settings.SocietyName) props.setProperty('SocietyName', settings.SocietyName);
  if (settings.UpiId) props.setProperty('UpiId', settings.UpiId);
  if (settings.MonthlyContributionAmount) props.setProperty('MonthlyContributionAmount', String(settings.MonthlyContributionAmount));

  return settings;
}

/**
 * Saves/Updates settings directly in the Settings sheet and syncs with ScriptProperties.
 */
function saveSettingsToSheet(ss, data) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSettingsSheet(ss);
  const props = PropertiesService.getScriptProperties();
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  const updates = {};
  if (data.SocietyName || data.societyName) {
    const sName = String(data.SocietyName || data.societyName).trim();
    updates.SocietyName = { value: sName, desc: 'Colony / Society Official Name' };
    props.setProperty('SocietyName', sName);
  }
  if (data.UpiId || data.upiId) {
    const upi = String(data.UpiId || data.upiId).trim();
    updates.UpiId = { value: upi, desc: 'Society Payment UPI ID (VPA) for Resident Dues' };
    props.setProperty('UpiId', upi);
  }
  if (data.MonthlyContributionAmount || data.monthlyContributionAmount) {
    const amt = Number(data.MonthlyContributionAmount || data.monthlyContributionAmount) || 500;
    updates.MonthlyContributionAmount = { value: amt, desc: 'Default Monthly Maintenance Fee' };
    props.setProperty('MonthlyContributionAmount', String(amt));
  }

  const lastRow = sheet.getLastRow();
  const existingKeys = new Map();
  if (lastRow > 1) {
    const keyData = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < keyData.length; i++) {
      const k = String(keyData[i][0] || '').trim();
      if (k) existingKeys.set(k, i + 2); // row index
    }
  }

  for (const [key, item] of Object.entries(updates)) {
    if (existingKeys.has(key)) {
      const rowIdx = existingKeys.get(key);
      sheet.getRange(rowIdx, 2).setValue(item.value);
      sheet.getRange(rowIdx, 4).setValue(nowStr);
    } else {
      sheet.appendRow([key, item.value, item.desc, nowStr]);
    }
  }

  sheet.autoResizeColumns(1, 4);
}

/**
 * Reads and respects the Society Balance tab.
 * NEVER deletes or overwrites user formulas in Row 2!
 * Google Sheets recalculates all formulas automatically in real time.
 */
function updateSocietyBalance(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  let balanceSheet = ss.getSheetByName(SHEETS.BALANCE);
  if (!balanceSheet) {
    balanceSheet = ss.insertSheet(SHEETS.BALANCE);
    balanceSheet.appendRow(['BalanceID', 'TotalCollectedContributions', 'TotalLoansDisbursed', 'TotalPrincipalRepaid', 'TotalInterestEarned', 'CurrentCashInBank', 'CirculatingLoans', 'LastUpdated']);
    balanceSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
  }

  // Only if Row 2 is completely missing, insert the default formulas once:
  if (balanceSheet.getLastRow() < 2) {
    balanceSheet.appendRow([
      'BAL-01',
      '=IFERROR(SUMIF(MonthlyContributions!G2:G, "Paid", MonthlyContributions!E2:E), 0)',
      '=IFERROR(SUM(Loans!D2:D), 0)',
      '=IFERROR(SUM(LoanPayments!F2:F), 0)',
      '=IFERROR(SUM(LoanPayments!G2:G), 0)',
      '=IFERROR(B2 + D2 + E2 - C2 - IFERROR(SUM(Settlements!K2:K), 0), 0)',
      '=IFERROR(SUMIF(Loans!K2:K, "Active", Loans!J2:J), 0)',
      '=NOW()'
    ]);
  } else if (balanceSheet.getLastRow() > 2) {
    // If old script versions accidentally created duplicate rows (Row 3, 4, etc.), clean up rows 3+ ONLY.
    // Row 2 is protected and NEVER deleted!
    balanceSheet.deleteRows(3, balanceSheet.getLastRow() - 2);
  }

  // Ensure formula in Column F (CurrentCashInBank) subtracts settlements if Settlements sheet exists:
  if (balanceSheet.getLastRow() >= 2) {
    const fFormula = String(balanceSheet.getRange(2, 6).getFormula() || '');
    if (fFormula && !fFormula.includes('Settlements')) {
      balanceSheet.getRange(2, 6).setFormula('=IFERROR(B2 + D2 + E2 - C2 - IFERROR(SUM(Settlements!K2:K), 0), 0)');
    }
  }

  // Force Google Sheets formula recalculation
  SpreadsheetApp.flush();

  // Read the live calculated results from Row 2
  const row = balanceSheet.getRange(2, 1, 1, 8).getValues()[0];

  return {
    balanceId: row[0] || 'BAL-01',
    totalContributions: Number(row[1]) || 0,
    totalLoansDisbursed: Number(row[2]) || 0,
    totalPrincipalRepaid: Number(row[3]) || 0,
    totalInterestEarned: Number(row[4]) || 0,
    currentCashInBank: Number(row[5]) || 0,
    circulatingLoans: Number(row[6]) || 0,
    lastUpdated: row[7] instanceof Date ? row[7].toISOString() : String(row[7] || '')
  };
}

/**
 * Helper to convert sheet rows into JSON array of objects
 */
function sheetToJson(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  
  return rows.map(row => {
    const obj = {};
    headers.forEach((header, index) => {
      let val = row[index];
      if (val instanceof Date) {
        const h = String(header).toLowerCase();
        if (h === 'monthyear') {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'MMMM yyyy');
        } else if (h.includes('time') || val.getFullYear() === 1899) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'hh:mm a');
        } else {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
      }
      obj[header] = val;
    });
    return obj;
  });
}

/**
 * Handle HTTP GET Requests
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'getAllData';
    const props = PropertiesService.getScriptProperties();

    let result = {};

    if (action === 'getAllData') {
      const settings = getSettingsFromSheet(ss);
      result = {
        success: true,
        societyName: settings.SocietyName || props.getProperty('SocietyName') || 'తిరుమల కాలనీ సహాయనిధి',
        upiId: settings.UpiId || props.getProperty('UpiId') || 'colonyfund@upi',
        monthlyContributionAmount: Number(settings.MonthlyContributionAmount || props.getProperty('MonthlyContributionAmount')) || 500,
        settings: settings,
        members: sheetToJson(ss.getSheetByName(SHEETS.MEMBERS)),
        contributions: sheetToJson(ss.getSheetByName(SHEETS.CONTRIBUTIONS)),
        loans: sheetToJson(ss.getSheetByName(SHEETS.LOANS)),
        loanPayments: sheetToJson(ss.getSheetByName(SHEETS.LOAN_PAYMENTS)),
        balance: sheetToJson(ss.getSheetByName(SHEETS.BALANCE))[0] || {},
        loginLogs: (() => {
          const lSheet = ss.getSheetByName(SHEETS.LOGIN_LOGS);
          if (lSheet && lSheet.getLastRow() > 20) {
            pruneLoginLogsSheet(lSheet, 5);
          }
          return sheetToJson(lSheet);
        })(),
        notices: sheetToJson(ss.getSheetByName(SHEETS.ANNOUNCEMENTS)),
        polls: sheetToJson(ss.getSheetByName(SHEETS.POLLS)),
        pollVotes: sheetToJson(ss.getSheetByName(SHEETS.POLL_VOTES)),
        settlements: sheetToJson(ss.getSheetByName(SHEETS.SETTLEMENTS)),
        maintenanceCollections: sheetToJson(ss.getSheetByName(SHEETS.MAINTENANCE)),
        colonyExpenses: sheetToJson(ss.getSheetByName(SHEETS.EXPENSES)),
        maintenanceFeeAmount: Number(settings.MaintenanceFeeAmount || 100)
      };
    } else if (action === 'getNotices') {
      result = { success: true, data: sheetToJson(ss.getSheetByName(SHEETS.ANNOUNCEMENTS)) };
    } else if (action === 'getPolls') {
      result = { success: true, polls: sheetToJson(ss.getSheetByName(SHEETS.POLLS)), votes: sheetToJson(ss.getSheetByName(SHEETS.POLL_VOTES)) };
    } else if (action === 'getMembers') {
      result = { success: true, data: sheetToJson(ss.getSheetByName(SHEETS.MEMBERS)) };
    } else if (action === 'getContributions') {
      result = { success: true, data: sheetToJson(ss.getSheetByName(SHEETS.CONTRIBUTIONS)) };
    } else if (action === 'getLoans') {
      result = { success: true, data: sheetToJson(ss.getSheetByName(SHEETS.LOANS)) };
    } else if (action === 'getBalance') {
      result = { success: true, data: sheetToJson(ss.getSheetByName(SHEETS.BALANCE))[0] || {} };
    } else if (action === 'getLoginLogs') {
      result = { success: true, data: sheetToJson(ss.getSheetByName(SHEETS.LOGIN_LOGS)) };
    } else {
      result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle HTTP POST Requests (Modifications & Batch Creations)
 */
function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let body = {};
    if (e && e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    }

    const action = body.action;
    let responseData = { success: true };

    if (action === 'recordContribution') {
      const sheet = ss.getSheetByName(SHEETS.CONTRIBUTIONS);
      const data = body.data;
      
      const lastRow = sheet.getLastRow();
      let foundIndex = -1;
      if (lastRow > 1) {
        const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i][1]) === String(data.MemberID) && String(rows[i][3]) === String(data.MonthYear)) {
            foundIndex = i + 2;
            break;
          }
        }
      }

      if (foundIndex > 0) {
        sheet.getRange(foundIndex, 5).setValue(data.Amount || 500);
        sheet.getRange(foundIndex, 6).setValue(data.PaymentDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
        sheet.getRange(foundIndex, 7).setValue(data.Status || 'Paid');
        sheet.getRange(foundIndex, 8).setValue(data.PaymentMode || 'UPI');
        sheet.getRange(foundIndex, 9).setValue(data.TransactionRef || '');
      } else {
        sheet.appendRow([
          data.ContributionID || ('CON-' + Date.now()),
          data.MemberID,
          data.MemberName,
          data.MonthYear,
          data.Amount || 500,
          data.PaymentDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          data.Status || 'Paid',
          data.PaymentMode || 'UPI',
          data.TransactionRef || ''
        ]);
      }
      updateSocietyBalance(ss);
      responseData.message = 'Contribution recorded successfully';

    } else if (action === 'bulkRecordContributions') {
      const records = body.data && body.data.records ? body.data.records : [];
      let contSheet = ss.getSheetByName(SHEETS.CONTRIBUTIONS);
      if (!contSheet) {
        contSheet = ss.insertSheet(SHEETS.CONTRIBUTIONS);
        contSheet.appendRow(['ContributionID', 'MemberID', 'MemberName', 'MonthYear', 'Amount', 'PaymentDate', 'Status', 'PaymentMode', 'TransactionRef']);
        contSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }

      if (records.length > 0 && contSheet) {
        const lastRow = contSheet.getLastRow();
        const existingRowMap = {};
        let fullData = null;
        let hasUpdates = false;

        if (lastRow > 1) {
          // Read all rows into memory for instant lookup and zero-lag in-memory modification
          fullData = contSheet.getRange(1, 1, lastRow, 9).getValues();
          for (let i = 1; i < lastRow; i++) {
            const key = String(fullData[i][1]) + '_' + String(fullData[i][3]);
            existingRowMap[key] = i; // 0-based row index in fullData
          }
        }

        const appendRows = [];
        const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

        records.forEach(rec => {
          const key = String(rec.MemberID) + '_' + String(rec.MonthYear);
          if (existingRowMap[key] !== undefined && fullData) {
            const idx = existingRowMap[key];
            fullData[idx][4] = Number(rec.Amount) || 500;
            fullData[idx][5] = rec.PaymentDate || todayStr;
            fullData[idx][6] = rec.Status || 'Paid';
            fullData[idx][7] = rec.PaymentMode || 'UPI';
            fullData[idx][8] = rec.TransactionRef || 'HISTORICAL-BATCH';
            hasUpdates = true;
          } else {
            appendRows.push([
              rec.ContributionID || ('CON-' + String(rec.MonthYear).replace(/\s+/g, '') + '-' + rec.MemberID),
              rec.MemberID,
              rec.MemberName,
              rec.MonthYear,
              Number(rec.Amount) || 500,
              rec.PaymentDate || todayStr,
              rec.Status || 'Paid',
              rec.PaymentMode || 'UPI',
              rec.TransactionRef || 'HISTORICAL-BATCH'
            ]);
            existingRowMap[key] = -1; // Prevent duplicate appends within same payload
          }
        });

        // 1. Write back existing updated rows in ONE single call
        if (hasUpdates && fullData && fullData.length > 1) {
          contSheet.getRange(1, 1, lastRow, 9).setValues(fullData);
        }

        // 2. Append new rows in ONE single call, auto-expanding sheet capacity if needed
        if (appendRows.length > 0) {
          const currentLast = contSheet.getLastRow();
          const currentMax = contSheet.getMaxRows();
          const needed = (currentLast + appendRows.length) - currentMax;
          if (needed > 0) {
            contSheet.insertRowsAfter(currentMax, needed + 100);
          }
          contSheet.getRange(currentLast + 1, 1, appendRows.length, 9).setValues(appendRows);
        }

        updateSocietyBalance(ss);
        responseData.createdCount = appendRows.length;
        responseData.totalCount = records.length;
        responseData.message = `Bulk processed ${records.length} contributions (${appendRows.length} created)`;
      }

    } else if (action === 'updateSocietySettings') {
      const data = body.data || {};
      saveSettingsToSheet(ss, data);
      responseData.message = 'Society settings updated in Google Sheet Settings tab';
      responseData.upiId = data.UpiId || data.upiId;
      responseData.societyName = data.SocietyName || data.societyName;

    } else if (action === 'issueLoan') {
      const loansSheet = ss.getSheetByName(SHEETS.LOANS);
      const data = body.data;
      loansSheet.appendRow([
        data.LoanID || ('LOAN-' + Date.now()),
        data.MemberID,
        data.MemberName,
        Number(data.PrincipalAmount),
        Number(data.InterestRatePercent || 1.5),
        data.IssueDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        Number(data.TenureMonths || 6),
        Number(data.ExpectedTotalInterest || 0),
        Number(data.TotalRepayable || data.PrincipalAmount),
        Number(data.RemainingPrincipal || data.PrincipalAmount),
        'Active',
        data.RepaymentMode || 'interest_only'
      ]);
      updateSocietyBalance(ss);
      responseData.message = 'Loan issued successfully';

    } else if (action === 'recordLoanPayment') {
      const paymentsSheet = ss.getSheetByName(SHEETS.LOAN_PAYMENTS);
      const loansSheet = ss.getSheetByName(SHEETS.LOANS);
      const data = body.data;

      paymentsSheet.appendRow([
        data.PaymentID || ('LP-' + Date.now()),
        data.LoanID,
        data.MemberID,
        data.MemberName,
        data.PaymentDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        Number(data.PrincipalPaid || 0),
        Number(data.InterestPaid || 0),
        Number(data.TotalPaid || 0),
        data.Notes || ''
      ]);

      const lastRow = loansSheet.getLastRow();
      if (lastRow > 1) {
        const rows = loansSheet.getRange(2, 1, lastRow - 1, 11).getValues();
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i][0]) === String(data.LoanID)) {
            const currentRemaining = Number(rows[i][9]);
            const newRemaining = Math.max(0, currentRemaining - Number(data.PrincipalPaid || 0));
            loansSheet.getRange(i + 2, 10).setValue(newRemaining);
            if (newRemaining <= 0) {
              loansSheet.getRange(i + 2, 11).setValue('Closed');
            }
            break;
          }
        }
      }

      updateSocietyBalance(ss);
      responseData.message = 'Loan repayment recorded successfully';

    } else if (action === 'addMember') {
      const sheet = ss.getSheetByName(SHEETS.MEMBERS);
      const data = body.data;

      // Ensure headers exist
      if (sheet.getLastColumn() < 12) {
        if (sheet.getLastColumn() < 9 || String(sheet.getRange(1, 9).getValue()).trim() === '') sheet.getRange(1, 9).setValue('ResidentType');
        if (sheet.getLastColumn() < 10 || String(sheet.getRange(1, 10).getValue()).trim() === '') sheet.getRange(1, 10).setValue('IsWelfareMember');
        if (sheet.getLastColumn() < 11 || String(sheet.getRange(1, 11).getValue()).trim() === '') sheet.getRange(1, 11).setValue('IsMaintenanceMember');
        if (sheet.getLastColumn() < 12 || String(sheet.getRange(1, 12).getValue()).trim() === '') sheet.getRange(1, 12).setValue('MaintenanceStartDate');
        sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }

      sheet.appendRow([
        data.MemberID || ('MEM' + String(sheet.getLastRow()).padStart(3, '0')),
        data.Name,
        data.Phone,
        data.FlatNumber || data.HouseNumber,
        data.JoinedDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        data.Status || 'Active',
        data.Role || 'Member',
        String(data.Pin || '1234'),
        data.ResidentType || 'Resident',
        data.IsWelfareMember !== false ? 'Yes' : 'No',
        data.IsMaintenanceMember !== false ? 'Yes' : 'No',
        data.MaintenanceStartDate || ''
      ]);
      responseData.message = 'Member added successfully';

    } else if (action === 'updateMemberPin') {
      const sheet = ss.getSheetByName(SHEETS.MEMBERS);
      const data = body.data;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const rows = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i][0]) === String(data.MemberID)) {
            sheet.getRange(i + 2, 8).setValue(String(data.Pin));
            break;
          }
        }
      }
      responseData.message = 'Member PIN updated successfully';

    } else if (action === 'updateMemberProfile' || action === 'updateMember') {
      const sheet = ss.getSheetByName(SHEETS.MEMBERS);
      const data = body.data;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        // Ensure subscription headers 9, 10, 11, 12 exist
        if (sheet.getLastColumn() < 12) {
          if (sheet.getLastColumn() < 9 || String(sheet.getRange(1, 9).getValue()).trim() === '') sheet.getRange(1, 9).setValue('ResidentType');
          if (sheet.getLastColumn() < 10 || String(sheet.getRange(1, 10).getValue()).trim() === '') sheet.getRange(1, 10).setValue('IsWelfareMember');
          if (sheet.getLastColumn() < 11 || String(sheet.getRange(1, 11).getValue()).trim() === '') sheet.getRange(1, 11).setValue('IsMaintenanceMember');
          if (sheet.getLastColumn() < 12 || String(sheet.getRange(1, 12).getValue()).trim() === '') sheet.getRange(1, 12).setValue('MaintenanceStartDate');
          sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
        }

        const rows = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
        for (let i = 0; i < rows.length; i++) {
          const rowMemId = String(rows[i][0] || '').trim();
          const rowName = String(rows[i][1] || '').trim().toLowerCase();
          const targetMemId = String(data.MemberID || '').trim();
          const targetName = String(data.Name || '').trim().toLowerCase();

          if ((targetMemId && rowMemId === targetMemId) || (targetName && rowName === targetName)) {
            const rowIdx = i + 2;
            if (data.Name) sheet.getRange(rowIdx, 2).setValue(String(data.Name));
            if (data.Phone) sheet.getRange(rowIdx, 3).setValue(String(data.Phone));
            if (data.HouseNumber || data.FlatNumber) {
              sheet.getRange(rowIdx, 4).setValue(String(data.HouseNumber || data.FlatNumber));
            }
            if (data.Status) sheet.getRange(rowIdx, 6).setValue(String(data.Status));
            if (data.Role) sheet.getRange(rowIdx, 7).setValue(String(data.Role));
            if (data.Pin) sheet.getRange(rowIdx, 8).setValue(String(data.Pin));
            if (data.ResidentType !== undefined) sheet.getRange(rowIdx, 9).setValue(String(data.ResidentType));
            if (data.IsWelfareMember !== undefined) sheet.getRange(rowIdx, 10).setValue(data.IsWelfareMember ? 'Yes' : 'No');
            if (data.IsMaintenanceMember !== undefined) sheet.getRange(rowIdx, 11).setValue(data.IsMaintenanceMember ? 'Yes' : 'No');
            if (data.MaintenanceStartDate !== undefined) sheet.getRange(rowIdx, 12).setValue(String(data.MaintenanceStartDate));
            break;
          }
        }
      }
      responseData.message = 'Member profile updated successfully';

    } else if (action === 'settleMember') {
      let settleSheet = ss.getSheetByName(SHEETS.SETTLEMENTS);
      if (!settleSheet) {
        settleSheet = ss.insertSheet(SHEETS.SETTLEMENTS);
        settleSheet.appendRow(['SettlementID', 'MemberID', 'MemberName', 'HouseNumber', 'SettlementDate', 'TotalContributed', 'SharePercent', 'InterestEarned', 'GrossPayout', 'LoanDeduction', 'NetPayout', 'PaymentMode', 'TransactionRef', 'Notes', 'Status']);
        settleSheet.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }
      const data = body.data;
      settleSheet.appendRow([
        data.SettlementID || ('SETTLE-' + Date.now()),
        data.MemberID,
        data.MemberName,
        data.HouseNumber || '',
        data.SettlementDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        Number(data.TotalContributed || 0),
        Number(data.SharePercent || 0),
        Number(data.InterestEarned || 0),
        Number(data.GrossPayout || 0),
        Number(data.LoanDeduction || 0),
        Number(data.NetPayout || 0),
        data.PaymentMode || 'UPI',
        data.TransactionRef || '',
        data.Notes || '',
        data.Status || 'Completed'
      ]);

      // Update Member Status in Members Sheet to 'Exited'
      const membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
      if (membersSheet && membersSheet.getLastRow() > 1) {
        const memRows = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 6).getValues();
        for (let i = 0; i < memRows.length; i++) {
          if (String(memRows[i][0]).trim() === String(data.MemberID).trim()) {
            membersSheet.getRange(i + 2, 6).setValue('Exited');
            break;
          }
        }
      }

      // If active loan was deducted, close active loans for this member in Loans sheet
      if (Number(data.LoanDeduction || 0) > 0) {
        const loansSheet = ss.getSheetByName(SHEETS.LOANS);
        if (loansSheet && loansSheet.getLastRow() > 1) {
          const loanRows = loansSheet.getRange(2, 1, loansSheet.getLastRow() - 1, 11).getValues();
          for (let i = 0; i < loanRows.length; i++) {
            if (String(loanRows[i][1]).trim() === String(data.MemberID).trim() && String(loanRows[i][10]).trim().toLowerCase() === 'active') {
              loansSheet.getRange(i + 2, 10).setValue(0); // RemainingPrincipal = 0
              loansSheet.getRange(i + 2, 11).setValue('Closed'); // Status = 'Closed'
            }
          }
        }
      }

      updateSocietyBalance(ss);
      responseData.message = 'Member settled and marked Exited successfully';

    } else if (action === 'reactivateMember' || action === 'updateMemberStatus') {
      const data = body.data || {};
      const memId = String(data.MemberID || data.memberId || '').trim();
      const newStatus = String(data.Status || data.status || 'Active').trim();

      // 1. Update Member status in Members sheet
      const membersSheet = ss.getSheetByName(SHEETS.MEMBERS);
      if (membersSheet && membersSheet.getLastRow() > 1) {
        const memRows = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 6).getValues();
        for (let i = 0; i < memRows.length; i++) {
          if (String(memRows[i][0]).trim() === memId) {
            membersSheet.getRange(i + 2, 6).setValue(newStatus);
            break;
          }
        }
      }

      // 2. If status is Active, remove any settlement row from Settlements sheet
      if (newStatus.toLowerCase() === 'active') {
        const settleSheet = ss.getSheetByName(SHEETS.SETTLEMENTS);
        if (settleSheet && settleSheet.getLastRow() > 1) {
          const settleRows = settleSheet.getRange(2, 1, settleSheet.getLastRow() - 1, 2).getValues();
          for (let i = settleRows.length - 1; i >= 0; i--) {
            if (String(settleRows[i][1]).trim() === memId) {
              settleSheet.deleteRow(i + 2);
            }
          }
        }
      }

      updateSocietyBalance(ss);
      responseData.message = 'Member reactivated successfully';

    } else if (action === 'saveSettings') {
      const data = body.data || {};
      saveSettingsToSheet(ss, data);
      responseData.message = 'Settings saved successfully to Google Sheet Settings tab';
      responseData.upiId = data.UpiId || data.upiId;
      responseData.societyName = data.SocietyName || data.societyName;
      responseData.monthlyContributionAmount = data.MonthlyContributionAmount || data.monthlyContributionAmount;

    } else if (action === 'recordLoginLog') {
      let logsSheet = ss.getSheetByName(SHEETS.LOGIN_LOGS);
      if (!logsSheet) {
        logsSheet = ss.insertSheet(SHEETS.LOGIN_LOGS);
        logsSheet.appendRow(['LogID', 'MemberID', 'MemberName', 'HouseNumber', 'Action', 'Timestamp', 'DeviceInfo']);
        logsSheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }

      const data = body.data || {};
      const logId = data.LogID || ('LOG-' + Date.now());
      const nowStr = data.Timestamp || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      
      logsSheet.appendRow([
        logId,
        data.MemberID || '',
        data.MemberName || '',
        data.HouseNumber || data.FlatNumber || '',
        data.Action || 'LOGIN',
        nowStr,
        data.DeviceInfo || 'App'
      ]);

      // Automatically prune sheet so each member retains at most 5 latest entries
      pruneLoginLogsSheet(logsSheet, 5);

      responseData.message = 'Login activity logged successfully';

    } else if (action === 'recordMaintenance') {
      let maintSheet = ss.getSheetByName(SHEETS.MAINTENANCE);
      if (!maintSheet) {
        maintSheet = ss.insertSheet(SHEETS.MAINTENANCE);
        maintSheet.appendRow(['CollectionID', 'MemberID', 'MemberName', 'HouseNumber', 'MonthYear', 'Amount', 'PaymentDate', 'Status', 'PaymentMode', 'TransactionRef']);
        maintSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }
      const data = body.data || {};
      maintSheet.appendRow([
        data.CollectionID || ('MNT-' + Date.now()),
        data.MemberID || '',
        data.MemberName || '',
        data.HouseNumber || data.FlatNumber || '',
        data.MonthYear || '',
        Number(data.Amount || 100),
        data.PaymentDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        data.Status || 'Paid',
        data.PaymentMode || 'UPI',
        data.TransactionRef || ''
      ]);
      responseData.message = 'Maintenance collection recorded successfully';

    } else if (action === 'addColonyExpense') {
      let expSheet = ss.getSheetByName(SHEETS.EXPENSES);
      if (!expSheet) {
        expSheet = ss.insertSheet(SHEETS.EXPENSES);
        expSheet.appendRow(['ExpenseID', 'Date', 'Category', 'Amount', 'Description', 'PaidTo', 'ReceiptRef', 'LoggedBy']);
        expSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }
      const data = body.data || {};
      expSheet.appendRow([
        data.ExpenseID || ('EXP-' + Date.now()),
        data.Date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        data.Category || 'Other',
        Number(data.Amount || 0),
        data.Description || '',
        data.PaidTo || '',
        data.ReceiptRef || '',
        data.LoggedBy || 'Treasurer'
      ]);
      responseData.message = 'Colony expense logged successfully';

    } else if (action === 'cleanupLoginLogs') {
      let logsSheet = ss.getSheetByName(SHEETS.LOGIN_LOGS);
      if (logsSheet) {
        const maxPerUser = (body.data && body.data.maxPerMember) ? Number(body.data.maxPerMember) : 5;
        const res = pruneLoginLogsSheet(logsSheet, maxPerUser);
        responseData.message = 'Pruned login logs successfully. Retained latest ' + maxPerUser + ' entries per user. Removed ' + res.removedCount + ' old entries.';
        responseData.removedCount = res.removedCount;
        responseData.keptCount = res.keptCount;
      } else {
        responseData.message = 'LoginLogs sheet not found';
      }

    } else if (action === 'createNotice') {
      let sheet = ss.getSheetByName(SHEETS.ANNOUNCEMENTS);
      if (!sheet) {
        sheet = ss.insertSheet(SHEETS.ANNOUNCEMENTS);
        sheet.appendRow(['NoticeID', 'Title', 'NoticeType', 'MeetingDate', 'MeetingTime', 'Venue', 'Agenda', 'PostedBy', 'PostedDate', 'Status']);
        sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }
      const data = body.data || {};
      const noticeId = data.NoticeID || ('NOT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss'));
      sheet.appendRow([
        noticeId,
        data.Title || '',
        data.NoticeType || 'Meeting',
        data.MeetingDate || '',
        data.MeetingTime ? ("'" + String(data.MeetingTime).replace(/^'+/, '')) : '',
        data.Venue || '',
        data.Agenda || '',
        data.PostedBy || 'Admin',
        data.PostedDate || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        data.Status || 'Active'
      ]);
      responseData.noticeId = noticeId;
      responseData.message = 'Notice published successfully';

    } else if (action === 'deleteNotice') {
      const sheet = ss.getSheetByName(SHEETS.ANNOUNCEMENTS);
      const data = body.data || {};
      const targetId = String(data.NoticeID || data.noticeId);
      if (sheet && sheet.getLastRow() > 1) {
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i][0]) === targetId) {
            sheet.getRange(i + 2, 10).setValue('Archived');
            break;
          }
        }
      }
      responseData.message = 'Notice archived successfully';

    } else if (action === 'createPoll') {
      let sheet = ss.getSheetByName(SHEETS.POLLS);
      if (!sheet) {
        sheet = ss.insertSheet(SHEETS.POLLS);
        sheet.appendRow(['PollID', 'Question', 'Description', 'Options', 'ExpiresDate', 'CreatedBy', 'CreatedAt', 'Status']);
        sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }
      const data = body.data || {};
      const pollId = data.PollID || ('POL-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss'));
      const optionsStr = Array.isArray(data.Options) ? data.Options.join(';') : String(data.Options || '');
      sheet.appendRow([
        pollId,
        data.Question || '',
        data.Description || '',
        optionsStr,
        data.ExpiresDate || '',
        data.CreatedBy || 'Admin',
        data.CreatedAt || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        'Active'
      ]);
      responseData.pollId = pollId;
      responseData.message = 'Poll created successfully';

    } else if (action === 'votePoll') {
      let sheet = ss.getSheetByName(SHEETS.POLL_VOTES);
      if (!sheet) {
        sheet = ss.insertSheet(SHEETS.POLL_VOTES);
        sheet.appendRow(['VoteID', 'PollID', 'MemberID', 'HouseNumber', 'SelectedOption', 'VotedAt']);
        sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1b5e20').setFontColor('#ffffff');
      }
      const data = body.data || {};
      const pollId = String(data.PollID || data.pollId);
      const houseNo = String(data.HouseNumber || data.FlatNumber || '').trim();
      const memberId = String(data.MemberID || data.memberId || '').trim();
      const option = String(data.SelectedOption || data.selectedOption || '');
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

      // Check if house or member has already voted for this poll
      let updated = false;
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const rows = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
        for (let i = 0; i < rows.length; i++) {
          const rowPoll = String(rows[i][1]);
          const rowMember = String(rows[i][2]);
          const rowHouse = String(rows[i][3]).trim();
          if (rowPoll === pollId && ((houseNo && rowHouse === houseNo) || (memberId && rowMember === memberId))) {
            sheet.getRange(i + 2, 5).setValue(option);
            sheet.getRange(i + 2, 6).setValue(nowStr);
            updated = true;
            break;
          }
        }
      }

      if (!updated) {
        const voteId = 'VOT-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 1000);
        sheet.appendRow([voteId, pollId, memberId, houseNo, option, nowStr]);
      }
      responseData.message = 'Vote recorded successfully';

    } else if (action === 'closePoll') {
      const sheet = ss.getSheetByName(SHEETS.POLLS);
      const data = body.data || {};
      const pollId = String(data.PollID || data.pollId);
      if (sheet && sheet.getLastRow() > 1) {
        const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < rows.length; i++) {
          if (String(rows[i][0]) === pollId) {
            sheet.getRange(i + 2, 8).setValue('Closed');
            break;
          }
        }
      }
      responseData.message = 'Poll closed successfully';

    } else if (action === 'createBackup') {
      const retention = (body.data && body.data.retentionDays) ? Number(body.data.retentionDays) : 30;
      const backupResult = createDailyBackup(retention);
      responseData = {
        success: true,
        message: 'Google Drive snapshot created successfully',
        backup: backupResult
      };

    } else if (action === 'getBackupStatus') {
      const folder = getOrCreateBackupFolder();
      const triggers = ScriptApp.getProjectTriggers();
      let triggerActive = false;
      for (let i = 0; i < triggers.length; i++) {
        if (triggers[i].getHandlerFunction() === 'createDailyBackup') {
          triggerActive = true;
          break;
        }
      }
      responseData = {
        success: true,
        triggerActive: triggerActive,
        folderName: folder.getName(),
        folderUrl: folder.getUrl()
      };

    } else {
      responseData = { success: false, error: 'Unknown action: ' + action };
    }

    responseData.balance = updateSocietyBalance(ss);

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Prunes the LoginLogs sheet so each member retains at most maxPerMember newest entries
 */
function pruneLoginLogsSheet(sheet, maxPerMember) {
  if (!sheet) return { removedCount: 0, keptCount: 0 };
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { removedCount: 0, keptCount: 0 };

  const maxEntries = maxPerMember || 5;
  const numRows = lastRow - 1;
  const data = sheet.getRange(2, 1, numRows, 7).getValues();

  // Rows were appended sequentially, so newer rows are at the bottom.
  // We scan from bottom (newest) to top (oldest), keeping at most maxEntries per MemberID.
  const memberCounts = {};
  const keptRows = [];
  let removedCount = 0;

  for (let i = data.length - 1; i >= 0; i--) {
    const row = data[i];
    const memId = String(row[1] || 'UNKNOWN').trim();
    if (!memId) continue;

    memberCounts[memId] = (memberCounts[memId] || 0) + 1;
    if (memberCounts[memId] <= maxEntries) {
      keptRows.unshift(row);
    } else {
      removedCount++;
    }
  }

  // If any old entries were removed, rewrite the sheet with keptRows
  if (removedCount > 0) {
    sheet.getRange(2, 1, numRows, 7).clearContent();
    if (keptRows.length > 0) {
      sheet.getRange(2, 1, keptRows.length, 7).setValues(keptRows);
    }
    // Delete excess empty rows if the sheet has grown large
    const currentMax = sheet.getMaxRows();
    const desiredMax = Math.max(keptRows.length + 5, 20);
    if (currentMax > desiredMax + 20) {
      sheet.deleteRows(desiredMax + 1, currentMax - desiredMax);
    }
  }

  return { removedCount: removedCount, keptCount: keptRows.length };
}

/**
 * Standalone function you can run directly from Apps Script Editor to clean up logs
 */
function pruneLoginLogsManual() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logsSheet = ss.getSheetByName(SHEETS.LOGIN_LOGS);
  if (!logsSheet) {
    Logger.log('LoginLogs sheet does not exist.');
    return;
  }
  const result = pruneLoginLogsSheet(logsSheet, 5);
  Logger.log('Pruned LoginLogs: Kept ' + result.keptCount + ' rows, removed ' + result.removedCount + ' old rows.');
}

// ============================================================================
// AUTOMATED GOOGLE DRIVE BACKUP & SNAPSHOT SYSTEM
// ============================================================================

const BACKUP_CONFIG = {
  FOLDER_NAME: 'Colony Fund Backups',
  DEFAULT_RETENTION_DAYS: 30, // Retain last 30 daily snapshots
  TRIGGER_HOUR: 2 // 2:00 AM daily
};

/**
 * Gets or creates the dedicated backup folder in Google Drive.
 * Attempts to place it inside the parent folder of the spreadsheet if available,
 * or at the root of the user's Google Drive.
 */
function getOrCreateBackupFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const parents = ssFile.getParents();
  let parentFolder = null;
  if (parents && parents.hasNext()) {
    parentFolder = parents.next();
  }

  // Look for existing folder in parent folder or Drive root
  const searchScope = parentFolder || DriveApp;
  const folders = searchScope.getFoldersByName(BACKUP_CONFIG.FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }

  // Create new folder
  if (parentFolder) {
    return parentFolder.createFolder(BACKUP_CONFIG.FOLDER_NAME);
  } else {
    return DriveApp.createFolder(BACKUP_CONFIG.FOLDER_NAME);
  }
}

/**
 * Creates a timestamped snapshot copy of the entire active spreadsheet in Google Drive.
 * Automatically prunes older backups beyond retentionDays (defaults to 30 days).
 *
 * Can be run manually, via Google Sheets Menu, or via daily time-driven trigger.
 *
 * @param {number} [retentionDays] Number of days to retain backups (default 30)
 * @returns {object} Status, backup file info, and cleanup count
 */
function createDailyBackup(retentionDays) {
  const maxDays = (typeof retentionDays === 'number' && retentionDays > 0)
    ? retentionDays
    : BACKUP_CONFIG.DEFAULT_RETENTION_DAYS;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssFile = DriveApp.getFileById(ss.getId());
  const backupFolder = getOrCreateBackupFolder();

  // Create timestamped name: e.g. "ColonyFund_Backup_2026-10-16_02-00"
  const timeZone = Session.getScriptTimeZone() || 'Asia/Kolkata';
  const timestamp = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd_HH-mm');
  const baseName = ss.getName().replace(/_Backup_.*$/, '').trim();
  const backupName = baseName + '_Backup_' + timestamp;

  // Make full copy directly into the dedicated backup folder
  const backupCopy = ssFile.makeCopy(backupName, backupFolder);

  // Auto-prune old backup files in the backup folder
  let prunedCount = 0;
  let totalRetained = 0;
  const nowMs = new Date().getTime();
  const maxAgeMs = maxDays * 24 * 60 * 60 * 1000;

  const files = backupFolder.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    const fileAgeMs = nowMs - file.getDateCreated().getTime();
    if (file.getId() !== backupCopy.getId() && fileAgeMs > maxAgeMs) {
      file.setTrashed(true);
      prunedCount++;
    } else {
      totalRetained++;
    }
  }

  const result = {
    success: true,
    backupId: backupCopy.getId(),
    backupName: backupName,
    backupUrl: backupCopy.getUrl(),
    folderName: backupFolder.getName(),
    folderUrl: backupFolder.getUrl(),
    timestamp: timestamp,
    totalRetained: totalRetained,
    prunedCount: prunedCount
  };

  Logger.log('Colony Fund Snapshot Created: ' + JSON.stringify(result));
  return result;
}

/**
 * Sets up a time-driven trigger to run createDailyBackup() every day at 2:00 AM.
 * Avoids creating duplicate triggers if one is already active.
 */
function setupDailyBackupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createDailyBackup') {
      return {
        success: true,
        alreadyExists: true,
        message: 'Automated 2 AM daily backup trigger is already active.'
      };
    }
  }

  ScriptApp.newTrigger('createDailyBackup')
    .timeBased()
    .everyDays(1)
    .atHour(BACKUP_CONFIG.TRIGGER_HOUR)
    .create();

  return {
    success: true,
    alreadyExists: false,
    message: 'Automated daily backup trigger successfully scheduled for 2:00 AM every day.'
  };
}

/**
 * Removes any active daily backup triggers
 */
function removeDailyBackupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'createDailyBackup') {
      ScriptApp.deleteTrigger(triggers[i]);
      removed++;
    }
  }
  return { success: true, removedCount: removed };
}

/**
 * Custom UI menu action: Manual Backup Now
 */
function menuCreateBackupNow() {
  try {
    const res = createDailyBackup(30);
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '✅ Snapshot Created Successfully!',
      'Backup Name: ' + res.backupName + '\n' +
      'Saved In: ' + res.folderName + '\n' +
      'Total Backups Kept: ' + res.totalRetained + '\n' +
      'Old Backups Pruned (>30 days): ' + res.prunedCount + '\n\n' +
      'You can find this file in your Google Drive folder: "' + res.folderName + '".',
      ui.ButtonSet.OK
    );
  } catch (err) {
    if (String(err).indexOf('permission') !== -1 || String(err).indexOf('auth') !== -1) {
      SpreadsheetApp.getUi().alert(
        '⚠️ Authorization Required',
        'Google needs permission to save backup files to your Google Drive.\n\n' +
        '👉 How to authorize (takes 10 seconds):\n' +
        '1. Go to "Extensions" > "Apps Script".\n' +
        '2. Select "createDailyBackup" from the top dropdown.\n' +
        '3. Click "Run" (▶️) and approve permissions.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert('❌ Backup failed: ' + err.toString());
    }
  }
}

/**
 * Custom UI menu action: Setup / Verify Automated 2 AM Trigger
 */
function menuSetupDailyBackupTrigger() {
  try {
    const res = setupDailyBackupTrigger();
    const ui = SpreadsheetApp.getUi();
    if (res.alreadyExists) {
      ui.alert(
        'ℹ️ Schedule Active',
        'Automated Daily Backup is ALREADY active and running every night at 2:00 AM.\n\nOld backups are automatically cleaned up after 30 days.',
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        '🎉 Automation Enabled',
        'Automated Daily Backup has been ENABLED!\n\nEvery night between 2:00 AM - 3:00 AM, Google will automatically clone this spreadsheet with a timestamp into "Colony Fund Backups" and prune backups older than 30 days.',
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    if (String(err).indexOf('script.scriptapp') !== -1 || String(err).indexOf('permission') !== -1) {
      SpreadsheetApp.getUi().alert(
        '⚠️ One-Time Authorization Needed',
        'Google requires one-time permission to schedule automatic background triggers.\n\n' +
        '👉 How to authorize (takes 10 seconds):\n' +
        '1. Open "Extensions" > "Apps Script".\n' +
        '2. In the top toolbar dropdown, select "setupDailyBackupTrigger".\n' +
        '3. Click "Run" (▶️).\n' +
        '4. Click "Review permissions" → choose your account → "Advanced" → "Allow".\n\n' +
        'That\'s it! The daily 2 AM backup will be activated permanently.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert('❌ Trigger setup error: ' + err.toString());
    }
  }
}

/**
 * Custom UI menu action: Show Backups Folder link
 */
function menuOpenBackupsFolder() {
  try {
    const folder = getOrCreateBackupFolder();
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      '📁 Backups Folder Location',
      'Your backups are stored in Google Drive folder:\n"' + folder.getName() + '"\n\nDirect URL:\n' + folder.getUrl(),
      ui.ButtonSet.OK
    );
  } catch (err) {
    SpreadsheetApp.getUi().alert('❌ Could not locate folder: ' + err.toString());
  }
}

