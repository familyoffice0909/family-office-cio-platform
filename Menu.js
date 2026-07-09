/************************************************************
 * Menu.gs
 * Family Office CIO Platform
 * Enterprise Menu Structure
 ************************************************************/

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  const platformMenu = ui.createMenu('Platform')
    .addItem('Bootstrap Platform', 'foBootstrap')
    .addItem('Platform Health Check', 'foRunPlatformHealthCheck')
    .addItem('Platform Integrity Check', 'foRunPlatformIntegrityCheck')
    .addItem('Run Data Validation', 'foRunDataValidation');

  const investmentsMenu = ui.createMenu('Investments')
    .addItem('Recommendation Engine Smoke Test', 'foRunRecommendationEngineSmokeTest')
    .addItem('Append Sample Recommendation', 'foAppendSampleRecommendation')
    .addSeparator()
    .addItem('Portfolio Engine Smoke Test', 'foRunPortfolioEngineSmokeTest')
    .addItem('Build Portfolio Snapshot', 'foBuildPortfolioSnapshot')
    .addItem('Rebuild Portfolio State', 'foRebuildPortfolioState')
    .addItem('Seed Known CDRs', 'foSeedKnownCDRs');

  const reportsMenu = ui.createMenu('Reports')
    .addItem('Archive Smoke Test Report', 'foArchiveSmokeTestReport');

  const adminMenu = ui.createMenu('Administration')
    .addItem('Show Platform Version', 'foShowVersion')
    .addItem('List Installed Triggers', 'foListTriggers')
    .addItem('Create Dashboard Backup', 'foCreateDashboardBackup');

  const diagnosticsMenu = ui.createMenu('Diagnostics')
    .addItem('Run Modular Smoke Test', 'foRunModularSmokeTest');

  ui.createMenu('Family Office CIO')
    .addSubMenu(platformMenu)
    .addSubMenu(investmentsMenu)
    .addSubMenu(reportsMenu)
    .addSubMenu(adminMenu)
    .addSubMenu(diagnosticsMenu)
    .addToUi();
}