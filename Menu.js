function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Family Office CIO')
    .addItem('Run Bootstrap', 'foBootstrap')
    .addItem('Run Platform Health Check', 'foRunPlatformHealthCheck')
    .addItem('Run Platform Integrity Check', 'foRunPlatformIntegrityCheck')
    .addItem('Show Version', 'foShowVersion')
    .addSeparator()
    .addItem('Append Sample Recommendation', 'foAppendSampleRecommendation')
    .addItem('Rebuild Portfolio State', 'foRebuildPortfolioState')
    .addItem('Seed Known CDRs', 'foSeedKnownCDRs')
    .addItem('Archive Smoke Test Report', 'foArchiveSmokeTestReport')
    .addSeparator()
    .addItem('List Triggers', 'foListTriggers')
    .addItem('Create Dashboard Backup', 'foCreateDashboardBackup')
    .addSeparator()
    .addItem('Run Modular Smoke Test', 'foRunModularSmokeTest')
    .addToUi();
}