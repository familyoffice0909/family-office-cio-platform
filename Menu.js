function onOpen() {

  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Family Office CIO')

    // ===============================
    // Platform
    // ===============================
    .addSubMenu(
      ui.createMenu('Platform')
        .addItem('Run Bootstrap', 'foBootstrap')
        .addItem('Run Platform Health Check', 'foRunPlatformHealthCheck')
        .addItem('Run Platform Integrity Check', 'foRunPlatformIntegrityCheck')
        .addItem('Run Data Validation', 'foRunDataValidation')
        .addSeparator()
        .addItem('Show Version', 'foShowVersion')
    )

    // ===============================
    // Investments
    // ===============================
    .addSubMenu(
      ui.createMenu('Investments')

        .addItem('Append Sample Recommendation', 'foAppendSampleRecommendation')

        .addSeparator()

        .addItem('Rebuild Portfolio Snapshot', 'foRebuildPortfolioState')

        .addItem('Run Market Data Refresh', 'foRunMarketDataRefresh')

        .addItem('Run Portfolio Valuation', 'foRunPortfolioValuation')

        .addItem('Run Portfolio Data Integrity', 'foRunPortfolioDataIntegrity')

        .addItem('Run Market Intelligence', 'foRunMarketIntelligence')

        .addItem('Run CIO Decision Engine', 'foRunCioDecisionEngine')

        .addItem('Run Executive Report', 'foRunExecutiveReportEngine')

        .addItem('Run Executive Dashboard', 'foRunExecutiveDashboardEngine')
    )

    // ===============================
    // Administration
    // ===============================
    .addSubMenu(
      ui.createMenu('Administration')

        .addItem('Seed Known CDRs', 'foSeedKnownCDRs')

        .addItem('Archive Smoke Test Report', 'foArchiveSmokeTestReport')

        .addItem('Create Dashboard Backup', 'foCreateDashboardBackup')

        .addItem('List Triggers', 'foListTriggers')
    )

    // ===============================
    // Smoke Tests
    // ===============================
    .addSubMenu(
      ui.createMenu('Smoke Tests')

        .addItem('Platform Smoke Test', 'foRunModularSmokeTest')

        .addItem('Market Data Smoke Test', 'foRunMarketDataSmokeTest')

        .addItem('Portfolio Valuation Smoke Test', 'foRunPortfolioValuationSmokeTest')

        .addItem('Portfolio Data Integrity Smoke Test', 'foRunPortfolioDataIntegritySmokeTest')

        .addItem('Autonomous CIO Smoke Test', 'foRunAutonomousCioOrchestratorSmokeTest')
    )

    // ===============================
    // Autonomous CIO
    // ===============================
    .addSubMenu(
      ui.createMenu('Autonomous CIO')

        .addItem('Run Autonomous CIO', 'foRunAutonomousCioOrchestrator')
    )

    .addToUi();

}