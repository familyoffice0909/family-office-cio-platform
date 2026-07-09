/*********************************************************
 * ModuleRegistry.gs
 * Wave 2.3.1
 *********************************************************/

const FO_MODULES = {
  HEALTH: foRunPlatformHealthCheck,
  INTEGRITY: foRunPlatformIntegrityCheck,
  VALIDATION: foRunDataValidation,
  MARKET_DATA: foRunMarketDataRefresh,
  VALUATION: foRunPortfolioValuation,
  PORTFOLIO_DATA_INTEGRITY: foRunPortfolioDataIntegrity,
  PERFORMANCE: foRunPortfolioPerformance,
  PORTFOLIO: foBuildPortfolioSnapshot,
  MARKET: foRunMarketIntelligence,
  CIO: foRunCioDecisionEngine,
  REPORT: foRunExecutiveReportEngine,
  DASHBOARD: foRunExecutiveDashboardEngine
};

function foGetModule(name) {
  if (!FO_MODULES[name]) {
    throw new Error('Module not registered: ' + name);
  }

  return FO_MODULES[name];
}