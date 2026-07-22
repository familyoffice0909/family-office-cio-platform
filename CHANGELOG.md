# Changelog

All notable changes to the Family Office CIO Platform are documented here.

The project follows [Semantic Versioning](https://semver.org/).

## [2.9.0] - 2026-07-21

### Added
- Materiality and Prioritization Intelligence for executive investment decisions
- Portfolio-level materiality assessment and prioritization logic
- Enhanced capital deployment priority evaluation
- Expanded investment decision-support intelligence
- Trend-aware materiality and prioritization integration
- New materiality and prioritization regression test suite

### Changed
- Enhanced `CapitalDeploymentPriorityEngine`
- Enhanced `ExecutiveDecisionIntegrationA233`
- Enhanced `InvestmentDecisionSupportEngine`
- Enhanced `InvestmentTrendEngine`
- Enhanced `PortfolioMaterialityEngine`
- Updated worksheet schema registry and related regression tests

### Validation
- Platform validator passed with zero warnings and zero errors
- ESLint passed
- Jest passed: 5 test suites and 46 tests
- Production promotion completed through pull request #51
- Production release published as `v2.9.0`

### Baseline
- Production repository tag: `v2.9.0`
- Production commit: `7c63aea594784a1cd607f8cd33fff7a54e81cf6a`
- Lab baseline tag: `v2.9.0-lab`
- Sprint 3.0.0 and subsequent development must branch from the certified v2.9.0 lab baseline

## [Unreleased]

### Added
- Enterprise engineering documentation
- Repository governance and release checklist
- Automated module, menu, orchestrator, version, manifest, duplicate-function, and secret validation
- CI-generated smoke-test inventory artifact

### Planned
- Buy Zone Intelligence hardening
- Decision explainability
- Recommendation change detection

## [1.0.0] - 2026-07-10

### Added
- Enterprise Apps Script architecture and module separation
- Configuration, logging, spreadsheet, version, backup, validation, and trigger services
- Bootstrap, platform health, platform integrity, and modular smoke-test framework
- Portfolio valuation, data integrity, performance, exposure, attribution, and reconciliation engines
- Market data integration and market symbol registry
- Recommendation, market intelligence, CIO decision, executive reporting, and dashboard engines
- Autonomous CIO orchestrator with run and step logging
- Buy Zone Intelligence baseline
- GitHub Actions continuous-integration baseline
- Production tag and GitHub release `v1.0.0`

### Known limitations
- Apps Script uses a seeded IBKR snapshot rather than direct live broker ingestion
- Automated GitHub-to-Apps-Script deployment requires credential hardening

