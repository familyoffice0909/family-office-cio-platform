function foUpsertMarketAccess(record) {
  const module = 'MarketAccessService';

  try {
    foInfo_(module, 'Start', 'Upserting market access record.');

    const ledger = foLedger_();

    const sheet = foEnsureSheet_(ledger, 'Canadian Market Access Library', [
      'Ticker',
      'Company',
      'Native Listing',
      'Canadian Common Share',
      'Canadian CDR',
      'Canadian ETF Alternative',
      'Preferred Vehicle',
      'Market Access Type',
      'Last Reviewed',
      'Liquidity Assessment',
      'Tracking Assessment',
      'CIO Notes'
    ]);

    const values = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).toUpperCase() === String(record.ticker).toUpperCase()) {
        targetRow = i + 1;
        break;
      }
    }

    const row = [
      record.ticker || '',
      record.company || '',
      record.nativeListing || '',
      record.canadianCommonShare || '',
      record.canadianCDR || '',
      record.canadianETF || '',
      record.preferredVehicle || '',
      record.marketAccessType || '',
      new Date(),
      record.liquidityAssessment || '',
      record.trackingAssessment || '',
      record.notes || ''
    ];

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    foInfo_(module, 'Complete', 'Market access record upserted: ' + (record.ticker || 'unknown'));

    return {
      status: 'SUCCESS',
      ticker: record.ticker || ''
    };

  } catch (error) {
    foError_(module, 'Failure', error);
    throw error;
  }
}

function foSeedKnownCDRs() {
  foUpsertMarketAccess({
    ticker: 'MU',
    company: 'Micron',
    nativeListing: 'MU US',
    canadianCDR: 'Micron HG CDR',
    preferredVehicle: 'Micron HG CDR',
    marketAccessType: 'CDR',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });

  foUpsertMarketAccess({
    ticker: 'AVGO',
    company: 'Broadcom',
    nativeListing: 'AVGO US',
    canadianCDR: 'Broadcom CDR',
    preferredVehicle: 'Broadcom CDR',
    marketAccessType: 'CDR',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });

  foUpsertMarketAccess({
    ticker: 'QCOM',
    company: 'Qualcomm',
    nativeListing: 'QCOM US',
    canadianCDR: 'Qualcomm HG CDR',
    preferredVehicle: 'Qualcomm HG CDR',
    marketAccessType: 'CDR',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });

  foUpsertMarketAccess({
    ticker: 'CLS',
    company: 'Celestica',
    nativeListing: 'CLS US / TSX context to verify',
    canadianCDR: 'Celestica CDR',
    preferredVehicle: 'Canadian access to be evaluated',
    marketAccessType: 'CDR/Common Share Review',
    liquidityAssessment: 'Pending',
    trackingAssessment: 'Pending',
    notes: 'Seeded CDR candidate.'
  });
}