/**
 * Sync renewals from Renewal Status tab back to TRACKER
 * When Final Status = "Renew", extend contract by 12 months
 */
function syncRenewals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = ss.getSheetByName(CONFIG.TRACKER_SHEET);
  const renewalSheet = ss.getSheetByName(CONFIG.RENEWAL_SHEET);
  
  if (!trackerSheet || !renewalSheet) {
    Logger.log('ERROR: Required sheets not found');
    return;
  }
  
  const renewalData = renewalSheet.getDataRange().getValues();
  let renewalsProcessed = 0;
  
  // Skip header row
  for (let i = 1; i < renewalData.length; i++) {
    const companyName = renewalData[i][CONFIG.RENEWAL_COLS.COMPANIES - 1];
    const finalStatus = renewalData[i][CONFIG.RENEWAL_COLS.FINAL_STATUS - 1];
    const newRenewalDate = renewalData[i][CONFIG.RENEWAL_COLS.NEW_RENEWAL_DATE - 1];
    
    // Only process if Final Status is "Renew"
    if (finalStatus === 'Renew') {
      
      // Find company in TRACKER
      const trackerData = trackerSheet.getDataRange().getValues();
      
      for (let j = 1; j < trackerData.length; j++) {
        const trackerCompanyName = trackerData[j][CONFIG.TRACKER_COLS.COMPANY_NAME - 1];
        
        if (trackerCompanyName === companyName) {
          
          // Get current contract end
          const currentEnd = trackerData[j][CONFIG.TRACKER_COLS.CONTRACT_END - 1];
          const currentEndDate = new Date(currentEnd);
          
          // Calculate new contract end (+12 months from current end)
          const newEndDate = new Date(currentEndDate);
          newEndDate.setFullYear(newEndDate.getFullYear() + 1);
          
          // Update contract end date
          trackerSheet.getRange(j + 1, CONFIG.TRACKER_COLS.CONTRACT_END).setValue(newEndDate);
          
          // Populate another 12 months of "paid" status
          const newStartMonth = new Date(currentEndDate);
          newStartMonth.setMonth(currentEndDate.getMonth() + 1); // Start from month after current end
          
          populate12MonthsFromDate(trackerSheet, j + 1, newStartMonth);
          
          Logger.log(`Extended contract for ${companyName} until ${newEndDate}`);
          
          // Update Renewal Status to mark as processed
          renewalSheet.getRange(i + 1, CONFIG.RENEWAL_COLS.CUSTOMER_STATUS).setValue('Renewed');
          
          renewalsProcessed++;
          break;
        }
      }
    }
  }
  
  Logger.log(`Renewals processed: ${renewalsProcessed}`);
  
  if (renewalsProcessed > 0) {
    SpreadsheetApp.getUi().alert(`✅ Processed ${renewalsProcessed} renewals`);
  }
}

/**
 * Populate 12 months from a specific start date
 */
function populate12MonthsFromDate(sheet, rowNumber, startDate) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  for (let monthOffset = 0; monthOffset < 12; monthOffset++) {
    const targetDate = new Date(startDate);
    targetDate.setMonth(startDate.getMonth() + monthOffset);
    
    const targetMonthYear = Utilities.formatDate(targetDate, Session.getScriptTimeZone(), 'MMM-yyyy');
    
    const colIndex = headers.findIndex(header => {
      if (typeof header === 'string') {
        return header.trim() === targetMonthYear;
      }
      return false;
    });
    
    if (colIndex !== -1) {
      const cell = sheet.getRange(rowNumber, colIndex + 1);
      cell.setValue('paid');
      
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['paid', 'renew', 'terminate', 'not proceed'], true)
        .build();
      cell.setDataValidation(rule);
    }
  }
}