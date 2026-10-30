// Standardized Loan Provider Type Categorization
// Merges 120+ unique lender variants, abbreviations, and spelling differences
// into clean standardized categories while preserving original API data.

export const STANDARDIZED_PROVIDERS_ORDER = [
  'Bajaj Finance',
  'Aditya Birla Finance',
  'Tata Capital',
  'Cholamandalam Finance',
  'L&T Finance',
  'Godrej Capital',
  'Poonawalla Fincorp',
  'Credit Saison',
  'SMFG India Credit',
  'InCred',
  'Lendingkart',
  'Shriram Finance',
  'Paysense',
  'HDFC Bank',
  'ICICI Bank',
  'IDFC FIRST Bank',
  'Axis Bank',
  'Bank of Baroda',
  'State Bank of India',
  'Bank of India',
  'Canara Bank',
  'IndusInd Bank',
  'Punjab National Bank',
  'Kotak Mahindra Bank',
  'HSBC Bank',
  'Deutsche Bank',
  'YES Bank',
  'Standard Chartered Bank',
  'Bandhan Bank',
  'Central Bank of India',
  'F2 Fintech – Lender Selection',
  'No Provider Available',
  'Multi-Lender',
  'Other / Unmapped',
] as const;

export type StandardizedProviderCategory = (typeof STANDARDIZED_PROVIDERS_ORDER)[number];

export function standardizeProvider(raw: any): StandardizedProviderCategory {
  if (raw === null || raw === undefined) return 'No Provider Available';
  const s = String(raw).trim();
  if (!s) return 'No Provider Available';

  const lower = s.toLowerCase();

  // Blank / empty / null placeholders
  if (
    [
      'none',
      'null',
      'na',
      'n/a',
      '-',
      '--',
      '0',
      '#n/a',
      'undefined',
      'no provider available',
      'not available',
    ].includes(lower)
  ) {
    return 'No Provider Available';
  }

  // Multi-Lender check: values containing multiple distinct lenders separated by commas, '+', '/', or 'and'
  const isMulti = (() => {
    if (lower.includes(',') || lower.includes('+')) return true;
    if (
      /\b(hdfc|icici|axis|sbi|tata|chola|abfl|l&t|godrej|bajaj|indusind|hsbc)\s*[\/]\s*(hdfc|icici|axis|sbi|tata|chola|abfl|l&t|godrej|bajaj|indusind|hsbc)/i.test(
        lower
      )
    ) {
      return true;
    }
    if (
      /\b(chola|tata|godrej|l&t|abfl|bajaj)\s+(and|&)\s+(aditya|chola|tata|godrej|l&t|abfl|bajaj)/i.test(
        lower
      )
    ) {
      return true;
    }
    return false;
  })();
  if (isMulti) return 'Multi-Lender';

  // 31. F2 Fintech – Lender Selection
  if (
    lower.includes('let f2') ||
    lower.includes('f2 fintech') ||
    lower.includes('f2fintech')
  ) {
    return 'F2 Fintech – Lender Selection';
  }

  // 1. Bajaj Finance (Bajaj Finance, Bajaj Finserv, Bajaj Market, Bajaj OD, etc.)
  if (
    /\bbajaj?\b/i.test(lower) ||
    /\bbabaj\b/i.test(lower) ||
    lower.startsWith('baja od') ||
    lower === 'baja'
  ) {
    return 'Bajaj Finance';
  }

  // 2. Aditya Birla Finance (Aditya Birla Finance, ABFL, ABCL, etc.)
  if (
    /\badity?a?\s*b[ir]{2}la\b/i.test(lower) ||
    /\babfl\b/i.test(lower) ||
    /\babcl\b/i.test(lower)
  ) {
    return 'Aditya Birla Finance';
  }

  // 3. Tata Capital (Tata Capital, Tata OD, Tata, etc.)
  if (lower.includes('tata')) {
    return 'Tata Capital';
  }

  // 4. Cholamandalam Finance (Chola, Chola Finance, Cholamandalam, Chola Mandalam, Chola OD, etc.)
  if (lower.includes('chola') || lower.includes('cholamandalam')) {
    return 'Cholamandalam Finance';
  }

  // 5. L&T Finance (L&T Finance, LNT, L&T OD, L&T, etc.)
  if (
    /\bl\s*&\s*t\b/i.test(lower) ||
    /\blnt\b/i.test(lower) ||
    lower.includes('l&t')
  ) {
    return 'L&T Finance';
  }

  // 6. Godrej Capital (Godrej Capital, Godrej OD, Godrej, etc.)
  if (lower.includes('godrej') || lower.includes('goderej')) {
    return 'Godrej Capital';
  }

  // 7. Poonawalla Fincorp (Poonawalla, Poonawala, Poonawalla Fincorp)
  if (lower.includes('poonawala') || lower.includes('poonawalla')) {
    return 'Poonawalla Fincorp';
  }

  // 8. Credit Saison (Credit Saison, Credit Session)
  if (lower.includes('credit saison') || lower.includes('credit session')) {
    return 'Credit Saison';
  }

  // 9. SMFG India Credit (SMFG, SMFG India Credit, Fullerton)
  if (lower.includes('smfg') || lower.includes('fullerton')) {
    return 'SMFG India Credit';
  }

  // 10. InCred (InCred, Incred)
  if (lower.includes('incred')) {
    return 'InCred';
  }

  // 11. Lendingkart (Lendingkart, Lending Cart)
  if (/\blending\s*k?c?art\b/i.test(lower)) {
    return 'Lendingkart';
  }

  // 12. Shriram Finance (Shriram, Shriram Finance)
  if (lower.includes('shriram')) {
    return 'Shriram Finance';
  }

  // 13. Paysense (Paysense)
  if (lower.includes('paysense')) {
    return 'Paysense';
  }

  // 14. HDFC Bank (HDFC, HDFC Bank)
  if (lower.includes('hdfc')) {
    return 'HDFC Bank';
  }

  // 15. ICICI Bank (ICICI, ICICI Bank)
  if (lower.includes('icici')) {
    return 'ICICI Bank';
  }

  // 16. IDFC FIRST Bank (IDFC, IDFC First Bank, IDFC First)
  if (lower.includes('idfc')) {
    return 'IDFC FIRST Bank';
  }

  // 17. Axis Bank (Axis, Axis Bank)
  if (lower.includes('axis')) {
    return 'Axis Bank';
  }

  // 18. Bank of Baroda (Bank of Baroda, BOB)
  if (lower.includes('bank of baroda') || lower === 'bob') {
    return 'Bank of Baroda';
  }

  // 19. State Bank of India (SBI, State Bank of India)
  if (lower.includes('state bank of india') || lower === 'sbi') {
    return 'State Bank of India';
  }

  // 20. Bank of India (BOI, Bank of India)
  if (lower.includes('bank of india') || lower === 'boi') {
    return 'Bank of India';
  }

  // 21. Canara Bank (Canara Bank, Canara)
  if (lower.includes('canara')) {
    return 'Canara Bank';
  }

  // 22. IndusInd Bank (IndusInd, IndusInd Bank)
  if (lower.includes('indusind')) {
    return 'IndusInd Bank';
  }

  // 23. Punjab National Bank (PNB, Punjab National Bank, LAP PNB)
  if (lower.includes('punjab national bank') || lower.includes('pnb')) {
    return 'Punjab National Bank';
  }

  // 24. Kotak Mahindra Bank (Kotak, Kotak Mahindra Bank)
  if (lower.includes('kotak')) {
    return 'Kotak Mahindra Bank';
  }

  // 25. HSBC Bank (HSBC, HSBC Bank)
  if (lower.includes('hsbc')) {
    return 'HSBC Bank';
  }

  // 26. Deutsche Bank (Deutsche Bank)
  if (lower.includes('deutsche')) {
    return 'Deutsche Bank';
  }

  // 27. YES Bank (YES Bank)
  if (lower.includes('yes bank')) {
    return 'YES Bank';
  }

  // 28. Standard Chartered Bank (Standard Chartered, Standard Chartered Bank)
  if (lower.includes('standard chartered')) {
    return 'Standard Chartered Bank';
  }

  // 29. Bandhan Bank (Bandhan Bank)
  if (lower.includes('bandhan')) {
    return 'Bandhan Bank';
  }

  // 30. Central Bank of India (Central Bank of India, CBI)
  if (lower.includes('central bank of india') || lower === 'cbi') {
    return 'Central Bank of India';
  }

  // 32. No Provider Available (fallback for explicit mention)
  if (lower.includes('no provider')) {
    return 'No Provider Available';
  }

  // 34. Other / Unmapped
  return 'Other / Unmapped';
}
