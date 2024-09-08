/** DCA Calculator (2k24 edition) */

import * as cheerio from 'cheerio';
import ExcelJS from 'exceljs';
import * as fs from 'fs';
import * as path from 'path';

import {
  FinanceApiResponse,
  FinancialData,
  isFinancialDataArray,
  ResponseBody,
} from './types';

/* eslint-disable no-console */

const SELECTED_INDUSTRY =
  'S&P 500 Companies - Sept. 8, 2024' satisfies keyof typeof TICKERS_BY_INDUSTRY;
const TICKERS_BY_INDUSTRY = {
  'S&P 500 Companies - Sept. 8, 2024': [
    'MMM',
    'AOS',
    'ABT',
    'ABBV',
    'ACN',
    'ADBE',
    'AMD',
    'AES',
    'AFL',
    'A',
    'APD',
    'ABNB',
    'AKAM',
    'ALB',
    'ARE',
    'ALGN',
    'ALLE',
    'LNT',
    'ALL',
    'GOOGL',
    'GOOG',
    'MO',
    'AMZN',
    'AMCR',
    'AEE',
    'AAL',
    'AEP',
    'AXP',
    'AIG',
    'AMT',
    'AWK',
    'AMP',
    'AME',
    'AMGN',
    'APH',
    'ADI',
    'ANSS',
    'AON',
    'APA',
    'AAPL',
    'AMAT',
    'APTV',
    'ACGL',
    'ADM',
    'ANET',
    'AJG',
    'AIZ',
    'T',
    'ATO',
    'ADSK',
    'ADP',
    'AZO',
    'AVB',
    'AVY',
    'AXON',
    'BKR',
    'BALL',
    'BAC',
    'BK',
    'BBWI',
    'BAX',
    'BDX',
    'BRK.B',
    'BBY',
    'BIO',
    'TECH',
    'BIIB',
    'BLK',
    'BX',
    'BA',
    'BKNG',
    'BWA',
    'BSX',
    'BMY',
    'AVGO',
    'BR',
    'BRO',
    'BF.B',
    'BLDR',
    'BG',
    'BXP',
    'CHRW',
    'CDNS',
    'CZR',
    'CPT',
    'CPB',
    'COF',
    'CAH',
    'KMX',
    'CCL',
    'CARR',
    'CTLT',
    'CAT',
    'CBOE',
    'CBRE',
    'CDW',
    'CE',
    'COR',
    'CNC',
    'CNP',
    'CF',
    'CRL',
    'SCHW',
    'CHTR',
    'CVX',
    'CMG',
    'CB',
    'CHD',
    'CI',
    'CINF',
    'CTAS',
    'CSCO',
    'C',
    'CFG',
    'CLX',
    'CME',
    'CMS',
    'KO',
    'CTSH',
    'CL',
    'CMCSA',
    'CAG',
    'COP',
    'ED',
    'STZ',
    'CEG',
    'COO',
    'CPRT',
    'GLW',
    'CPAY',
    'CTVA',
    'CSGP',
    'COST',
    'CTRA',
    'CRWD',
    'CCI',
    'CSX',
    'CMI',
    'CVS',
    'DHR',
    'DRI',
    'DVA',
    'DAY',
    'DECK',
    'DE',
    'DAL',
    'DVN',
    'DXCM',
    'FANG',
    'DLR',
    'DFS',
    'DG',
    'DLTR',
    'D',
    'DPZ',
    'DOV',
    'DOW',
    'DHI',
    'DTE',
    'DUK',
    'DD',
    'EMN',
    'ETN',
    'EBAY',
    'ECL',
    'EIX',
    'EW',
    'EA',
    'ELV',
    'EMR',
    'ENPH',
    'ETR',
    'EOG',
    'EPAM',
    'EQT',
    'EFX',
    'EQIX',
    'EQR',
    'ESS',
    'EL',
    'ETSY',
    'EG',
    'EVRG',
    'ES',
    'EXC',
    'EXPE',
    'EXPD',
    'EXR',
    'XOM',
    'FFIV',
    'FDS',
    'FICO',
    'FAST',
    'FRT',
    'FDX',
    'FIS',
    'FITB',
    'FSLR',
    'FE',
    'FI',
    'FMC',
    'F',
    'FTNT',
    'FTV',
    'FOXA',
    'FOX',
    'BEN',
    'FCX',
    'GRMN',
    'IT',
    'GE',
    'GEHC',
    'GEV',
    'GEN',
    'GNRC',
    'GD',
    'GIS',
    'GM',
    'GPC',
    'GILD',
    'GPN',
    'GL',
    'GDDY',
    'GS',
    'HAL',
    'HIG',
    'HAS',
    'HCA',
    'DOC',
    'HSIC',
    'HSY',
    'HES',
    'HPE',
    'HLT',
    'HOLX',
    'HD',
    'HON',
    'HRL',
    'HST',
    'HWM',
    'HPQ',
    'HUBB',
    'HUM',
    'HBAN',
    'HII',
    'IBM',
    'IEX',
    'IDXX',
    'ITW',
    'INCY',
    'IR',
    'PODD',
    'INTC',
    'ICE',
    'IFF',
    'IP',
    'IPG',
    'INTU',
    'ISRG',
    'IVZ',
    'INVH',
    'IQV',
    'IRM',
    'JBHT',
    'JBL',
    'JKHY',
    'J',
    'JNJ',
    'JCI',
    'JPM',
    'JNPR',
    'K',
    'KVUE',
    'KDP',
    'KEY',
    'KEYS',
    'KMB',
    'KIM',
    'KMI',
    'KKR',
    'KLAC',
    'KHC',
    'KR',
    'LHX',
    'LH',
    'LRCX',
    'LW',
    'LVS',
    'LDOS',
    'LEN',
    'LLY',
    'LIN',
    'LYV',
    'LKQ',
    'LMT',
    'L',
    'LOW',
    'LULU',
    'LYB',
    'MTB',
    'MRO',
    'MPC',
    'MKTX',
    'MAR',
    'MMC',
    'MLM',
    'MAS',
    'MA',
    'MTCH',
    'MKC',
    'MCD',
    'MCK',
    'MDT',
    'MRK',
    'META',
    'MET',
    'MTD',
    'MGM',
    'MCHP',
    'MU',
    'MSFT',
    'MAA',
    'MRNA',
    'MHK',
    'MOH',
    'TAP',
    'MDLZ',
    'MPWR',
    'MNST',
    'MCO',
    'MS',
    'MOS',
    'MSI',
    'MSCI',
    'NDAQ',
    'NTAP',
    'NFLX',
    'NEM',
    'NWSA',
    'NWS',
    'NEE',
    'NKE',
    'NI',
    'NDSN',
    'NSC',
    'NTRS',
    'NOC',
    'NCLH',
    'NRG',
    'NUE',
    'NVDA',
    'NVR',
    'NXPI',
    'ORLY',
    'OXY',
    'ODFL',
    'OMC',
    'ON',
    'OKE',
    'ORCL',
    'OTIS',
    'PCAR',
    'PKG',
    'PANW',
    'PARA',
    'PH',
    'PAYX',
    'PAYC',
    'PYPL',
    'PNR',
    'PEP',
    'PFE',
    'PCG',
    'PM',
    'PSX',
    'PNW',
    'PNC',
    'POOL',
    'PPG',
    'PPL',
    'PFG',
    'PG',
    'PGR',
    'PLD',
    'PRU',
    'PEG',
    'PTC',
    'PSA',
    'PHM',
    'QRVO',
    'PWR',
    'QCOM',
    'DGX',
    'RL',
    'RJF',
    'RTX',
    'O',
    'REG',
    'REGN',
    'RF',
    'RSG',
    'RMD',
    'RVTY',
    'ROK',
    'ROL',
    'ROP',
    'ROST',
    'RCL',
    'SPGI',
    'CRM',
    'SBAC',
    'SLB',
    'STX',
    'SRE',
    'NOW',
    'SHW',
    'SPG',
    'SWKS',
    'SJM',
    'SW',
    'SNA',
    'SOLV',
    'SO',
    'LUV',
    'SWK',
    'SBUX',
    'STT',
    'STLD',
    'STE',
    'SYK',
    'SMCI',
    'SYF',
    'SNPS',
    'SYY',
    'TMUS',
    'TROW',
    'TTWO',
    'TPR',
    'TRGP',
    'TGT',
    'TEL',
    'TDY',
    'TFX',
    'TER',
    'TSLA',
    'TXN',
    'TXT',
    'TMO',
    'TJX',
    'TSCO',
    'TT',
    'TDG',
    'TRV',
    'TRMB',
    'TFC',
    'TYL',
    'TSN',
    'USB',
    'UBER',
    'UDR',
    'ULTA',
    'UNP',
    'UAL',
    'UPS',
    'URI',
    'UNH',
    'UHS',
    'VLO',
    'VTR',
    'VLTO',
    'VRSN',
    'VRSK',
    'VZ',
    'VRTX',
    'VTRS',
    'VICI',
    'V',
    'VST',
    'VMC',
    'WRB',
    'GWW',
    'WAB',
    'WBA',
    'WMT',
    'DIS',
    'WBD',
    'WM',
    'WAT',
    'WEC',
    'WFC',
    'WELL',
    'WST',
    'WDC',
    'WY',
    'WMB',
    'WTW',
    'WYNN',
    'XEL',
    'XYL',
    'YUM',
    'ZBRA',
    'ZBH',
    'ZTS',
  ],
  'Fundamentally Strong Stocks For Long Term': [
    'DIVISLAB',
    'CGPOWER.NS',
    'HDFCAMC.NS',
    'GICRE.NS',
    'BDL',
    'GSK',
    'NAM-INDIA.NS',
    'LLOYDSME.BO',
    '%5EBSESN',
    'SKFINDIA.NS',
  ],
  'Best Long-Term Stocks of September 2024': [
    'T',
    'CVS',
    'F',
    'ALL',
    'KHC',
    'KR',
    'DFS',
    'HIG',
    'FE',
  ],
  gopro: ['GPRO'],
  'Largest Companies In The Utilities Sector': [
    'NEE',
    'SO',
    'DUK',
    'GEV',
    'AEP',
    'CEG',
    'PCG',
    'SRE',
    'D',
    'PEG',
    'EXC',
    'ED',
    'XEL',
    'EIX',
    'WEC',
    'AWK',
    'ETR',
    'DTE',
    'FE',
    'VST',
    'ES',
    'PPL',
    'AEE',
    'CMS',
    'ATO',
    'CNP',
    'NRG',
    'LNT',
    'NI',
    'EVRG',
    'AGR',
    'AES',
    'WTRG',
    'PNW',
    'OGE',
    'TLN',
    'IDA',
    'CWEN-A',
    'UGI',
    'POR',
    'SWX',
    'NJR',
    'ORA',
    'BKH',
    'OGS',
    'SR',
    'TXNM',
    'ALE',
    'NWE',
    'FLNC',
    'OTTR',
    'CWT',
    'MGEE',
    'AWR',
    'AVA',
    'CPK',
    'AY',
    'NFE',
    'NEP',
    'RNW',
    'SJW',
    'NWN',
    'CTRI',
    'KEN',
    'HE',
    'SPH',
    'MSEX',
    'UTL',
    'ARIS',
    'OKLO',
    'YORW',
    'GNE',
    'CWCO',
    'ARTNA',
    'GWRS',
    'PCYO',
    'RGCO',
    'ELLO',
    'MCPB',
  ],
  'Largest Companies In The Technology Sector (Top 100)': [
    'AAPL',
    'MSFT',
    'NVDA',
    'AVGO',
    'ORCL',
    'ADBE',
    'CRM',
    'AMD',
    'ACN',
    'CSCO',
    'IBM',
    'TXN',
    'QCOM',
    'INTU',
    'NOW',
    'UBER',
    'AMAT',
    'ADP',
    'PANW',
    'ADI',
    'ANET',
    'FI',
    'MU',
    'LRCX',
    'KLAC',
    'INTC',
    'DELL',
    'APH',
    'MSI',
    'SNPS',
    'PLTR',
    'CDNS',
    'WDAY',
    'MRVL',
    'CRWD',
    'ROP',
    'NXPI',
    'FTNT',
    'ADSK',
    'TTD',
    'PAYX',
    'FIS',
    'TEL',
    'FICO',
    'TEAM',
    'MCHP',
    'MPWR',
    'SQ',
    'CTSH',
    'IT',
    'SNOW',
    'DDOG',
    'GLW',
    'GRMN',
    'HPQ',
    'ON',
    'CDW',
    'APP',
    'ANSS',
    'NET',
    'HUBS',
    'KEYS',
    'TYL',
    'FTV',
    'IOT',
    'BR',
    'ZS',
    'NTAP',
    'SMCI',
    'HPE',
    'GFS',
    'FSLR',
    'MSTR',
    'GDDY',
    'CPAY',
    'LDOS',
    'CHKP',
    'WDC',
    'MDB',
    'ZM',
    'STX',
    'TER',
    'PTC',
    'TDY',
    'VRSN',
    'SSNC',
    'ENTG',
    'ZBRA',
    'SWKS',
    'GEN',
    'NTNX',
    'MANH',
    'AKAM',
    'DT',
    'PSTG',
    'ENPH',
    'BSY',
    'AZPN',
    'GWRE',
    'TRMB',
  ],
  'Beverages - Non-Alcoholic': [
    'AKO-A',
    'AKO-B',
    'BRFH',
    'CCEP',
    'CELH',
    'COCO',
    'COKE',
    'FIZZ',
    'KDP',
    'KO',
    'KOF',
    'MNST',
    'PEP',
    'PRMW',
    'SHOT',
    'STKL',
    'ZVIA',
  ],
  'S&P 500 Health Care': [
    'LLY',
    'UNH',
    'JNJ',
    'ABBV',
    'MRK',
    'TMO',
    'ABT',
    'DHR',
    'AMGN',
    'ISRG',
    'PFE',
    'SYK',
    'ELV',
    'REGN',
    'VRTX',
    'BSX',
    'MDT',
    'HCA',
    'CI',
    'BMY',
    'GILD',
    'ZTS',
    'CVS',
    'BDX',
    'MCK',
    'COR',
    'IQV',
    'HUM',
    'EW',
    'A',
    'IDXX',
    'GEHC',
    'CNC',
    'RMD',
    'MTD',
    'BIIB',
    'MRNA',
    'DXCM',
    'CAH',
    'STE',
    'WST',
    'ZBH',
    'COO',
    'BAX',
    'WAT',
    'MOH',
    'HOLX',
    'LH',
    'DGX',
    'ALGN',
    'PODD',
    'RVTY',
    'UHS',
    'VTRS',
    'DVA',
    'INCY',
    'TFX',
    'SOLV',
    'TECH',
    'CTLT',
    'CRL',
    'BIO',
    'HSIC',
  ],
  'S&P 500 Consumer Staples': [
    'WMT',
    'PG',
    'COST',
    'KO',
    'PEP',
    'PM',
    'MDLZ',
    'MO',
    'CL',
    'TGT',
    'KDP',
    'KMB',
    'MNST',
    'STZ',
    'KVUE',
    'KHC',
    'GIS',
    'HSY',
    'SYY',
    'KR',
    'EL',
    'ADM',
    'K',
    'CHD',
    'TSN',
    'MKC',
    // 'BF.B',
    'CLX',
    'DG',
    'HRL',
    'CAG',
    'CPB',
    'DLTR',
    'BG',
    'SJM',
    'TAP',
    'LW',
    'WBA',
  ],
};

const tickers = TICKERS_BY_INDUSTRY[SELECTED_INDUSTRY];

const excelMappings = {
  balanceSheet: {
    inventory: { C: 'C3', D: 'D3', E: 'E3', F: 'F3' },
    ppe: { C: 'C4', D: 'D4', E: 'E4', F: 'F4' },
    goodwill: { C: 'C5', D: 'D5', E: 'E5', F: 'F5' },
    'total assets': { C: 'C6', D: 'D6', E: 'E6', F: 'F6' },
    'current liabilities': { C: 'C7', D: 'D7', E: 'E7', F: 'F7' },
    'long term debt': { C: 'C8', D: 'D8', E: 'E8', F: 'F8' },
    'total liabilities': { C: 'C9', D: 'D9', E: 'E9', F: 'F9' },
    'treasury stock': { C: 'C10', D: 'D10', E: 'E10', F: 'F10' },
    'preferred stock': { C: 'C11', D: 'D11', E: 'E11', F: 'F11' },
    'retained earnings': { C: 'C12', D: 'D12', E: 'E12', F: 'F12' },
    'total equity': { C: 'C13', D: 'D13', E: 'E13', F: 'F13' },
  },
  incomeStatement: {
    'r&d': { C: 'C15', D: 'D15', E: 'E15', F: 'F15' },
  },
  cashFlow: {
    'net operating cash flow': { C: 'C17', D: 'D17', E: 'E17', F: 'F17' },
  },
  years: { C: 'C2', D: 'D2', E: 'E2', F: 'F2' }, // year headers
};

const config = {
  fetchIncomeStatement: true,
  fetchBalanceSheet: true,
  fetchCashFlow: true,
};

const MAX_RETRIES = 10; // Maximum number of retries
const BATCH_SIZE = 1; // Number of requests per batch
const RETRY_DELAY_MS = 10000; // Base delay between retries
const BATCH_DELAY_MS = 1250; // Delay between batches of requests

// Timer variables to track elapsed time
let elapsedTime = 0; // in milliseconds
let isSleeping = false; // Flag to indicate if we're in the sleep period
const TIMER_INTERVAL = 1000; // 1 second interval for tracking time
const FORCE_SLEEP_INTERVAL_MS = 30000; // Force sleep every 30 seconds
const FORCE_SLEEP_DURATION_MS = 30000; // Sleep duration is 30 seconds

const incomeStatementURL =
  'https://finance.yahoo.com/quote/<TICKER>/financials/';
const balanceSheetURL =
  'https://finance.yahoo.com/quote/<TICKER>/balance-sheet/';
const cashFlowURL = 'https://finance.yahoo.com/quote/<TICKER>/cash-flow/';

const userAgent =
  'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36';

const reportMappings = {
  balanceSheet: {
    inventory: 'annualInventory',
    ppe: 'annualNetPPE',
    goodwill: 'annualGoodwill',
    'total assets': 'annualTotalAssets',
    'current liabilities': 'annualCurrentLiabilities',
    'long term debt': 'annualTotalNonCurrentLiabilitiesNetMinorityInterest',
    'total liabilities': 'annualTotalLiabilitiesNetMinorityInterest',
    'treasury stock': 'annualTreasuryStock',
    'preferred stock': 'annualPreferredStock',
    'retained earnings': 'annualRetainedEarnings',
    'total equity': 'annualTotalEquityGrossMinorityInterest',
  },
  incomeStatement: {
    'r&d': 'annualResearchAndDevelopment',
  },
  cashFlow: {
    'net operating cash flow': 'annualOperatingCashFlow',
  },
};

// Timer to track and reset every 30 seconds
const startTimer = () => {
  const intervalId = setInterval(async () => {
    elapsedTime += TIMER_INTERVAL;

    if (elapsedTime >= FORCE_SLEEP_INTERVAL_MS) {
      isSleeping = true; // Set sleeping flag to true
      console.log(
        `Forcing sleep for ${FORCE_SLEEP_DURATION_MS / 1000} seconds...`,
      );
      await sleep(FORCE_SLEEP_DURATION_MS); // Sleep for the designated time
      isSleeping = false; // Reset the sleeping flag after sleep
      elapsedTime = 0; // Reset the timer after sleep
    }
  }, TIMER_INTERVAL);

  return intervalId; // Return the interval ID for potential clearing later
};

// Stop the timer when necessary
const stopTimer = (intervalId: NodeJS.Timeout) => {
  clearInterval(intervalId);
};

const sleep = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

const extractScriptWithDataURL = (html: string) => {
  const $ = cheerio.load(html);

  // Filter to find the script tag with the required data-url
  const scriptTag = $('script').filter((_, el) => {
    const dataUrl = $(el).attr('data-url');
    return !!(
      dataUrl &&
      dataUrl.startsWith(
        'https://query1.finance.yahoo.com/ws/fundamentals-timeseries',
      )
    );
  });

  if (scriptTag.length) {
    return scriptTag.html(); // Return the content inside the <script> tag
  }
  return null; // If no matching script tag is found
};

const parseJSONFromScript = (scriptContent: string): ResponseBody | null => {
  try {
    // Assuming the script content contains valid JSON directly
    const jsonData: FinanceApiResponse = JSON.parse(scriptContent);

    // Assert the type for the body
    if (typeof jsonData.body === 'string') {
      return JSON.parse(jsonData.body) as ResponseBody;
    } else {
      return jsonData.body ?? null;
    }
  } catch (error) {
    console.error('Error parsing JSON from script tag:', error);
    return null;
  }
};

// Define the structure of the final report
interface Report {
  [date: string]: {
    [field: string]: number;
  };
}

interface FinalReport {
  balanceSheet: Report;
  incomeStatement: Report;
  cashFlow: Report;
}

// Modify finalReport to hold results for each ticker
const finalReport: Record<string, FinalReport> = {};

import { exec } from 'child_process';
import { promisify } from 'util';

const execPromise = promisify(exec);

const runCurlCommand = async (
  url: string,
  ticker: string,
  reportType: 'balanceSheet' | 'incomeStatement' | 'cashFlow',
  attempt = 1,
  retryDelay = RETRY_DELAY_MS,
  currentIndex: number,
  totalTickers: number,
): Promise<void> => {
  const command = `curl -s --max-time 10 '${url}' -H 'user-agent: ${userAgent}'`;
  console.log(`Running: ${command}`);

  console.log(
    `Processing ticker ${currentIndex + 1} of ${totalTickers}: ${ticker}`,
  );

  try {
    const { stdout, stderr } = await execPromise(command, {
      maxBuffer: 1024 * 1000 * 20,
    });

    if (stderr) {
      console.error(`stderr: ${stderr}`);
      throw new Error(
        `[stderr] Error fetching ${reportType} for ${ticker}: ${stderr}`,
      );
    }

    // Extract content inside the matching <script> tag
    const extractedContent = extractScriptWithDataURL(stdout);

    if (extractedContent) {
      console.log(
        `\n--- Extracted Script Content for ${reportType} (${ticker}) ---\n`,
      );
      const parsedData = parseJSONFromScript(extractedContent);

      if (parsedData) {
        console.log(
          `\n--- Parsed JSON Data for ${reportType} (${ticker}) ---\n`,
        );
        handleParsedData(parsedData, ticker, reportType);
      } else {
        throw new Error(`Failed to parse JSON for ${reportType} (${ticker})`);
      }
    } else {
      throw new Error(
        `No matching <script> tag found for ${reportType} (${ticker})`,
      );
    }
  } catch (err) {
    if (attempt <= MAX_RETRIES) {
      console.warn(
        `Attempt ${attempt} failed for ${ticker} (${reportType}). Retrying in ${retryDelay} ms...`,
      );
      await sleep(retryDelay); // Wait before retrying
      return runCurlCommand(
        url,
        ticker,
        reportType,
        attempt + 1,
        retryDelay,
        currentIndex,
        totalTickers,
      ); // Exponential backoff
    } else {
      console.error(
        `Failed after ${MAX_RETRIES + 1} attempts for ${ticker} (${reportType}). Error:`,
        err,
      );
    }
  }
};

const processInBatches = async (tickers: string[]) => {
  const totalTickers = tickers.length;

  // Start the timer
  const intervalId = startTimer();

  for (let i = 0; i < totalTickers; i += BATCH_SIZE) {
    // Wait for sleep to finish if we're in the sleeping period
    while (isSleeping) {
      console.log('Sleeping, waiting for the sleep period to finish...');
      await sleep(1000); // Check every second whether sleep period is over
    }

    const batch = tickers.slice(i, i + BATCH_SIZE);

    const promises: Promise<void>[] = [];
    for (const [currentIndex, ticker] of batch.entries()) {
      const globalIndex = i + currentIndex; // Current global index

      if (config.fetchIncomeStatement) {
        const incomeStatementURLFinal = incomeStatementURL.replace(
          '<TICKER>',
          ticker,
        );
        promises.push(
          runCurlCommand(
            incomeStatementURLFinal,
            ticker,
            'incomeStatement',
            1,
            RETRY_DELAY_MS,
            globalIndex,
            totalTickers,
          ).catch((err) => {
            console.error(
              `Error fetching Income Statement for ${ticker}:`,
              err,
            );
          }),
        );
      }

      if (config.fetchBalanceSheet) {
        const balanceSheetURLFinal = balanceSheetURL.replace(
          '<TICKER>',
          ticker,
        );
        promises.push(
          runCurlCommand(
            balanceSheetURLFinal,
            ticker,
            'balanceSheet',
            1,
            RETRY_DELAY_MS,
            globalIndex,
            totalTickers,
          ).catch((err) => {
            console.error(`Error fetching Balance Sheet for ${ticker}:`, err);
          }),
        );
      }

      if (config.fetchCashFlow) {
        const cashFlowURLFinal = cashFlowURL.replace('<TICKER>', ticker);
        promises.push(
          runCurlCommand(
            cashFlowURLFinal,
            ticker,
            'cashFlow',
            1,
            RETRY_DELAY_MS,
            globalIndex,
            totalTickers,
          ).catch((err) => {
            console.error(`Error fetching Cash Flow for ${ticker}:`, err);
          }),
        );
      }
    }

    // Wait for all requests in the batch to finish
    await Promise.all(promises);

    // Wait before processing the next batch
    console.log(`Waiting ${BATCH_DELAY_MS} ms before the next batch...`);
    await sleep(BATCH_DELAY_MS);
  }

  // Stop the timer once processing is complete
  stopTimer(intervalId);
};

/**
 * Function to format the final report by reordering the date keys
 * in reverse chronological order and limiting to the most recent 4 years.
 */
const formatFinalReport = (
  finalReport: Record<string, FinalReport>,
): Record<string, FinalReport> => {
  // Helper function to sort keys by date and limit to 4
  const reorderAndLimitDates = (data: Report): Report => {
    const sortedKeys = Object.keys(data)
      .sort((a, b) => new Date(b).getTime() - new Date(a).getTime()) // Sort by date, most recent first
      .slice(0, 4); // Only take the most recent 4 keys

    const newData: Report = {};
    for (const key of sortedKeys) {
      newData[key] = data[key];
    }
    return newData;
  };

  // Create a new final report with reordered and limited dates
  const formattedReport: Record<string, FinalReport> = {};

  // Iterate through each ticker in the finalReport
  for (const ticker in finalReport) {
    formattedReport[ticker] = {
      balanceSheet: reorderAndLimitDates(finalReport[ticker].balanceSheet),
      incomeStatement: reorderAndLimitDates(
        finalReport[ticker].incomeStatement,
      ),
      cashFlow: reorderAndLimitDates(finalReport[ticker].cashFlow),
    };
  }

  return formattedReport;
};

// Handler function for parsed data
const handleParsedData = (
  data: ResponseBody,
  ticker: string,
  reportType: 'balanceSheet' | 'incomeStatement' | 'cashFlow',
) => {
  console.log(`Handling parsed data for ${reportType} (${ticker})`);

  const mappings = reportMappings[reportType as keyof typeof reportMappings];

  if (!mappings) {
    console.log(`No mappings found for ${reportType}`);
    return;
  }

  // Initialize the ticker key if not already initialized
  if (!finalReport[ticker]) {
    finalReport[ticker] = {
      balanceSheet: {},
      incomeStatement: {},
      cashFlow: {},
    };
  }

  Object.keys(mappings).forEach((field) => {
    const key = mappings[field as keyof typeof mappings];
    data?.timeseries?.result?.forEach((result) => {
      const fieldData = result[key];
      if (isFinancialDataArray(fieldData as FinancialData[])) {
        fieldData?.forEach((entry) => {
          if (entry) {
            const asOfDate = entry?.asOfDate;
            const reportedValue = entry?.reportedValue?.raw;

            if (asOfDate && reportedValue !== undefined) {
              // Store the data in the finalReport object under the appropriate ticker
              if (!finalReport[ticker][reportType]) {
                finalReport[ticker][reportType] = {};
              }
              if (!finalReport[ticker][reportType][asOfDate]) {
                finalReport[ticker][reportType][asOfDate] = {};
              }
              finalReport[ticker][reportType][asOfDate][field] = reportedValue;
            }
          }
        });
      }
    });
  });
};

// Function to apply mappings to the Excel sheet
const applyMappingsToSheet = (
  worksheet: ExcelJS.Worksheet,
  reportData: FinalReport,
  years: string[],
) => {
  // Apply years in header
  worksheet.getCell(excelMappings.years.C).value = years[0];
  worksheet.getCell(excelMappings.years.D).value = years[1];
  worksheet.getCell(excelMappings.years.E).value = years[2];
  worksheet.getCell(excelMappings.years.F).value = years[3];

  // Balance sheet
  for (const [field, cells] of Object.entries(excelMappings.balanceSheet)) {
    if (reportData.balanceSheet) {
      worksheet.getCell(cells.C).value =
        reportData.balanceSheet[years[0]]?.[field] || 0;
      worksheet.getCell(cells.D).value =
        reportData.balanceSheet[years[1]]?.[field] || 0;
      worksheet.getCell(cells.E).value =
        reportData.balanceSheet[years[2]]?.[field] || 0;
      worksheet.getCell(cells.F).value =
        reportData.balanceSheet[years[3]]?.[field] || 0;
    }
  }

  // Income statement
  for (const [field, cells] of Object.entries(excelMappings.incomeStatement)) {
    if (reportData.incomeStatement) {
      worksheet.getCell(cells.C).value =
        reportData.incomeStatement[years[0]]?.[field] || 0;
      worksheet.getCell(cells.D).value =
        reportData.incomeStatement[years[1]]?.[field] || 0;
      worksheet.getCell(cells.E).value =
        reportData.incomeStatement[years[2]]?.[field] || 0;
      worksheet.getCell(cells.F).value =
        reportData.incomeStatement[years[3]]?.[field] || 0;
    }
  }

  // Cash flow
  for (const [field, cells] of Object.entries(excelMappings.cashFlow)) {
    if (reportData.cashFlow) {
      worksheet.getCell(cells.C).value =
        reportData.cashFlow[years[0]]?.[field] || 0;
      worksheet.getCell(cells.D).value =
        reportData.cashFlow[years[1]]?.[field] || 0;
      worksheet.getCell(cells.E).value =
        reportData.cashFlow[years[2]]?.[field] || 0;
      worksheet.getCell(cells.F).value =
        reportData.cashFlow[years[3]]?.[field] || 0;
    }
  }
};

// Function to auto-resize the column widths based on content with wider columns
const autoResizeColumns = (worksheet: ExcelJS.Worksheet) => {
  worksheet.columns.forEach((column) => {
    let maxLength = 14; // Set a slightly wider default minimum width
    if (column)
      /** @ts-expect-error: column is not undefined */
      column.eachCell({ includeEmpty: true }, (cell) => {
        const columnLength = cell.value ? cell.value.toString().length : 0;
        if (columnLength > maxLength) {
          maxLength = columnLength;
        }
      });
    column.width = maxLength + 5; // Increase padding to make columns wider
  });
};

// Function to update the existing "Summary" sheet and reference score percent from each stock sheet
const updateSummaryPage = async (
  workbook: ExcelJS.Workbook,
  tickers: string[],
) => {
  // Get the existing Summary sheet
  const summarySheet = workbook.getWorksheet('Summary');

  if (!summarySheet) {
    console.error('Summary sheet not found.');
    return;
  }

  // Iterate over each ticker and populate column A (tickers) and column B (score percentages)
  tickers.forEach((ticker, index) => {
    const rowNumber = 2 + index; // Starting from row 2
    const stockSheet = workbook.getWorksheet(`${ticker} Results`);

    if (stockSheet) {
      // Set the ticker name in column A (A2, A3, A4, etc.) with a hyperlink to the respective worksheet (A1 cell)
      summarySheet.getCell(`A${rowNumber}`).value = {
        text: ticker, // Display text for the hyperlink
        hyperlink: `#'${ticker} Results'!A1`, // Hyperlink to the respective worksheet's A1 cell
      };

      // Create a reference to cell I2 of the stock sheet for the score percent in column B
      summarySheet.getCell(`B${rowNumber}`).value = {
        formula: `'${ticker} Results'!I2`, // Reference the I2 cell from the stock sheet
      };
    }
  });

  // Auto-resize the columns for better visibility
  summarySheet.getColumn(1).width = 15; // Ticker column width
  summarySheet.getColumn(2).width = 20; // Score percent column width
};

// Main function to read, duplicate, and apply data to Excel
const processExcelTemplate = async (
  formattedReport: Record<string, FinalReport>,
) => {
  const workbook = new ExcelJS.Workbook();
  const templatePath = path.resolve('./input/TICKER_TEMPLATE.xlsx');

  if (!fs.existsSync(templatePath)) {
    console.error('Template file not found.');
    return;
  }

  // Read the template
  await workbook.xlsx.readFile(templatePath);

  const tickersWithData = [];

  // Iterate over each ticker in the report
  for (const ticker in formattedReport) {
    const reportData = formattedReport[ticker];

    // Sort years in reverse chronological order and take the most recent 4
    const years = Object.keys(reportData.balanceSheet)
      .sort((a, b) => new Date(b).getTime() - new Date(a).getTime())
      .slice(0, 4);

    // Log the final results for each ticker
    console.log(
      `Final results for ${ticker}:`,
      JSON.stringify(reportData, null, 2),
    );

    // Get the template worksheet
    const templateSheet = workbook.getWorksheet('<TICKER> Results');
    if (!templateSheet) {
      console.error('Template sheet "<TICKER> Results" not found.');
      return;
    }

    // Duplicate and rename the sheet, and set default zoom to 200%
    const newSheet = workbook.addWorksheet(`${ticker} Results`, {
      properties: { tabColor: { argb: 'FF00FF00' } }, // Adjust if needed
      views: [{ state: 'normal', zoomScale: 200 }], // Set default zoom to 200%
    });

    templateSheet.eachRow((row, rowIndex) => {
      const newRow = newSheet.getRow(rowIndex);
      newRow.values = row.values;

      // Copy styles
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        newCell.style = cell.style; // Copy styles
        if (cell.formula) {
          newCell.value = { formula: cell.formula, result: cell.result }; // Copy formulas
        }
      });
    });

    // Apply data to the new sheet
    applyMappingsToSheet(newSheet, reportData, years);

    // Auto-resize the columns based on content
    autoResizeColumns(newSheet);

    // Add ticker to the list of tickers with data
    tickersWithData.push(ticker);
  }

  // Update the existing summary page with the new tickers and score references
  await updateSummaryPage(workbook, tickersWithData);

  // Save the updated workbook to the output directory
  const outputPath = path.resolve(`./output/${SELECTED_INDUSTRY}.xlsx`);
  await workbook.xlsx.writeFile(outputPath);
  console.log(`Excel file saved to ${outputPath}`);
};

// Main function
const main = async () => {
  await processInBatches(tickers); // Process tickers in batches

  // Format the final report (sort dates and limit to 4 most recent)
  const formattedReport = formatFinalReport(finalReport);

  // Process the Excel template and apply data
  await processExcelTemplate(formattedReport);
};

main().catch((error) => {
  console.error('Unhandled error:', error.message);
  // process.exit(1); // Ensure the process exits gracefully on unhandled errors
});

// on sig int, exit gracefully (export to excel)
process.on('SIGINT', () => {
  console.log('Exiting gracefully...');

  const formattedReport = formatFinalReport(finalReport);
  console.log('Final report:', JSON.stringify(formattedReport, null, 2));
  console.log('Exporting to Excel...');
  processExcelTemplate(formattedReport).then(() => {
    console.log('Excel export complete.');
    process.exit(0);
  });
});
