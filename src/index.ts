/** DCA Calculator (2k24 edition) */

import { exec } from 'child_process';
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

const tickers = [
  'AAPL',
  'MSFT',
  'AMZN',
  'NVDA',
  'GOOGL',
  'GOOG',
  // 'BRK.B',
  // 'META',
  // 'TSLA',
  // 'UNH',
  // 'JNJ',
  // 'V',
  // 'PG',
  // 'XOM',
  // 'MA',
  // 'LLY',
  // 'HD',
  // 'MRK',
  // 'ABBV',
  // 'CVX',
  // 'PEP',
  // 'KO',
  // 'PFE',
  // 'BAC',
  // 'AVGO',
  // 'COST',
  // 'MCD',
  // 'TMO',
  // 'CSCO',
  // 'DHR',
  // 'WMT',
  // 'ACN',
  // 'ABT',
  // 'LIN',
  // 'ADBE',
  // 'VZ',
  // 'DIS',
  // 'NFLX',
  // 'NEE',
  // 'TXN',
  // 'PM',
  // 'MS',
  // 'UNP',
  // 'HON',
  // 'AMD',
  // 'RTX',
  // 'AMGN',
  // 'CRM',
  // 'LOW',
  // 'INTC',
  // 'SCHW',
  // 'BMY',
  // 'SPGI',
  // 'IBM',
  // 'COP',
  // 'MDT',
  // 'GS',
  // 'BLK',
  // 'ELV',
  // 'CVS',
  // 'GE',
  // 'CAT',
  // 'LMT',
  // 'AXP',
  // 'PLD',
  // 'UPS',
  // 'DE',
  // 'T',
  // 'CB',
  // 'ISRG',
  // 'MMC',
  // 'INTU',
  // 'NOW',
  // 'MO',
  // 'DUK',
  // 'MDLZ',
  // 'PYPL',
  // 'TMUS',
  // 'ADP',
  // 'SYK',
  // 'AMT',
  // 'BKNG',
  // 'C',
  // 'ZTS',
  // 'TJX',
  // 'APD',
  // 'TGT',
  // 'BDX',
  // 'REGN',
  // 'PNC',
  // 'CCI',
  // 'SO',
  // 'CI',
  // 'PGR',
  // 'GM',
  // 'USB',
  // 'GILD',
  // 'EW',
  // 'ADI',
  // 'MU',
  // 'ITW',
  // 'F',
  // 'NOC',
  // 'ETN',
  // 'CL',
  // 'NSC',
  // 'WM',
  // 'NKE',
  // 'ECL',
  // 'AON',
  // 'D',
  // 'ROP',
  // 'CSX',
  // 'SBUX',
  // 'TFC',
  // 'AEP',
  // 'MRNA',
  // 'HUM',
  // 'SHW',
  // 'EMR',
  // 'MAR',
  // 'CMCSA',
  // 'ORLY',
  // 'FDX',
  // 'CDNS',
  // 'BDX',
  // 'FISV',
  // 'HCA',
  // 'LRCX',
  // 'KMB',
  // 'AIG',
  // 'DG',
  // 'CTVA',
  // 'PH',
  // 'MPC',
  // 'EXC',
  // 'CME',
  // 'MCO',
  // 'APTV',
  // 'ADSK',
  // 'SPG',
  // 'OXY',
  // 'CHTR',
  // 'VRTX',
  // 'DOW',
  // 'KLAC',
  // 'TRV',
  // 'MCHP',
  // 'GIS',
  // 'STZ',
  // 'SNPS',
  // 'PSX',
  // 'SYY',
  // 'BK',
  // 'MCK',
  // 'MTD',
  // 'VLO',
  // 'HLT',
  // 'AFL',
  // 'ALL',
  // 'YUM',
  // 'SRE',
  // 'IQV',
  // 'AZO',
  // 'MNST',
  // 'KHC',
  // 'PAYX',
  // 'HES',
  // 'IDXX',
  // 'PCAR',
  // 'A',
  // 'CTAS',
  // 'STT',
  // 'PXD',
  // 'EOG',
  // 'DLR',
  // 'ADM',
  // 'WEC',
  // 'ODFL',
  // 'PPG',
  // 'MSCI',
  // 'GLW',
  // 'LHX',
  // 'ED',
  // 'WTW',
  // 'KR',
  // 'RMD',
  // 'BKR',
  // 'RSG',
  // 'NEM',
  // 'DHI',
  // 'FTNT',
  // 'PRU',
  // 'ORCL',
  // 'CARR',
  // 'GWW',
  // 'MKC',
  // 'ANET',
  // 'WMB',
  // 'KEYS',
  // 'VICI',
  // 'WBA',
  // 'ROST',
  // 'CDW',
  // 'ZBH',
  // 'O',
  // 'TDG',
  // 'TT',
  // 'NUE',
  // 'EFX',
  // 'XEL',
  // 'WELL',
  // 'AMP',
  // 'AEE',
  // 'CTSH',
  // 'HSY',
  // 'CMS',
  // 'PPL',
  // 'DLTR',
  // 'BAX',
  // 'NDAQ',
  // 'DFS',
  // 'EBAY',
  // 'TSCO',
  // 'VRSK',
  // 'LVS',
  // 'ZBRA',
  // 'SBAC',
  // 'DGX',
  // 'AWK',
  // 'DTE',
  // 'MPWR',
  // 'EXR',
  // 'FITB',
  // 'SIVB',
  // 'HAL',
  // 'VFC',
  // 'FRC',
  // 'EFTR',
  // 'KMX',
  // 'EXPE',
  // 'ANSS',
  // 'MLM',
  // 'CNP',
  // 'PEAK',
  // 'WY',
  // 'TER',
  // 'CINF',
  // 'AOS',
  // 'FDS',
  // 'IRM',
  // 'SWK',
  // 'PFG',
  // 'HRL',
  // 'TYL',
  // 'GPC',
  // 'MKTX',
  // 'NTRS',
  // 'ETSY',
  // 'HST',
  // 'ROL',
  // 'KSU',
  // 'WHR',
  // 'CHRW',
  // 'LKQ',
  // 'FTV',
  // 'BRO',
  // 'TPR',
  // 'ETR',
  // 'RJF',
  // 'CMS',
  // 'PKG',
  // 'LNT',
  // 'CF',
  // 'RCL',
  // 'CLX',
  // 'SWKS',
  // 'BXP',
  // 'ESS',
  // 'CEG',
  // 'XYL',
  // 'VTR',
  // 'TSN',
  // 'QRVO',
  // 'CBRE',
  // 'MAS',
  // 'AES',
  // 'HIG',
  // 'WAT',
  // 'ZION',
  // 'NVR',
  // 'COO',
  // 'ALLE',
  // 'SJM',
  // 'NDSN',
  // 'J',
  // 'WRB',
  // 'TXT',
  // 'UHS',
  // 'TROW',
  // 'HII',
  // 'NLSN',
  // 'CTLT',
  // 'TTWO',
  // 'WU',
  // 'DRI',
  // 'PHM',
  // 'PWR',
  // 'LUV',
  // 'OKE',
  // 'BBY',
  // 'PTC',
  // 'MRO',
  // 'ATO',
  // 'INVH',
  // 'FE',
  // 'PENN',
  // 'OMC',
  // 'NLOK',
  // 'FAST',
  // 'CBOE',
  // 'JBHT',
  // 'AES',
  // 'NI',
  // 'DOV',
  // 'FFIV',
  // 'STE',
  // 'FIS',
  // 'TAP',
  // 'WYNN',
  // 'IVZ',
  // 'XRX',
  // 'NWL',
  // 'LUMN',
  // 'LEG',
  // 'CMA',
  // 'HPQ',
  // 'RL',
  // 'PBCT',
  // 'MOS',
];

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

const MAX_RETRIES = 3; // Maximum number of retries
const BATCH_SIZE = 10; // Number of requests per batch
const RETRY_DELAY_MS = 2000; // Base delay between retries
const BATCH_DELAY_MS = 5000; // Delay between batches of requests

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

// Promisified runCurlCommand with retry logic and exponential backoff
const runCurlCommand = async (
  url: string,
  ticker: string,
  reportType: 'balanceSheet' | 'incomeStatement' | 'cashFlow',
  attempt = 1,
  retryDelay = RETRY_DELAY_MS,
): Promise<void> => {
  try {
    const command = `curl -s '${url}' -H 'user-agent: ${userAgent}'`;
    console.log(`Running: ${command}`);

    const { exec } = await import('child_process');
    exec(
      command,
      { maxBuffer: 1024 * 1000 * 20 },
      async (error, stdout, stderr) => {
        if (error) {
          throw new Error(
            `Error fetching ${reportType} for ${ticker}: ${error.message}`,
          );
        }
        if (stderr) {
          console.error(`stderr: ${stderr}`);
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
            console.log(`Failed to parse JSON for ${reportType} (${ticker})`);
          }
        } else {
          console.log(
            `No matching <script> tag found for ${reportType} (${ticker})`,
          );
        }
      },
    );
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
        retryDelay * 2,
      ); // Exponential backoff
    } else {
      console.error(
        `Failed after ${MAX_RETRIES + 1} attempts for ${ticker} (${reportType}). Error:`,
        err,
      );
      throw err; // Throw the error after exhausting retries
    }
  }
};

// Function to process tickers in batches with a delay between batches
const processInBatches = async (tickers: string[]) => {
  for (let i = 0; i < tickers.length; i += BATCH_SIZE) {
    const batch = tickers.slice(i, i + BATCH_SIZE);

    const promises: Promise<void>[] = [];
    for (const ticker of batch) {
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
          runCurlCommand(balanceSheetURLFinal, ticker, 'balanceSheet').catch(
            (err) => {
              console.error(`Error fetching Balance Sheet for ${ticker}:`, err);
            },
          ),
        );
      }

      if (config.fetchCashFlow) {
        const cashFlowURLFinal = cashFlowURL.replace('<TICKER>', ticker);
        promises.push(
          runCurlCommand(cashFlowURLFinal, ticker, 'cashFlow').catch((err) => {
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
  }

  // Save the updated workbook to the output directory
  const outputPath = path.resolve(`./output/TICKER_RESULTS.xlsx`);
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

main();
