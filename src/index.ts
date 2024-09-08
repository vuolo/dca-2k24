/** DCA Calculator (2k24 edition) */

import { exec } from 'child_process';
import * as cheerio from 'cheerio';
import * as xlsx from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

import {
  FinanceApiResponse,
  FinancialData,
  isFinancialDataArray,
  ResponseBody,
} from './types';

/* eslint-disable no-console */

const tickers = ['NVDA', 'TSLA', 'AAPL', 'GOOGL', 'AMZN'];

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

// Promisified runCurlCommand
const runCurlCommand = (
  url: string,
  ticker: string,
  reportType: 'balanceSheet' | 'incomeStatement' | 'cashFlow',
): Promise<void> => {
  return new Promise((resolve, reject) => {
    const command = `curl -s '${url}' -H 'user-agent: ${userAgent}'`;
    console.log(`Running: ${command}`);

    exec(command, { maxBuffer: 1024 * 1000 * 20 }, (error, stdout, stderr) => {
      if (error) {
        console.error(
          `Error fetching ${reportType} for ${ticker}: ${error.message}`,
        );
        reject(error);
        return;
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

        // Try to parse the JSON from the script content
        const parsedData = parseJSONFromScript(extractedContent);

        if (parsedData) {
          console.log(
            `\n--- Parsed JSON Data for ${reportType} (${ticker}) ---\n`,
          );
          handleParsedData(parsedData, ticker, reportType); // Pass to handler function
          resolve();
        } else {
          console.log(`Failed to parse JSON for ${reportType} (${ticker})`);
          reject(new Error('Failed to parse JSON'));
        }
      } else {
        console.log(
          `No matching <script> tag found for ${reportType} (${ticker})`,
        );
        reject(new Error('No matching <script> tag found'));
      }
    });
  });
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
const applyMappingsToSheet = (
  worksheet: any,
  reportData: FinalReport,
  years: string[],
) => {
  // Apply years in header
  worksheet[excelMappings.years.C] = { v: years[0] };
  worksheet[excelMappings.years.D] = { v: years[1] };
  worksheet[excelMappings.years.E] = { v: years[2] };
  worksheet[excelMappings.years.F] = { v: years[3] };

  // Balance sheet
  for (const [field, cells] of Object.entries(excelMappings.balanceSheet)) {
    if (reportData.balanceSheet) {
      worksheet[cells.C] = {
        v: reportData.balanceSheet[years[0]]?.[field] || '',
      };
      worksheet[cells.D] = {
        v: reportData.balanceSheet[years[1]]?.[field] || '',
      };
      worksheet[cells.E] = {
        v: reportData.balanceSheet[years[2]]?.[field] || '',
      };
      worksheet[cells.F] = {
        v: reportData.balanceSheet[years[3]]?.[field] || '',
      };
    }
  }

  // Income statement
  for (const [field, cells] of Object.entries(excelMappings.incomeStatement)) {
    if (reportData.incomeStatement) {
      worksheet[cells.C] = {
        v: reportData.incomeStatement[years[0]]?.[field] || '',
      };
      worksheet[cells.D] = {
        v: reportData.incomeStatement[years[1]]?.[field] || '',
      };
      worksheet[cells.E] = {
        v: reportData.incomeStatement[years[2]]?.[field] || '',
      };
      worksheet[cells.F] = {
        v: reportData.incomeStatement[years[3]]?.[field] || '',
      };
    }
  }

  // Cash flow
  for (const [field, cells] of Object.entries(excelMappings.cashFlow)) {
    if (reportData.cashFlow) {
      worksheet[cells.C] = { v: reportData.cashFlow[years[0]]?.[field] || '' };
      worksheet[cells.D] = { v: reportData.cashFlow[years[1]]?.[field] || '' };
      worksheet[cells.E] = { v: reportData.cashFlow[years[2]]?.[field] || '' };
      worksheet[cells.F] = { v: reportData.cashFlow[years[3]]?.[field] || '' };
    }
  }
};

// Main function to read, duplicate, and apply data to Excel
const processExcelTemplate = async (
  formattedReport: Record<string, FinalReport>,
) => {
  const templatePath = path.resolve('./input/TICKER_TEMPLATE.xlsx');

  if (!fs.existsSync(templatePath)) {
    console.error('Template file not found.');
    return;
  }

  // Read the Excel file
  const workbook = xlsx.readFile(templatePath);

  // Iterate over each ticker in the report
  for (const ticker in formattedReport) {
    const reportData = formattedReport[ticker];

    // Sort years in reverse chronological order and take the most recent 4
    const years = Object.keys(reportData.balanceSheet)
      .sort((a, b) => new Date(b).getTime() - new Date(a).getTime())
      .slice(0, 4);

    // Get the template worksheet
    const templateSheetName = '<TICKER> Results';
    const worksheet = workbook.Sheets[templateSheetName];
    if (!worksheet) {
      console.error(`Template sheet "${templateSheetName}" not found.`);
      return;
    }

    // Duplicate and rename the sheet
    const newSheetName = `${ticker} Results`;
    const jsonSheet = xlsx.utils.sheet_to_json(worksheet, {
      header: 1,
      raw: true,
    });
    const aoaSheet = jsonSheet.map((row: any) => Object.values(row));
    const newSheet = xlsx.utils.aoa_to_sheet(aoaSheet);

    // Apply data to the new sheet
    applyMappingsToSheet(newSheet, reportData, years);

    // Add the new sheet to the workbook
    xlsx.utils.book_append_sheet(workbook, newSheet, newSheetName);
  }

  // Save the updated workbook to the output directory
  const outputPath = path.resolve(`./output/TICKER_RESULTS.xlsx`);
  xlsx.writeFile(workbook, outputPath);
  console.log(`Excel file saved to ${outputPath}`);
};

// After generating the final report, call the processing function
const main = async () => {
  const promises: Promise<void>[] = [];

  for (const ticker of tickers) {
    if (config.fetchIncomeStatement) {
      const incomeStatementURLFinal = incomeStatementURL.replace(
        '<TICKER>',
        ticker,
      );
      promises.push(
        runCurlCommand(incomeStatementURLFinal, ticker, 'incomeStatement'),
      );
    }

    if (config.fetchBalanceSheet) {
      const balanceSheetURLFinal = balanceSheetURL.replace('<TICKER>', ticker);
      promises.push(
        runCurlCommand(balanceSheetURLFinal, ticker, 'balanceSheet'),
      );
    }

    if (config.fetchCashFlow) {
      const cashFlowURLFinal = cashFlowURL.replace('<TICKER>', ticker);
      promises.push(runCurlCommand(cashFlowURLFinal, ticker, 'cashFlow'));
    }
  }

  // Wait for all curl commands to resolve
  await Promise.all(promises);

  // Format the final report (sort dates and limit to 4 most recent)
  const formattedReport = formatFinalReport(finalReport);

  // Process the Excel template and apply data
  await processExcelTemplate(formattedReport);
};

main();
