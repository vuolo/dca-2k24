/** DCA Calculator (2k24 edition) */

import { exec } from 'child_process';
import * as cheerio from 'cheerio';
import {
  FinanceApiResponse,
  FinancialData,
  isFinancialDataArray,
  ResponseBody,
} from './types';

/* eslint-disable no-console */

const config = {
  fetchIncomeStatement: true,
  fetchBalanceSheet: true,
  fetchCashFlow: true,
};

const tickers = ['NVDA'];

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

// Object to hold the final structured result
const finalReport: FinalReport = {
  balanceSheet: {},
  incomeStatement: {},
  cashFlow: {},
};

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
              // Store the data in the finalReport object
              if (!finalReport[reportType]) {
                finalReport[reportType] = {};
              }
              if (!finalReport[reportType][asOfDate]) {
                finalReport[reportType][asOfDate] = {};
              }
              finalReport[reportType][asOfDate][field] = reportedValue;
            }
          }
        });
      }
    });
  });
};

/**
 * Function to format the final report by reordering the date keys
 * in reverse chronological order and limiting to the most recent 4 years.
 */
const formatFinalReport = (finalReport: FinalReport): FinalReport => {
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
  const formattedReport: FinalReport = {
    balanceSheet: reorderAndLimitDates(finalReport.balanceSheet),
    incomeStatement: reorderAndLimitDates(finalReport.incomeStatement),
    cashFlow: reorderAndLimitDates(finalReport.cashFlow),
  };

  return formattedReport;
};

// At the end of the main function, call the formatter and log the result
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

  // Log the final structured report before formatting
  console.log('\n--- Final Structured Report ---\n');
  console.log(JSON.stringify(finalReport, null, 2));

  // Format the final report (sort dates and limit to 4 most recent)
  const formattedReport = formatFinalReport(finalReport);

  // Log the formatted report
  console.log('\n--- Formatted Final Report (Most Recent 4 Years) ---\n');
  console.log(JSON.stringify(formattedReport, null, 2));
};

main();
