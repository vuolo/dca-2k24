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

const tickers = [
  // 'COST',
  'NVDA',
  // 'AAPL',
  // 'TSLA'
];

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

// Handler function for parsed data
const handleParsedData = (
  data: ResponseBody,
  ticker: string,
  reportType: string,
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
        console.log(`\n**${field.toUpperCase()}** (${key}):`);
        fieldData?.forEach((entry) => {
          if (entry) {
            const asOfDate = entry?.asOfDate;
            const reportedValue = entry?.reportedValue?.raw;
            if (asOfDate && reportedValue !== undefined) {
              console.log(
                `  As of Date: ${asOfDate} - Reported Value: ${reportedValue}`,
              );
            }
          }
        });
      }
    });
  });
};

const runCurlCommand = (url: string, ticker: string, reportType: string) => {
  const command = `curl -s '${url}' -H 'user-agent: ${userAgent}'`;
  console.log(`Running: ${command}`);

  exec(command, { maxBuffer: 1024 * 1000 * 20 }, (error, stdout, stderr) => {
    if (error) {
      console.error(
        `Error fetching ${reportType} for ${ticker}: ${error.message}`,
      );
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
      } else {
        console.log(`Failed to parse JSON for ${reportType} (${ticker})`);
      }
    } else {
      console.log(
        `No matching <script> tag found for ${reportType} (${ticker})`,
      );
    }
  });
};

const main = async () => {
  for (const ticker of tickers) {
    if (config.fetchIncomeStatement) {
      const incomeStatementURLFinal = incomeStatementURL.replace(
        '<TICKER>',
        ticker,
      );
      runCurlCommand(incomeStatementURLFinal, ticker, 'incomeStatement');
    }

    if (config.fetchBalanceSheet) {
      const balanceSheetURLFinal = balanceSheetURL.replace('<TICKER>', ticker);
      runCurlCommand(balanceSheetURLFinal, ticker, 'balanceSheet');
    }

    if (config.fetchCashFlow) {
      const cashFlowURLFinal = cashFlowURL.replace('<TICKER>', ticker);
      runCurlCommand(cashFlowURLFinal, ticker, 'cashFlow');
    }
  }
};

main();
