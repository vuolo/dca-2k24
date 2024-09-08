/** DCA Calculator (2k24 edition) */

const tickers = ['NVDA', 'COST'];

const incomeStatementURL =
  'https://finance.yahoo.com/quote/<TICKER>/financials';
const balanceSheetURL =
  'https://finance.yahoo.com/quote/<TICKER>/balance-sheet';
const cashFlowURL = 'https://finance.yahoo.com/quote/<TICKER>/cash-flow';

const headers = {
  'user-agent':
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36',
};

const main = async () => {
  for (const ticker of tickers) {
    const [incomeStatement_response, balanceSheet_response, cashFlow_response] =
      await Promise.all([
        fetch(incomeStatementURL.replace('<TICKER>', ticker)),
        fetch(balanceSheetURL.replace('<TICKER>', ticker)),
        fetch(cashFlowURL.replace('<TICKER>', ticker)),
      ]);
  }
};

void main();
