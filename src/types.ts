export type ReportedValue = {
  raw?: number;
  fmt?: string;
};

export type FinancialData = {
  dataId?: number;
  asOfDate?: string;
  periodType?: string;
  currencyCode?: string;
  reportedValue?: ReportedValue;
};

export type FinancialMeta = {
  symbol?: string[];
  type?: string[];
};

export type TimeseriesItem = {
  meta?: FinancialMeta;
  timestamp?: number[];
} & Record<string, (FinancialData | null)[] | undefined>;

export type TimeseriesResult = {
  result?: TimeseriesItem[];
};

export type ResponseBody = {
  timeseries?: TimeseriesResult;
};

export type FinanceApiResponse = {
  status?: number;
  statusText?: string;
  headers?: Record<string, any>;
  body?: string | ResponseBody; // body can be a JSON string or the parsed object
};

// Type guards to validate if the data is of a specific type
export const isFinancialDataArray = (
  data: FinancialData[] | null | undefined,
): data is FinancialData[] => {
  return Array.isArray(data);
};

export const isFinancialData = (
  data: FinancialData | null | undefined,
): data is FinancialData => {
  return data !== null && data !== undefined;
};
