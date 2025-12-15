export interface BankRow {
  __rowNum__: number;
  [key: string]: any;
}

export interface RuleStats {
  rule: number | string;
  count: number;
}

export interface ConfigMap {
  nameMap: Record<string, string>;
  amountMap: Record<string, string>;
}

export interface ProcessingResult {
  processedData: BankRow[];
  stats: RuleStats[];
  vlookupData: any[];
  rule3Mismatches: any[];
}
