export type NullOption = "drop" | "fill0" | "fillMean" | "fillMode" | "leave";
export type BlankOption = "toNull" | "drop" | "leave";
export type ZeroOption = "leave" | "toNull" | "replaceMean" | "replaceMedian";
export type DuplicateOption = "remove" | "keep";

export interface CleaningConfig {
  nulls: NullOption;
  blanks: BlankOption;
  zeros: ZeroOption;
  duplicates: DuplicateOption;
}

export interface FileData {
  id: string;
  name: string;
  type: string;
  data: any[]; // Array of objects
  originalData: any[];
  metadata: {
    rows: number;
    cols: number;
    nullCount: number;
    blankCount: number;
    uniqueCount: number;
    columnStats: Record<string, {
      nulls: number;
      blanks: number;
      uniques: number;
      type: "numeric" | "string" | "other";
    }>;
  };
}

export interface RunHistory {
  id: string;
  timestamp: number;
  fileName: string;
  config: CleaningConfig;
  rowsRemoved: number;
  changes: string[];
}
