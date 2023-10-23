export type BaseAddress = {
  col: number;
  row: number;
};

export type Address = BaseAddress & {
  sheetName?: string;
  address: string;
  $col$row: string;
};

export type Range = {
  tl: BaseAddress;
  br: BaseAddress;
};

export type Dimension = {
  width: number;
  height: number;
};

export type TemplateInfo = {
  top: number;
  left: number;
  bottom: number;
  right: number;
  /**
   * @returns e.g: 'Sheet1'
   */
  sheetName: string;
  /**
   * @return e.g: "B2:D6"
   */
  dimension: string;
  tl: Address;
  br: Address;
};

export type MergeRange = {
  model: { top: number; left: number; bottom: number; right: number };
};
