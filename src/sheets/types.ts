export type SheetCellValue = string | number | boolean | Date;
export type SheetRow = SheetCellValue[];
export type GenericObject = Record<string, SheetCellValue>;
export type HeaderMap = Record<string, number>;

