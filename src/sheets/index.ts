// Spreadsheet utilities
export {
  getSpreadsheet,
  createSheet,
  getSheet,
} from "./spreadsheet";

// Write operations
export {
  appendObject,
  appendObjects,
  updateObject,
  updateObjects,
  deleteObject,
  deleteObjects,
  deleteObjectsByFilter,
  clearAll,
  upsertObject,
  upsertObjects,
  replaceAll,
} from "./write";

// Read operations
export {
  getAllObjects,
  getObject,
  getObjectBatch,
  getHeaderMap,
  filterObjects,
  findObject,
  findObjectIndex,
  countObjects,
  getFirst,
  getLast,
  exists,
  sortObjects,
  getObjectsPaginated,
  sum,
  average,
  min,
  max,
  groupBy,
  getDistinctValues,
  filterByColumn,
} from "./read";

// Formatting operations
export {
  trim,
  trimRows,
  trimColumns,
} from "./formatting";

// Types
export type {
  SheetCellValue,
  SheetRow,
  GenericObject,
  HeaderMap,
} from "./types";

