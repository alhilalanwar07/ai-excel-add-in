import type { AIMasterPayload, ActionType } from "./gemini";

export interface PendingActionLike {
  name: ActionType;
  args: Record<string, unknown>;
  details: AIMasterPayload;
}

export interface ActionValidationResult {
  isValid: boolean;
  normalizedAction: PendingActionLike;
  warnings: string[];
  errors: string[];
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function asFiniteNumber(value: unknown): number | null {
  if (typeof value !== "number" || !Number.isFinite(value)) return null;
  return value;
}

function normalizeIntegerField(
  args: Record<string, unknown>,
  field: string,
  fallback: number,
  minValue: number,
  maxValue: number,
  warnings: string[]
): number {
  const raw = args[field];
  const parsed = asFiniteNumber(raw);
  if (parsed === null) {
    if (raw !== undefined) {
      warnings.push(`${field} dinormalisasi dari ${String(raw)} menjadi ${fallback}.`);
    }
    return fallback;
  }

  const rounded = Math.floor(parsed);
  const clamped = Math.max(minValue, Math.min(maxValue, rounded));
  if (clamped !== rounded) {
    warnings.push(`${field} dinormalisasi dari ${rounded} menjadi ${clamped}.`);
  }
  return clamped;
}

function normalizeEnumField(
  args: Record<string, unknown>,
  field: string,
  fallback: string,
  allowed: string[],
  warnings: string[]
): string {
  const raw = args[field];
  if (!isNonEmptyString(raw)) {
    if (raw !== undefined) {
      warnings.push(`${field} dinormalisasi dari ${String(raw)} menjadi ${fallback}.`);
    }
    return fallback;
  }

  const normalized = raw.trim().toLowerCase();
  if (!allowed.includes(normalized)) {
    warnings.push(`${field} dinormalisasi dari ${raw} menjadi ${fallback}.`);
    return fallback;
  }
  return normalized;
}

function validateWriteFormula(args: Record<string, unknown>, errors: string[]): void {
  if (!isNonEmptyString(args.address)) errors.push("write_formula membutuhkan address.");
  if (!isNonEmptyString(args.formula)) errors.push("write_formula membutuhkan formula.");
}

function validateBulkWriteFormulas(args: Record<string, unknown>, errors: string[]): void {
  const hasRangePayload = isNonEmptyString(args.rangeAddress) && Array.isArray(args.formulas);
  const hasItemsPayload = Array.isArray(args.items) && args.items.length > 0;

  if (!hasRangePayload && !hasItemsPayload) {
    errors.push("bulk_write_formulas membutuhkan rangeAddress+formulas atau items[].");
    return;
  }

  if (hasItemsPayload) {
    const items = args.items as Array<Record<string, unknown>>;
    const validItemCount = items.filter((item) => isNonEmptyString(item.address) && isNonEmptyString(item.formula)).length;
    if (validItemCount === 0) {
      errors.push("bulk_write_formulas items[] harus berisi address dan formula yang valid.");
    }
  }
}

function validatePivotSummary(args: Record<string, unknown>, warnings: string[], errors: string[]): Record<string, unknown> {
  if (!isNonEmptyString(args.sourceRange)) {
    errors.push("pivot_summary membutuhkan sourceRange.");
  }

  const normalized: Record<string, unknown> = { ...args };
  normalized.groupByColumn = normalizeIntegerField(args, "groupByColumn", 0, 0, 200, warnings);
  normalized.valueColumn = normalizeIntegerField(args, "valueColumn", 1, 0, 200, warnings);
  normalized.topN = normalizeIntegerField(args, "topN", 10, 1, 5000, warnings);

  const minValueRaw = asFiniteNumber(args.minValue);
  if (minValueRaw === null) {
    if (args.minValue !== undefined) {
      warnings.push(`minValue dinormalisasi dari ${String(args.minValue)} menjadi 0.`);
    }
    normalized.minValue = 0;
  } else {
    const clamped = Math.max(0, minValueRaw);
    if (clamped !== minValueRaw) {
      warnings.push(`minValue dinormalisasi dari ${minValueRaw} menjadi ${clamped}.`);
    }
    normalized.minValue = clamped;
  }

  normalized.aggregation = normalizeEnumField(args, "aggregation", "sum", ["sum", "avg", "count"], warnings);
  normalized.chartType = normalizeEnumField(args, "chartType", "column", ["column", "bar", "line", "pie"], warnings);
  normalized.chartPreset = normalizeEnumField(args, "chartPreset", "compact", ["compact", "presentation", "executive"], warnings);

  if (typeof args.sortDescending !== "boolean") {
    if (args.sortDescending !== undefined) {
      warnings.push(`sortDescending dinormalisasi dari ${String(args.sortDescending)} menjadi true.`);
    }
    normalized.sortDescending = true;
  }

  if (!isNonEmptyString(args.outputSheetName)) {
    if (args.outputSheetName !== undefined) {
      warnings.push(`outputSheetName dinormalisasi dari ${String(args.outputSheetName)} menjadi PivotSummary.`);
    }
    normalized.outputSheetName = "PivotSummary";
  }

  return normalized;
}

function validateDataManipulation(args: Record<string, unknown>, warnings: string[], errors: string[]): Record<string, unknown> {
  const normalized: Record<string, unknown> = { ...args };
  if (!isNonEmptyString(args.range)) {
    errors.push("data_manipulation membutuhkan range.");
  }

  normalized.operation = normalizeEnumField(
    args,
    "operation",
    "sort",
    ["sort", "filter", "remove_duplicates", "delete"],
    warnings
  );
  return normalized;
}

export function validateAndNormalizeAction(action: PendingActionLike): ActionValidationResult {
  const warnings: string[] = [];
  const errors: string[] = [];
  let normalizedArgs: Record<string, unknown> = { ...action.args };

  switch (action.name) {
    case "write_formula":
      validateWriteFormula(normalizedArgs, errors);
      break;
    case "bulk_write_formulas":
      validateBulkWriteFormulas(normalizedArgs, errors);
      break;
    case "pivot_summary":
      normalizedArgs = validatePivotSummary(normalizedArgs, warnings, errors);
      break;
    case "data_manipulation":
      normalizedArgs = validateDataManipulation(normalizedArgs, warnings, errors);
      break;
    case "insert_data":
      if (!isNonEmptyString(normalizedArgs.startAddress)) errors.push("insert_data membutuhkan startAddress.");
      if (!Array.isArray(normalizedArgs.dataValues)) errors.push("insert_data membutuhkan dataValues array.");
      break;
    case "format_range":
    case "clear_range":
      if (!isNonEmptyString(normalizedArgs.address)) errors.push(`${action.name} membutuhkan address.`);
      break;
    case "chart":
      if (!isNonEmptyString(normalizedArgs.data_range)) errors.push("chart membutuhkan data_range.");
      normalizedArgs.chart_type = normalizeEnumField(
        normalizedArgs,
        "chart_type",
        "column",
        ["column", "bar", "line", "pie"],
        warnings
      );
      break;
    case "analysis":
    case "clarification":
      break;
    default:
      errors.push(`Aksi ${String(action.name)} belum didukung validator.`);
      break;
  }

  const normalizedDetails: AIMasterPayload = {
    ...action.details,
    execution_payload: normalizedArgs,
  };

  return {
    isValid: errors.length === 0,
    normalizedAction: {
      ...action,
      args: normalizedArgs,
      details: normalizedDetails,
    },
    warnings,
    errors,
  };
}
