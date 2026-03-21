export type TelemetryStage = "request_start" | "schema" | "ai_call" | "action_execute";
export type TelemetryOutcome = "success" | "failure" | "retry";
export type TelemetryEventName =
  | "request_start"
  | "schema_done"
  | "ai_done"
  | "action_done"
  | "failure"
  | "retry_attempt";

export interface TelemetryEvent {
  eventName: TelemetryEventName;
  request_id: string;
  timestamp: string;
  stage: TelemetryStage;
  source: "taskpane" | "proxy" | "backend";
  outcome: TelemetryOutcome;
  latency_ms: number;
  model_used?: string;
  action_type?: string;
  error_class?: "validation" | "network" | "model" | "execution" | "unknown";
  error_message?: string;
  input_type?: "text" | "image" | "text+image";
  user_intent_summary?: string;
  sheet_count?: number;
  sample_range_count?: number;
  ai_latency_ms?: number;
  schema_latency_ms?: number;
  action_latency_ms?: number;
  token_input_estimate?: number;
  token_output_estimate?: number;
  action_count?: number;
  retry_attempt_number?: number;
  retry_reason?: string;
  max_retry_allowed?: number;
  is_recoverable?: boolean;
  error_code?: string;
}

const TELEMETRY_WINDOW_EVENT = "excel-ai-telemetry";
const DEFAULT_TELEMETRY_ENDPOINT = "http://localhost:3001/api/telemetry/events";
const TELEMETRY_BATCH_SIZE = 20;
const TELEMETRY_FLUSH_INTERVAL_MS = 2500;

let telemetryForwarderStarted = false;
let telemetryQueue: TelemetryEvent[] = [];
let telemetryFlushTimer: number | null = null;
let telemetryFlushInFlight = false;

export function createTelemetryRequestId(): string {
  return `req_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

export function classifyError(message: string): TelemetryEvent["error_class"] {
  const lower = message.toLowerCase();
  if (lower.includes("timeout") || lower.includes("network") || lower.includes("fetch")) return "network";
  if (lower.includes("validation") || lower.includes("invalid") || lower.includes("payload")) return "validation";
  if (lower.includes("model") || lower.includes("ai") || lower.includes("proxy")) return "model";
  if (lower.includes("excel") || lower.includes("execute") || lower.includes("worksheet")) return "execution";
  return "unknown";
}

export function emitTelemetryEvent(partial: Omit<TelemetryEvent, "timestamp">): void {
  const event: TelemetryEvent = {
    ...partial,
    timestamp: new Date().toISOString(),
    latency_ms: Number.isFinite(partial.latency_ms) ? Math.max(0, Math.round(partial.latency_ms)) : 0,
  };

  try {
    window.dispatchEvent(new CustomEvent(TELEMETRY_WINDOW_EVENT, { detail: event }));
  } catch {
    // Ignore dispatch failures in environments without CustomEvent support.
  }

  // Console output for immediate baseline collection during Sprint 1.
  console.info("[Telemetry]", event);
}

function getTelemetryEndpoint(): string {
  return localStorage.getItem("telemetry_endpoint") || DEFAULT_TELEMETRY_ENDPOINT;
}

async function flushTelemetryQueue(): Promise<void> {
  if (telemetryFlushInFlight || telemetryQueue.length === 0) {
    return;
  }

  telemetryFlushInFlight = true;
  const batch = telemetryQueue.slice(0, TELEMETRY_BATCH_SIZE);

  try {
    const response = await fetch(getTelemetryEndpoint(), {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ events: batch }),
    });

    if (!response.ok) {
      throw new Error(`Telemetry proxy ${response.status}`);
    }

    telemetryQueue = telemetryQueue.slice(batch.length);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.warn("[Telemetry] flush gagal:", message);
  } finally {
    telemetryFlushInFlight = false;
  }
}

function ensureFlushTimer(): void {
  if (telemetryFlushTimer !== null) {
    return;
  }

  telemetryFlushTimer = window.setInterval(() => {
    void flushTelemetryQueue();
  }, TELEMETRY_FLUSH_INTERVAL_MS);
}

export function startTelemetryForwarder(): void {
  if (telemetryForwarderStarted) {
    return;
  }
  telemetryForwarderStarted = true;

  window.addEventListener(TELEMETRY_WINDOW_EVENT, (event: Event) => {
    const customEvent = event as CustomEvent<TelemetryEvent>;
    const payload = customEvent.detail;
    if (!payload) return;

    telemetryQueue.push(payload);
    if (telemetryQueue.length >= TELEMETRY_BATCH_SIZE) {
      void flushTelemetryQueue();
    }
  });

  ensureFlushTimer();
}
