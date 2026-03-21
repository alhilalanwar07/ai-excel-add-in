import * as React from "react";
import { useState, useEffect, useRef, useCallback, useMemo } from "react";
import {
  makeStyles,
  Input,
  Button,
  tokens,
  Title3,
  Body1,
  Spinner,
  Text,
} from "@fluentui/react-components";
import { SendRegular, KeyRegular, DocumentRegular } from "@fluentui/react-icons";
import { getWorkbookSchema, sendAICommand, invalidateWorkbookSchemaCache } from "../utils/gemini";
import { excelService } from "../services/ExcelService";
import type { AIMasterPayload, ActionType } from "../utils/gemini";
import { PROMPT_TEMPLATES } from "../constants/promptTemplates";
import {
  classifyError,
  createTelemetryRequestId,
  emitTelemetryEvent,
  startTelemetryForwarder,
} from "../utils/telemetry";
import { validateAndNormalizeAction } from "../utils/actionValidator";

type TemplateLevel = "Basic" | "Advanced" | "Automation";

// ─── Types ────────────────────────────────────────────────────────────────────

interface PendingActionData {
  name: ActionType;
  args: Record<string, unknown>;
  details: AIMasterPayload;
}

interface Message {
  /** ID unik — dipakai untuk update tertarget, menghindari bug index */
  id: string;
  role: "user" | "ai" | "system";
  text: string;
  actionOutput?: string;
  isPendingAwaitingConfirmation?: boolean;
  pendingActionData?: PendingActionData;
  imageBase64?: string;
  recommendedPrompt?: string;
  recommendedParams?: Array<{ field: string; value: string }>;
}

interface RecoverySuggestion {
  userMessage: string;
  recommendedPrompt: string;
  recommendedParams: Array<{ field: string; value: string }>;
}

type ModelId =
  | "gemini-2.5-flash"
  | "gemini-1.5-pro"
  | "gemini-2.5-pro"
  | "qwen/qwen3.5-397b-a17b"
  | "qwen/qwen3-coder-480b-a35b-instruct"
  | "meta/llama-3.2-90b-vision-instruct";

const MODEL_OPTIONS: { value: ModelId; label: string }[] = [
  { value: "gemini-2.5-flash", label: "Gemini 2.5 Flash (Gratis / Tercepat)" },
  { value: "gemini-1.5-pro",   label: "Gemini 1.5 Pro (Lebih Pintar)" },
  { value: "gemini-2.5-pro",   label: "Gemini 2.5 Pro (Berbayar)" },
  { value: "qwen/qwen3.5-397b-a17b", label: "NVIDIA Qwen 3.5 397B (Super Pintar)" },
  { value: "qwen/qwen3-coder-480b-a35b-instruct", label: "NVIDIA Qwen3 Coder 480B (Khusus Coding/Logic)" },
  { value: "meta/llama-3.2-90b-vision-instruct", label: "NVIDIA Llama 3.2 90B Vision (Reasoning Visual)" },
];

const MAX_SELF_CORRECTION = 3;

// ─── Styles ───────────────────────────────────────────────────────────────────

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    backgroundColor: tokens.colorBrandBackground,
    color: tokens.colorNeutralForegroundOnBrand,
  },
  title: { color: tokens.colorNeutralForegroundOnBrand },
  apiKeyRow: { display: "flex", gap: "8px", alignItems: "center" },
  chatArea: {
    flexGrow: 1,
    overflowY: "auto",
    padding: "16px",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  messageRow: {
    display: "flex",
    flexDirection: "column",
    maxWidth: "85%",
  },
  userRow: { alignSelf: "flex-end", alignItems: "flex-end" },
  aiRow:   { alignSelf: "flex-start", alignItems: "flex-start" },
  bubble: {
    padding: "10px 14px",
    borderRadius: "8px",
    wordBreak: "break-word",
  },
  userBubble: {
    backgroundColor: tokens.colorBrandBackground2,
    color: tokens.colorNeutralForeground1,
    borderBottomRightRadius: "2px",
  },
  aiBubble: {
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    borderBottomLeftRadius: "2px",
  },
  systemText: {
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
    marginTop: "4px",
  },
  inputArea: {
    padding: "16px",
    display: "flex",
    gap: "8px",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  inputField: { flexGrow: 1 },
  quickPromptCard: {
    margin: "8px 16px 0",
    padding: "10px",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: "8px",
    backgroundColor: tokens.colorNeutralBackground2,
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
  quickPromptActions: {
    display: "flex",
    gap: "8px",
  },
  quickPromptSelectors: {
    display: "flex",
    gap: "8px",
  },
});

// ─── Helpers ──────────────────────────────────────────────────────────────────

function mkId(): string {
  return crypto.randomUUID();
}

function mkMsg(partial: Omit<Message, "id">): Message {
  return { id: mkId(), ...partial };
}

function isDestructiveAction(data: PendingActionData): boolean {
  return (
    data.name === "clear_range" ||
    (data.name === "data_manipulation" && data.args["operation"] === "delete")
  );
}

function formatMs(value: number): string {
  return `${Math.round(value)}ms`;
}

function summarizeIntent(promptText: string): string {
  const normalized = promptText.replace(/\s+/g, " ").trim();
  return normalized.length > 80 ? `${normalized.slice(0, 77)}...` : normalized;
}

function getSchemaStats(contextData: string): { sheetCount: number; sampleRangeCount: number } {
  try {
    const parsed = JSON.parse(contextData) as {
      sheets?: unknown[];
      activeSheetInfo?: { selectionSample?: unknown[][] } | null;
    };
    const sheetCount = Array.isArray(parsed.sheets) ? parsed.sheets.length : 0;
    const sampleRangeCount = Array.isArray(parsed.activeSheetInfo?.selectionSample)
      ? parsed.activeSheetInfo!.selectionSample.length
      : 0;
    return { sheetCount, sampleRangeCount };
  } catch {
    return { sheetCount: 0, sampleRangeCount: 0 };
  }
}

function buildNormalizationMessage(warnings: string[]): string {
  const normalizedLines = warnings.map((warning) => {
    const match = warning.match(/^(\w+) dinormalisasi dari (.+) menjadi (.+)\.$/i);
    if (!match) return `- ${warning}`;

    const field = match[1];
    const fromValue = match[2];
    const toValue = match[3];
    return `- ${field}: ${fromValue} -> ${toValue} (gunakan ${toValue} untuk hasil konsisten)`;
  });

  return `⚙️ Penyesuaian otomatis diterapkan:\n${normalizedLines.join("\n")}`;
}

function buildRecommendedPromptFromWarnings(warnings: string[]): string | null {
  const recommendedPairs: string[] = [];

  warnings.forEach((warning) => {
    const match = warning.match(/^(\w+) dinormalisasi dari (.+) menjadi (.+)\.$/i);
    if (!match) return;

    const field = match[1];
    const toValue = match[3];
    recommendedPairs.push(`${field} ${toValue}`);
  });

  if (recommendedPairs.length === 0) return null;
  return `Gunakan parameter rekomendasi berikut pada perintah berikutnya: ${recommendedPairs.join(", ")}.`;
}

function extractRecommendedParams(warnings: string[]): Array<{ field: string; value: string }> {
  const params: Array<{ field: string; value: string }> = [];

  warnings.forEach((warning) => {
    const match = warning.match(/^(\w+) dinormalisasi dari (.+) menjadi (.+)\.$/i);
    if (!match) return;
    params.push({ field: match[1], value: match[3] });
  });

  return params;
}

function applyRecommendationsToPrompt(
  basePrompt: string,
  params: Array<{ field: string; value: string }>
): string {
  if (!basePrompt.trim()) {
    return params.map((p) => `${p.field} ${p.value}`).join(", ");
  }

  let updated = basePrompt;
  for (const param of params) {
    const escapedField = param.field.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const existingPattern = new RegExp(`\\b${escapedField}\\s+([^,.;\\n]+)`, "i");
    if (existingPattern.test(updated)) {
      updated = updated.replace(existingPattern, `${param.field} ${param.value}`);
    } else {
      updated = `${updated.replace(/\s+$/, "")}, ${param.field} ${param.value}`;
    }
  }

  return updated;
}

function buildRecoverySuggestion(errorClass: ReturnType<typeof classifyError>, errorMessage: string): RecoverySuggestion {
  const base = {
    userMessage: `Saran pemulihan: ${errorMessage}`,
    recommendedPrompt: "",
    recommendedParams: [] as Array<{ field: string; value: string }>,
  };

  if (errorClass === "validation") {
    return {
      userMessage: "Input belum valid. Sistem menyiapkan parameter aman agar bisa dieksekusi.",
      recommendedPrompt: "Gunakan parameter valid untuk aksi yang sama dengan target range aktif dan batas aman.",
      recommendedParams: [
        { field: "topN", value: "10" },
        { field: "minValue", value: "0" },
        { field: "sortDirection", value: "desc" },
      ],
    };
  }

  if (errorClass === "network") {
    return {
      userMessage: "Koneksi timeout. Coba request lebih ringkas atau retry dengan timeout lebih longgar.",
      recommendedPrompt: "Jalankan ulang dengan payload ringkas dan retry policy network.",
      recommendedParams: [
        { field: "timeoutMs", value: "45000" },
        { field: "maxRetries", value: "2" },
      ],
    };
  }

  if (errorClass === "model") {
    return {
      userMessage: "Model utama gagal. Disarankan fallback ke model cepat untuk melanjutkan flow.",
      recommendedPrompt: "Gunakan model fallback dan format respons action yang lebih ketat.",
      recommendedParams: [
        { field: "model", value: "gemini-2.5-flash" },
      ],
    };
  }

  if (errorClass === "execution") {
    return {
      userMessage: "Eksekusi Excel gagal. Coba targetkan range aktif terlebih dahulu.",
      recommendedPrompt: "Validasi sheet aktif dan jalankan ulang aksi pada range yang lebih sempit.",
      recommendedParams: [
        { field: "target_scope", value: "active_sheet" },
      ],
    };
  }

  return {
    ...base,
    userMessage: "Terjadi error tak terklasifikasi. Coba ulang dengan instruksi lebih spesifik.",
    recommendedPrompt: "Ulangi perintah dengan parameter eksplisit dan target range/sheet yang jelas.",
    recommendedParams: [],
  };
}

// ─── Component ────────────────────────────────────────────────────────────────

const App: React.FC = () => {
  const styles = useStyles();

  const [selectedModel, setSelectedModel] = useState<ModelId>("gemini-2.5-pro");
  const [apiKey,        setApiKey]        = useState("");
  const [isKeySaved,    setIsKeySaved]    = useState(false);
  const [messages,      setMessages]      = useState<Message[]>([
    mkMsg({ role: "ai", text: "Halo! Saya AI Agent Excel Anda. Masukkan API Key untuk mulai." }),
  ]);
  const [prompt,     setPrompt]     = useState("");
  const [attachedImage, setAttachedImage] = useState<string | null>(null);
  const [isLoading,  setIsLoading]  = useState(false);
  const [selectedTemplateLevel, setSelectedTemplateLevel] = useState<TemplateLevel>("Basic");
  const [selectedTemplateId, setSelectedTemplateId] = useState<string>(PROMPT_TEMPLATES[0]?.id ?? "");
  const lastUserPromptRef = useRef<string>("");

  // ── FIX #1: Ref yang selalu sinkron dengan state messages terbaru.
  // Semua callback membaca ini — bukan `messages` langsung — agar tidak stale.
  const messagesRef = useRef<Message[]>(messages);
  useEffect(() => { messagesRef.current = messages; }, [messages]);

  // ── Auto-scroll ke pesan terbaru
  const chatBottomRef = useRef<HTMLDivElement>(null);
  useEffect(() => {
    chatBottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  // ── Sinkronkan API key dari localStorage setiap model berubah
  useEffect(() => {
    const isNvidiaModel = selectedModel.includes("qwen") || selectedModel.startsWith("meta");
    const key = isNvidiaModel ? "nvidia_api_key" : "gemini_api_key";
    const saved = localStorage.getItem(key) ?? "";
    setApiKey(saved);
    setIsKeySaved(saved.length > 0);
  }, [selectedModel]);

  useEffect(() => {
    startTelemetryForwarder();
  }, []);

  useEffect(() => {
    const eventType = Office.EventType.DocumentSelectionChanged;
    const handler = () => invalidateWorkbookSchemaCache();

    Office.context.document.addHandlerAsync(eventType, handler);

    return () => {
      Office.context.document.removeHandlerAsync(eventType, { handler });
    };
  }, []);

  // ─── appendMsg ───────────────────────────────────────────────────────────
  // Tambah pesan ke state DAN ke ref secara sinkron.
  // Mengembalikan snapshot terbaru agar langsung bisa dipakai sebagai chatHistory.
  const appendMsg = useCallback((msg: Message): Message[] => {
    const next = [...messagesRef.current, msg];
    messagesRef.current = next;       // sinkron sebelum re-render
    setMessages(next);
    return next;
  }, []);

  // ─── updateMsgById ───────────────────────────────────────────────────────
  // FIX #2: Update via ID bukan index array — tidak ada risiko off-by-one.
  const updateMsgById = useCallback((id: string, patch: Partial<Message>) => {
    setMessages((prev) =>
      prev.map((m) => (m.id === id ? { ...m, ...patch } : m))
    );
  }, []);

  // ─── executeAction ───────────────────────────────────────────────────────
  // Satu fungsi tunggal yang menangani: eksekusi → success loop → self-correction.
  // FIX #3: finally selalu memanggil setIsLoading(false) tanpa syarat retryCount.
  const executeAction = useCallback(
    async (
      msgId: string,
      actionData: PendingActionData,
      retryCount: number,
      _historySnapshot: Message[]
    ): Promise<void> => {
      if (retryCount === 0) setIsLoading(true);

      try {
        const actionStartedAt = performance.now();
        const executionResult = await excelService.executeAction(
          actionData.name,
          actionData.args,
          actionData.details
        );
        invalidateWorkbookSchemaCache();

        emitTelemetryEvent({
          eventName: "action_done",
          request_id: msgId,
          stage: "action_execute",
          source: "taskpane",
          outcome: "success",
          latency_ms: performance.now() - actionStartedAt,
          action_type: actionData.name,
          action_count: 1,
          action_latency_ms: performance.now() - actionStartedAt,
        });

        if (executionResult.warnings.length > 0) {
          const recommendedPrompt = buildRecommendedPromptFromWarnings(executionResult.warnings);
          const recommendedParams = extractRecommendedParams(executionResult.warnings);
          appendMsg(
            mkMsg({
              role: "system",
              text: buildNormalizationMessage(executionResult.warnings),
              actionOutput: `[Normalization warnings: ${executionResult.warnings.length}]`,
              recommendedPrompt: recommendedPrompt ?? undefined,
              recommendedParams: recommendedParams.length > 0 ? recommendedParams : undefined,
            })
          );
        }

        const successMsg = mkMsg({
          role: "system",
          text: `✅ ${actionData.name} berhasil.${retryCount > 0 ? ` (Setelah ${retryCount}x perbaikan)` : ""}`,
          actionOutput: `[System: ${actionData.name} executed. Proceed if needed.]`,
        });

        // appendMsg mengembalikan snapshot terbaru — langsung pakai sebagai chatHistory
        const historyAfterSuccess = appendMsg(successMsg);

        // ── Continuous Agent Loop ──────────────────────────────────────────
        try {
          const ctxData = await getWorkbookSchema();
          const loopResp = await sendAICommand({
            // Pesan eksplisit agar AI tidak bingung dengan string kosong
            userMessage:
              "[SYSTEM] Aksi sebelumnya berhasil. Lanjutkan ke langkah berikutnya jika ada, atau beri tahu user bahwa tugas selesai.",
            apiKey,
            contextData: ctxData,
            chatHistory: historyAfterSuccess, // FIX #1: bukan `messages` (stale)
            model: selectedModel,
            requestId: msgId,
            onRetryAttempt: ({ attempt, maxRetries, reason }) => {
              emitTelemetryEvent({
                eventName: "retry_attempt",
                request_id: msgId,
                stage: "ai_call",
                source: "taskpane",
                outcome: "retry",
                latency_ms: 0,
                model_used: selectedModel,
                retry_attempt_number: attempt,
                max_retry_allowed: maxRetries,
                retry_reason: reason,
              });
            },
          });

          await processAIResponse(loopResp, historyAfterSuccess);
        } catch (loopErr) {
          console.warn("[ContinuousLoop] Gagal:", loopErr);
          // Tidak fatal — user sudah melihat pesan sukses di atas
        }
      } catch (execError) {
        const errMsg =
          execError instanceof Error ? execError.message : String(execError);

        if (retryCount < MAX_SELF_CORRECTION) {
          appendMsg(
            mkMsg({
              role: "ai",
              text: `⚠️ Gagal: ${errMsg}. Mencoba perbaikan otomatis (${retryCount + 1}/${MAX_SELF_CORRECTION})...`,
            })
          );

          try {
            const ctxData = await getWorkbookSchema();
            const fixResp = await sendAICommand({
              userMessage: "",
              apiKey,
              contextData: ctxData,
              chatHistory: messagesRef.current, // FIX #1: baca ref, bukan closure
              model: selectedModel,
              errorFeedback: errMsg,
              requestId: msgId,
              onRetryAttempt: ({ attempt, maxRetries, reason }) => {
                emitTelemetryEvent({
                  eventName: "retry_attempt",
                  request_id: msgId,
                  stage: "ai_call",
                  source: "taskpane",
                  outcome: "retry",
                  latency_ms: 0,
                  model_used: selectedModel,
                  retry_attempt_number: attempt,
                  max_retry_allowed: maxRetries,
                  retry_reason: reason,
                });
              },
            });

            if (fixResp.functionCall) {
              const fixedAction: PendingActionData = {
                name: fixResp.functionCall.name,
                args: fixResp.functionCall.args,
                details: fixResp.functionCall.parsedDetails,
              };
              const validatedFix = validateAndNormalizeAction(fixedAction);
              if (!validatedFix.isValid) {
                const validationError = validatedFix.errors.join(" ");
                const recovery = buildRecoverySuggestion("validation", validationError);
                emitTelemetryEvent({
                  eventName: "failure",
                  request_id: msgId,
                  stage: "action_execute",
                  source: "taskpane",
                  outcome: "failure",
                  latency_ms: 0,
                  action_type: fixedAction.name,
                  error_class: "validation",
                  error_message: validationError,
                  is_recoverable: false,
                });
                appendMsg(
                  mkMsg({
                    role: "ai",
                    text: `❌ Payload aksi tidak valid: ${validationError}\n\n${recovery.userMessage}`,
                    recommendedPrompt: recovery.recommendedPrompt,
                    recommendedParams: recovery.recommendedParams,
                  })
                );
                return;
              }

              if (validatedFix.warnings.length > 0) {
                const recommendedPrompt = buildRecommendedPromptFromWarnings(validatedFix.warnings);
                const recommendedParams = extractRecommendedParams(validatedFix.warnings);
                appendMsg(
                  mkMsg({
                    role: "system",
                    text: buildNormalizationMessage(validatedFix.warnings),
                    actionOutput: `[Validator warnings: ${validatedFix.warnings.length}]`,
                    recommendedPrompt: recommendedPrompt ?? undefined,
                    recommendedParams: recommendedParams.length > 0 ? recommendedParams : undefined,
                  })
                );
              }

              await executeAction(msgId, validatedFix.normalizedAction, retryCount + 1, messagesRef.current);
            } else {
              appendMsg(mkMsg({ role: "ai", text: "❌ AI tidak menemukan solusi perbaikan." }));
            }
          } catch (fixErr) {
            const fixMsg = fixErr instanceof Error ? fixErr.message : String(fixErr);
            appendMsg(mkMsg({ role: "ai", text: `❌ Gagal menghubungi API saat perbaikan: ${fixMsg}` }));
          }
        } else {
          emitTelemetryEvent({
            eventName: "failure",
            request_id: msgId,
            stage: "action_execute",
            source: "taskpane",
            outcome: "failure",
            latency_ms: 0,
            action_type: actionData.name,
            error_class: classifyError(errMsg),
            error_message: errMsg,
            is_recoverable: false,
          });

          appendMsg(
            mkMsg({
              role: "ai",
              text: `❌ Dibatalkan setelah ${MAX_SELF_CORRECTION}x gagal: ${errMsg}`,
            })
          );
        }
      } finally {
        // FIX #3: Selalu reset — tidak peduli retryCount atau rekursi
        setIsLoading(false);
      }
    },
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [apiKey, selectedModel, appendMsg]
  );

  // ─── processAIResponse ───────────────────────────────────────────────────
  // Tangani functionCall / text dari respons AI.
  // Dipakai di handleSend dan di continuous loop agar tidak duplikasi kode.
  const processAIResponse = useCallback(
    async (
      response: Awaited<ReturnType<typeof sendAICommand>>,
      _currentHistory: Message[]
    ): Promise<void> => {
      if (!response.functionCall) {
        emitTelemetryEvent({
          eventName: "action_done",
          request_id: _currentHistory[_currentHistory.length - 1]?.id ?? mkId(),
          stage: "action_execute",
          source: "taskpane",
          outcome: "success",
          latency_ms: 0,
          action_type: "analysis",
          action_count: 1,
        });
        appendMsg(mkMsg({ role: "ai", text: response.textResponse }));
        return;
      }

      const { name, args, parsedDetails: details } = response.functionCall;
      const actionData: PendingActionData = { name, args, details };
      const validatedAction = validateAndNormalizeAction(actionData);
      if (!validatedAction.isValid) {
        const validationError = validatedAction.errors.join(" ");
        const recovery = buildRecoverySuggestion("validation", validationError);
        emitTelemetryEvent({
          eventName: "failure",
          request_id: _currentHistory[_currentHistory.length - 1]?.id ?? mkId(),
          stage: "action_execute",
          source: "taskpane",
          outcome: "failure",
          latency_ms: 0,
          action_type: actionData.name,
          error_class: "validation",
          error_message: validationError,
          is_recoverable: false,
        });
        appendMsg(
          mkMsg({
            role: "ai",
            text: `❌ Payload aksi tidak valid: ${validationError}\n\n${recovery.userMessage}`,
            recommendedPrompt: recovery.recommendedPrompt,
            recommendedParams: recovery.recommendedParams,
          })
        );
        return;
      }

      if (validatedAction.warnings.length > 0) {
        const recommendedPrompt = buildRecommendedPromptFromWarnings(validatedAction.warnings);
        const recommendedParams = extractRecommendedParams(validatedAction.warnings);
        appendMsg(
          mkMsg({
            role: "system",
            text: buildNormalizationMessage(validatedAction.warnings),
            actionOutput: `[Validator warnings: ${validatedAction.warnings.length}]`,
            recommendedPrompt: recommendedPrompt ?? undefined,
            recommendedParams: recommendedParams.length > 0 ? recommendedParams : undefined,
          })
        );
      }

      const safeActionData = validatedAction.normalizedAction;
      const previewText =
        safeActionData.details.preview_description ||
        response.textResponse ||
        `Akan menjalankan: **${name}**`;

      if (safeActionData.details.requires_confirmation) {
        // Tunggu konfirmasi user — tambah pesan pending lalu berhenti
        appendMsg(
          mkMsg({
            role: "ai",
            text: previewText,
            isPendingAwaitingConfirmation: true,
            pendingActionData: safeActionData,
          })
        );
      } else {
        // Auto-execute (hanya aksi safe: write_formula, analysis, dll.)
        const autoMsg = mkMsg({
          role: "ai",
          text: previewText,
          isPendingAwaitingConfirmation: false,
          pendingActionData: safeActionData,
        });
        const historyWithAuto = appendMsg(autoMsg);
        await executeAction(autoMsg.id, safeActionData, 0, historyWithAuto);
      }
    },
    [appendMsg, executeAction]
  );

  // ─── handleSend ──────────────────────────────────────────────────────────

  const handleSend = useCallback(async () => {
    if (!prompt.trim() || !isKeySaved || isLoading) return;

    const text = prompt.trim();
    const imagePayload = attachedImage;
    const requestId = createTelemetryRequestId();
    lastUserPromptRef.current = text;
    
    setPrompt("");
    setAttachedImage(null);

    const userMsg = mkMsg({ role: "user", text, imageBase64: imagePayload || undefined });
    const historyWithUser = appendMsg(userMsg);

    emitTelemetryEvent({
      eventName: "request_start",
      request_id: requestId,
      stage: "request_start",
      source: "taskpane",
      outcome: "success",
      latency_ms: 0,
      input_type: imagePayload ? (text ? "text+image" : "image") : "text",
      user_intent_summary: summarizeIntent(text || "image request"),
      model_used: selectedModel,
    });

    setIsLoading(true);
    try {
      const totalStartedAt = performance.now();
      const schemaStartedAt = performance.now();
      const ctxData = await getWorkbookSchema();
      const schemaDuration = performance.now() - schemaStartedAt;
      const schemaStats = getSchemaStats(ctxData);

      emitTelemetryEvent({
        eventName: "schema_done",
        request_id: requestId,
        stage: "schema",
        source: "taskpane",
        outcome: "success",
        latency_ms: schemaDuration,
        model_used: selectedModel,
        schema_latency_ms: schemaDuration,
        sheet_count: schemaStats.sheetCount,
        sample_range_count: schemaStats.sampleRangeCount,
      });

      const aiStartedAt = performance.now();
      const response = await sendAICommand({
        userMessage: text,
        apiKey,
        contextData: ctxData,
        chatHistory: historyWithUser, // FIX #1: snapshot terbaru
        model: selectedModel,
        requestId,
        imageBase64: imagePayload || undefined,
        onRetryAttempt: ({ attempt, maxRetries, reason }) => {
          emitTelemetryEvent({
            eventName: "retry_attempt",
            request_id: requestId,
            stage: "ai_call",
            source: "taskpane",
            outcome: "retry",
            latency_ms: 0,
            model_used: selectedModel,
            retry_attempt_number: attempt,
            max_retry_allowed: maxRetries,
            retry_reason: reason,
          });
        },
      });
      const aiDuration = performance.now() - aiStartedAt;
      const totalDuration = performance.now() - totalStartedAt;

      emitTelemetryEvent({
        eventName: "ai_done",
        request_id: requestId,
        stage: "ai_call",
        source: "taskpane",
        outcome: "success",
        latency_ms: aiDuration,
        model_used: response.meta?.modelUsed ?? selectedModel,
        ai_latency_ms: aiDuration,
      });

      await processAIResponse(response, historyWithUser);

      appendMsg(
        mkMsg({
          role: "system",
          text: `⏱ Schema ${formatMs(schemaDuration)} • AI ${formatMs(aiDuration)} • Total ${formatMs(totalDuration)} • Model ${response.meta?.modelUsed ?? selectedModel} • Attempt ${response.meta?.attempts ?? 1}`,
          actionOutput: `perf:schema=${Math.round(schemaDuration)}ms;ai=${Math.round(aiDuration)}ms;total=${Math.round(totalDuration)}ms;model=${response.meta?.modelUsed ?? selectedModel};attempt=${response.meta?.attempts ?? 1}`,
        })
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      const errorClass = classifyError(msg);
      const recovery = buildRecoverySuggestion(errorClass, msg);
      emitTelemetryEvent({
        eventName: "failure",
        request_id: requestId,
        stage: "ai_call",
        source: "taskpane",
        outcome: "failure",
        latency_ms: 0,
        model_used: selectedModel,
        error_class: errorClass,
        error_message: msg,
      });
      appendMsg(
        mkMsg({
          role: "ai",
          text: `❌ Error: ${msg}\n\n${recovery.userMessage}`,
          recommendedPrompt: recovery.recommendedPrompt,
          recommendedParams: recovery.recommendedParams,
        })
      );
    } finally {
      setIsLoading(false);
    }
  }, [prompt, isKeySaved, isLoading, apiKey, selectedModel, appendMsg, processAIResponse]);

  // ─── handleApproveAction ─────────────────────────────────────────────────

  const handleApproveAction = useCallback(
    (msgId: string, actionData: PendingActionData) => {
      updateMsgById(msgId, { isPendingAwaitingConfirmation: false });
      // FIX #1 & #2: Kirim ref snapshot, tidak ada index aritmatika
      void executeAction(msgId, actionData, 0, messagesRef.current);
    },
    [updateMsgById, executeAction]
  );

  // ─── handleRejectAction ──────────────────────────────────────────────────

  const handleRejectAction = useCallback(
    (msgId: string) => {
      setMessages((prev) =>
        prev.map((m) =>
          m.id === msgId
            ? { ...m, isPendingAwaitingConfirmation: false, text: m.text + "\n\n*(Dibatalkan oleh Pengguna)*" }
            : m
        )
      );
      appendMsg(mkMsg({ role: "user", text: "Batalkan aksi tadi." }));
    },
    [appendMsg]
  );

  // ─── handleUndo ──────────────────────────────────────────────────────────

  const handleUndo = useCallback(async () => {
    try {
      const success = await excelService.undoLastAction();
      if (success) invalidateWorkbookSchemaCache();
      appendMsg(
        mkMsg({
          role: "ai",
          text: success
            ? "↩️ Aksi terakhir berhasil di-undo."
            : "⚠️ Tidak ada aksi yang bisa di-undo.",
        })
      );
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      appendMsg(mkMsg({ role: "ai", text: `❌ Gagal Undo: ${msg}` }));
    }
  }, [appendMsg]);

  // ─── handleSaveKey ───────────────────────────────────────────────────────

  const handleSaveKey = useCallback(() => {
    if (!isKeySaved && apiKey.trim()) {
      const isNvidiaModel = selectedModel.includes("qwen") || selectedModel.startsWith("meta");
      const storageKey = isNvidiaModel ? "nvidia_api_key" : "gemini_api_key";
      localStorage.setItem(storageKey, apiKey.trim());
    }
    setIsKeySaved((v) => !v);
  }, [isKeySaved, apiKey, selectedModel]);

  const filteredTemplates = useMemo(
    () => PROMPT_TEMPLATES.filter((template) => template.level === selectedTemplateLevel),
    [selectedTemplateLevel]
  );

  useEffect(() => {
    if (filteredTemplates.length === 0) {
      setSelectedTemplateId("");
      return;
    }

    const stillExists = filteredTemplates.some((template) => template.id === selectedTemplateId);
    if (!stillExists) {
      setSelectedTemplateId(filteredTemplates[0].id);
    }
  }, [filteredTemplates, selectedTemplateId]);

  const selectedTemplate = filteredTemplates.find((t) => t.id === selectedTemplateId) ?? filteredTemplates[0];

  const handleUseTemplate = useCallback(() => {
    if (!selectedTemplate) return;
    setPrompt(selectedTemplate.prompt);
  }, [selectedTemplate]);

  const handleCopyTemplate = useCallback(async () => {
    if (!selectedTemplate) return;
    try {
      await navigator.clipboard.writeText(selectedTemplate.prompt);
      appendMsg(
        mkMsg({
          role: "system",
          text: `Template disalin: ${selectedTemplate.title}`,
          actionOutput: `[Template copied: ${selectedTemplate.id}]`,
        })
      );
    } catch {
      appendMsg(
        mkMsg({
          role: "ai",
          text: "Gagal menyalin ke clipboard. Gunakan tombol Gunakan ke Input lalu copy manual.",
        })
      );
    }
  }, [appendMsg, selectedTemplate]);

  // ─── Render ───────────────────────────────────────────────────────────────

  return (
    <div className={styles.root}>
      {/* ── Header ── */}
      <header className={styles.header}>
        <Title3 className={styles.title}>Excel × AI Agent</Title3>

        <select
          value={selectedModel}
          onChange={(e) => setSelectedModel(e.target.value as ModelId)}
          style={{ width: "100%", padding: "4px", borderRadius: "4px" }}
          disabled={isLoading}
        >
          {MODEL_OPTIONS.map((opt) => (
            <option key={opt.value} value={opt.value}>
              {opt.label}
            </option>
          ))}
        </select>

        <div className={styles.apiKeyRow}>
          <Input
            type="password"
            placeholder={
              selectedModel.includes("qwen") || selectedModel.startsWith("meta")
                ? "Masukkan NVIDIA API Key..."
                : "Masukkan Gemini API Key..."
            }
            value={apiKey}
            onChange={(_, d) => setApiKey(d.value)}
            disabled={isKeySaved || isLoading}
            contentBefore={<KeyRegular />}
            style={{ flexGrow: 1 }}
            autoComplete="off"
          />
          <Button
            appearance={isKeySaved ? "secondary" : "primary"}
            onClick={handleSaveKey}
            disabled={isLoading}
          >
            {isKeySaved ? "Edit" : "Simpan"}
          </Button>
        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 4 }}>
          <Button
            appearance="subtle"
            size="small"
            onClick={handleUndo}
            disabled={isLoading}
          >
            ↩️ Undo
          </Button>
        </div>
      </header>

      <div className={styles.quickPromptCard}>
        <Text size={300} weight="semibold">Template Prompt (Copy-Paste)</Text>
        <div className={styles.quickPromptSelectors}>
          <select
            value={selectedTemplateLevel}
            onChange={(e) => setSelectedTemplateLevel(e.target.value as TemplateLevel)}
            style={{ width: "45%", padding: "4px", borderRadius: "4px" }}
            disabled={isLoading}
          >
            <option value="Basic">Basic</option>
            <option value="Advanced">Advanced</option>
            <option value="Automation">Automation</option>
          </select>
          <select
            value={selectedTemplateId}
            onChange={(e) => setSelectedTemplateId(e.target.value)}
            style={{ width: "55%", padding: "4px", borderRadius: "4px" }}
            disabled={isLoading}
          >
            {filteredTemplates.map((template) => (
              <option key={template.id} value={template.id}>
                {template.title}
              </option>
            ))}
          </select>
        </div>
        <Text size={200}>{selectedTemplate?.prompt ?? "Pilih template prompt."}</Text>
        <div className={styles.quickPromptActions}>
          <Button appearance="secondary" onClick={handleUseTemplate} disabled={isLoading || !selectedTemplate}>
            Gunakan ke Input
          </Button>
          <Button appearance="primary" onClick={handleCopyTemplate} disabled={isLoading || !selectedTemplate}>
            Copy Prompt
          </Button>
        </div>
      </div>

      {/* ── Chat Area ── */}
      <div className={styles.chatArea}>
        {messages.map((msg) => (
          <div
            key={msg.id}
            className={`${styles.messageRow} ${
              msg.role === "user" ? styles.userRow : styles.aiRow
            }`}
          >
            <div
              className={`${styles.bubble} ${
                msg.role === "user" ? styles.userBubble : styles.aiBubble
              }`}
            >
              <Body1 style={{ whiteSpace: "pre-line" }}>{msg.text}</Body1>

              {msg.recommendedPrompt && (
                <div style={{ marginTop: 8 }}>
                  <Button
                    size="small"
                    appearance="secondary"
                    onClick={() => {
                      const mergedPrompt = msg.recommendedParams && msg.recommendedParams.length > 0
                        ? applyRecommendationsToPrompt(
                            lastUserPromptRef.current || prompt,
                            msg.recommendedParams
                          )
                        : msg.recommendedPrompt!;
                      setPrompt(mergedPrompt);
                      emitTelemetryEvent({
                        eventName: "action_done",
                        request_id: msg.id,
                        stage: "action_execute",
                        source: "taskpane",
                        outcome: "success",
                        latency_ms: 0,
                        action_type: "recovery_prompt_apply",
                        action_count: 1,
                      });
                    }}
                    disabled={isLoading}
                  >
                    Terapkan ke Prompt Terakhir
                  </Button>
                </div>
              )}

              {/* Safety / Dry-Run Confirmation Panel */}
              {msg.isPendingAwaitingConfirmation && msg.pendingActionData && (
                <ConfirmationPanel
                  actionData={msg.pendingActionData}
                  isLoading={isLoading}
                  onApprove={() => handleApproveAction(msg.id, msg.pendingActionData!)}
                  onReject={() => handleRejectAction(msg.id)}
                />
              )}
            </div>

            {msg.actionOutput && (
              <span className={styles.systemText}>{msg.actionOutput}</span>
            )}
          </div>
        ))}

        {isLoading && (
          <div className={`${styles.messageRow} ${styles.aiRow}`}>
            <Spinner size="tiny" label="AI sedang berpikir..." />
          </div>
        )}

        {/* Anchor auto-scroll */}
        <div ref={chatBottomRef} />
      </div>

      {/* ── Input Area ── */}
      <div style={{ padding: "0 16px" }}>
          {attachedImage && (
             <div style={{ display: 'inline-block', position: 'relative', marginBottom: '8px' }}>
                <img src={attachedImage} alt="attachment" style={{ height: '60px', borderRadius: '4px', border: '1px solid #ddd' }} />
                <Button appearance="subtle" size="small" onClick={() => setAttachedImage(null)} style={{ position: 'absolute', top: -5, right: -5, background: 'white', borderRadius: '50%', minWidth: '20px', padding: 0 }}>❌</Button>
             </div>
          )}
      </div>
      <div className={styles.inputArea}>
        <input 
          type="file" 
          accept="image/*" 
          id="image-attach-input" 
          style={{ display: 'none' }} 
          onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) {
                  const reader = new FileReader();
                  reader.onload = (ev) => setAttachedImage(ev.target?.result as string);
                  reader.readAsDataURL(file);
              }
          }}
        />
        <Button 
          icon={<DocumentRegular />} 
          appearance="subtle" 
          onClick={() => document.getElementById("image-attach-input")?.click()} 
          title="Unggah Gambar / Faktur"
        />
        <Input
          className={styles.inputField}
          placeholder={
            isKeySaved ? "Ketik pesan atau unggah gambar..." : "Simpan API Key terlebih dahulu"
          }
          value={prompt}
          onChange={(_, d) => setPrompt(d.value)}
          disabled={!isKeySaved || isLoading}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) handleSend();
          }}
        />
        <Button
          icon={<SendRegular />}
          appearance="primary"
          onClick={handleSend}
          disabled={!isKeySaved || isLoading || (!prompt.trim() && !attachedImage)}
        />
      </div>
    </div>
  );
};

// ─── ConfirmationPanel (sub-component terpisah) ───────────────────────────────
// Dipisahkan agar re-render App tidak memicu ulang seluruh panel

interface ConfirmationPanelProps {
  actionData: PendingActionData;
  isLoading: boolean;
  onApprove: () => void;
  onReject: () => void;
}

const ConfirmationPanel: React.FC<ConfirmationPanelProps> = React.memo(
  ({ actionData, isLoading, onApprove, onReject }) => {
    const destructive = isDestructiveAction(actionData);

    return (
      <div
        style={{
          marginTop: 14,
          padding: 12,
          background: destructive ? "#ffebee" : "#e8f5e9",
          borderRadius: 6,
          border: `1px solid ${destructive ? "#ffcdd2" : "#c8e6c9"}`,
        }}
      >
        <Text
          size={300}
          weight="semibold"
          style={{
            display: "block",
            marginBottom: 8,
            color: destructive ? "#c62828" : "#2e7d32",
          }}
        >
          {destructive
            ? "⚠️ Peringatan: Operasi Destruktif"
            : "✅ Konfirmasi Tindakan AI"}
        </Text>

        <div style={{ marginBottom: 10, fontSize: 14, lineHeight: 1.5, color: "#333" }}>
          {actionData.details.preview_description ||
            "Membutuhkan konfirmasi untuk mengeksekusi aksi ini."}
        </div>

        <details style={{ marginBottom: 10 }}>
          <summary style={{ fontSize: 12, cursor: "pointer", color: "#666" }}>
            Detail Teknis (JSON)
          </summary>
          <pre
            style={{
              fontSize: 11,
              whiteSpace: "pre-wrap",
              maxHeight: 150,
              overflowY: "auto",
              margin: "8px 0 0",
              padding: 8,
              background: "rgba(255,255,255,0.8)",
              borderRadius: 4,
              border: "1px solid #e0e0e0",
            }}
          >
            {`Tipe: ${actionData.name}\nArgs: ${JSON.stringify(actionData.args, null, 2)}`}
          </pre>
        </details>

        <div style={{ display: "flex", gap: 8 }}>
          <Button
            appearance="primary"
            style={destructive ? { backgroundColor: "#d32f2f" } : undefined}
            onClick={onApprove}
            disabled={isLoading}
          >
            Ya, Jalankan
          </Button>
          <Button appearance="secondary" onClick={onReject} disabled={isLoading}>
            Tolak
          </Button>
        </div>
      </div>
    );
  }
);

export default App;