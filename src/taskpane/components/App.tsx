import * as React from "react";
import { useState, useEffect, useRef, useCallback } from "react";
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
import { getWorkbookSchema, sendAICommand } from "../utils/gemini";
import { excelService } from "../services/ExcelService";
import type { AIMasterPayload, ActionType } from "../utils/gemini";

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
        await excelService.executeAction(
          actionData.name,
          actionData.args,
          actionData.details
        );

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
            });

            if (fixResp.functionCall) {
              const fixedAction: PendingActionData = {
                name: fixResp.functionCall.name,
                args: fixResp.functionCall.args,
                details: fixResp.functionCall.parsedDetails,
              };
              await executeAction(msgId, fixedAction, retryCount + 1, messagesRef.current);
            } else {
              appendMsg(mkMsg({ role: "ai", text: "❌ AI tidak menemukan solusi perbaikan." }));
            }
          } catch (fixErr) {
            const fixMsg = fixErr instanceof Error ? fixErr.message : String(fixErr);
            appendMsg(mkMsg({ role: "ai", text: `❌ Gagal menghubungi API saat perbaikan: ${fixMsg}` }));
          }
        } else {
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
        appendMsg(mkMsg({ role: "ai", text: response.textResponse }));
        return;
      }

      const { name, args, parsedDetails: details } = response.functionCall;
      const actionData: PendingActionData = { name, args, details };
      const previewText =
        details.preview_description ||
        response.textResponse ||
        `Akan menjalankan: **${name}**`;

      if (details.requires_confirmation) {
        // Tunggu konfirmasi user — tambah pesan pending lalu berhenti
        appendMsg(
          mkMsg({
            role: "ai",
            text: previewText,
            isPendingAwaitingConfirmation: true,
            pendingActionData: actionData,
          })
        );
      } else {
        // Auto-execute (hanya aksi safe: write_formula, analysis, dll.)
        const autoMsg = mkMsg({
          role: "ai",
          text: previewText,
          isPendingAwaitingConfirmation: false,
          pendingActionData: actionData,
        });
        const historyWithAuto = appendMsg(autoMsg);
        await executeAction(autoMsg.id, actionData, 0, historyWithAuto);
      }
    },
    [appendMsg, executeAction]
  );

  // ─── handleSend ──────────────────────────────────────────────────────────

  const handleSend = useCallback(async () => {
    if (!prompt.trim() || !isKeySaved || isLoading) return;

    const text = prompt.trim();
    const imagePayload = attachedImage;
    
    setPrompt("");
    setAttachedImage(null);

    const userMsg = mkMsg({ role: "user", text, imageBase64: imagePayload || undefined });
    const historyWithUser = appendMsg(userMsg);

    setIsLoading(true);
    try {
      const ctxData = await getWorkbookSchema();
      const response = await sendAICommand({
        userMessage: text,
        apiKey,
        contextData: ctxData,
        chatHistory: historyWithUser, // FIX #1: snapshot terbaru
        model: selectedModel,
        imageBase64: imagePayload || undefined
      });

      await processAIResponse(response, historyWithUser);
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      appendMsg(mkMsg({ role: "ai", text: `❌ Error: ${msg}` }));
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