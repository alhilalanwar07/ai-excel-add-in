import express from 'express';
import cors from 'cors';
import axios from 'axios';

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));

// Proxy endpoint untuk Gemini API
app.post('/api/gemini/generate', async (req, res) => {
    try {
        const { apiKey, model, messages, tools } = req.body;
        
        if (!apiKey) {
            res.status(400).json({ error: "Gemini API Key is required" });
            return;
        }

        const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
        
        // Memformat payload array messages (OpenAI style) ke format native Gemini (content.parts)
        const geminiContents = messages.map((m: any) => {
             return {
                 role: m.role === 'assistant' ? 'model' : 'user',
                 parts: [{ text: m.content }]
             }
        });
        
        // Memformat Tools
        const geminiTools = tools ? [{
             functionDeclarations: tools.map((t: any) => t.function || t)
        }] : undefined;

        const payload = {
            contents: geminiContents,
            tools: geminiTools
        };

        const response = await axios.post(url, payload, {
            headers: { "Content-Type": "application/json" },
            timeout: 120000 // 120 seconds timeout protection for large prompts
        });

        // Mapping respons format Native Gemini ke OpenAI Chat Completions Style
        const candidate = response.data.candidates?.[0];
        const functionCallObj = candidate?.content?.parts?.find((p: any) => p.functionCall)?.functionCall;
        const textObj = candidate?.content?.parts?.find((p: any) => p.text)?.text;

        const mappedResponse = {
             choices: [
                  {
                       message: {
                            role: "assistant",
                            content: textObj || "",
                            tool_calls: functionCallObj ? [
                                 {
                                     type: "function",
                                     function: {
                                          name: functionCallObj.name,
                                          arguments: JSON.stringify(functionCallObj.args)
                                     }
                                 }
                            ] : undefined
                       }
                  }
             ]
        };

        res.json(mappedResponse);
    } catch (error: any) {
        // Fallback untuk Rate Limit Gemini
        if (error.response?.status === 429 || error.response?.data?.error?.status === "RESOURCE_EXHAUSTED") {
            const errorMsg = "Gemini API Limit (429): Quota habis atau terlalu sering meminta dalam 1 menit. Harap gunakan model NVIDIA Qwen untuk sementara, atau tingkatkan langganan Anda.";
            console.error(errorMsg);
            
            // Kita bungkus sebagai response valid agar ditangkap catch loop di UI tanpa ngerusak parsing Axios
            res.status(429).json({ error: errorMsg });
            return;
        }

        console.error("Gemini Proxy Error:", error.response?.data || error.message);
        res.status(500).json({ error: error.response?.data || error.message });
    }
});

// Proxy endpoint untuk NVIDIA API (Qwen dll)
app.post('/api/nvidia/generate', async (req, res) => {
    try {
        const { apiKey, model, messages, tools } = req.body;
        
        if (!apiKey) {
            res.status(400).json({ error: "NVIDIA API Key is required" });
            return;
        }

        const url = 'https://integrate.api.nvidia.com/v1/chat/completions';
        const payload = { model, messages, tools, temperature: 0.1, max_tokens: 3000 };

        const response = await axios.post(url, payload, {
            headers: { 
                "Content-Type": "application/json",
                "Authorization": `Bearer ${apiKey}`
            },
            timeout: 120000 // 120 seconds timeout consideration
        });

        res.json(response.data);
    } catch (error: any) {
        console.error("NVIDIA Proxy Error:", error.response?.data || error.message);
        res.status(500).json({ error: error.response?.data || error.message });
    }
});

const PORT = 3001;
app.listen(PORT, () => {
    console.log(`Backend Proxy Server is running on http://localhost:${PORT}`);
});
