import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import OpenAI from "openai";

// FreedomForged AI for Outlook - BYOK Chat Endpoint
// Version: 1.1.1 - Routes user API keys through server to avoid CORS

interface ChatMessage {
    role: "user" | "assistant" | "system";
    content: string;
}

const corsHeaders = {
    "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json"
};

// BYOK Chat endpoint - Uses user's API key (routes through server to avoid CORS)
app.http("byok-chat", {
    methods: ["POST", "OPTIONS"],
    authLevel: "anonymous",
    route: "chat",
    handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
        // Handle CORS preflight
        if (request.method === "OPTIONS") {
            return { status: 204, headers: corsHeaders };
        }

        try {
            const body = await request.json() as {
                provider: string;
                apiKey: string;
                model: string;
                messages: ChatMessage[];
            };
            
            const { provider, apiKey, model, messages } = body;

            if (!apiKey) {
                return {
                    status: 400,
                    headers: corsHeaders,
                    body: JSON.stringify({ error: "API key is required" })
                };
            }

            if (!messages || messages.length === 0) {
                return {
                    status: 400,
                    headers: corsHeaders,
                    body: JSON.stringify({ error: "Messages are required" })
                };
            }

            // Create OpenAI client with user's key
            let client: OpenAI;
            let modelToUse = model;

            if (provider === 'github') {
                client = new OpenAI({
                    baseURL: "https://models.inference.ai.azure.com",
                    apiKey: apiKey
                });
            } else if (provider === 'openai') {
                client = new OpenAI({
                    apiKey: apiKey
                });
            } else if (provider === 'xai') {
                client = new OpenAI({
                    baseURL: "https://api.x.ai/v1",
                    apiKey: apiKey
                });
            } else {
                return {
                    status: 400,
                    headers: corsHeaders,
                    body: JSON.stringify({ error: `Unsupported provider: ${provider}` })
                };
            }

            context.log(`BYOK request: provider=${provider}, model=${modelToUse}`);

            const completion = await client.chat.completions.create({
                model: modelToUse,
                messages: messages,
                max_tokens: 1000,
                temperature: 0.7
            });

            const response = completion.choices[0]?.message?.content || "No response generated";

            return {
                status: 200,
                headers: corsHeaders,
                body: JSON.stringify({
                    message: response,
                    model: modelToUse,
                    provider: provider,
                    usage: completion.usage
                })
            };
        } catch (error: any) {
            context.error("BYOK Chat error:", error);
            
            // Return appropriate error message
            let errorMessage = "Failed to process request";
            let statusCode = 500;
            
            if (error.status === 401 || error.message?.includes("401") || error.message?.includes("Unauthorized")) {
                errorMessage = "Invalid API key";
                statusCode = 401;
            } else if (error.status === 429) {
                errorMessage = "Rate limit exceeded";
                statusCode = 429;
            } else if (error.message) {
                errorMessage = error.message;
            }

            return {
                status: statusCode,
                headers: corsHeaders,
                body: JSON.stringify({ error: errorMessage })
            };
        }
    }
});
