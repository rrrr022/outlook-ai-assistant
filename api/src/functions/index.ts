import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import OpenAI from "openai";

// Import other function modules to register them
import "./chat";
import "./health";

// FreedomForged AI for Outlook - Backend API
// Version: 1.1.1 - Added BYOK support for GitHub Models (CORS bypass)
// Last updated: Deployment trigger

// Initialize OpenAI client for GitHub Models (server key mode)
const openai = new OpenAI({
    baseURL: "https://models.inference.ai.azure.com",
    apiKey: process.env.GITHUB_TOKEN || ""
});

interface ChatMessage {
    role: "user" | "assistant" | "system";
    content: string;
}

interface ChatRequest {
    message: string;
    context?: {
        subject?: string;
        sender?: string;
        body?: string;
    };
    history?: ChatMessage[];
}

// CORS headers
const corsHeaders = {
    "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json"
};

// Health check endpoint
app.http("health", {
    methods: ["GET", "OPTIONS"],
    authLevel: "anonymous",
    route: "health",
    handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
        if (request.method === "OPTIONS") {
            return { status: 204, headers: corsHeaders };
        }
        
        return {
            status: 200,
            headers: corsHeaders,
            body: JSON.stringify({
                status: "healthy",
                timestamp: new Date().toISOString(),
                version: "1.0.0"
            })
        };
    }
});

// Main chat endpoint (server-key mode)
app.http("main-chat", {
    methods: ["POST", "OPTIONS"],
    authLevel: "anonymous",
    route: "ai/chat",
    handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
        // Handle CORS preflight
        if (request.method === "OPTIONS") {
            return { status: 204, headers: corsHeaders };
        }

        try {
            const body = await request.json() as ChatRequest;
            const { message, context: emailContext, history = [] } = body;

            if (!message) {
                return {
                    status: 400,
                    headers: corsHeaders,
                    body: JSON.stringify({ error: "Message is required" })
                };
            }

            // Build system prompt
            let systemPrompt = `You are an AI assistant integrated into Microsoft Outlook. You help users manage their emails, calendar, and tasks efficiently.

Your capabilities:
- Summarize emails
- Draft email responses
- Suggest actions based on email content
- Help organize and prioritize emails
- Provide calendar suggestions
- Help manage tasks

Be concise, professional, and helpful. Format responses with markdown when appropriate.`;

            // Add email context if available
            if (emailContext) {
                systemPrompt += `\n\nCurrent Email Context:`;
                if (emailContext.subject) systemPrompt += `\n- Subject: ${emailContext.subject}`;
                if (emailContext.sender) systemPrompt += `\n- From: ${emailContext.sender}`;
                if (emailContext.body) systemPrompt += `\n- Body Preview: ${emailContext.body.substring(0, 500)}...`;
            }

            // Build messages array
            const messages: ChatMessage[] = [
                { role: "system", content: systemPrompt },
                ...history.slice(-10), // Keep last 10 messages for context
                { role: "user", content: message }
            ];

            // Call GitHub Models API
            const completion = await openai.chat.completions.create({
                model: process.env.GITHUB_MODEL || "gpt-4o",
                messages: messages,
                max_tokens: 1000,
                temperature: 0.7
            });

            const response = completion.choices[0]?.message?.content || "I apologize, but I couldn't generate a response.";

            // Generate suggestions based on the context
            const suggestions = generateSuggestions(message, emailContext);

            return {
                status: 200,
                headers: corsHeaders,
                body: JSON.stringify({
                    response,
                    suggestions
                })
            };

        } catch (error: any) {
            context.log("Chat error:", error);
            
            return {
                status: 500,
                headers: corsHeaders,
                body: JSON.stringify({
                    error: "Failed to process chat request",
                    details: error.message
                })
            };
        }
    }
});

// Email actions endpoint
app.http("email-actions", {
    methods: ["POST", "OPTIONS"],
    authLevel: "anonymous",
    route: "email/actions",
    handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
        if (request.method === "OPTIONS") {
            return { status: 204, headers: corsHeaders };
        }

        try {
            const body = await request.json() as { action: string; emailContent: any };
            const { action, emailContent } = body;

            let prompt = "";
            switch (action) {
                case "summarize":
                    prompt = `Summarize this email concisely:\n\nSubject: ${emailContent.subject}\nFrom: ${emailContent.sender}\n\n${emailContent.body}`;
                    break;
                case "draft_reply":
                    prompt = `Draft a professional reply to this email:\n\nSubject: ${emailContent.subject}\nFrom: ${emailContent.sender}\n\n${emailContent.body}`;
                    break;
                case "extract_tasks":
                    prompt = `Extract any action items or tasks from this email:\n\nSubject: ${emailContent.subject}\nFrom: ${emailContent.sender}\n\n${emailContent.body}`;
                    break;
                case "analyze_sentiment":
                    prompt = `Analyze the tone and sentiment of this email:\n\nSubject: ${emailContent.subject}\nFrom: ${emailContent.sender}\n\n${emailContent.body}`;
                    break;
                default:
                    return {
                        status: 400,
                        headers: corsHeaders,
                        body: JSON.stringify({ error: "Unknown action" })
                    };
            }

            const completion = await openai.chat.completions.create({
                model: process.env.GITHUB_MODEL || "gpt-4o",
                messages: [
                    { role: "system", content: "You are a helpful email assistant. Be concise and professional." },
                    { role: "user", content: prompt }
                ],
                max_tokens: 800,
                temperature: 0.7
            });

            return {
                status: 200,
                headers: corsHeaders,
                body: JSON.stringify({
                    result: completion.choices[0]?.message?.content || "Unable to process request",
                    action
                })
            };

        } catch (error: any) {
            context.log("Email action error:", error);
            return {
                status: 500,
                headers: corsHeaders,
                body: JSON.stringify({ error: "Failed to process email action" })
            };
        }
    }
});

// Helper function to generate suggestions
function generateSuggestions(message: string, context?: ChatRequest["context"]): string[] {
    const suggestions: string[] = [];
    const lowerMessage = message.toLowerCase();

    if (context?.body) {
        suggestions.push("📝 Summarize this email");
        suggestions.push("✉️ Draft a reply");
        suggestions.push("✅ Extract action items");
    }

    if (lowerMessage.includes("meeting") || lowerMessage.includes("schedule")) {
        suggestions.push("📅 Check my calendar");
        suggestions.push("🕐 Find available time slots");
    }

    if (lowerMessage.includes("task") || lowerMessage.includes("todo")) {
        suggestions.push("📋 Show my tasks");
        suggestions.push("➕ Create a new task");
    }

    return suggestions.slice(0, 4);
}

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
                    response,
                    model: modelToUse,
                    provider
                })
            };

        } catch (error: any) {
            context.log("BYOK Chat error:", error);
            
            // Parse error message for better user feedback
            let errorMessage = error.message || "Connection failed";
            
            if (errorMessage.includes("401") || errorMessage.includes("Unauthorized")) {
                errorMessage = "Invalid API key. Please check your key and try again.";
            } else if (errorMessage.includes("403") || errorMessage.includes("Forbidden")) {
                errorMessage = "Access denied. Your API key may not have access to this model.";
            } else if (errorMessage.includes("429")) {
                errorMessage = "Rate limited. Please wait a moment and try again.";
            } else if (errorMessage.includes("404")) {
                errorMessage = "Model not found. Please select a different model.";
            }
            
            return {
                status: 500,
                headers: corsHeaders,
                body: JSON.stringify({
                    error: errorMessage,
                    details: error.message
                })
            };
        }
    }
});
