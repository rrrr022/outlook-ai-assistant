import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";

const corsHeaders = {
    "Access-Control-Allow-Origin": process.env.FRONTEND_URL || "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization",
    "Content-Type": "application/json"
};

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
                version: "1.1.1"
            })
        };
    }
});
