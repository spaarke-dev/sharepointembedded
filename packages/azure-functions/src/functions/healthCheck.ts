import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";

export async function healthCheck(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    context.log("Health check function processed a request.");
    
    // Simple health check - no auth required
    return {
        status: 200,
        headers: {
            "Content-Type": "application/json"
        },
        body: JSON.stringify({
            status: "healthy",
            timestamp: new Date().toISOString(),
            containerTypeId: process.env.APP_CONTAINER_TYPE_ID
        })
    };
}

app.http("healthCheck", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: healthCheck
});