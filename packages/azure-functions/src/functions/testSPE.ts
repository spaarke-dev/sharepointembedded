import { app, HttpRequest, HttpResponseInit, InvocationContext } from "@azure/functions";
import { ConfidentialClientApplication } from "@azure/msal-node";

const msalConfig = {
    auth: {
        clientId: process.env.APP_CLIENT_ID!,
        clientSecret: process.env.APP_CLIENT_SECRET!,
        authority: process.env.APP_AUTHORITY!
    }
};

const cca = new ConfidentialClientApplication(msalConfig);

// Test auth flow - simplified to just verify tokens work
export async function testListContainerTypes(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const authHeader = request.headers.get("authorization");
        if (!authHeader) {
            return { 
                status: 401, 
                body: JSON.stringify({ error: "No authorization header" })
            };
        }
        
        // Just verify we got a token and can get an app token
        const appTokenResponse = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        return {
            status: 200,
            body: JSON.stringify({
                test: "List Container Types",
                success: true,
                message: "Auth flow working",
                hasUserToken: true,
                hasAppToken: !!appTokenResponse.accessToken,
                tokenLength: appTokenResponse.accessToken?.length,
                containerTypeId: process.env.APP_CONTAINER_TYPE_ID
            })
        };
    } catch (error: any) {
        return {
            status: 500,
            body: JSON.stringify({
                test: "List Container Types",
                success: false,
                error: error.message
            })
        };
    }
}

// Debug Graph - test what we can actually access
export async function debugGraph(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        // Use app-only token
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const tests = [];
        
        // Test 1: Basic Graph access
        const usersResponse = await fetch("https://graph.microsoft.com/v1.0/users?$top=1", {
            headers: { "Authorization": `Bearer ${response.accessToken}` }
        });
        tests.push({
            endpoint: "/users",
            status: usersResponse.status,
            success: usersResponse.ok
        });
        
        // Test 2: Try container endpoint
        const containersResponse = await fetch(
            `https://graph.microsoft.com/beta/storage/fileStorage/containers?$filter=containerTypeId eq ${process.env.APP_CONTAINER_TYPE_ID}`,
            {
                headers: { "Authorization": `Bearer ${response.accessToken}` }
            }
        );
        
        let containerData;
        try {
            containerData = await containersResponse.json();
        } catch {
            containerData = await containersResponse.text();
        }
        
        tests.push({
            endpoint: "/storage/fileStorage/containers",
            status: containersResponse.status,
            success: containersResponse.ok,
            data: containerData
        });
        
        return {
            status: 200,
            body: JSON.stringify({
                success: true,
                tokenScopes: response.scopes,
                tests: tests
            })
        };
    } catch (error: any) {
        return {
            status: 500,
            body: JSON.stringify({
                success: false,
                error: error.message
            })
        };
    }
}

// Keep other test functions...
export async function testListContainers(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    return {
        status: 200,
        body: JSON.stringify({
            test: "List Containers",
            success: true,
            message: "Auth header received",
            containerTypeId: process.env.APP_CONTAINER_TYPE_ID,
            hasAuth: !!request.headers.get("authorization")
        })
    };
}

export async function testCreateContainer(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    return {
        status: 200,
        body: JSON.stringify({
            test: "Create Container",
            success: true,
            config: {
                hasClientId: !!process.env.APP_CLIENT_ID,
                hasClientSecret: !!process.env.APP_CLIENT_SECRET,
                hasAuthority: !!process.env.APP_AUTHORITY,
                hasContainerTypeId: !!process.env.APP_CONTAINER_TYPE_ID,
                containerTypeId: process.env.APP_CONTAINER_TYPE_ID
            }
        })
    };
}

// Register functions
app.http("testListContainerTypes", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: testListContainerTypes
});

app.http("testListContainers", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: testListContainers
});

app.http("testCreateContainer", {
    methods: ["POST"],
    authLevel: "anonymous",
    handler: testCreateContainer
});

app.http("debugGraph", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: debugGraph
});