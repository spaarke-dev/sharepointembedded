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

// List containers - using app-only auth instead of OBO
export async function listContainers(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        context.log("Listing containers using app-only auth...");
        
        // Use client credentials (app-only) instead of OBO
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const url = `https://graph.microsoft.com/beta/storage/fileStorage/containers?$filter=containerTypeId eq ${process.env.APP_CONTAINER_TYPE_ID}`;
        context.log("Calling:", url);
        
        const graphResponse = await fetch(url, {
            headers: {
                "Authorization": `Bearer ${response.accessToken}`
            }
        });
        
        const result = await graphResponse.text();
        let jsonResult;
        try {
            jsonResult = JSON.parse(result);
        } catch {
            jsonResult = { rawResponse: result };
        }
        
        return {
            status: 200,
            body: JSON.stringify({
                success: graphResponse.ok,
                status: graphResponse.status,
                data: jsonResult
            })
        };
    } catch (error: any) {
        context.error("Error listing containers:", error);
        return {
            status: 500,
            body: JSON.stringify({
                success: false,
                error: error.message
            })
        };
    }
}

// Create container - using app-only auth
export async function createContainer(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        context.log("Creating container using app-only auth...");
        
        // Use client credentials (app-only) instead of OBO
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const containerData = {
            displayName: `Test Container ${new Date().getTime()}`,
            description: "Created via SPE test",
            containerTypeId: process.env.APP_CONTAINER_TYPE_ID
        };
        
        context.log("Container payload:", containerData);
        
        const graphResponse = await fetch("https://graph.microsoft.com/beta/storage/fileStorage/containers", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${response.accessToken}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(containerData)
        });
        
        const result = await graphResponse.text();
        let jsonResult;
        try {
            jsonResult = JSON.parse(result);
        } catch {
            jsonResult = { rawResponse: result };
        }
        
        context.log("Graph response status:", graphResponse.status);
        context.log("Graph response:", jsonResult);
        
        return {
            status: 200,
            body: JSON.stringify({
                success: graphResponse.ok,
                status: graphResponse.status,
                data: jsonResult,
                containerTypeId: process.env.APP_CONTAINER_TYPE_ID
            })
        };
    } catch (error: any) {
        context.error("Error creating container:", error);
        return {
            status: 500,
            body: JSON.stringify({
                success: false,
                error: error.message,
                stack: error.stack
            })
        };
    }
}

// Keep upload file as before (it might need the user context)
export async function uploadFile(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        context.log("Uploading file...");
        
        const body = await request.json() as any;
        const containerId = body?.containerId;
        
        if (!containerId) {
            return { 
                status: 400, 
                body: JSON.stringify({ error: "containerId required in request body" }) 
            };
        }
        
        // Use app-only auth
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const fileName = `test-${new Date().getTime()}.txt`;
        const fileContent = "This is a test file uploaded via SPE API";
        
        const uploadUrl = `https://graph.microsoft.com/v1.0/drives/${containerId}/root:/${fileName}:/content`;
        
        const graphResponse = await fetch(uploadUrl, {
            method: "PUT",
            headers: {
                "Authorization": `Bearer ${response.accessToken}`,
                "Content-Type": "text/plain"
            },
            body: fileContent
        });
        
        const result = await graphResponse.text();
        let jsonResult;
        try {
            jsonResult = JSON.parse(result);
        } catch {
            jsonResult = { rawResponse: result };
        }
        
        return {
            status: 200,
            body: JSON.stringify({
                success: graphResponse.ok,
                status: graphResponse.status,
                fileName: fileName,
                data: jsonResult
            })
        };
    } catch (error: any) {
        context.error("Error uploading file:", error);
        return {
            status: 500,
            body: JSON.stringify({
                success: false,
                error: error.message
            })
        };
    }
}

// Register functions
app.http("createContainer", {
    methods: ["POST"],
    authLevel: "anonymous",
    handler: createContainer
});

app.http("listContainers", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: listContainers
});

app.http("uploadFile", {
    methods: ["POST"],
    authLevel: "anonymous",
    handler: uploadFile
});