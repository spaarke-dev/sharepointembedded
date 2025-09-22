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

// Create or verify container type
export async function createContainerType(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        context.log("Creating/verifying container type...");
        
        // Get app-only token
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const containerTypeId = process.env.APP_CONTAINER_TYPE_ID;
        
        // First, try to get the container type to see if it exists
        const getUrl = `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes/${containerTypeId}`;
        context.log("Checking if container type exists:", getUrl);
        
        const checkResponse = await fetch(getUrl, {
            headers: {
                "Authorization": `Bearer ${response.accessToken}`
            }
        });
        
        if (checkResponse.ok) {
            const existingType = await checkResponse.json();
            return {
                status: 200,
                body: JSON.stringify({
                    success: true,
                    message: "Container type already exists",
                    data: existingType
                })
            };
        }
        
        // If it doesn't exist, create it
        context.log("Container type not found, attempting to create...");
        
        const createUrl = "https://graph.microsoft.com/beta/storage/fileStorage/containerTypes";
        const containerTypeData = {
            containerTypeId: containerTypeId,
            displayName: "SPE Test Container Type",
            description: "Container type for SharePoint Embedded testing",
            owningApplicationId: process.env.APP_CLIENT_ID
        };
        
        const createResponse = await fetch(createUrl, {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${response.accessToken}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(containerTypeData)
        });
        
        const result = await createResponse.text();
        let jsonResult;
        try {
            jsonResult = JSON.parse(result);
        } catch {
            jsonResult = { rawResponse: result };
        }
        
        return {
            status: createResponse.ok ? 200 : createResponse.status,
            body: JSON.stringify({
                success: createResponse.ok,
                status: createResponse.status,
                data: jsonResult
            })
        };
    } catch (error: any) {
        context.error("Error with container type:", error);
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

// Get container type permissions
export async function getContainerTypePermissions(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const containerTypeId = process.env.APP_CONTAINER_TYPE_ID;
        const url = `https://graph.microsoft.com/beta/storage/fileStorage/containerTypes/${containerTypeId}/permissions`;
        
        const permResponse = await fetch(url, {
            headers: {
                "Authorization": `Bearer ${response.accessToken}`
            }
        });
        
        const result = await permResponse.json();
        
        return {
            status: 200,
            body: JSON.stringify({
                success: permResponse.ok,
                data: result
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

// Register functions
app.http("createContainerType", {
    methods: ["POST"],
    authLevel: "anonymous",
    handler: createContainerType
});

app.http("getContainerTypePermissions", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: getContainerTypePermissions
});