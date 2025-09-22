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

// Direct test of container creation to verify registration
export async function testRegistrationStatus(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        context.log("Testing container type registration status...");
        
        // Get app-only token
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const containerTypeId = process.env.APP_CONTAINER_TYPE_ID;
        const timestamp = Date.now();
        
        // Test 1: Create a container (the ultimate test)
        const containerData = {
            displayName: `Registration Test ${timestamp}`,
            description: "Testing if registration worked",
            containerTypeId: containerTypeId
        };
        
        context.log("Attempting to create container with type:", containerTypeId);
        
        const createResponse = await fetch("https://graph.microsoft.com/beta/storage/fileStorage/containers", {
            method: "POST",
            headers: {
                "Authorization": `Bearer ${response.accessToken}`,
                "Content-Type": "application/json"
            },
            body: JSON.stringify(containerData)
        });
        
        const createResult = await createResponse.json();
        
        if (createResponse.ok) {
            // Registration worked! Clean up the test container
            const containerId = createResult.id;
            
            // Delete the test container
            await fetch(`https://graph.microsoft.com/beta/storage/fileStorage/containers/${containerId}`, {
                method: "DELETE",
                headers: {
                    "Authorization": `Bearer ${response.accessToken}`
                }
            });
            
            return {
                status: 200,
                body: JSON.stringify({
                    registrationStatus: "✅ CONFIRMED - Registration successful!",
                    success: true,
                    message: "Container type is properly registered. Your app can now create and manage containers.",
                    testContainerCreated: true,
                    testContainerDeleted: true,
                    details: {
                        containerTypeId: containerTypeId,
                        appId: process.env.APP_CLIENT_ID,
                        timestamp: new Date().toISOString()
                    }
                })
            };
        } else {
            // Check the error to determine registration status
            const errorCode = createResult.error?.code;
            const errorMessage = createResult.error?.message;
            
            let registrationStatus = "❌ NOT REGISTERED";
            let recommendation = "";
            
            if (errorCode === "accessDenied") {
                registrationStatus = "❌ NOT REGISTERED - Access Denied";
                recommendation = "Run the registration script again or wait for propagation.";
            } else if (errorMessage?.includes("container type")) {
                registrationStatus = "❌ Container Type Issue";
                recommendation = "Container type may not exist or is misconfigured.";
            } else {
                registrationStatus = "⚠️ UNKNOWN STATUS";
                recommendation = "Check the error details below.";
            }
            
            return {
                status: 200,
                body: JSON.stringify({
                    registrationStatus: registrationStatus,
                    success: false,
                    message: "Container creation failed - registration may not be complete.",
                    recommendation: recommendation,
                    error: createResult.error,
                    details: {
                        containerTypeId: containerTypeId,
                        appId: process.env.APP_CLIENT_ID,
                        httpStatus: createResponse.status,
                        timestamp: new Date().toISOString()
                    }
                })
            };
        }
    } catch (error: any) {
        context.error("Error testing registration:", error);
        return {
            status: 500,
            body: JSON.stringify({
                registrationStatus: "❌ TEST FAILED",
                success: false,
                error: error.message,
                stack: error.stack
            })
        };
    }
}

// Register function
app.http("testRegistrationStatus", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: testRegistrationStatus
});