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

// Comprehensive container type verification
export async function verifyContainerType(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        context.log("Verifying container type registration...");
        
        // Get app-only token
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        const containerTypeId = process.env.APP_CONTAINER_TYPE_ID;
        const results: any = {
            containerTypeId: containerTypeId,
            clientId: process.env.APP_CLIENT_ID,
            timestamp: new Date().toISOString(),
            checks: []
        };
        
        // Test 1: Try to list containers with this container type
        context.log("Test 1: Listing containers with container type filter...");
        // Fix: Remove quotes around GUID in filter
        const listUrl = `https://graph.microsoft.com/beta/storage/fileStorage/containers?$filter=containerTypeId eq ${containerTypeId}`;
        
        try {
            const listResponse = await fetch(listUrl, {
                headers: {
                    "Authorization": `Bearer ${response.accessToken}`,
                    "ConsistencyLevel": "eventual"
                }
            });
            
            const listData = await listResponse.json();
            results.checks.push({
                test: "List containers with containerTypeId filter",
                success: listResponse.ok,
                status: listResponse.status,
                message: listResponse.ok ? "Can query containers with this type" : "Cannot query containers",
                data: listData
            });
            
            if (listResponse.ok && listData.value) {
                results.containerCount = listData.value.length;
                results.existingContainers = listData.value.map((c: any) => ({
                    id: c.id,
                    displayName: c.displayName,
                    createdDateTime: c.createdDateTime
                }));
            }
        } catch (error: any) {
            results.checks.push({
                test: "List containers with containerTypeId filter",
                success: false,
                error: error.message
            });
        }
        
        // Test 2: Try to create a test container with this type
        context.log("Test 2: Attempting to create a container with the type...");
        const createUrl = "https://graph.microsoft.com/beta/storage/fileStorage/containers";
        const testContainerData = {
            displayName: `Verification Test ${Date.now()}`,
            description: "Temporary container for verification",
            containerTypeId: containerTypeId
        };
        
        try {
            const createResponse = await fetch(createUrl, {
                method: "POST",
                headers: {
                    "Authorization": `Bearer ${response.accessToken}`,
                    "Content-Type": "application/json"
                },
                body: JSON.stringify(testContainerData)
            });
            
            const createData = await createResponse.json();
            results.checks.push({
                test: "Create container with containerTypeId",
                success: createResponse.ok,
                status: createResponse.status,
                message: createResponse.ok ? "Successfully created container" : "Failed to create container",
                data: createData
            });
            
            if (createResponse.ok && createData.id) {
                results.newContainerId = createData.id;
                
                // If creation succeeded, try to delete it to clean up
                const deleteUrl = `https://graph.microsoft.com/beta/storage/fileStorage/containers/${createData.id}`;
                const deleteResponse = await fetch(deleteUrl, {
                    method: "DELETE",
                    headers: {
                        "Authorization": `Bearer ${response.accessToken}`
                    }
                });
                
                results.checks.push({
                    test: "Delete test container",
                    success: deleteResponse.ok || deleteResponse.status === 204,
                    status: deleteResponse.status,
                    message: "Cleanup of test container"
                });
            }
        } catch (error: any) {
            results.checks.push({
                test: "Create container with containerTypeId",
                success: false,
                error: error.message
            });
        }
        
        // Test 3: Check if we can list all container types (this might fail due to permissions)
        context.log("Test 3: Trying to list all container types...");
        const typesUrl = "https://graph.microsoft.com/beta/storage/fileStorage/containerTypes";
        
        try {
            const typesResponse = await fetch(typesUrl, {
                headers: {
                    "Authorization": `Bearer ${response.accessToken}`
                }
            });
            
            const typesData = await typesResponse.json();
            results.checks.push({
                test: "List all container types",
                success: typesResponse.ok,
                status: typesResponse.status,
                message: typesResponse.ok ? "Can list container types" : "Cannot list container types (might be normal)",
                data: typesData
            });
        } catch (error: any) {
            results.checks.push({
                test: "List all container types",
                success: false,
                error: error.message
            });
        }
        
        // Test 4: Try listing ALL containers (no filter)
        context.log("Test 4: Listing all containers without filter...");
        const allContainersUrl = "https://graph.microsoft.com/beta/storage/fileStorage/containers";
        
        try {
            const allResponse = await fetch(allContainersUrl, {
                headers: {
                    "Authorization": `Bearer ${response.accessToken}`
                }
            });
            
            const allData = await allResponse.json();
            results.checks.push({
                test: "List all containers (no filter)",
                success: allResponse.ok,
                status: allResponse.status,
                message: allResponse.ok ? "Can list containers without filter" : "Cannot list any containers",
                containerCount: allData.value?.length || 0,
                data: allResponse.ok ? { count: allData.value?.length || 0 } : allData
            });
        } catch (error: any) {
            results.checks.push({
                test: "List all containers (no filter)",
                success: false,
                error: error.message
            });
        }
        const canCreateContainers = results.checks.find((c: any) => c.test === "Create container with containerTypeId")?.success;
        const canListContainers = results.checks.find((c: any) => c.test === "List containers with containerTypeId filter")?.success;
        
        results.summary = {
            isRegistered: canCreateContainers || canListContainers,
            canCreateContainers: canCreateContainers,
            canListContainers: canListContainers,
            recommendation: getRecommendation(results)
        };
        
        return {
            status: 200,
            headers: {
                "Content-Type": "application/json"
            },
            body: JSON.stringify(results)
        };
    } catch (error: any) {
        context.error("Error verifying container type:", error);
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

function getRecommendation(results: any): string {
    const canCreate = results.checks.find((c: any) => c.test === "Create container with containerTypeId")?.success;
    const canList = results.checks.find((c: any) => c.test === "List containers with containerTypeId filter")?.success;
    
    if (canCreate && canList) {
        return "✅ Container Type is properly registered and working!";
    } else if (canList && !canCreate) {
        return "⚠️ Can list but not create containers. Check if the app owns the Container Type.";
    } else if (!canList && !canCreate) {
        const createError = results.checks.find((c: any) => c.test === "Create container with containerTypeId")?.data?.error;
        if (createError?.message?.includes("not found")) {
            return "❌ Container Type not found. It needs to be registered via SharePoint admin.";
        } else if (createError?.message?.includes("accessDenied")) {
            return "❌ Access denied. Check permissions and Container Type ownership.";
        }
        return "❌ Container Type appears to not be registered or accessible.";
    }
    return "⚠️ Partial functionality detected. Review the detailed results.";
}

// Get token claims for debugging
export async function getTokenClaims(request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> {
    try {
        // Get app token
        const response = await cca.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        
        // Decode the token (basic parsing - in production use proper JWT library)
        const tokenParts = response.accessToken.split('.');
        const claims = JSON.parse(Buffer.from(tokenParts[1], 'base64').toString());
        
        // Remove sensitive data
        delete claims.appid;
        delete claims.oid;
        delete claims.tid;
        
        return {
            status: 200,
            body: JSON.stringify({
                success: true,
                roles: claims.roles || [],
                scp: claims.scp || "No delegated scopes (app-only token)",
                aud: claims.aud,
                iss: claims.iss,
                ver: claims.ver,
                hasFileStorageContainerSelected: claims.roles?.includes("FileStorageContainer.Selected"),
                allClaims: claims
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
app.http("verifyContainerType", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: verifyContainerType
});

app.http("getTokenClaims", {
    methods: ["GET"],
    authLevel: "anonymous",
    handler: getTokenClaims
});