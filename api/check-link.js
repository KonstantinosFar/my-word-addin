const { app } = require('@azure/functions');

app.http('check-link', {
    methods: ['POST'],
    authLevel: 'anonymous',
    handler: async (request, context) => {
        try {
            // 1. Get the URL sent from your Word Add-in
            const body = await request.json();
            const urlToTest = body.url;

            // 2. SECRET SAUCE: Grab the key from Azure Settings
            // This 'AZURE_FUNCTION_KEY' is a placeholder. 
            // You will paste the real key in the Azure Portal later.
            const secretKey = process.env.AZURE_FUNCTION_KEY; 

            // 3. Talk to your original "locked" Azure Function
            const originalFunctionUrl = "https://wordlinkfunc-cede-faccezaka0gxckdk.canadacentral-01.azurewebsites.net/api/check-link";

            const response = await fetch(originalFunctionUrl, {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "x-functions-key": secretKey // Using the hidden key
                },
                body: JSON.stringify({ url: urlToTest })
            });

            const data = await response.json();

            // 4. Send the result back to Word
            return { jsonBody: data };

        } catch (error) {
            context.log(`Error: ${error.message}`);
            return { 
                status: 500, 
                jsonBody: { error: "The Middle-Man failed to connect." } 
            };
        }
    }
});
