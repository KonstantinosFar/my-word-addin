Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("checkLinksBtn").onclick = scanAndHighlightLinks;
    }
});

async function scanAndHighlightLinks() {
    const statusDiv = document.getElementById("status");
    statusDiv.innerText = "Scanning...";

    await Word.run(async (context) => {
        // Search the document for anything starting with http or https
        // Using Word's wildcard search syntax
        const searchResults = context.document.body.search("<http*>", { matchWildcards: true });
        searchResults.load("items");
        await context.sync();

        let brokenCount = 0;

        for (let i = 0; i < searchResults.items.length; i++) {
            const range = searchResults.items[i];
            range.load("text");
            await context.sync();

            // Clean up the extracted text (remove trailing punctuation if any)
            const urlText = range.text.replace(/[.,;!?]$/, '').trim();
            
            statusDiv.innerText = `Checking: ${urlText}`;

            // Send to your Azure backend
            const isBroken = await checkUrlWithAzure(urlText);

            if (isBroken) {
                range.font.highlightColor = "red";
                brokenCount++;
            }
        }

        await context.sync();
        statusDiv.innerText = `Done! Found ${brokenCount} broken link(s).`;
    }).catch(function (error) {
        console.log("Error: " + error);
        statusDiv.innerText = "An error occurred.";
    });
}

// Function to call your Azure Web App
async function checkUrlWithAzure(url) {
    try {
        // REPLACE THIS URL WITH YOUR AZURE WEB APP URL
        const azureEndpoint = "https://wordadd.azurewebsites.net/api/check-link";
        
        const response = await fetch(azureEndpoint, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url: url })
        });
        
        const data = await response.json();
        return data.broken;
    } catch (e) {
        console.error("Backend error", e);
        return true; // Treat as broken if backend fails
    }
}