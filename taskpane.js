Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("checkLinksBtn").onclick = scanAndHighlightLinks;
    }
});

async function scanAndHighlightLinks() {
    const statusDiv = document.getElementById("status");
    statusDiv.innerText = "Scanning...";

    await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        // 1. Use pure JS Regex to find the FULL URLs instead of Word's wildcard search
        const text = body.text;
        const urlRegex = /(https?:\/\/[^\s]+)/g;
        const urls = text.match(urlRegex);

        if (!urls) {
            statusDiv.innerText = "No links found.";
            return;
        }

        let brokenCount = 0;

        // 2. Loop through the full URLs
        for (let url of urls) {
            const cleanUrl = url.replace(/[.,;!?]$/, '').trim();
            statusDiv.innerText = `Checking: ${cleanUrl}`;

            const isBroken = await checkUrlWithAzure(cleanUrl);

            if (isBroken) {
                // 3. Search for the EXACT full string to highlight it all
                const searchResults = body.search(cleanUrl, { matchCase: false });
                searchResults.load("items");
                await context.sync();

                for (let i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].font.highlightColor = "red";
                }
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
        const azureEndpoint = "https://wordlinkfunc-cede-faccezaka0gxckdk.canadacentral-01.azurewebsites.net/api/check-link";
        
        const response = await fetch(azureEndpoint, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url: url })
        });
        
        const data = await response.json();
        
        // This makes sure it works whether your backend sends { broken: true } or { status: "broken" }
        return data.broken === true || data.status === "broken";
    } catch (e) {
        console.error("Backend error", e);
        return true; 
    }
}
