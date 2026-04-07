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

// Function to call your Azure Function
async function checkUrlWithAzure(url) {
    try {
        const azureEndpoint = "https://wordlinkfunc-cede-faccezaka0gxckdk.canadacentral-01.azurewebsites.net/api/check-link";
        
        const response = await fetch(azureEndpoint, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url: url })
        });
        
        if (!response.ok) {
            console.log("Server responded with error status: " + response.status);
            return false; // Don't highlight red if the server itself is down
        }

        const data = await response.json();
        console.log("Check result for " + url + ":", data);

        // A link is BROKEN only if data.ok is explicitly false
        return data.ok === false; 

    } catch (e) {
        console.error("Connection error to Azure:", e);
        // Change this to FALSE for now. 
        // If it's TRUE, any connection hiccup turns the whole document red.
        return false; 
    }
}
