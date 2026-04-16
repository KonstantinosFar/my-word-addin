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

        const urlRegex = /(https?:\/\/[^\s]+)/g;
        const urls = body.text.match(urlRegex);

        if (!urls) {
            statusDiv.innerText = "No links found.";
            return;
        }

        for (let url of urls) {
            const cleanUrl = url.replace(/[.,;!?]$/, '').trim();
            const isBroken = await checkUrlWithAzure(cleanUrl);

            if (isBroken) {
                const searchResults = body.search(cleanUrl);
                searchResults.load("items");
                await context.sync();
                searchResults.items.forEach(item => {
                    item.font.highlightColor = "red";
                });
            }
        }
        await context.sync();
        statusDiv.innerText = "Done scanning.";
    });
}

async function checkUrlWithAzure(url) {
    const azureEndpoint = "https://wordlinkfunc-cede-faccezaka0gxckdk.canadacentral-01.azurewebsites.net/api/check-link";
    const functionKey = "m9iyydRH2rs5-fGo3YI0a0MyWwWVkWq3zf637SeroPKRAzFuPTc5LQ==";

    try {
        const response = await fetch(azureEndpoint, {
            method: "POST",
            headers: { 
                "Content-Type": "application/json",
                "x-functions-key": functionKey 
            },
            body: JSON.stringify({ url: url })
        });
        const data = await response.json();
        return data.ok === false;
    } catch (e) {
        return false;
    }
}
