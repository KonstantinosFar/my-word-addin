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
                searchResults.items.forEach(item => item.font.highlightColor = "red");
            }
        }
        await context.sync();
        statusDiv.innerText = "Done!";
    });
}

async function checkUrlWithAzure(url) {
    try {
        const response = await fetch("/api/check-link", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url: url })
        });
        const data = await response.json();
        return data.ok === false;
    } catch (e) {
        return false;
    }
}
