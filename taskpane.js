Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("checkLinksBtn").onclick = scanAndHighlightLinks;
    }
});

async function scanAndHighlightLinks() {
    const statusDiv = document.getElementById("status");
    const listContainer = document.getElementById("broken-links-container");
    const listUl = document.getElementById("broken-links-list");

    statusDiv.innerText = "Scanning...";
    listUl.innerHTML = ""; 
    listContainer.style.display = "none";

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

        let brokenCount = 0;
        for (let url of urls) {
            const cleanUrl = url.replace(/[.,;!?]$/, '').trim();
            statusDiv.innerText = `Checking: ${cleanUrl}`;

            const isBroken = await checkUrlWithAzure(cleanUrl);

            if (isBroken) {
                brokenCount++;
                listContainer.style.display = "block";
                const li = document.createElement("li");
                li.innerHTML = `<a href="#" style="color:red;">🔍 ${cleanUrl}</a>`;
                listUl.appendChild(li);

                const searchResults = body.search(cleanUrl);
                searchResults.load("items");
                await context.sync();
                searchResults.items.forEach(item => item.font.highlightColor = "red");
            }
        }
        await context.sync();
        statusDiv.innerText = `Done! Found ${brokenCount} broken link(s).`;
    });
}

async function checkUrlWithAzure(url) {
    try {
        // This relative path tells Azure to use your /api folder automatically
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
