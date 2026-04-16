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

        const text = body.text;
        const urlRegex = /(https?:\/\/[^\s]+)/g;
        const urls = text.match(urlRegex);

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
                li.style.marginBottom = "10px";

                const a = document.createElement("a");
                a.href = "#"; 
                a.innerText = "🔍 " + cleanUrl;
                a.style.color = "#d13438"; 
                a.style.textDecoration = "underline";
                a.style.cursor = "pointer";

                a.onclick = async (e) => {
                    e.preventDefault();
                    await jumpToLinkInDoc(cleanUrl);
                };

                li.appendChild(a);
                listUl.appendChild(li);

                const searchResults = body.search(cleanUrl, { matchCase: false });
                searchResults.load("items");
                await context.sync();
                for (let i = 0; i < searchResults.items.length; i++) {
                    searchResults.items[i].font.highlightColor = "red";
                }
            }
        }
        await context.sync();
        statusDiv.innerText = `Done! Found ${brokenCount} broken link(s).`;
    }).catch(function (error) {
        console.error(error);
        statusDiv.innerText = "Error scanning links.";
    });
}

async function jumpToLinkInDoc(linkText) {
    await Word.run(async (context) => {
        const results = context.document.body.search(linkText, { matchCase: false });
        results.load("items");
        await context.sync();
        if (results.items.length > 0) {
            results.items[0].select();
        }
    });
}

async function checkUrlWithAzure(url) {
    try {
        // Now using your relative API path. NO SECRET KEY HERE!
        const response = await fetch("/api/check-link", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url: url })
        });
        
        const data = await response.json();
        return data.ok === false; 
    } catch (e) {
        console.error("Connection error:", e);
        return false; 
    }
}
