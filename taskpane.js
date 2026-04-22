Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("checkLinksBtn").onclick = scanAndHighlightLinks;
    }
});

async function scanAndHighlightLinks() {
    const statusDiv = document.getElementById("status");
    const listUl = document.getElementById("broken-links-list");
    const listContainer = document.getElementById("broken-links-container");

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
                li.style.marginBottom = "10px";

                // Create the clickable link for the sidebar
                const a = document.createElement("a");
                a.href = "#";
                a.innerText = "➡️ " + cleanUrl;
                a.style.color = "#0078d4";
                a.style.textDecoration = "underline";
                a.style.cursor = "pointer";

                // Trigger the jump function on click
                a.onclick = async (e) => {
                    e.preventDefault();
                    await jumpToLinkInDoc(cleanUrl);
                };

                li.appendChild(a);
                listUl.appendChild(li);

                // Highlight the text in red immediately
                const searchResults = body.search(cleanUrl);
                searchResults.load("items");
                await context.sync();
                searchResults.items.forEach(item => {
                    item.font.highlightColor = "red";
                });
            }
        }
        await context.sync();
        statusDiv.innerText = `Done! Found ${brokenCount} broken link(s).`;
    });
}

// The function that scrolls Word to the link
async function jumpToLinkInDoc(linkText) {
    await Word.run(async (context) => {
        const results = context.document.body.search(linkText, { matchCase: false });
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
            // This selects the text and moves the view to it
            results.items[0].select();
        }
    }).catch(function (error) {
        console.error("Navigation error: " + error.message);
    });
}

async function checkUrlWithAzure(url) {
    // Using your Canadian Azure Function endpoint
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
