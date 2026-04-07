Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("checkLinksBtn").onclick = scanAndHighlightLinks;
    }
});

async function scanAndHighlightLinks() {
    const statusDiv = document.getElementById("status");
    const listContainer = document.getElementById("broken-links-container");
    const listUl = document.getElementById("broken-links-list");

    // 1. Reset UI for a new scan
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
            // Clean common punctuation off the end of the URL
            const cleanUrl = url.replace(/[.,;!?]$/, '').trim();
            statusDiv.innerText = `Checking: ${cleanUrl}`;

            const isBroken = await checkUrlWithAzure(cleanUrl);

            if (isBroken) {
                brokenCount++;
                
                // 2. Reveal the list container
                listContainer.style.display = "block";

                // 3. Create the list item for the sidebar
                const li = document.createElement("li");
                const a = document.createElement("a");
                a.href = cleanUrl;
                a.target = "_blank"; 
                a.innerText = cleanUrl;
                
                li.appendChild(a);
                listUl.appendChild(li);

                // 4. Highlight the link in the Word document
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
        console.log("Error: " + error);
        statusDiv.innerText = "An error occurred during scanning.";
    });
}

async function checkUrlWithAzure(url) {
    try {
        // Use your verified long URL
        const azureEndpoint = "https://wordlinkfunc-cede-faccezaka0gxckdk.canadacentral-01.azurewebsites.net/api/check-link";
        
        const response = await fetch(azureEndpoint, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ url: url })
        });
        
        if (!response.ok) return false;

        const data = await response.json();
        // Return true if the link is broken (ok is false)
        return data.ok === false; 

    } catch (e) {
        console.error("Backend error", e);
        return false; 
    }
}
