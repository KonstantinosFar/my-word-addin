Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("checkLinksBtn").onclick = scanAndHighlightLinks;
    }
});

async function scanAndHighlightLinks() {
    const statusDiv = document.getElementById("status");
    const listContainer = document.getElementById("broken-links-container");
    const listUl = document.getElementById("broken-links-list");

    // Reset UI
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
                
                // Show container and add to list immediately
                listContainer.style.display = "block";
                const li = document.createElement("li");
                const a = document.createElement("a");
                a.href = cleanUrl;
                a.target = "_blank"; 
                a.innerText = cleanUrl;
                li.appendChild(a);
                listUl.appendChild(li);

                // Highlight in Word
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
        statusDiv.innerText = "Error: Check browser console (F12).";
    });
}

async function checkUrlWithAzure(url) {
    try {
        const azureEndpoint = "https://wordlinkfunc-cede-faccezaka0gxckdk.canadacentral-01.azurewebsites.net/api/check-link";
        
        // REPLACE 'YOUR_KEY_HERE' with the secret key you just copied
        const functionKey = "m9iyydRH2rs5-fGo3YI0a0MyWwWVkWq3zf637SeroPKRAzFuPTc5LQ==";

        const response = await fetch(azureEndpoint, {
            method: "POST",
            headers: { 
                "Content-Type": "application/json",
                "x-functions-key": functionKey // This is the secret password
            },
            body: JSON.stringify({ url: url })
        });
        
        const data = await response.json();
        return data.ok === false; 

    } catch (e) {
        return false; 
    }
}
