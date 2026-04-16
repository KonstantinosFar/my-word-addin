Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("checkLinksBtn").onclick = scanAndHighlightLinks;
    }
});

/**
 * Scans the document for links and checks them via Azure
 */
async function scanAndHighlightLinks() {
    const statusDiv = document.getElementById("status");
    const listContainer = document.getElementById("broken-links-container");
@@ -41,11 +44,24 @@

                // Show container and add to list immediately
                listContainer.style.display = "block";
                
                const li = document.createElement("li");
                li.style.marginBottom = "10px";

                // Create the clickable "Navigation" link
                const a = document.createElement("a");
                a.href = cleanUrl;
                a.target = "_blank"; 
                a.innerText = cleanUrl;
                a.href = "#"; // Prevent page refresh
                a.innerText = "🔍 " + cleanUrl;
                a.style.color = "#d13438"; // Microsoft Red
                a.style.textDecoration = "underline";
                a.style.cursor = "pointer";

                // This is the "Jump" trigger
                a.onclick = async (e) => {
                    e.preventDefault();
                    await jumpToLinkInDoc(cleanUrl);
                };

                li.appendChild(a);
                listUl.appendChild(li);

@@ -68,26 +84,45 @@
    });
}

/**
 * Scrolls the document to the first instance of the broken link
 */
async function jumpToLinkInDoc(linkText) {
    await Word.run(async (context) => {
        const results = context.document.body.search(linkText, { matchCase: false });
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
            // .select() highlights the text and scrolls the window to it
            results.items[0].select();
        }
    }).catch(function (error) {
        console.error("Jump error: " + error.message);
    });
}

/**
 * Communicates with your Azure Function
 */
async function checkUrlWithAzure(url) {
    try {
        const azureEndpoint = "https://wordlinkfunc-cede-faccezaka0gxckdk.canadacentral-01.azurewebsites.net/api/check-link";

        // 1. Paste your 'default' key from the Azure App Keys screen here
        // Using your verified key
        const functionKey = "m9iyydRH2rs5-fGo3YI0a0MyWwWVkWq3zf637SeroPKRAzFuPTc5LQ=="; 

        const response = await fetch(azureEndpoint, {
            method: "POST",
            headers: { 
                "Content-Type": "application/json",
                // 2. This is the magic line that lets the Add-in through the lock
                "x-functions-key": functionKey 
            },
            body: JSON.stringify({ url: url })
        });

        // If the key is wrong, Azure will still return 401
        if (response.status === 401) {
            console.error("The function key in taskpane.js does not match Azure.");
            console.error("Auth failed. Key mismatch.");
            return false;
        }
