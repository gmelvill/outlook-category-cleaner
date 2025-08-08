Office.onReady(function () {
    if (Office.context.mailbox) {
        getCategories();
    }
});

async function getCategories() {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, async function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            const accessToken = result.value;
            const headers = new Headers();
            headers.append("Authorization", `Bearer ${accessToken}`);

            const catResponse = await fetch(
                "https://graph.microsoft.com/v1.0/me/outlook/masterCategories",
                { headers }
            );
            const categories = await catResponse.json();

            let unused = [];
            for (let cat of categories.value) {
                const used = await isCategoryUsed(cat.displayName, headers);
                if (!used) unused.push(cat.displayName);
            }

            renderCategoryList(unused, headers);
        } else {
            document.getElementById("categoryList").textContent = "Error getting token.";
        }
    });
}

async function isCategoryUsed(categoryName, headers) {
    const msgResponse = await fetch(
        `https://graph.microsoft.com/v1.0/me/messages?$top=1&$filter=categories/any(c:c eq '${categoryName}')`,
        { headers }
    );
    const messages = await msgResponse.json();
    return messages.value.length > 0;
}

function renderCategoryList(unused, headers) {
    const listDiv = document.getElementById("categoryList");
    listDiv.innerHTML = "";
    if (unused.length === 0) {
        listDiv.textContent = "No unused categories found.";
        return;
    }

    unused.forEach(cat => {
        const div = document.createElement("div");
        div.className = "category-item";
        div.textContent = cat;
        const btn = document.createElement("button");
        btn.textContent = "Delete";
        btn.onclick = () => deleteCategory(cat, headers, div);
        div.appendChild(btn);
        listDiv.appendChild(div);
    });
}

async function deleteCategory(name, headers, element) {
    await fetch(
        `https://graph.microsoft.com/v1.0/me/outlook/masterCategories/${encodeURIComponent(name)}`,
        { method: "DELETE", headers }
    );
    element.remove();
}
