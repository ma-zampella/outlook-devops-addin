/* global Office */

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        loadEmailData();
        document.getElementById("createBtn").addEventListener("click", createWorkItem);
        document.getElementById("synthesizeBtn").addEventListener("click", synthesizeDescription);
        document.getElementById("toggleSetup").addEventListener("click", toggleSetup);
        document.getElementById("savePat").addEventListener("click", savePat);
        document.getElementById("saveGhToken").addEventListener("click", saveGhToken);
    }
});

// --- Email loading ---

var emailBodyFull = "";

function loadEmailData() {
    var item = Office.context.mailbox.item;
    document.getElementById("titleInput").value = item.subject || "";

    item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailBodyFull = result.value;
            document.getElementById("descInput").value = emailBodyFull.substring(0, 500);
        } else {
            document.getElementById("descInput").value = "(impossibile leggere il body)";
        }
    });
}

// --- Settings ---

function toggleSetup() {
    var panel = document.getElementById("setupPanel");
    panel.classList.toggle("visible");
}

function savePat() {
    var pat = document.getElementById("patInput").value.trim();
    if (pat) {
        localStorage.setItem("devops_pat", pat);
        showStatus("PAT DevOps salvato!", "success");
        document.getElementById("patInput").value = "";
    }
}

function saveGhToken() {
    var token = document.getElementById("ghTokenInput").value.trim();
    if (token) {
        localStorage.setItem("github_token", token);
        showStatus("Token GitHub salvato!", "success");
        document.getElementById("ghTokenInput").value = "";
    }
}

function getPat() {
    return localStorage.getItem("devops_pat");
}

function getGhToken() {
    return localStorage.getItem("github_token");
}

// --- LLM Synthesis ---

function synthesizeDescription() {
    var ghToken = getGhToken();
    if (!ghToken) {
        showStatus("Configura il token GitHub nelle impostazioni!", "error");
        toggleSetup();
        return;
    }

    var text = emailBodyFull || document.getElementById("descInput").value;
    if (!text.trim()) {
        showStatus("Nessun testo da sintetizzare", "error");
        return;
    }

    var btn = document.getElementById("synthesizeBtn");
    btn.disabled = true;
    btn.textContent = "Sintesi in corso...";

    var prompt = "Sei un assistente che crea descrizioni sintetiche per task di lavoro. " +
        "Data la seguente email, scrivi una descrizione concisa (max 3-4 frasi) per un work item " +
        "che catturi l'azione richiesta, il contesto e le eventuali scadenze. " +
        "Rispondi SOLO con la descrizione, senza prefissi.\n\nEmail:\n" + text.substring(0, 3000);

    var payload = JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        max_tokens: 300
    });

    fetch("https://models.inference.ai.azure.com/chat/completions", {
        method: "POST",
        headers: {
            "Authorization": "Bearer " + ghToken,
            "Content-Type": "application/json"
        },
        body: payload
    })
    .then(function (response) {
        if (!response.ok) {
            return response.text().then(function (t) { throw new Error("LLM HTTP " + response.status + ": " + t); });
        }
        return response.json();
    })
    .then(function (data) {
        var content = data.choices[0].message.content;
        document.getElementById("descInput").value = content;
        showStatus("Descrizione sintetizzata con AI", "success");
    })
    .catch(function (err) {
        showStatus("Errore sintesi: " + err.message, "error");
    })
    .finally(function () {
        btn.disabled = false;
        btn.textContent = "✨ Sintetizza con AI";
    });
}

// --- Email as attachment (.eml) ---

function getEmailAsEml() {
    return new Promise(function (resolve, reject) {
        Office.context.mailbox.item.getAsFileAsync(function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value); // Base64 string
            } else {
                reject(new Error("getAsFileAsync fallito: " + result.error.message));
            }
        });
    });
}

function uploadAttachment(pat, project, base64Eml, fileName) {
    var url = "https://dev.azure.com/Ivecogrp/" + project +
        "/_apis/wit/attachments?fileName=" + encodeURIComponent(fileName) + "&api-version=7.1";

    // Decode Base64 to binary
    var binary = atob(base64Eml);
    var bytes = new Uint8Array(binary.length);
    for (var i = 0; i < binary.length; i++) {
        bytes[i] = binary.charCodeAt(i);
    }

    return fetch(url, {
        method: "POST",
        headers: {
            "Content-Type": "application/octet-stream",
            "Authorization": "Basic " + btoa(":" + pat)
        },
        body: bytes.buffer
    })
    .then(function (response) {
        if (!response.ok) {
            return response.text().then(function (t) { throw new Error("Upload attachment HTTP " + response.status); });
        }
        return response.json();
    });
}

// --- Work Item creation ---

function createWorkItem() {
    var pat = getPat();
    if (!pat) {
        showStatus("Configura prima il PAT nelle impostazioni!", "error");
        toggleSetup();
        return;
    }

    var title = document.getElementById("titleInput").value.trim();
    var description = document.getElementById("descInput").value.trim();
    var board = document.getElementById("boardSelect").value;
    var assignTo = document.getElementById("assignSelect").value;

    if (!title) {
        showStatus("Il titolo è obbligatorio", "error");
        return;
    }

    var btn = document.getElementById("createBtn");
    btn.disabled = true;
    btn.textContent = "Creazione in corso...";

    // Step 1: Get email as .eml
    getEmailAsEml()
        .then(function (base64Eml) {
            // Step 2: Upload to first project (attachment is shared in org)
            var firstProject = (board === "ops") ? "Reply%20Operation" : "Reply%20Development%20Activities";
            return uploadAttachment(pat, firstProject, base64Eml, title.substring(0, 50) + ".eml")
                .then(function (attachResult) {
                    return attachResult.url;
                });
        })
        .catch(function () {
            // Se getAsFileAsync non è supportato (Mailbox < 1.14), procedi senza allegato
            return null;
        })
        .then(function (attachmentUrl) {
            var promises = [];

            if (board === "dev" || board === "both") {
                promises.push(
                    callDevOpsApi(pat, "Reply%20Development%20Activities", "Product%20Backlog%20Item", title, description, assignTo, attachmentUrl)
                );
            }
            if (board === "ops" || board === "both") {
                promises.push(
                    callDevOpsApi(pat, "Reply%20Operation", "Task", title, description, assignTo, attachmentUrl)
                );
            }

            return Promise.all(promises);
        })
        .then(function (results) {
            var ids = results.map(function (r) { return "#" + r.id; }).join(", ");
            showStatus("Work item creato: " + ids, "success");
        })
        .catch(function (err) {
            showStatus("Errore: " + err.message, "error");
        })
        .finally(function () {
            btn.disabled = false;
            btn.textContent = "Crea Work Item";
        });
}

function callDevOpsApi(pat, project, workItemType, title, description, assignTo, attachmentUrl) {
    var url = "https://dev.azure.com/Ivecogrp/" + project + "/_apis/wit/workitems/$" + workItemType + "?api-version=7.1";

    var body = [
        { op: "add", path: "/fields/System.Title", value: title },
        { op: "add", path: "/fields/System.Description", value: description },
        { op: "add", path: "/fields/System.AssignedTo", value: assignTo }
    ];

    if (attachmentUrl) {
        body.push({
            op: "add",
            path: "/relations/-",
            value: {
                rel: "AttachedFile",
                url: attachmentUrl,
                attributes: { comment: "Email originale" }
            }
        });
    }

    return fetch(url, {
        method: "PATCH",
        headers: {
            "Content-Type": "application/json-patch+json",
            "Authorization": "Basic " + btoa(":" + pat)
        },
        body: JSON.stringify(body)
    })
    .then(function (response) {
        if (!response.ok) {
            return response.text().then(function (text) {
                throw new Error("HTTP " + response.status + ": " + text);
            });
        }
        return response.json();
    });
}

function showStatus(message, type) {
    var el = document.getElementById("statusMsg");
    el.textContent = message;
    el.className = "status " + type;
}
