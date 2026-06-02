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

    var subject = document.getElementById("titleInput").value;

    var prompt = "You are an assistant that analyzes emails and produces two outputs.\n" +
        "OUTPUT 1 - TITLE: Determine if the email is related to one of these brands: Iveco, IVG, FPT. " +
        "If you can identify the brand, output: MAIL / BRAND / original_subject. " +
        "If you cannot identify the brand, output: MAIL / original_subject. " +
        "Example: MAIL / FPT / Issue on Sitecore\n" +
        "OUTPUT 2 - DESCRIPTION: Write a concise summary (3-4 sentences max) in ENGLISH of the email content, " +
        "capturing the action requested, context and any deadlines.\n\n" +
        "Reply ONLY in this exact JSON format (no markdown, no code blocks):\n" +
        '{"title": "MAIL / ... / ...", "description": "..."}\n\n' +
        "Email subject: " + subject + "\n" +
        "Email body:\n" + text.substring(0, 3000);

    var payload = JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        max_tokens: 400
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
        var content = data.choices[0].message.content.trim();
        try {
            // Remove markdown code block if present
            content = content.replace(/^```json\s*/, "").replace(/```\s*$/, "").trim();
            var parsed = JSON.parse(content);
            document.getElementById("titleInput").value = parsed.title;
            document.getElementById("descInput").value = "PLEASE CHECK MAIL ATTACHED\n\n" + parsed.description;
        } catch (e) {
            // Fallback: use raw content as description
            document.getElementById("descInput").value = "PLEASE CHECK MAIL ATTACHED\n\n" + content;
        }
        showStatus("Titolo e descrizione generati con AI", "success");
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
                        .then(function (r) { r._project = "Reply%20Development%20Activities"; return r; })
                );
            }
            if (board === "ops" || board === "both") {
                promises.push(
                    callDevOpsApi(pat, "Reply%20Operation", "Task", title, description, assignTo, attachmentUrl)
                        .then(function (r) { r._project = "Reply%20Operation"; return r; })
                );
            }

            return Promise.all(promises);
        })
        .then(function (results) {
            var links = results.map(function (r) {
                var url = "https://dev.azure.com/Ivecogrp/" + r._project + "/_workitems/edit/" + r.id;
                return '<a href="' + url + '" target="_blank">#' + r.id + ' (' + r._project.replace(/%20/g, ' ') + ')</a>';
            }).join("<br/>");
            showStatusHtml("Work item creato:<br/>" + links, "success");

            // Find Reply Operation task link for the draft
            var opsResult = null;
            for (var i = 0; i < results.length; i++) {
                if (results[i]._project === "Reply%20Operation") {
                    opsResult = results[i];
                    break;
                }
            }
            var useResult = opsResult || results[0];
            var taskLink = "https://dev.azure.com/Ivecogrp/" + useResult._project + "/_workitems/edit/" + useResult.id;

            // Generate draft reply
            generateDraftReply(taskLink, useResult.id);
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

function showStatusHtml(html, type) {
    var el = document.getElementById("statusMsg");
    el.innerHTML = html;
    el.className = "status " + type;
}

// --- Draft Reply ---

function generateDraftReply(taskLink, taskId) {
    var ghToken = getGhToken();
    if (!ghToken) return;

    var emailText = emailBodyFull.substring(0, 1500);
    var subject = document.getElementById("titleInput").value;

    var prompt = "You must detect the language of the following email (Italian or English) and reply in the SAME language.\n" +
        "Write a brief, professional reply email saying:\n" +
        "- We have received their message and are looking into the issue\n" +
        "- We will keep them updated on the progress\n" +
        "- Reference the internal tracking task: " + taskLink + "\n\n" +
        "Keep it short (3-5 sentences). Do NOT include subject line. Output ONLY the email body text.\n\n" +
        "Original email subject: " + subject + "\n" +
        "Original email body:\n" + emailText;

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
        if (!response.ok) throw new Error("LLM error");
        return response.json();
    })
    .then(function (data) {
        var draftBody = data.choices[0].message.content.trim();
        setDraftReply(draftBody);
    })
    .catch(function (err) {
        showStatus("Draft non generato: " + err.message, "error");
    });
}

function setDraftReply(bodyText) {
    var item = Office.context.mailbox.item;
    var htmlBody = bodyText.replace(/\n/g, "<br/>");
    try {
        item.displayReplyForm({
            htmlBody: htmlBody,
            callback: function (result) {
                if (result.status === Office.AsyncResultStatus.Failed) {
                    showStatusHtml(document.getElementById("statusMsg").innerHTML +
                        "<br/><br/><b>Draft (copia manualmente):</b><br/>" + htmlBody, "success");
                }
            }
        });
    } catch (e) {
        // Fallback: show draft text in the status area
        showStatusHtml(document.getElementById("statusMsg").innerHTML +
            "<br/><br/><b>Draft risposta:</b><br/><div style='background:#fff;padding:8px;border:1px solid #ccc;margin-top:4px;'>" +
            htmlBody + "</div>", "success");
    }
}
