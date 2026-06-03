/* global Office */

Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        loadEmailData();
        document.getElementById("createBtn").addEventListener("click", createWorkItem);
        document.getElementById("synthesizeBtn").addEventListener("click", runSynthesis);
        document.getElementById("toggleSetup").addEventListener("click", toggleSetup);
        document.getElementById("savePat").addEventListener("click", savePat);
        document.getElementById("saveGhToken").addEventListener("click", saveGhToken);
        document.getElementById("boardSelect").addEventListener("change", toggleOpsFields);
        toggleOpsFields();
    }
});

// --- Email loading ---

var emailBodyFull = "";
var emailSubject = "";

function loadEmailData() {
    var item = Office.context.mailbox.item;
    emailSubject = item.subject || "";
    document.getElementById("titleInput").value = emailSubject;

    // Get full conversation body (includes entire thread)
    item.body.getAsync(Office.CoercionType.Text, function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            emailBodyFull = result.value;
            // Auto-synthesize on load
            runSynthesis();
        } else {
            document.getElementById("descInput").value = "(impossibile leggere il body)";
        }
    });
}

// --- Settings ---

function toggleOpsFields() {
    var board = document.getElementById("boardSelect").value;
    var wrapper = document.getElementById("opsFieldsWrapper");
    wrapper.style.display = (board === "ops" || board === "both") ? "block" : "none";
}

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

// --- LLM Synthesis (integrated into creation flow) ---

function runSynthesis() {
    var ghToken = getGhToken();
    if (!ghToken) {
        document.getElementById("descInput").value = "Configura il token GitHub per generare titolo e descrizione con AI.";
        return;
    }

    var text = emailBodyFull;
    if (!text || !text.trim()) {
        document.getElementById("descInput").value = "In attesa del body email...";
        return;
    }

    var btn = document.getElementById("synthesizeBtn");
    btn.disabled = true;
    btn.textContent = "Sintesi in corso...";
    document.getElementById("descInput").value = "Generazione AI in corso...";

    synthesizeWithAI(emailSubject, text)
        .then(function (synthesized) {
            document.getElementById("titleInput").value = synthesized.title;
            document.getElementById("descInput").value = synthesized.description;
        })
        .catch(function (err) {
            document.getElementById("descInput").value = "Errore AI: " + err.message + "\n\n" + emailBodyFull.substring(0, 500);
        })
        .finally(function () {
            btn.disabled = false;
            btn.textContent = "🔄 Rigenera con AI";
        });
}

function synthesizeWithAI(subject, bodyText) {
    var ghToken = getGhToken();
    if (!ghToken) {
        return Promise.reject(new Error("Token GitHub non configurato"));
    }

    var prompt = "You are an assistant that analyzes email threads and produces two outputs.\n" +
        "You will receive the FULL email thread (including previous replies). Use the entire context to understand the issue.\n\n" +
        "OUTPUT 1 - TITLE: Determine if the email is related to one of these brands: Iveco, IVG, FPT. " +
        "If you can identify the brand, output: MAIL / BRAND / original_subject. " +
        "If you cannot identify the brand, output: MAIL / original_subject. " +
        "Example: MAIL / FPT / Issue on Sitecore\n" +
        "OUTPUT 2 - DESCRIPTION: Write a concise summary (3-4 sentences max) in ENGLISH of the email thread, " +
        "capturing the action requested, context and any deadlines. Consider the full conversation history.\n\n" +
        "Reply ONLY in this exact JSON format (no markdown, no code blocks):\n" +
        '{"title": "MAIL / ... / ...", "description": "..."}\n\n' +
        "Email subject: " + subject + "\n" +
        "Full email thread:\n" + bodyText.substring(0, 4000);

    var payload = JSON.stringify({
        model: "gpt-4o-mini",
        messages: [{ role: "user", content: prompt }],
        max_tokens: 400
    });

    return fetch("https://models.inference.ai.azure.com/chat/completions", {
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
        content = content.replace(/^```json\s*/, "").replace(/```\s*$/, "").trim();
        try {
            var parsed = JSON.parse(content);
            return {
                title: parsed.title,
                description: "PLEASE CHECK MAIL ATTACHED\n\n" + parsed.description
            };
        } catch (e) {
            return {
                title: "MAIL / " + subject,
                description: "PLEASE CHECK MAIL ATTACHED\n\n" + content
            };
        }
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

    var board = document.getElementById("boardSelect").value;
    var assignTo = document.getElementById("assignSelect").value;

    var btn = document.getElementById("createBtn");
    btn.disabled = true;
    btn.textContent = "Creazione work item...";

    var title = document.getElementById("titleInput").value.trim();
    var description = document.getElementById("descInput").value.trim();

    if (!title) {
        showStatus("Il titolo è obbligatorio", "error");
        btn.disabled = false;
        btn.textContent = "Crea Work Item";
        return;
    }

    // Step 1: Get email as .eml
    getEmailAsEml()
        .then(function (base64Eml) {
            var firstProject = (board === "ops") ? "Reply%20Operation" : "Reply%20Development%20Activities";
            return uploadAttachment(pat, firstProject, base64Eml, title.substring(0, 50) + ".eml")
                .then(function (attachResult) {
                    return attachResult.url;
                });
        })
        .catch(function () {
            return null;
        })
        .then(function (attachmentUrl) {
            var promises = [];

            if (board === "dev" || board === "both") {
                promises.push(
                    callDevOpsApi(pat, "Reply%20Development%20Activities", "Product%20Backlog%20Item", title, description, assignTo, attachmentUrl, false)
                        .then(function (r) { r._project = "Reply%20Development%20Activities"; return r; })
                );
            }
            if (board === "ops" || board === "both") {
                var opsType = document.getElementById("opsTypeSelect").value;
                promises.push(
                    callDevOpsApi(pat, "Reply%20Operation", opsType, title, description, assignTo, attachmentUrl, true)
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

            // Generate draft reply all
            btn.textContent = "Generazione bozza...";
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

function callDevOpsApi(pat, project, workItemType, title, description, assignTo, attachmentUrl, isOpsTask) {
    var url = "https://dev.azure.com/Ivecogrp/" + project + "/_apis/wit/workitems/$" + workItemType + "?api-version=7.1";

    var body = [
        { op: "add", path: "/fields/System.Title", value: title },
        { op: "add", path: "/fields/System.Description", value: description },
        { op: "add", path: "/fields/System.AssignedTo", value: assignTo }
    ];

    // Add Reply Operation required fields
    if (isOpsTask) {
        var app = document.getElementById("appSelect").value;
        var severity = document.getElementById("severitySelect").value;

        body.push({ op: "add", path: "/fields/Microsoft.VSTS.TCM.ReproSteps", value: "-" });
        body.push({ op: "add", path: "/fields/Custom.Source", value: "1 MAIL" });
        body.push({ op: "add", path: "/fields/Custom.Region", value: "Global" });
        body.push({ op: "add", path: "/fields/Custom.OriginalApplication", value: app });
        body.push({ op: "add", path: "/fields/Microsoft.VSTS.Common.Severity", value: severity });
        body.push({ op: "add", path: "/fields/Custom.Impact", value: "D - Post-Release (LOW customer impact)" });
    }

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

    var emailText = emailBodyFull.substring(0, 2000);
    var subject = document.getElementById("titleInput").value;

    var prompt = "Analyze the following email thread. Identify the language of the LAST message from the customer/external sender (not internal replies). " +
        "Reply in THAT language (English or Italian).\n\n" +
        "Write a brief, professional reply saying:\n" +
        "- We have received their message and are looking into the issue\n" +
        "- We will keep them updated on the progress\n" +
        "- Include this exact text as a clickable reference: [Task #" + taskId + "](" + taskLink + ")\n\n" +
        "Keep it short (3-5 sentences). Do NOT include subject line, greetings like 'Dear' or sign-off. Output ONLY the body paragraph(s).\n\n" +
        "Original email subject: " + subject + "\n" +
        "Email thread:\n" + emailText;

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
        setDraftReplyAll(draftBody, taskLink, taskId);
    })
    .catch(function (err) {
        showStatusHtml(document.getElementById("statusMsg").innerHTML +
            "<br/><span style='color:#a80000;'>Draft non generato: " + err.message + "</span>", "success");
    });
}

function setDraftReplyAll(bodyText, taskLink, taskId) {
    var item = Office.context.mailbox.item;

    // Convert markdown link to HTML anchor and ensure clickable link
    var htmlBody = bodyText.replace(/\n/g, "<br/>");
    // Replace markdown-style links [text](url) with HTML anchors
    htmlBody = htmlBody.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>');
    // If the link wasn't inserted by LLM as markdown, ensure it's there as clickable
    if (htmlBody.indexOf(taskLink) === -1 && htmlBody.indexOf("Task #" + taskId) === -1) {
        htmlBody += '<br/><br/>Task reference: <a href="' + taskLink + '">Task #' + taskId + '</a>';
    }
    // Also replace plain URL text with clickable link (if LLM put it as plain text)
    htmlBody = htmlBody.replace(new RegExp('(?<!href=")(' + taskLink.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')(?!</a>)', 'g'),
        '<a href="' + taskLink + '">Task #' + taskId + '</a>');

    // Wrap in Aptos 12 black font
    htmlBody = '<div style="font-family: Aptos, Calibri, sans-serif; font-size: 12pt; color: black;">' + htmlBody + '</div>';

    try {
        item.displayReplyAllForm(htmlBody);
    } catch (e) {
        // Fallback: show draft in the panel
        showStatusHtml(document.getElementById("statusMsg").innerHTML +
            "<br/><br/><b>Draft risposta (Reply All):</b><br/><div style='background:#fff;padding:8px;border:1px solid #ccc;margin-top:4px;font-family:Aptos,sans-serif;font-size:12pt;'>" +
            htmlBody + "</div>", "success");
    }
}
