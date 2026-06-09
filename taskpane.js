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

    var prompt = "You are a Technical Business Analyst. Your job is to analyze email threads and convert them into clear, actionable DevOps tickets for developers.\n" +
    "You will receive the FULL email thread. Use the entire context to extract ONLY the technical issue and the task to be done.\n\n" +
    "OUTPUT 1 - TITLE: Determine if the email is related to one of these brands: Iveco, IVG, FPT. " +
    "If you can identify the brand, output: MAIL / BRAND / original_subject. " +
    "If you cannot identify the brand, output: MAIL / original_subject. " +
    "Example: MAIL / FPT / Issue on Sitecore\n" +
    "OUTPUT 2 - DESCRIPTION: Write a technical description (3-4 sentences max) in ENGLISH focused EXCLUSIVELY on the actual problem, bug, or feature request. " +
    "DO NOT summarize the conversation history, DO NOT mention who wrote to whom, and DO NOT include pleasantries or phrases like 'The user states that'. " +
    "Focus ONLY on: What is broken/requested, where it happens, and what the expected behavior is, so a developer can immediately understand the scope of work.\n\n" +
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
        .catch(function (e) {
            console.warn("Attachment skipped:", e.message);
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
            var opsResult = null;
            var devResult = null;
            for (var i = 0; i < results.length; i++) {
                if (results[i]._project === "Reply%20Operation") {
                    opsResult = results[i];
                }
                if (results[i]._project === "Reply%20Development%20Activities") {
                    devResult = results[i];
                }
            }

            if (board === "both" && opsResult && devResult) {
                return linkWorkItemsParentChild(
                    pat,
                    devResult._project,
                    devResult.id,
                    opsResult._project,
                    opsResult.id
                )
                .then(function () {
                    return { results: results, opsResult: opsResult, linkCreated: true, linkError: null };
                })
                .catch(function (err) {
                    return { results: results, opsResult: opsResult, linkCreated: false, linkError: err.message };
                });
            }

            return { results: results, opsResult: opsResult, linkCreated: false, linkError: null };
        })
        .then(function (context) {
            var results = context.results;
            var links = results.map(function (r) {
                var url = "https://dev.azure.com/Ivecogrp/" + r._project + "/_workitems/edit/" + r.id;
                return '<a href="' + url + '" target="_blank">#' + r.id + ' (' + r._project.replace(/%20/g, ' ') + ')</a>';
            }).join("<br/>");
            var statusHtml = "Work item creato:<br/>" + links;
            if (context.linkCreated) {
                statusHtml += "<br/><br/>Relazione creata: task Operation Parent del task Reply Development Activities.";
            } else if (context.linkError) {
                statusHtml += "<br/><br/><span style='color:#a80000;'>Work item creati ma link Parent/Child non creato: " + context.linkError + "</span>";
            }
            showStatusHtml(statusHtml, "success");

            // Find Reply Operation task link for the draft
            var opsResult = context.opsResult;
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
        var environment = document.getElementById("environmentSelect").value;

        body.push({ op: "add", path: "/fields/Custom.StepstoReproduce", value: "-" });
        body.push({ op: "add", path: "/fields/Custom.Source1", value: "1 MAIL" });
        body.push({ op: "add", path: "/fields/Custom.Region_INC", value: "Global" });
        body.push({ op: "add", path: "/fields/Custom.Applications", value: app });
        body.push({ op: "add", path: "/fields/Custom.ReplyApplication", value: app });
        body.push({ op: "add", path: "/fields/Custom.Severity1", value: severity });
        body.push({ op: "add", path: "/fields/Custom.Environment_INC", value: environment });
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
    .catch(function (networkErr) {
        throw new Error("Rete/CORS bloccato per " + workItemType + ": " + networkErr.message);
    })
    .then(function (response) {
        if (!response.ok) {
            return response.text().then(function (text) {
                throw new Error(workItemType + " HTTP " + response.status + ": " + text);
            });
        }
        return response.json();
    });
}

function linkWorkItemsParentChild(pat, childProject, childId, parentProject, parentId) {
    var childApiUrl = "https://dev.azure.com/Ivecogrp/" + childProject +
        "/_apis/wit/workitems/" + childId + "?api-version=7.1";
    var parentApiReference = "https://dev.azure.com/Ivecogrp/" + parentProject +
        "/_apis/wit/workItems/" + parentId;

    var body = [{
        op: "add",
        path: "/relations/-",
        value: {
            rel: "System.LinkTypes.Hierarchy-Reverse",
            url: parentApiReference,
            attributes: { comment: "Linked automatically by Outlook add-in" }
        }
    }];

    return fetch(childApiUrl, {
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
                throw new Error("Link Parent/Child HTTP " + response.status + ": " + text);
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

    var prompt = "SYSTEM INSTRUCTION:\n" +
    "You are an automated support assistant. Your task is to write a short, professional email acknowledgment.\n\n" +

    "CRITICAL RULES:\n" +
    "1. LANGUAGE: Detect the language of the VERY LAST email in the thread below. You MUST reply ONLY in that exact language (e.g., Italian if the last email is in Italian, English if it is in English).\n" +
    "2. STYLE & BREVITY: Keep the central body paragraph extremely concise (max 2 sentences). Do not add conversational fluff.\n\n" +

    "REQUIRED STRUCTURE (Translate all elements to the detected language):\n" +
    "1. GREETING: Start with an appropriate greeting based on the time of day (e.g., 'Buongiorno / Buonasera' in Italian, or 'Good morning / Good afternoon' in English).\n" +
    "2. LINE BREAK\n" +
    "3. BODY: State that we received the message, are looking into the issue, and will keep them updated. Include this exact clickable reference: [Task #" + taskId + "](" + taskLink + ")\n" +
    "4. LINE BREAK\n" +
    "5. SIGN-OFF & SIGNATURE: Close with a professional thank you (e.g., 'Grazie e un cordiale saluto' / 'Thank you and best regards') followed by the name 'Marcello'.\n\n" +

    "EMAIL DATA TO ANALYZE:\n" +
    "Original Subject: " + subject + "\n" +
    "Thread History:\n" + emailText + "\n\n" +

    "OUTPUT DIRECTIVE: Generate ONLY the structured email according to the rules above. Do not include subject lines or extra text."

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
