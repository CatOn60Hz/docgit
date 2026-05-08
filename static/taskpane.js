let documentPath = "";

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        bindButtons();
        bindDiffModeToggle();
        detectDocumentPath();
    }
});

// ── Button bindings ──────────────────────────────────────────────────────────
function bindButtons() {
    document.getElementById("btnInit").onclick = async () => {
        await runCommand("init");
        refreshAll();
    };

    document.getElementById("btnCommit").onclick = async () => {
        const msg = document.getElementById("commitMsg").value.trim();
        if (!msg) { showOutput("Please enter a commit message."); return; }

        showOutput("Saving document...");
        const saved = await saveDocument();
        if (!saved) {
            showOutput("Note: Auto-save was not available. Committing current file on disk...");
            await new Promise(r => setTimeout(r, 800));
        }

        await runCommand(`commit -m ${JSON.stringify(msg)}`);
        document.getElementById("commitMsg").value = "";
        refreshAll();
    };

    document.getElementById("btnStatus").onclick = () => runCommand("status");
    document.getElementById("btnLog").onclick    = () => runCommand("log");
    document.getElementById("btnGraph").onclick  = () => runCommand("graph");

    document.getElementById("btnDiff").onclick = async () => {
        if (!documentPath) { showOutput("Document path not set."); return; }
        const mode = document.querySelector('input[name="diffMode"]:checked').value;
        const body = { filepath: documentPath, mode };

        if (mode === "branches") {
            body.branch_a = document.getElementById("diffBranchA").value;
            body.branch_b = document.getElementById("diffBranchB").value;
            if (!body.branch_a || !body.branch_b) { showOutput("Select both branches for diff."); return; }
        } else if (mode === "commits") {
            body.commit_a = document.getElementById("diffCommitA").value;
            body.commit_b = document.getElementById("diffCommitB").value;
            if (!body.commit_a || !body.commit_b) { showOutput("Select both commits for diff."); return; }
        }

        showOutput("Building diff...");
        try {
            const r = await apiFetch("/api/diff", body);
            if (r.error) { showOutput("Error: " + r.error); return; }
            showOutput("Opening diff viewer...");
            window.open(r.diff_url, "_blank");
        } catch(e) { showOutput("Diff failed: " + e); }
    };

    document.getElementById("btnBranch").onclick = async () => {
        const name = document.getElementById("branchName").value.trim();
        if (!name) { showOutput("Enter a branch name."); return; }
        await runCommand(`branch ${JSON.stringify(name)}`);
        document.getElementById("branchName").value = "";
        refreshAll();
    };

    document.getElementById("btnSwitch").onclick = async () => {
        const t = document.getElementById("switchSelect").value;
        if (!t) { showOutput("Select a branch to switch to."); return; }
        await runCommand(`switch ${JSON.stringify(t)}`);
        refreshAll();
    };

    document.getElementById("btnMerge").onclick = async () => {
        const t = document.getElementById("mergeSelect").value;
        if (!t) { showOutput("Select a branch to merge."); return; }
        await runCommand(`merge ${JSON.stringify(t)}`);
        refreshAll();
    };

    // Two-step rollback: first click arms it, second click fires
    let rollbackArmed = false;
    document.getElementById("btnRollback").onclick = async () => {
        const cid = document.getElementById("rollbackSelect").value;
        if (!cid) { showOutput("Select a commit to roll back to."); return; }

        if (!rollbackArmed) {
            rollbackArmed = true;
            document.getElementById("btnRollback").textContent = "Click again to confirm";
            document.getElementById("btnRollback").style.background = "#a4262c";
            document.getElementById("btnRollback").style.color = "white";
            setTimeout(() => {
                rollbackArmed = false;
                document.getElementById("btnRollback").textContent = "Rollback";
                document.getElementById("btnRollback").style.background = "";
                document.getElementById("btnRollback").style.color = "";
            }, 4000);
            return;
        }

        rollbackArmed = false;
        document.getElementById("btnRollback").textContent = "Rollback";
        document.getElementById("btnRollback").style.background = "";
        document.getElementById("btnRollback").style.color = "";

        const result = await runCommand(`rollback ${cid}`);
        refreshAll();

        // After rollback the file on disk is restored — user must close+reopen
        showOutput((document.getElementById("output").innerText) +
            "\n\nIMPORTANT: Close and reopen the Word document to see the restored version.");
    };

    document.getElementById("btnSetPath").onclick = () => {
        const manual = document.getElementById("manualPath").value.trim();
        if (manual) {
            documentPath = manual;
            document.getElementById("docPath").innerText = documentPath;
            document.getElementById("pathWarning").style.display = "none";
            refreshAll();
        }
    };

    document.getElementById("btnRefresh").onclick = refreshAll;
}

// ── Diff mode toggle ─────────────────────────────────────────────────────────
function bindDiffModeToggle() {
    document.querySelectorAll('input[name="diffMode"]').forEach(radio => {
        radio.onchange = () => {
            const mode = radio.value;
            document.getElementById("diffBranchPanel").style.display = mode === "branches" ? "block" : "none";
            document.getElementById("diffCommitPanel").style.display = mode === "commits"  ? "block" : "none";
        };
    });
}

// ── Path detection ───────────────────────────────────────────────────────────
function detectDocumentPath() {
    try {
        const url = Office.context.document.url;
        if (url && url.length > 0) {
            documentPath = normalizeWordPath(url);
            document.getElementById("docPath").innerText = documentPath;
            document.getElementById("pathWarning").style.display = "none";
            refreshAll();
            return;
        }
    } catch(e) {}
    Office.context.document.getFilePropertiesAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded && result.value && result.value.url) {
            documentPath = normalizeWordPath(result.value.url);
            document.getElementById("docPath").innerText = documentPath;
            document.getElementById("pathWarning").style.display = "none";
            refreshAll();
        } else {
            document.getElementById("docPath").innerText = "Could not detect path.";
            document.getElementById("pathWarning").style.display = "block";
        }
    });
}

function normalizeWordPath(filepath) {
    if (!filepath) return "";
    if (filepath.startsWith("file:///")) {
        filepath = decodeURIComponent(filepath.substring(8));
        filepath = filepath.replace(/\//g, "\\");
    }
    return filepath;
}

// ── Save ─────────────────────────────────────────────────────────────────────
function saveDocument() {
    return new Promise((resolve) => {
        // Timeout after 3 seconds in case saveAsync never calls back
        const tid = setTimeout(() => resolve(false), 3000);
        try {
            Office.context.document.saveAsync((result) => {
                clearTimeout(tid);
                resolve(result.status === Office.AsyncResultStatus.Succeeded);
            });
        } catch(e) {
            clearTimeout(tid);
            resolve(false);
        }
    });
}

// ── Refresh branches + commits dropdowns ─────────────────────────────────────
async function refreshAll() {
    if (!documentPath) return;
    await Promise.all([refreshBranches(), refreshCommits()]);
}

async function refreshBranches() {
    if (!documentPath) return;
    try {
        const data = await apiFetch("/api/branches", { filepath: documentPath });
        if (!data.branches) return;

        const sel  = (id, placeholder) => {
            const el = document.getElementById(id);
            el.innerHTML = `<option value="">${placeholder}</option>`;
            return el;
        };

        const switchSel = sel("switchSelect", "Switch to branch...");
        const mergeSel  = sel("mergeSelect",  "Merge into current...");
        const diffA     = sel("diffBranchA",  "Branch A...");
        const diffB     = sel("diffBranchB",  "Branch B...");

        data.branches.forEach(b => {
            const isCurrent = b === data.current;
            switchSel.add(new Option(isCurrent ? `✓ ${b} (current)` : b, b, false, false));
            if (!isCurrent) mergeSel.add(new Option(b, b));
            diffA.add(new Option(b, b));
            diffB.add(new Option(b, b));
        });

        document.getElementById("currentBranch").innerText = data.current;
        document.getElementById("branchBadge").style.display = "block";
    } catch(e) {}
}

async function refreshCommits() {
    if (!documentPath) return;
    try {
        const data = await apiFetch("/api/commits", { filepath: documentPath });
        if (!data.commits) return;

        const rollSel  = document.getElementById("rollbackSelect");
        const diffComA = document.getElementById("diffCommitA");
        const diffComB = document.getElementById("diffCommitB");

        rollSel.innerHTML  = '<option value="">Select commit to restore...</option>';
        diffComA.innerHTML = '<option value="">Commit A...</option>';
        diffComB.innerHTML = '<option value="">Commit B...</option>';

        data.commits.forEach((c, i) => {
            const label = `${c.id}  ${c.timestamp}  ${c.message}`;
            rollSel.add(new Option(label, c.id));   // ALL commits available for rollback
            diffComA.add(new Option(label, c.id));
            diffComB.add(new Option(label, c.id));
        });
    } catch(e) {}
}

// ── Command runner ───────────────────────────────────────────────────────────
async function runCommand(commandStr) {
    if (!documentPath) {
        detectDocumentPath();
        showOutput("Document path not set. Enter it manually above.");
        return;
    }
    showOutput("Running: docgit " + commandStr + "\n...");
    try {
        const data = await apiFetch("/api/command", { command: commandStr, filepath: documentPath });
        showOutput(data.error ? "Error: " + data.error : (data.output || "(command completed)"));
    } catch(err) {
        showOutput("Connection failed:\n" + err.toString() +
                   "\n\nMake sure docgit_server is running.\n" +
                   "Open https://127.0.0.1:5000 in Edge if you see a cert warning.");
    }
}

// ── Shared fetch helper ──────────────────────────────────────────────────────
async function apiFetch(path, body) {
    const controller = new AbortController();
    const tid = setTimeout(() => controller.abort(), 15000);
    try {
        const r = await fetch("https://127.0.0.1:5000" + path, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
            signal: controller.signal
        });
        clearTimeout(tid);
        return await r.json();
    } catch(e) {
        clearTimeout(tid);
        throw e;
    }
}

function showOutput(text) {
    document.getElementById("output").innerText = text;
}
