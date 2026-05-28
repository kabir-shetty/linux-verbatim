/**
 * Startup.js — OnlyOffice macro equivalent of Startup.bas (VBA)
 *
 * PORTING NOTES
 * ─────────────────────────────────────────────────────────────────────────────
 * VBA concept              → JS / OnlyOffice equivalent
 * ─────────────────────────────────────────────────────────────────────────────
 * GetSetting / SaveSetting → localStorage  (persistent key-value store)
 * ThisDocument.Variables   → document CustomProperties via the OO SDK
 * MsgBox(…, vbYesNo)       → window.confirm / Api.ShowMessageBox (OO)
 * Application.Documents    → no direct equivalent; tracked via a module-level
 *                            variable or sessionStorage
 * ActiveDocument.UpdateStyles → not exposed in OO macro API; stubbed
 * Keyboard shortcuts        → not configurable from OO macros; stubbed
 * Mac / Windows #If directives → navigator.platform checks
 * Protected View            → no equivalent in OO; guard removed
 * Audio recording           → no equivalent in OO; stubbed
 * UI.ShowForm               → replaced with Api.ShowMessageBox prompts
 * DateDiff / DatePart        → native JS Date arithmetic
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * 
// ─── Helpers: persistent settings (replaces VBA GetSetting / SaveSetting) ────

/**
 * Reads a setting from localStorage.
 * Key is namespaced as "AppName|Section|Key".
 * @param {string} app      e.g. "Verbatim"
 * @param {string} section  e.g. "Admin"
 * @param {string} key      e.g. "FirstRun"
 * @param {*}      def      default value if not found
 */
function getSetting(app, section, key, def) {
    var stored = localStorage.getItem(app + "|" + section + "|" + key);
    if (stored === null) return def;
    // Coerce stored strings back to booleans / numbers where sensible
    if (stored === "true")  return true;
    if (stored === "false") return false;
    var num = Number(stored);
    if (!isNaN(num) && stored.trim() !== "") return num;
    return stored;
}

/**
 * Writes a setting to localStorage.
 */
function saveSetting(app, section, key, value) {
    localStorage.setItem(app + "|" + section + "|" + key, String(value));
}

// ─── Helpers: document custom properties (replaces ThisDocument.Variables) ───

/**
 * Adds a custom property to the active document.
 * Equivalent to: ThisDocument.Variables.Add Name:="x", Value:="y"
 */
function addDocVariable(name, value) {
    try {
        var oDoc = Api.GetDocument();
        // Custom properties are the closest OO equivalent to VBA doc variables
        oDoc.SetCustomProperty(name, value);
    } catch (e) {
        // Non-fatal: log and continue
        console.log("addDocVariable failed for '" + name + "': " + e);
    }
}

// ─── Helpers: OS / platform detection (replaces VBA #If Mac) ─────────────────

function isMac() {
    return navigator.platform.toUpperCase().indexOf("MAC") >= 0;
}

// ─── Helpers: date arithmetic (replaces VBA DateDiff / DatePart) ──────────────

/** Returns the difference in whole days between two Date objects. */
function dateDiffDays(dateA, dateB) {
    var msPerDay = 1000 * 60 * 60 * 24;
    return Math.floor((dateB - dateA) / msPerDay);
}

/** Returns the day-of-week number (1 = Sun … 7 = Sat), matching VBA DatePart("w",…). */
function datePartWeekday(date) {
    return date.getDay() + 1; // JS: 0=Sun, VBA: 1=Sun
}

// ─── Globals (module-level state, replaces Globals module) ───────────────────

var Globals = {
    ActiveSpeechDoc:   "",
    RecordAudioToggle: false,
    DocumentCount:     0,          // incremented in autoOpen / decremented in autoClose

    InitializeGlobals: function () {
        // Place any one-time initialisation here.
        // Equivalent to Globals.InitializeGlobals in VBA.
        this.ActiveSpeechDoc   = sessionStorage.getItem("ActiveSpeechDoc") || "";
        this.RecordAudioToggle = getSetting("Verbatim", "Audio", "RecordToggle", false);
        console.log("Globals initialised.");
    }
};

// ─── Settings stubs (replaces Settings module) ───────────────────────────────

var Settings = {
    /**
     * Returns a version string.
     * Replace the hard-coded value with a dynamic lookup if needed.
     */
    GetVersion: function () {
        return getSetting("Verbatim", "Admin", "Version", "1.0.0");
    },

    /**
     * Checks for updates.
     * In VBA this opened a web dialog; here we show a message box stub.
     */
    UpdateCheck: function () {
        Api.ShowMessageBox(
            "Update Check",
            "Checking for Verbatim updates… (stub — implement HTTP fetch here)",
            0  // OK button only
        );
        saveSetting("Verbatim", "Profile", "LastUpdateCheck", new Date().toISOString());
    },

    /**
     * Resets keyboard shortcuts.
     * OnlyOffice macros cannot modify keyboard bindings — this is a stub.
     */
    ResetKeyboardShortcuts: function () {
        console.log("ResetKeyboardShortcuts: not supported in OO macros — skipped.");
    },

    /**
     * Removes Verbatim customisations from the Normal template.
     * Stub — implement if/when OO exposes template editing.
     */
    UnverbatimizeNormal: function () {
        console.log("UnverbatimizeNormal: stub — implement as needed.");
    },

    /**
     * Imports custom code from an external file.
     * Stub — implement using OO plugin / file API if needed.
     */
    ImportCustomCode: function (notify) {
        if (notify) {
            console.log("ImportCustomCode: stub — implement as needed.");
        }
    }
};

// ─── Troubleshooting stubs (replaces Troubleshooting module) ─────────────────

var Troubleshooting = {
    CheckDocx: function (notify) {
        // In VBA this warned if the file was saved as .doc instead of .docx.
        // OO documents are always in ODF/OOXML format; check is not required.
        console.log("CheckDocx: not applicable in OnlyOffice — skipped.");
        return true;
    },

    CheckSaveFormat: function (notify) {
        console.log("CheckSaveFormat: not applicable in OnlyOffice — skipped.");
        return true;
    },

    InstallCheckTemplateName: function () {
        // Return true (pass) — template management differs in OO
        return true;
    },

    InstallCheckTemplateLocation: function () {
        return true;
    }
};

// ─── View stubs (replaces View module) ───────────────────────────────────────

var View = {
    DefaultView: function () {
        // OnlyOffice does not expose view-mode switching in the macro API.
        console.log("DefaultView: not supported in OO macros — skipped.");
    }
};

// ─── Plugins stubs (replaces Plugins module) ─────────────────────────────────

var Plugins = {
    NavPaneCycle: function () {
        console.log("NavPaneCycle: not supported in OO macros — skipped.");
    }
};

// ─── UI stubs (replaces UI module) ───────────────────────────────────────────

var UI = {
    /**
     * Shows a named form.
     * VBA opened a UserForm; here we use a message box as a placeholder.
     * Replace each case with a real OO Plugin panel or dialog if required.
     */
    ShowForm: function (formName) {
        Api.ShowMessageBox(
            formName,
            "Opening form: " + formName + " (stub — implement real UI here)",
            0
        );
    }
};

// ─── Audio stubs (replaces Audio module) ─────────────────────────────────────

var Audio = {
    SaveRecord: function () {
        console.log("Audio.SaveRecord: not supported in OO macros — skipped.");
    }
};

// ═══════════════════════════════════════════════════════════════════════════════
// Core lifecycle functions  (direct equivalents of the VBA Startup module)
// ═══════════════════════════════════════════════════════════════════════════════

/**
 * Equivalent to VBA AutoOpen().
 * Call this when a document is opened.
 */
function AutoOpen() {
    Start();
}

/**
 * Equivalent to VBA AutoNew().
 * Call this when a new document is created from the template.
 */
function AutoNew() {
    try {
        // Add document variables / custom properties
        addDocVariable("Creator",        getSetting("Verbatim", "Profile", "Name", ""));
        addDocVariable("Team",           getSetting("Verbatim", "Profile", "SchoolName", ""));
        addDocVariable("VerbatimVersion",Settings.GetVersion());
        addDocVariable("OS",             navigator.platform);
        addDocVariable("OSVersion",      navigator.appVersion);
        // Note: there is no direct Word version equivalent in OO macros.
        addDocVariable("WordVersion",    "OnlyOffice");

        // Mark the document as saved (suppress the dirty flag)
        // OO does not expose Document.Saved directly; a no-op here.

        Start();
    } catch (e) {
        console.log("AutoNew error: " + e);
    }
}

/**
 * Equivalent to VBA AutoClose().
 * Call this when a document is about to close.
 */
function AutoClose() {
    try {
        var oDoc = Api.GetDocument();

        // If this was the active speech doc, clear the reference
        var docName = oDoc.GetFileName ? oDoc.GetFileName() : "";
        if (Globals.ActiveSpeechDoc === docName) {
            Globals.ActiveSpeechDoc = "";
            sessionStorage.setItem("ActiveSpeechDoc", "");
        }

        // Check doc format warnings (stubbed for OO)
        if (getSetting("Verbatim", "Admin", "SuppressDocCheck", false) === false) {
            Troubleshooting.CheckDocx({ Notify: true });
            Troubleshooting.CheckSaveFormat({ Notify: true });
        }

        // If this is the last document and audio recording is on, prompt to save
        // OO doesn't report document count easily; we track it ourselves.
        Globals.DocumentCount = Math.max(0, Globals.DocumentCount - 1);
        if (Globals.DocumentCount === 0 && Globals.RecordAudioToggle === true) {
            var stopRecording = window.confirm(
                'Audio recording appears to be active.\n\n' +
                'Stop and save recording now?\n\n' +
                'Click Cancel to discard the recording.'
            );
            if (stopRecording) {
                Audio.SaveRecord();
            }
        }
    } catch (e) {
        console.log("AutoClose error: " + e);
    }
}

/**
 * Equivalent to VBA Start().
 * The main initialisation routine called by both AutoOpen and AutoNew.
 */
function Start() {
    try {
        Globals.InitializeGlobals();
        Globals.DocumentCount++;

        // Set default view and navigation pane
        View.DefaultView();
        // DocumentMap (navigation pane) — not directly accessible in OO macro API

        // Refresh document styles from template if the setting is on
        if (getSetting("Verbatim", "Admin", "AutoUpdateStyles", true) === true) {
            // ActiveDocument.UpdateStyles() — not available in OO macro API; skipped.
            console.log("AutoUpdateStyles: style refresh not supported in OO macros.");
        }

        // Prevent Word/OO creating new styles automatically
        if (getSetting("Verbatim", "Admin", "SuppressStyleChecks", false) === false) {
            // Options.AutoFormatAsYouTypeDefineStyles — no OO equivalent; skipped.
            console.log("SuppressStyleChecks: auto-style restriction not supported in OO macros.");
        }

        // Nav pane cycle on startup
        if (getSetting("Verbatim", "View", "NPCStartup", false) === true) {
            Plugins.NavPaneCycle();
        }

        // ── First run ──────────────────────────────────────────────────────────
        if (getSetting("Verbatim", "Admin", "FirstRun", true) === true) {
            FirstRun();
            return;
        }

        // ── Install checks (only on the very first document opened) ───────────
        if (
            getSetting("Verbatim", "Admin", "SuppressInstallChecks", false) === false &&
            Globals.DocumentCount === 1
        ) {
            var nameOk     = Troubleshooting.InstallCheckTemplateName();
            var locationOk = Troubleshooting.InstallCheckTemplateLocation();

            if (!nameOk || !locationOk) {
                var openTroubleshooter = window.confirm(
                    "Verbatim appears to be installed incorrectly.\n\n" +
                    "Would you like to open the Troubleshooter?\n\n" +
                    "(This message can be suppressed in the Verbatim settings.)"
                );
                if (openTroubleshooter) {
                    UI.ShowForm("Troubleshooter");
                    return;
                }
            }
        }

        // ── Weekly update check (Wednesdays) ──────────────────────────────────
        if (getSetting("Verbatim", "Profile", "AutomaticUpdates", true) === true) {
            var lastCheck    = getSetting("Verbatim", "Profile", "LastUpdateCheck", null);
            var lastCheckDate = lastCheck ? new Date(lastCheck) : new Date(0);
            var now          = new Date();

            if (dateDiffDays(lastCheckDate, now) > 6) {
                // datePartWeekday: 4 = Wednesday (VBA) → getDay() 3 = Wednesday (JS 0-indexed)
                if (now.getDay() === 3) {   // Wednesday
                    Settings.UpdateCheck();
                    return;
                }
            }
        }

        // ── Custom code import ─────────────────────────────────────────────────
        if (getSetting("Verbatim", "Admin", "ImportCustomCode", false) === true) {
            Settings.ImportCustomCode(true);
        }

        // ── Mac: reset keybindings if PC shortcuts are detected ────────────────
        if (isMac()) {
            // VBA iterated KeyBindings to check for "Shift+2" mapped to PasteText.
            // OO macros cannot inspect or set key bindings — stub only.
            console.log("Mac keybinding check: not supported in OO macros — skipped.");
        }

    } catch (e) {
        console.log("Start() error: " + e);
    }
}

/**
 * Equivalent to VBA FirstRun().
 * Runs once to initialise settings and launch the setup wizard.
 */
function FirstRun() {
    // Mark first run as done
    saveSetting("Verbatim", "Admin", "FirstRun", false);

    // Clear out old installation artefacts
    Settings.UnverbatimizeNormal();

    // Remove old / sensitive registry keys (now localStorage entries)
    saveSetting("Verbatim", "Main", "TabroomUsername", "");
    saveSetting("Verbatim", "Main", "TabroomPassword", "");
    saveSetting("Verbatim", "Main", "GmailUsername",   "");
    saveSetting("Verbatim", "Main", "GmailPassword",   "");

    // Set up keyboard shortcuts
    Settings.ResetKeyboardShortcuts();

    // Launch setup wizard
    UI.ShowForm("Setup");
}

// ─── OnlyOffice macro entry point ─────────────────────────────────────────────
/**
 * "Main" is the function OnlyOffice calls when you run the macro.
 * Change the call below to whichever lifecycle event you need to test.
 */
function Main() {
    AutoOpen();
}