// Drag one file onto this script to generate <filename>.data.txt with metadata info.
// Version: 2025.002

(function main() {
    var shellApp = WScript.CreateObject("Shell.Application");
    var shell = WScript.CreateObject("WScript.Shell");
    var fso = WScript.CreateObject("Scripting.FileSystemObject");

    if (WScript.Arguments.Count() !== 1) {
        WScript.Echo("Please drag one file onto this script.");
        WScript.Quit();
    }

    var fullPath = WScript.Arguments.Item(0);
    if (!fso.FileExists(fullPath)) {
        WScript.Echo("File not found: " + fullPath);
        WScript.Quit();
    }

    var folderPath = fso.GetParentFolderName(fullPath) + "\\";
    var fileName = fso.GetFileName(fullPath);
    var outPath = folderPath + fileName + ".data.txt";

    var log = createLogFile(fso, outPath);
    if (!log) return;

    logShellMeta(shellApp, log, folderPath, fileName);
    logPowerShellMeta(shell, log, fullPath);

    log.WriteLine("\nDiagnostics complete.");
    log.Close();

    WScript.Echo("Metadata output done.");
    shell.Run('"' + outPath + '"');
})();

// === FUNCTIONS ===

function createLogFile(fso, path) {
    try {
        return fso.CreateTextFile(path, true, true);
    } catch (e) {
        WScript.Echo("Cannot create output file: " + path + "\n" + e.message);
        return null;
    }
}

function logShellMeta(shellApp, log, folderPath, fileName) {
    log.WriteLine("[Shell.NameSpace] Listing headers and file values\n");

    var ns = shellApp.NameSpace(folderPath);
    if (!ns) {
        log.WriteLine("ERROR: Unable to open Shell.NameSpace for folder.");
        return;
    }

    var fileShell = ns.ParseName(fileName);
    if (!fileShell) {
        log.WriteLine("ERROR: ns.ParseName returned null for file: " + fileName);
        return;
    }

    log.WriteLine("---- Available headers (index : header) ----");
    var hdrs = {};
    for (var i = 0; i < 500; i++) {
        try {
            var h = ns.GetDetailsOf(null, i);
            if (h && String(h).trim().length > 0) {
                hdrs[i] = String(h);
                log.WriteLine(i + " : " + h);
            }
        } catch (e) {}
    }

    log.WriteLine("\n---- Values for this file (index : header = value) ----");
    for (var j = 0; j < 500; j++) {
        try {
            var headerName = hdrs[j] ? hdrs[j] : "(header unknown)";
            var v = ns.GetDetailsOf(fileShell, j);
            if (v === null || v === undefined) v = "";
            log.WriteLine(j + " : " + headerName + " = [" + String(v) + "]");
        } catch (e2) {
            log.WriteLine(j + " : ERROR reading index " + j + " - " + e2.message);
        }
    }

    log.WriteLine("\n");
}

function logPowerShellMeta(shell, log, fullPath) {
    log.WriteLine("[PowerShell/System.Drawing] Reading Image.PropertyItems and EXIF tags...");

    var psPath = fullPath.replace(/'/g, "''");
    var ps = "";
    ps += "$path = '" + psPath + "'; ";
    ps += "try { Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue; ";
    ps += "$img = [System.Drawing.Image]::FromFile($path); ";
    ps += "if ($img -eq $null) { Write-Output 'PS_ERROR:Failed to load image' } else { ";
    ps += "$props = $img.PropertyItems; ";
    ps += "if ($props) { foreach ($p in $props) { ";
    ps += " $hex = '{0:X}' -f $p.Id; ";
    ps += " $enc = New-Object System.Text.ASCIIEncoding; ";
    ps += " $raw = $enc.GetString($p.Value) -replace [char]0, '' ; ";
    ps += " Write-Output ('PROP_ID:' + $p.Id + ' HEX:' + $hex + ' LEN:' + $p.Len + ' Type:' + $p.Type + ' VAL:' + $raw); ";
    ps += " } Write-Output 'DONE_PROPS'; } else { Write-Output 'NO_PROPERTY_ITEMS'; } } ";
    ps += "} catch { Write-Output ('PS_ERROR:' + $_.Exception.Message) } finally { if ($img) { $img.Dispose() } }";

    var cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command \"" + ps + "\"";

    try {
        var exec = shell.Exec(cmd);
        var out = readAll(exec.StdOut);
        var err = readAll(exec.StdErr);
        log.WriteLine("PowerShell stdout:\n" + out);
        if (err && err.trim().length > 0) log.WriteLine("PowerShell stderr:\n" + err);
        log.WriteLine("PowerShell exitCode: " + exec.ExitCode);
    } catch (psEx) {
        log.WriteLine("ERROR launching PowerShell: " + psEx.message);
    }

    log.WriteLine("\n");
}

function readAll(stream) {
    var result = "";
    try {
        while (!stream.AtEndOfStream) {
            result += stream.ReadLine() + "\n";
        }
    } catch (e) {}
    return result;
}
