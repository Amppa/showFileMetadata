// Drag one file onto this script to generate <filename>.data.txt with metadata info.
// Version: 2025.004

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
    var ext = getExtension(fileName);
    var outPath = folderPath + fileName + ".data.txt";

    var log = createLogFile(fso, outPath);
    if (!log) return;

    logHeader(log, fileName, ext);

    // 1) PowerShell metadata first (image or video)
    if (isSupportedBySystemDrawing(ext)) {
        logPowerShellImageMeta(shell, log, fullPath);
    } else if (isVideoExtension(ext)) {
        logPowerShellVideoMeta(shell, log, fullPath);
    } else {
        log.WriteLine("[PowerShell] Skipped for extension: " + ext);
        log.WriteLine("");
    }

    // 2) Shell.NameSpace metadata next
    logShellMeta(shellApp, log, folderPath, fileName);

    log.WriteLine("\nDiagnostics complete.");
    log.Close();

    WScript.Echo("Metadata output done.");
    shell.Run('"' + outPath + '"');
})();

// === CONFIG ===

function getExtension(filename) {
    var idx = filename.lastIndexOf(".");
    if (idx === -1) return "";
    return filename.substring(idx + 1).toLowerCase();
}

function isImageExtension(ext) {
    var imgs = ["jpg","jpeg","png","bmp","gif","tiff","tif","heic","heif"];
    for (var i = 0; i < imgs.length; i++) {
        if (imgs[i] === ext) return true;
    }
    return false;
}

function isVideoExtension(ext) {
    var vids = ["mp4","mov","mkv","avi","wmv","flv","webm"];
    for (var i = 0; i < vids.length; i++) {
        if (vids[i] === ext) return true;
    }
    return false;
}

function isSupportedBySystemDrawing(ext) {
    var supported = ["jpg","jpeg","png","bmp","gif","tiff","tif"];
    for (var i = 0; i < supported.length; i++) {
        if (supported[i] === ext) return true;
    }
    return false;
}

// === CORE LOGGING ===

function createLogFile(fso, path) {
    try {
        return fso.CreateTextFile(path, true, true);
    } catch (e) {
        WScript.Echo("Cannot create output file: " + path + "\n" + e.message);
        return null;
    }
}

function logHeader(log, fileName, ext) {
    log.WriteLine("=== Metadata dump for: " + fileName + " ===");
    log.WriteLine("File type (extension): " + (ext || "(none)"));
    log.WriteLine("========================================================\n");
}

// === PowerShell for images ===
function logPowerShellImageMeta(shell, log, fullPath) {
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

    runPowerShell(shell, log, ps);
}

// === PowerShell for videos ===
function logPowerShellVideoMeta(shell, log, fullPath) {
    log.WriteLine("[PowerShell/Shell.Application] Reading video metadata...");

    var psPath = fullPath.replace(/'/g, "''");
    var ps = "";
    ps += "$file = '" + psPath + "'; ";
    ps += "$folder = Split-Path $file; $name = Split-Path $file -Leaf; ";
    ps += "$sh = New-Object -ComObject Shell.Application; ";
    ps += "$ns = $sh.Namespace($folder); ";
    ps += "$item = $ns.ParseName($name); ";
    ps += "if ($item -eq $null) { Write-Output 'PS_ERROR:Cannot parse file'; exit } ";
    ps += "for ($i=0; $i -lt 400; $i++) { ";
    ps += " $key = $ns.GetDetailsOf($null, $i); ";
    ps += " if (![string]::IsNullOrWhiteSpace($key)) { ";
    ps += " $val = $ns.GetDetailsOf($item, $i); ";
    ps += " if ($val) { Write-Output ($i.ToString('000') + ' : ' + $key + ' = ' + $val) } ";
    ps += " } ";
    ps += "}";

    runPowerShell(shell, log, ps);
}

// === Shared PowerShell runner ===
function runPowerShell(shell, log, ps) {
    var cmd = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "' + ps + '"';
    try {
        var exec = shell.Exec(cmd);
        var out = readAll(exec.StdOut);
        var err = readAll(exec.StdErr);
        log.WriteLine("PowerShell stdout:\n" + out);
        if (err && err.trim().length > 0) log.WriteLine("PowerShell stderr:\n" + err);
        log.WriteLine("PowerShell exitCode: " + exec.ExitCode);
    } catch (ex) {
        log.WriteLine("ERROR launching PowerShell: " + ex.message);
    }
    log.WriteLine("");
}

// === Shell section ===
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

// === Stream reader ===
function readAll(stream) {
    var result = "";
    try {
        while (!stream.AtEndOfStream) {
            result += stream.ReadLine() + "\n";
        }
    } catch (e) {}
    return result;
}
