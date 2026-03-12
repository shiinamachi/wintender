const GUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const CANONICAL_IDENTITY_PATTERN = /^(path|process|guid|runtime):(.*)$/i;
const WINDOWS_DRIVE_PATH_PATTERN = /^[a-z]:[\\/]/i;
const WINDOWS_UNC_PATH_PATTERN = /^\\\\[^\\]+\\[^\\]+/;
const WINDOW_HANDLE_ICON_ID_PATTERN = /^(\d+):(.+)$/;

function normalizeText(text) {
  return typeof text === "string" ? text.trim() : "";
}

function normalizeWindowsPath(path) {
  return normalizeText(path).replace(/\//g, "\\").replace(/\\+/g, "\\").toLowerCase();
}

function looksLikeWindowsPath(value) {
  return (
    WINDOWS_DRIVE_PATH_PATTERN.test(value) ||
    WINDOWS_UNC_PATH_PATTERN.test(value) ||
    value.includes("\\") ||
    value.includes("/")
  );
}

function normalizeIdentityValue(kind, value) {
  const rawValue = normalizeText(value);

  if (!rawValue) {
    return "";
  }

  if (kind === "path") {
    const normalizedPath = normalizeWindowsPath(rawValue);
    return normalizedPath ? `path:${normalizedPath}` : "";
  }

  return `${kind}:${rawValue.toLowerCase()}`;
}

function getResolvedProcessInfoForIcon(
  icon,
  resolvedProcessInfoByIconId = new Map(),
) {
  const iconId = normalizeText(icon?.id);

  if (!iconId || typeof resolvedProcessInfoByIconId?.get !== "function") {
    return null;
  }

  return resolvedProcessInfoByIconId.get(iconId) ?? null;
}

function escapePowerShellSingleQuoted(value) {
  return String(value).replace(/'/g, "''");
}

export function parseTrayIconId(iconId) {
  const rawId = normalizeText(iconId);

  if (!rawId) {
    return {
      kind: "unknown",
      rawId: "",
      windowHandle: "",
    };
  }

  if (GUID_PATTERN.test(rawId)) {
    return {
      kind: "guid",
      rawId,
      windowHandle: "",
    };
  }

  const windowHandleMatch = rawId.match(WINDOW_HANDLE_ICON_ID_PATTERN);

  if (windowHandleMatch) {
    return {
      kind: "window_handle",
      rawId,
      windowHandle: windowHandleMatch[1],
    };
  }

  return {
    kind: "runtime",
    rawId,
    windowHandle: "",
  };
}

export function normalizeTrayAppIdentity(identity) {
  const rawIdentity = normalizeText(identity);

  if (!rawIdentity) {
    return "";
  }

  const canonicalMatch = rawIdentity.match(CANONICAL_IDENTITY_PATTERN);

  if (canonicalMatch) {
    return normalizeIdentityValue(
      canonicalMatch[1].toLowerCase(),
      canonicalMatch[2],
    );
  }

  if (looksLikeWindowsPath(rawIdentity)) {
    return normalizeIdentityValue("path", rawIdentity);
  }

  if (GUID_PATTERN.test(rawIdentity)) {
    return normalizeIdentityValue("guid", rawIdentity);
  }

  if (WINDOW_HANDLE_ICON_ID_PATTERN.test(rawIdentity)) {
    return normalizeIdentityValue("runtime", rawIdentity);
  }

  return normalizeIdentityValue("process", rawIdentity);
}

export function sanitizePinnedTrayApps(values) {
  if (!Array.isArray(values)) {
    return [];
  }

  return [...new Set(values.map((value) => normalizeTrayAppIdentity(value)).filter(Boolean))];
}

export function extractLegacyPinnedTrayTooltips(rawSettings) {
  if (!Array.isArray(rawSettings?.pinnedTrayTooltips)) {
    return [];
  }

  return [...new Set(rawSettings.pinnedTrayTooltips.map((value) => normalizeText(value)).filter(Boolean))];
}

export function resolveTrayAppIdentity(
  icon,
  resolvedProcessInfo = null,
) {
  const providerProcessPath = normalizeText(icon?.processPath);
  const resolvedProcessPath = normalizeText(resolvedProcessInfo?.processPath);

  if (providerProcessPath || resolvedProcessPath) {
    return normalizeIdentityValue(
      "path",
      providerProcessPath || resolvedProcessPath,
    );
  }

  const providerProcessName = normalizeText(icon?.processName);
  const resolvedProcessName = normalizeText(resolvedProcessInfo?.processName);

  if (providerProcessName || resolvedProcessName) {
    return normalizeIdentityValue(
      "process",
      providerProcessName || resolvedProcessName,
    );
  }

  const parsedIconId = parseTrayIconId(icon?.id);

  if (parsedIconId.kind === "guid") {
    return normalizeIdentityValue("guid", parsedIconId.rawId);
  }

  if (parsedIconId.rawId) {
    return normalizeIdentityValue("runtime", parsedIconId.rawId);
  }

  return "";
}

export function formatTrayAppIdentity(identity) {
  const normalizedIdentity = normalizeTrayAppIdentity(identity);

  if (!normalizedIdentity) {
    return "";
  }

  return normalizedIdentity.replace(CANONICAL_IDENTITY_PATTERN, (_, __, value) =>
    value,
  );
}

export function getDetectedTrayApps(
  icons,
  resolvedProcessInfoByIconId = new Map(),
) {
  const detectedApps = [];
  const seenIdentities = new Set();

  for (const icon of icons ?? []) {
    const fallbackIdentity = resolveTrayAppIdentity(icon, null);
    const identity = resolveTrayAppIdentity(
      icon,
      getResolvedProcessInfoForIcon(icon, resolvedProcessInfoByIconId),
    );

    if (!identity || seenIdentities.has(identity)) {
      continue;
    }

    seenIdentities.add(identity);
    const tooltip = normalizeText(icon?.tooltip);
    const label = formatTrayAppIdentity(identity);

    detectedApps.push({
      identity,
      fallbackIdentity,
      label,
      tooltip,
      displayLabel: tooltip || label,
    });
  }

  return detectedApps.sort((left, right) =>
    left.displayLabel.localeCompare(right.displayLabel),
  );
}

export function migrateLegacyPinnedTrayTooltips(
  legacyTooltips,
  icons,
  resolvedProcessInfoByIconId = new Map(),
) {
  const nextPinnedTrayApps = [];
  const seenIdentities = new Set();
  const sanitizedLegacyTooltips = [...new Set(
    (Array.isArray(legacyTooltips) ? legacyTooltips : [])
      .map((value) => normalizeText(value))
      .filter(Boolean),
  )];

  for (const tooltip of sanitizedLegacyTooltips) {
    const matchedIcon = (icons ?? []).find(
      (icon) => normalizeText(icon?.tooltip) === tooltip,
    );

    if (!matchedIcon) {
      continue;
    }

    const identity = resolveTrayAppIdentity(
      matchedIcon,
      getResolvedProcessInfoForIcon(
        matchedIcon,
        resolvedProcessInfoByIconId,
      ),
    );

    if (!identity || seenIdentities.has(identity)) {
      continue;
    }

    seenIdentities.add(identity);
    nextPinnedTrayApps.push(identity);
  }

  return nextPinnedTrayApps;
}

export function upgradePinnedTrayApps(
  pinnedTrayApps,
  icons,
  resolvedProcessInfoByIconId = new Map(),
) {
  const nextPinnedTrayApps = [...sanitizePinnedTrayApps(pinnedTrayApps)];

  for (const icon of icons ?? []) {
    const fallbackIdentity = resolveTrayAppIdentity(icon, null);
    const resolvedIdentity = resolveTrayAppIdentity(
      icon,
      getResolvedProcessInfoForIcon(icon, resolvedProcessInfoByIconId),
    );

    if (
      !fallbackIdentity ||
      !resolvedIdentity ||
      fallbackIdentity === resolvedIdentity
    ) {
      continue;
    }

    const fallbackIndex = nextPinnedTrayApps.indexOf(fallbackIdentity);

    if (fallbackIndex === -1) {
      continue;
    }

    nextPinnedTrayApps.splice(fallbackIndex, 1, resolvedIdentity);
  }

  return sanitizePinnedTrayApps(nextPinnedTrayApps);
}

export function getTrayAppLookupIds(
  icons,
  resolvedProcessInfoByIconId = new Map(),
) {
  const iconIds = [];
  const seenIconIds = new Set();

  for (const icon of icons ?? []) {
    const iconId = normalizeText(icon?.id);

    if (!iconId || seenIconIds.has(iconId)) {
      continue;
    }

    seenIconIds.add(iconId);

    if (
      normalizeText(icon?.processPath) ||
      normalizeText(icon?.processName) ||
      typeof resolvedProcessInfoByIconId?.get !== "function"
    ) {
      continue;
    }

    const cachedProcessInfo = resolvedProcessInfoByIconId.get(iconId);

    if (cachedProcessInfo) {
      continue;
    }

    if (parseTrayIconId(iconId).kind === "window_handle") {
      iconIds.push(iconId);
    }
  }

  return iconIds;
}

export function buildTrayAppLookupCommand(iconIds) {
  const lookupTargets = [...new Set(
    (Array.isArray(iconIds) ? iconIds : [])
      .map((iconId) => parseTrayIconId(iconId))
      .filter((entry) => entry.kind === "window_handle")
      .map((entry) => entry.rawId),
  )];

  if (!lookupTargets.length) {
    return "";
  }

  const targetsLiteral = lookupTargets
    .map((iconId) => {
      const parsedIconId = parseTrayIconId(iconId);
      return `[pscustomobject]@{Id='${escapePowerShellSingleQuoted(
        parsedIconId.rawId,
      )}';Handle='${escapePowerShellSingleQuoted(
        parsedIconId.windowHandle,
      )}'}`;
    })
    .join(",");

  return [
    "[Console]::OutputEncoding=[System.Text.Encoding]::UTF8",
    "if (-not ('WintenderNative' -as [type])) { Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class WintenderNative { [DllImport(\"user32.dll\")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }' }",
    `$targets=@(${targetsLiteral})`,
    "$items=@()",
    "foreach($target in $targets){",
    "$processIdForTrayLookup=[uint32]0",
    "[void][WintenderNative]::GetWindowThreadProcessId([IntPtr]([int64]$target.Handle), [ref]$processIdForTrayLookup)",
    "if($processIdForTrayLookup -le 0){ continue }",
    "try {",
    "$process=Get-Process -Id $processIdForTrayLookup -ErrorAction Stop",
    "$path=$process.Path",
    "$name=$process.ProcessName",
    "if($name -and $name -notmatch '\\.exe$'){ $name = \"$name.exe\" }",
    "$items += [pscustomobject]@{ id=$target.Id; processPath=$path; processName=$name }",
    "} catch { }",
    "}",
    "@{ items = $items } | ConvertTo-Json -Compress -Depth 4",
  ].join("; ");
}

export function parseTrayAppLookupResponse(stdout) {
  const rawOutput = normalizeText(stdout);

  if (!rawOutput) {
    return new Map();
  }

  try {
    const parsed = JSON.parse(rawOutput);
    const items = Array.isArray(parsed?.items)
      ? parsed.items
      : Array.isArray(parsed)
        ? parsed
        : [];
    const resolvedProcessInfoByIconId = new Map();

    for (const item of items) {
      const iconId = normalizeText(item?.id);

      if (!iconId) {
        continue;
      }

      resolvedProcessInfoByIconId.set(iconId, {
        processPath: normalizeText(item?.processPath),
        processName: normalizeText(item?.processName),
      });
    }

    return resolvedProcessInfoByIconId;
  } catch (error) {
    console.error("Failed to parse tray app lookup response", error);
    return new Map();
  }
}
