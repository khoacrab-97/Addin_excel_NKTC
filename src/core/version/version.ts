/** Thông tin phiên bản phát hành đọc từ `version.json` trên máy chủ. */
export interface ReleaseInfo {
  version: string;
  /** Thời điểm build (ISO 8601), nếu có. */
  builtAt?: string;
}

const VERSION_PATTERN = /^(\d+)\.(\d+)\.(\d+)$/;

/** Tách `x.y.z` thành 3 số; `undefined` nếu sai định dạng. */
export function parseVersion(version: string): [number, number, number] | undefined {
  const match = VERSION_PATTERN.exec(version.trim());
  if (!match) return undefined;
  return [Number(match[1]), Number(match[2]), Number(match[3])];
}

/** So sánh 2 phiên bản `x.y.z`: âm nếu a < b, 0 nếu bằng, dương nếu a > b. Sai định dạng → `undefined`. */
export function compareVersions(a: string, b: string): number | undefined {
  const left = parseVersion(a);
  const right = parseVersion(b);
  if (!left || !right) return undefined;
  for (let i = 0; i < 3; i++) {
    const diff = (left[i] ?? 0) - (right[i] ?? 0);
    if (diff !== 0) return diff;
  }
  return 0;
}

/** Đọc nội dung `version.json` (dữ liệu bên ngoài → kiểm tra đầy đủ). */
export function parseReleaseInfo(text: string): ReleaseInfo | undefined {
  let raw: unknown;
  try {
    raw = JSON.parse(text);
  } catch {
    return undefined;
  }
  if (typeof raw !== "object" || raw === null) return undefined;
  const { version, builtAt } = raw as Record<string, unknown>;
  if (typeof version !== "string" || parseVersion(version) === undefined) return undefined;
  const info: ReleaseInfo = { version: version.trim() };
  if (typeof builtAt === "string" && builtAt !== "") info.builtAt = builtAt;
  return info;
}

export type UpdateStatus =
  | { kind: "latest"; current: string; latest: string }
  | { kind: "available"; current: string; latest: string }
  | { kind: "newer-than-release"; current: string; latest: string }
  | { kind: "unknown"; current: string };

/** Kết luận trạng thái cập nhật từ phiên bản đang chạy và phiên bản phát hành (P3.2). */
export function evaluateUpdate(current: string, release: ReleaseInfo | undefined): UpdateStatus {
  if (!release) return { kind: "unknown", current };
  const diff = compareVersions(current, release.version);
  if (diff === undefined) return { kind: "unknown", current };
  if (diff < 0) return { kind: "available", current, latest: release.version };
  if (diff > 0) return { kind: "newer-than-release", current, latest: release.version };
  return { kind: "latest", current, latest: release.version };
}
