import { useEffect, useState } from "react";
import { type UpdateStatus, evaluateUpdate } from "../../../core/version/version";
import { fetchReleaseInfo } from "../../../version/updateCheck";
import { APP_VERSION } from "../../../version/version";

/** Kiểm tra cập nhật: tự chạy một lần khi mở add-in (P3.3) và chạy lại theo yêu cầu (P3.2). */
export function useUpdateCheck() {
  const [status, setStatus] = useState<UpdateStatus>();
  const [checking, setChecking] = useState(false);

  const check = (): Promise<void> =>
    fetchReleaseInfo().then((release) => {
      setStatus(evaluateUpdate(APP_VERSION, release));
    });

  useEffect(() => {
    void fetchReleaseInfo().then((release) => {
      setStatus(evaluateUpdate(APP_VERSION, release));
    });
  }, []);

  const recheck = (): void => {
    setChecking(true);
    void check().finally(() => {
      setChecking(false);
    });
  };

  return { status, checking, recheck };
}
