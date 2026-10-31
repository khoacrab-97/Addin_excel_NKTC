/** Đường dẫn tài liệu và file mẫu hiển thị trong add-in. */
const REPO_DOCS = "https://github.com/khoacrab-97/Addin_excel_NKTC/blob/main/docs";

export const LINKS = {
  installGuide: `${REPO_DOCS}/HUONG_DAN_CAI_DAT.md`,
  updateGuide: `${REPO_DOCS}/HUONG_DAN_CAI_DAT.md#5-c%E1%BA%ADp-nh%E1%BA%ADt`,
  formTargets: `${REPO_DOCS}/BUSINESS_SPEC.md#f23-bi%E1%BB%83u-m%E1%BA%ABu-excel-t%C3%ACnh-h%C3%ACnh-thi-c%C3%B4ng-h%C3%A0ng-ng%C3%A0y-templatesmau_nktcxlsx`,
  /** File mẫu được phát hành cùng add-in (đường dẫn tương đối với taskpane.html). */
  formTemplate: "templates/Mau_NKTC.xlsx",
  sampleSchedule: "templates/Mau_TienDo.xlsx",
} as const;
