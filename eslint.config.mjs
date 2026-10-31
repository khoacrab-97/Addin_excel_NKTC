import js from "@eslint/js";
import prettier from "eslint-config-prettier";
import reactHooks from "eslint-plugin-react-hooks";
import globals from "globals";
import tseslint from "typescript-eslint";

export default tseslint.config(
  { ignores: ["dist/", "node_modules/", "coverage/"] },
  js.configs.recommended,
  {
    files: ["src/**/*.{ts,tsx}"],
    extends: [...tseslint.configs.strictTypeChecked],
    languageOptions: {
      globals: globals.browser,
      parserOptions: { projectService: true, tsconfigRootDir: import.meta.dirname },
    },
    plugins: { "react-hooks": reactHooks },
    rules: {
      ...reactHooks.configs.recommended.rules,
      // Biểu thức template có số là cách viết phổ biến (ví dụ `1.${minor}`).
      "@typescript-eslint/restrict-template-expressions": ["error", { allowNumber: true }],
    },
  },
  {
    // Tầng core phải độc lập với Office.js / Excel (CLAUDE.md §3).
    files: ["src/core/**/*.ts"],
    rules: {
      "no-restricted-globals": [
        "error",
        { name: "Office", message: "core không được phụ thuộc Office.js." },
        { name: "Excel", message: "core không được phụ thuộc Excel API." },
      ],
      "no-restricted-imports": [
        "error",
        {
          patterns: [
            { group: ["**/excel/**", "**/taskpane/**"], message: "core không được import tầng excel/taskpane." },
          ],
        },
      ],
    },
  },
  {
    files: ["*.config.{js,mjs}", "webpack.config.js"],
    languageOptions: { globals: globals.node, sourceType: "commonjs" },
  },
  {
    // Excel giả lập: script trình duyệt viết theo ES5 (không qua bundler).
    files: ["scripts/smoke/fake-excel.js"],
    languageOptions: { globals: globals.browser, sourceType: "script" },
  },
  {
    files: ["scripts/smoke/run-*.js"],
    languageOptions: { globals: globals.node, sourceType: "commonjs" },
  },
  {
    files: ["*.mjs"],
    languageOptions: { sourceType: "module" },
  },
  prettier,
);
