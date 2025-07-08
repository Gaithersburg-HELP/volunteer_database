import { FlatCompat } from "@eslint/eslintrc";
import globals from "globals";
import pluginJs from "@eslint/js";
import eslintConfigPrettier from "eslint-config-prettier";
import path from "path";
import { fileURLToPath } from "url";

// mimic CommonJS variables -- not needed if using CommonJS
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const compat = new FlatCompat({
  baseDirectory: __dirname
});

/** @type {import('eslint').Linter.Config[]} */
export default [
  eslintConfigPrettier,
  pluginJs.configs.recommended,
  { files: ["src/*.js"] },
  {
    languageOptions: {
      globals: { ...globals.browser },
      ecmaVersion: 2020,
      sourceType: "script"
    }
  },
  { ignores: ["node_modules/**", "src/lib/*", "eslint.config.mjs"] },
  {
    rules: {
      eqeqeq: "error",
      strict: "off",
      "jsdoc/require-jsdoc": "off",
      "no-unused-vars": "off",
      "no-undef": "off",
      "no-labels": "off",
      "default-case": "off",
      "no-underscore-dangle": "off"
    }
  }
];
