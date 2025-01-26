import { FlatCompat } from "@eslint/eslintrc";
import globals from "globals";
import pluginJs from "@eslint/js";
import es5 from "eslint-plugin-es5";
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
  ...compat.extends("plugin:es5/no-es2015"),
  eslintConfigPrettier,
  pluginJs.configs.recommended,
  { files: ["src/*.js"] },
  {
    languageOptions: {
      globals: { ...globals.browser },
      ecmaVersion: 5,
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
