// eslint.config.js
import js from "@eslint/js";
import { defineConfig } from "eslint/config";
import prettierPlugin from "eslint-plugin-prettier";
import prettierConfig from "eslint-config-prettier";

export default defineConfig([
  {
    files: ["**/*.js"],
    plugins: {
      js,
      prettier: prettierPlugin
    },
    extends: ["js/recommended"],
    rules: {
      semi: "error",
      "prefer-const": "error",
      eqeqeq: "error",
      strict: "off",
      "jsdoc/require-jsdoc": "off",
      "no-unused-vars": "off",
      "no-undef": "off",
      "no-labels": "off",
      "default-case": "off",
      "no-underscore-dangle": "off",
      ...prettierConfig.rules
    }
  }
]);
