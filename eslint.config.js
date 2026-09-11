import typescriptEslint from "@typescript-eslint/eslint-plugin";
import eslintConfigPrettier from "eslint-config-prettier";

export default [
    ...typescriptEslint.configs["flat/recommended"],
    eslintConfigPrettier,
    {
        files: ["src/**/*.ts"],
        languageOptions: {
            parserOptions: {
                ecmaFeatures: {
                    jsx: true
                }
            }
        }
    }
];