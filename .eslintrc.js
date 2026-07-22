require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/default'],
  // Jest specs are compiled and checked by ts-jest (see jest.config.js), not by
  // the SPFx tsc/lint pipeline, so they are excluded from tsconfig and here too.
  ignorePatterns: ['src/**/__tests__/**', '**/*.test.ts'],
  parserOptions: { tsconfigRootDir: __dirname },
  overrides: [
    {
      files: ['*.ts', '*.tsx'],
      parser: '@typescript-eslint/parser',
      parserOptions: {
        project: './tsconfig.json',
        ecmaFeatures: { jsx: true },
        tsconfigRootDir: __dirname
      }
    }
  ]
};
