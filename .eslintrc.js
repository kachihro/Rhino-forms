require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/react'],
  parserOptions: { tsconfigRootDir: __dirname },
  overrides: [
    {
      files: ['*.ts', '*.tsx'],
      parser: '@typescript-eslint/parser',
      parserOptions: {
        project: './tsconfig.json',
        ecmaVersion: 2018,
        sourceType: 'module'
      },
      rules: {
        '@rushstack/no-new-null': 1,
        '@typescript-eslint/no-explicit-any': 0,
        '@typescript-eslint/no-floating-promises': 1,
        '@typescript-eslint/explicit-function-return-type': 0,
        '@microsoft/spfx/import-requires-chunk-name': 1,
        '@microsoft/spfx/no-require-ensure': 2,
        '@microsoft/spfx/pair-react-dom-render-unmount': 1
      }
    }
  ]
};
