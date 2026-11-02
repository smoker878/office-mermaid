import officeConfig from "office-addin-lint/config/eslint.config.mjs";

// The Office preset omits JSX/TSX from its file patterns. Keep its rules for React too.
export default officeConfig.map((config) =>
  config.files ? { ...config, files: ["**/*.{js,jsx,mjs,cjs,ts,tsx,cts,mts}"] } : config
);
