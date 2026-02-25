/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_GEMINI_API_KEY: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
```

The file structure should be:
```;
src / vite - env.d.ts;
App.tsx;
main.tsx;
index.css;
