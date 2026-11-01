/** Configuração usada para gerar o CSS embutido no index.html */
module.exports = {
  content: ['./index.html'],
  darkMode: 'class',
  theme: {
    extend: {
      fontFamily: {
        sans: ['"Space Grotesk"', 'system-ui', 'sans-serif'],
        mono: ['"JetBrains Mono"', 'ui-monospace', 'monospace'],
      },
      colors: {
        ink: { 950: '#07070c', 900: '#0d0d16', 800: '#141421', 700: '#1e1e2e' },
        lime: { 300: '#d9fb6f', 400: '#c6f432', 500: '#a8d61c' },
      },
    },
  },
};
