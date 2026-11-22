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
        brand: { 200: '#c7d7ff', 300: '#8db0ff', 400: '#5b8cff', 500: '#3d6ff0' },
      },
    },
  },
};
