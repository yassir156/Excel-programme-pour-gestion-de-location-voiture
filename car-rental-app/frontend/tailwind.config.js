/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,jsx}'],
  darkMode: 'class',
  theme: {
    extend: {
      colors: {
        navy: {
          950: '#0b1220',
          900: '#0f172a',
          800: '#152238',
          700: '#1e2f4d',
          600: '#28406b',
        },
        brand: {
          50: '#eef4ff',
          100: '#dbe7ff',
          200: '#b8cffe',
          300: '#8aaefc',
          400: '#5686f6',
          500: '#2f5fed',
          600: '#1f45cf',
          700: '#1c37a6',
          800: '#1c3184',
          900: '#1b2c6a',
        },
        accent: {
          400: '#34d399',
          500: '#10b981',
          600: '#059669',
        },
      },
      fontFamily: {
        sans: ['Inter', 'Segoe UI', 'system-ui', 'sans-serif'],
      },
      boxShadow: {
        soft: '0 2px 10px rgba(15, 23, 42, 0.06)',
        card: '0 1px 3px rgba(15, 23, 42, 0.08), 0 1px 2px rgba(15,23,42,0.04)',
      },
      borderRadius: {
        xl2: '1rem',
      },
    },
  },
  plugins: [],
};
