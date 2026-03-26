/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        brand: {
          DEFAULT: '#0f172a',
          card: '#1e293b',
          border: '#334155',
          accent: '#3b82f6',
        },
        up: '#22c55e',
        down: '#ef4444',
      },
    },
  },
  plugins: [],
}
