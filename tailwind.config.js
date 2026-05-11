/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./frontend/**/*.{html,js,vue}"
  ],
  theme: {
    extend: {
      colors: {
        dark: '#0f172a',
        darker: '#1e293b',
        primary: '#f1f5f9',
        secondary: '#94a3b8',
        muted: '#64748b'
      }
    },
  },
  plugins: [],
}

