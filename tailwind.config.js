/** @type {import('tailwindcss').Config} */
module.exports = {
  darkMode: 'class',
  content: ['./src/**/*.{html,ts}'],
  theme: {
    extend: {
      colors: {
        primary: 'rgb(var(--color-primary) / <alpha-value>)',
        'background-light': '#f6f6f8',
        'background-dark': 'rgb(var(--color-bg-dark) / <alpha-value>)',
        'sidebar-dark': 'rgb(var(--color-sidebar-dark) / <alpha-value>)',
      },
      fontFamily: {
        display: ['Inter', 'sans-serif'],
      },
    },
  },
  plugins: [],
};
