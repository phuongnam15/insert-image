/** @type {import('tailwindcss').Config} */
import plugin from 'tailwindcss/plugin';
export default {
  content: ["./public/index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: {
    extend: {
      fontFamily: {
        popi: "Poppins, sans-serif",
        meri: "Merriweather, serif",
        nutito: "Nunito, sans-serif",
      },
      borderWidth: {
        1: "1px",
      },
      boxShadow: {
        'custom-white': '0 0 3px rgba(255, 255, 255, 0.5)',
        'custom-inset': 'inset 0 -4px 20px rgba(124, 57, 237, 0.5)',
        'custom-inset-pink': 'inset 0 -4px 20px rgba(219, 112, 147, 0.5)',
        'custom-inset-green': 'inset 0 -4px 20px rgba(78, 184, 76, 0.5)',
        'custom-inset-gray': 'inset 0 -4px 20px rgba(53, 53, 91, 0.9)',
        'custom-gray': '0 4px 20px rgba(53, 53, 91, 0.9)',
        'custom-inset-2': '0 4px 20px rgba(111, 65, 210, 0.5)',
      },
      backgroundColor: {
        'custom-hover': 'rgba(124, 57, 237, 0.25)',
        'custom-hover2': 'rgba(219, 112, 147, 0.25)',
        'custom-hover3': 'rgba(53, 53, 91, 0.25)',
        'custom-hover4': 'rgba(78, 184, 76, 0.25)'
      },
      letterSpacing: {
        wider1: '0.07em',
      },
      backgroundImage: {
        'coin': "url('../src/assets/backgrounds/coin1.png')",
      },
    },
  },
  plugins: [
    plugin(function ({ theme, addUtilities }) {
      const neonUtilities = {};
      const colors = theme('colors');
      for(const color in colors) {
        if(typeof colors[color] === 'object') {
          neonUtilities[`.neon-${color}`] = {
            boxShadow: `inset 0 -4px 20px ${color}`
          }
        }
      }
      addUtilities(neonUtilities);
    })
  ],
};
