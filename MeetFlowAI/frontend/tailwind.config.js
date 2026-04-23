/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./app/**/*.{js,jsx}", "./components/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: {
        midnight: "#0b1020",
      },
    },
  },
  plugins: [require("@tailwindcss/typography")],
};
