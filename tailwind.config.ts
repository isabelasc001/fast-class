import type { Config } from "tailwindcss";
import tailwindcssAnimate from "tailwindcss-animate";

export default {
  darkMode: ["class"],
  content: [
    "./pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./components/**/*.{js,ts,jsx,tsx,mdx}",
    "./app/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {
      colors: {
        fast: {
          bg: "var(--color-bg)",
          surface: "var(--color-surface)",
          "surface-muted": "var(--color-surface-muted)",
          text: "var(--color-text)",
          heading: "var(--color-heading)",
          "text-muted": "var(--color-text-muted)",
          border: "var(--color-border)",
          "border-soft": "var(--color-border-soft)",
          primary: "var(--color-primary)",
          "primary-hover": "var(--color-primary-hover)",
          accent: "var(--color-accent)",
          success: "var(--color-success)",
          warning: "var(--color-warning)",
          error: "var(--color-error)",
          info: "var(--color-info)",
        },
        chart: {
          "1": "var(--color-primary)",
          "2": "var(--color-accent)",
          "3": "var(--color-success)",
          "4": "var(--color-warning)",
          "5": "var(--color-info)",
        },
      },
      borderRadius: {
        lg: "var(--radius-lg)",
        md: "var(--radius-md)",
        sm: "var(--radius-sm)",
      },
    },
  },
  plugins: [tailwindcssAnimate],
} satisfies Config;
