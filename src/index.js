import { createRoot } from "react-dom/client";
import App from "./App";
import { createTheme, ThemeProvider } from "@mui/material";

const theme = createTheme({
    palette: {
        primary: {
            main: "#013e87"
        },
        secondary: {
            main: "#2e74c9"
        },
    },
    typography:{
        h1: {
            fontSize: "3rem",
            fontWeight: 600,
        },
        h2: {
            fontSize: "1.75rem",
            fontWeight: 600,
        },
        h3: {
            fontSize: "1.5rem",
            fontWeight: 600,
        },
    }
})

const container = document.getElementById("app");
const root = createRoot(container)
root.render(
    <ThemeProvider theme={theme}>
        <App />
    </ThemeProvider>
);