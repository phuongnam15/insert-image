import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import App from "./App.js";
import { HashRouter } from "react-router-dom";
import { NotifyContextProvider } from "./contexts/notifyContext.jsx";

const root = ReactDOM.createRoot(document.getElementById("root"));
root.render(
  <HashRouter>
    <NotifyContextProvider>
        <App />
    </NotifyContextProvider>
  </HashRouter>,
);
