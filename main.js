const { app, BrowserWindow, Menu, ipcMain } = require("electron");
const path = require("path");
require("dotenv").config();
const { exec } = require("child_process");
const isDev = process.env.NODE_ENV === "development";
const { run } = require("./index");

function createWindow() {
  mainWindow = new BrowserWindow({
    title: "",
    width: 500,
    height: 600,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      preload: path.join(__dirname, `preload.js`),
    },
  });

  const startUrl = isDev
    ? "http://localhost:3000"
    : `file://${path.join(__dirname, "./ui/build/index.html")}`;

  mainWindow.loadURL(startUrl);

  // mainWindow.webContents.openDevTools();

  Menu.setApplicationMenu(null);
}
app.whenReady().then(() => {
  createWindow();

  mainWindow.on("closed", () => {
    mainWindow = null;
  });

  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

ipcMain.on("run-insert", (event, data) => {
  try {
    run()
    event.sender.send("run-insert", {});
  } catch (error) {
    event.sender.send("run-insert", { error: error.message });
  }
});
