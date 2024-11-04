const { contextBridge, ipcRenderer } =  require("electron");

contextBridge.exposeInMainWorld("ipcRenderer", {
  send: (chanel, data) => ipcRenderer.send(chanel, data),
  on: (chanel, func) =>
    ipcRenderer.on(chanel, func),
  once: (chanel, func) =>
    ipcRenderer.once(chanel, func),
  removeListener: (chanel, func) =>
    ipcRenderer.removeListener(chanel, func),
});