const { app, BrowserWindow, dialog, ipcMain } = require('electron')
const path = require('path')
const fs = require("fs");
const CreateWord = require("./word").CreateWord;

function createWindow () {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    }
  })

  win.loadFile('index.html')
}

app.whenReady().then(() => {
  createWindow()
  
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow()
    }
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

ipcMain.on('get-file-name', async (event, fileName) => {
  const images = fs.readdirSync("./images").filter(file => file.endsWith(".jpg") || file.endsWith(".jpeg") || file.endsWith(".png") || file.endsWith(".HEIC"));
  CreateWord(images, fileName);
  event.reply('file-created', 'Todo listo');
})
