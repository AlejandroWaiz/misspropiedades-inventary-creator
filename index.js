const { app, BrowserWindow, dialog, ipcMain } = require('electron')
const path = require('path')
const fs = require("fs");
const CreateWord = require("./word").CreateWord;
const os = require("os");

app.disableHardwareAcceleration();

function createWindow () {
  const win = new BrowserWindow({
    width: 600,
    height: 500,
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

ipcMain.on('get-file-name', async (event, { fileName, imagesFolder }) => {
  try {
    await CreateWord(event, fileName, imagesFolder);
    event.reply('file-created', 'Todo listo');
  } catch (error) {
    console.error('Error al crear el archivo Word:', error);
    event.reply('file-created', 'Hubo un error al crear el archivo: ' + error.message);
  }
});
