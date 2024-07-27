const { app, BrowserWindow, dialog, ipcMain } = require('electron')
const path = require('path')
const fs = require("fs");
const CreateWord = require("./word").CreateWord;

function createWindow () {
  const win = new BrowserWindow({
    width: 600,
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
  

  let images;
try {
  //Para local
  //images = fs.readdirSync("./images").filter(file => file.endsWith(".jpg") || file.endsWith(".jpeg") || file.endsWith(".png") || file.endsWith(".HEIC"));
  images = fs.readdirSync("./resources/app/images").filter(file => file.endsWith(".jpg") || file.endsWith(".jpeg") || file.endsWith(".png") || file.endsWith(".HEIC"));
} catch (error) {
  console.error('Error al leer los archivos de la carpeta:', error);
  event.reply('file-created', 'Hubo un error al leer los archivos de la carpeta');
  return;
}
  
  try {

    await CreateWord(event, images, fileName);
    event.reply('file-created', 'Todo listo');

  } catch (error) {

    console.error('Error al crear el archivo Word:', error);
    event.reply('file-created', 'Hubo un error al crear el archivo: ' + error.message);

  }

});
