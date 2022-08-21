const { app, BrowserWindow, dialog } = require('electron');
const { ipcMain } = require('electron');
const fs = require('fs')
const path = require('path')
let main_ui = null;

function createWindow () {
  main_ui = new BrowserWindow({
    width: 1500,
    height: 800,
    minWidth: 1000,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: true
    }
  })

  main_ui.setMenu(null);
  main_ui.loadFile('index.html');
  main_ui.on("close", (a, b) => {
    main_ui.webContents.send('close', {});
  });
  // main_ui.webContents.openDevTools();
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})

ipcMain.on('read-path', (event, arg) => {
  const fpath = dialog.showOpenDialogSync(main_ui, {
    filters: [{
      name: 'excel files',
      extensions: ['xlsx']
    }]
  });
  if (fpath) {
    main_ui.webContents.send('read-path', fpath[0]);
  }
});