const { contextBridge } = require('electron');

contextBridge.exposeInMainWorld('carRentalApp', {
  isElectron: true,
  apiBaseUrl: 'http://127.0.0.1:4310/api',
});
