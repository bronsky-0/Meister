const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('MeisterDesktop', {
    getInfo: function() {
        return ipcRenderer.invoke('desktop:getInfo');
    },
    startServer: function() {
        return ipcRenderer.invoke('desktop:startServer');
    },
    stopServer: function() {
        return ipcRenderer.invoke('desktop:stopServer');
    },
    copyText: function(text) {
        return ipcRenderer.invoke('desktop:copyText', text);
    },
    openExternal: function(url) {
        return ipcRenderer.invoke('desktop:openExternal', url);
    },
    getLogs: function() {
        return ipcRenderer.invoke('desktop:getLogs');
    },
    openLogFile: function() {
        return ipcRenderer.invoke('desktop:openLogFile');
    }
});
