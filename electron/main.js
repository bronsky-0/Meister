const { app, BrowserWindow, ipcMain, shell, clipboard } = require('electron');
const path = require('path');
const fs = require('fs');
const { createSyncServer } = require('../sync-server');

const APP_ROOT = path.join(__dirname, '..');
const DEFAULT_PORT = 41235;

let mainWindow = null;
let serverInstance = null;
let logFilePath = '';

function appendLogLine(line) {
    if (!logFilePath) return;
    try {
        fs.appendFileSync(logFilePath, line + '\n', 'utf8');
    } catch (e) {
        // ignore file write errors
    }
}

function getServerInfo() {
    if (!serverInstance || !serverInstance.isRunning()) {
        return { running: false, port: DEFAULT_PORT, urls: [], lanUrls: [], ips: [], localUrl: '' };
    }
    return serverInstance.getInfo();
}

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1280,
        height: 800,
        minWidth: 900,
        minHeight: 600,
        title: 'Meister',
        backgroundColor: '#000000',
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: false
        }
    });

    mainWindow.loadFile(path.join(APP_ROOT, 'secretary_terminal_updated .html'));

    mainWindow.on('closed', function() {
        mainWindow = null;
    });
}

async function startEmbeddedServer() {
    if (serverInstance && serverInstance.isRunning()) {
        return getServerInfo();
    }

    serverInstance = createSyncServer({
        root: APP_ROOT,
        port: DEFAULT_PORT,
        silent: true,
        onLog: function(_entry, line) {
            appendLogLine(line);
        }
    });

    await serverInstance.start();
    return getServerInfo();
}

async function stopEmbeddedServer() {
    if (!serverInstance) return { running: false };
    await serverInstance.stop();
    return getServerInfo();
}

app.whenReady().then(function() {
    logFilePath = path.join(app.getPath('userData'), 'meister-server.log');
    appendLogLine('=== Meister ' + app.getVersion() + ' started ===');

    ipcMain.handle('desktop:getInfo', function() {
        return {
            isDesktop: true,
            version: app.getVersion(),
            server: getServerInfo(),
            logFilePath: logFilePath
        };
    });

    ipcMain.handle('desktop:startServer', async function() {
        try {
            return { ok: true, server: await startEmbeddedServer(), logFilePath: logFilePath };
        } catch (err) {
            appendLogLine('[error] startServer failed: ' + (err.message || err));
            return { ok: false, error: err.message || String(err), logFilePath: logFilePath };
        }
    });

    ipcMain.handle('desktop:stopServer', async function() {
        try {
            return { ok: true, server: await stopEmbeddedServer(), logFilePath: logFilePath };
        } catch (err) {
            appendLogLine('[error] stopServer failed: ' + (err.message || err));
            return { ok: false, error: err.message || String(err), logFilePath: logFilePath };
        }
    });

    ipcMain.handle('desktop:getLogs', function() {
        if (serverInstance && serverInstance.getLogs) {
            return {
                entries: serverInstance.getLogs(),
                logFilePath: logFilePath
            };
        }
        let fileTail = '';
        try {
            if (fs.existsSync(logFilePath)) {
                fileTail = fs.readFileSync(logFilePath, 'utf8');
            }
        } catch (e) {
            fileTail = '';
        }
        return { entries: [], logFilePath: logFilePath, fileTail: fileTail };
    });

    ipcMain.handle('desktop:openLogFile', function() {
        if (logFilePath && fs.existsSync(logFilePath)) {
            shell.showItemInFolder(logFilePath);
            return true;
        }
        shell.openPath(app.getPath('userData'));
        return true;
    });

    ipcMain.handle('desktop:copyText', function(_event, text) {
        clipboard.writeText(text || '');
        return true;
    });

    ipcMain.handle('desktop:openExternal', function(_event, url) {
        if (url) shell.openExternal(url);
        return true;
    });

    createWindow();

    app.on('activate', function() {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', async function() {
    if (serverInstance) {
        await serverInstance.stop();
    }
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('before-quit', async function() {
    if (serverInstance) {
        await serverInstance.stop();
    }
});
