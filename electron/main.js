const { app, BrowserWindow, ipcMain, shell, clipboard, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { createSyncServer } = require('../sync-server');

const APP_ROOT = path.join(__dirname, '..');
const DEFAULT_PORT = 41235;
const HTML_FILE = 'secretary_terminal.html';

let mainWindow = null;
let serverInstance = null;
let logFilePath = '';
let isQuitting = false;

function ensureLogFilePath() {
    if (!logFilePath) {
        logFilePath = path.join(app.getPath('userData'), 'meister-server.log');
    }
    return logFilePath;
}

function appendLogLine(line) {
    try {
        fs.appendFileSync(ensureLogFilePath(), line + '\n', 'utf8');
    } catch (e) {
        // ignore file write errors
    }
}

function logMessage(message) {
    appendLogLine('[' + new Date().toISOString() + '] ' + message);
}

function showStartupError(title, message) {
    logMessage('[error] ' + title + ': ' + message);
    dialog.showErrorBox(title, message + '\n\nПодробности: ' + ensureLogFilePath());
}

function getServerInfo() {
    if (!serverInstance || !serverInstance.isRunning()) {
        return { running: false, port: DEFAULT_PORT, urls: [], lanUrls: [], ips: [], localUrl: '' };
    }
    return serverInstance.getInfo();
}

function stopEmbeddedServer() {
    if (!serverInstance || !serverInstance.isRunning()) {
        serverInstance = null;
        return Promise.resolve({ running: false });
    }
    return serverInstance.stop().then(function(info) {
        serverInstance = null;
        logMessage('Sync-server stopped');
        return info || { running: false };
    }).catch(function(err) {
        serverInstance = null;
        logMessage('[error] stop server: ' + (err.message || err));
        return { running: false };
    });
}

function shutdownApp() {
    if (isQuitting) return;
    isQuitting = true;
    stopEmbeddedServer().finally(function() {
        app.exit(0);
    });
}

function createWindow() {
    const htmlPath = path.join(APP_ROOT, HTML_FILE);
    if (!fs.existsSync(htmlPath)) {
        showStartupError(
            'Meister — ошибка запуска',
            'Не найден файл интерфейса:\n' + htmlPath
        );
        app.quit();
        return;
    }

    mainWindow = new BrowserWindow({
        width: 1280,
        height: 800,
        minWidth: 900,
        minHeight: 600,
        title: 'Meister',
        backgroundColor: '#000000',
        show: false,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolation: true,
            nodeIntegration: false,
            sandbox: false
        }
    });

    mainWindow.once('ready-to-show', function() {
        mainWindow.show();
    });

    mainWindow.webContents.on('did-fail-load', function(_event, errorCode, errorDescription, validatedURL) {
        showStartupError(
            'Meister — ошибка загрузки',
            errorDescription + ' (' + errorCode + ')\n' + validatedURL
        );
    });

    mainWindow.loadFile(htmlPath).catch(function(err) {
        showStartupError('Meister — ошибка запуска', err.message || String(err));
        app.quit();
    });

    mainWindow.on('close', function() {
        stopEmbeddedServer();
    });

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

process.on('uncaughtException', function(err) {
    logMessage('[uncaughtException] ' + (err && err.stack ? err.stack : String(err)));
});

process.on('unhandledRejection', function(reason) {
    logMessage('[unhandledRejection] ' + String(reason));
});

app.on('before-quit', function(event) {
    if (isQuitting) return;
    event.preventDefault();
    shutdownApp();
});

app.on('window-all-closed', function() {
    stopEmbeddedServer();
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.whenReady().then(function() {
    logMessage('=== Meister ' + app.getVersion() + ' started ===');
    logMessage('APP_ROOT: ' + APP_ROOT);
    logMessage('Platform: ' + process.platform + ' ' + process.arch);

    ipcMain.handle('desktop:getInfo', function() {
        return {
            isDesktop: true,
            version: app.getVersion(),
            server: getServerInfo(),
            logFilePath: ensureLogFilePath()
        };
    });

    ipcMain.handle('desktop:startServer', async function() {
        try {
            return { ok: true, server: await startEmbeddedServer(), logFilePath: ensureLogFilePath() };
        } catch (err) {
            logMessage('[error] startServer failed: ' + (err.message || err));
            return { ok: false, error: err.message || String(err), logFilePath: ensureLogFilePath() };
        }
    });

    ipcMain.handle('desktop:stopServer', async function() {
        try {
            const server = await stopEmbeddedServer();
            return { ok: true, server: server, logFilePath: ensureLogFilePath() };
        } catch (err) {
            logMessage('[error] stopServer failed: ' + (err.message || err));
            return { ok: false, error: err.message || String(err), logFilePath: ensureLogFilePath() };
        }
    });

    ipcMain.handle('desktop:getLogs', function() {
        if (serverInstance && serverInstance.getLogs) {
            return {
                entries: serverInstance.getLogs(),
                logFilePath: ensureLogFilePath()
            };
        }
        let fileTail = '';
        try {
            const logPath = ensureLogFilePath();
            if (fs.existsSync(logPath)) {
                fileTail = fs.readFileSync(logPath, 'utf8');
            }
        } catch (e) {
            fileTail = '';
        }
        return { entries: [], logFilePath: ensureLogFilePath(), fileTail: fileTail };
    });

    ipcMain.handle('desktop:quit', function() {
        shutdownApp();
        return { ok: true };
    });

    ipcMain.handle('desktop:openLogFile', function() {
        const logPath = ensureLogFilePath();
        if (logPath && fs.existsSync(logPath)) {
            shell.showItemInFolder(logPath);
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
