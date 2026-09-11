const { contextBridge } = require('electron')

contextBridge.exposeInMainWorld('studentTalkDesktop', {
  isDesktop: true,
})