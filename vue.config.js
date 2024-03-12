const { defineConfig } = require('@vue/cli-service')
module.exports = defineConfig({
  transpileDependencies: true,
  outputDir:'dist',
  pluginOptions: {
    electronBuilder: {
      builderOptions: {
        // build配置在此处
        // options placed here will be merged with default configuration and passed to electron-builder
        "appId": "com.example.yourapp",
        "productName": "toExcel",
        "win": {
          "target": [
            "nsis"
          ],
          "icon": "image/icon.ico"
        }
      }
    },
  },
})
