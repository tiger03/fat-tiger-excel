<template>
  <ImportButton buttonText="导入原始表" type="origin"/>
</template>

<script>
// import HelloWorld from './components/HelloWorld.vue'
import ImportButton from './components/ImportButton.vue'

export default {
  name: 'App',
  components: {
    // HelloWorld,
    ImportButton,
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
