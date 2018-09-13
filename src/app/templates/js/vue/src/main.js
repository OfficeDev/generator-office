// The Vue build version to load with the `import` command
// (runtime-only or standalone) has been set in webpack.base.conf with an alias.
/* eslint-disable no-multiple-empty-lines,no-unused-vars,semi */
import Vue from 'vue'
import App from './App'

<%- imports %>
/* eslint-enable no-multiple-empty-lines,no-unused-vars,semi */
Vue.config.productionTip = false

/* eslint-disable no-new */
const Office = window.Office
Office.initialize = () => {
  new Vue({
    el: '#app',
    components: {App},
    template: '<App/>'
  })
}
