<template>
<div id="app">
  <header class="ms-welcome__header ms-bgColor-neutralLighter ms-u-fadeIn500">
    <img width="90" height="90" src="../assets/logo-filled.png" alt="<%= projectInternalName %>" title="<%= projectInternalName %>" />
    <h1 class="ms-fontSize-su ms-fontWeight-light ms-fontColor-neutralPrimary">{{welcomeMessage}}</h1>
  </header>
  <main class="ms-welcome__main">
     <h2 class="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20"> Discover what <%= projectInternalName %> can do for you today! </h2>
     <ul class="ms-List ms-welcome__features ms-u-slideUpIn10">
        <li class="ms-ListItem">
            <i class="ms-Icon ms-Icon--Ribbon"></i>
            <span class="ms-font-m ms-fontColor-neutralPrimary">Achieve more with Office integration</span>
        </li>
        <li class="ms-ListItem">
            <i class="ms-Icon ms-Icon--Unlock"></i>
            <span class="ms-font-m ms-fontColor-neutralPrimary">Unlock features and functionality</span>
        </li>
        <li class="ms-ListItem">
            <i class="ms-Icon ms-Icon--Design"></i>
            <span class="ms-font-m ms-fontColor-neutralPrimary">Create and visualize like a pro</span>
        </li>
    </ul>
    <br />
    <br />
    <p class="ms-font-l">Modify the source files, then click <b>Run</b>.</p>
    <button v-on:click="onClickRun">Run</button>
  </main>
</div>
</template>

<script>
/* eslint-disable indent,semi,no-extra-semi,no-trailing-spaces,no-multiple-empty-lines,no-unreachable,keyword-spacing,padded-blocks,no-return-await */
export default {
  name: 'App',
  data: function () {
    return {
      welcomeMessage: 'Welcome'
    }
  },
  methods: {
    onClickRun: async () => {
      <%- snippet %>
    }
  }
}
/* eslint-enable indent,semi,no-extra-semi,no-trailing-spaces,no-multiple-empty-lines,no-unreachable,keyword-spacing,padded-blocks,no-return-await */
</script>

<style>
#content-header {
    background: #2a8dd4;
    color: #fff;
    position: absolute;
    top: 0;
    left: 0;
    width: 100%;
    height: 80px;
    overflow: hidden;
}

#content-main {
    background: #fff;
    position: fixed;
    top: 80px;
    left: 0;
    right: 0;
    bottom: 0;
    overflow: auto;
}

.padding {
    padding: 15px;
}
</style>
