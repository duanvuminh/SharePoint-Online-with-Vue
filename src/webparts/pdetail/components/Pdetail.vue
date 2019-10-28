<template>
  <v-app :id="$style.keep">
    <v-content app>
      <v-container>
        <v-layout align-center justify-center>
          <router-view />
        </v-layout>
      </v-container>
    </v-content>
  </v-app>
</template>

<style lang="scss" module>
:global {
  .v-application--wrap {
    min-height: 0 !important;
  }
  input[type="text"] {
    background: transparent !important;
  }
  :after,
  :before {
    vertical-align: baseline !important;
  }
  html {
    overflow: auto !important;
  }
}
#keep {
  :after,
  :before {
    vertical-align: inherit !important;
  }
}
</style>


