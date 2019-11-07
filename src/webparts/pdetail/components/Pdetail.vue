<template>
  <v-app :id="$style.keep">
    <v-content app>
      <div class="cs_container pt-5">
        <v-layout align-center justify-center>
          <router-view />
        </v-layout>
      </div>
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
  .cs_container {
      margin: 0 auto;
    }
  /* Extra small devices (phones, 600px and down) */
  @media only screen and (max-width: 600px) {
    .cs_container {
      max-width: 95%;
    }
  }

  /* Small devices (portrait tablets and large phones, 600px and up) */
  @media only screen and (min-width: 600px) {
    .cs_container {
      max-width: 90%;
    }
  }

  /* Medium devices (landscape tablets, 768px and up) */
  @media only screen and (min-width: 768px) {
    .cs_container {
      max-width: 80%;
    }
  }

  /* Large devices (laptops/desktops, 992px and up) */
  @media only screen and (min-width: 992px) {
    .cs_container {
      max-width: 70%;
    }
  }

  /* Extra large devices (large laptops and desktops, 1200px and up) */
  @media only screen and (min-width: 1200px) {
    .cs_container {
      max-width: 60%;
    }
  }
}
#keep {
  :after,
  :before {
    vertical-align: inherit !important;
  }
}
</style>


