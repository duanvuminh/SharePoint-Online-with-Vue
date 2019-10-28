<template>
  <v-flex style="min-height:50vh" :class="$style.mywrap">
    <v-card style="overflow:hidden">
      <v-list>
        <v-list-item-group mandatory color="indigo">
          <v-list-item style="top:-10px">
            <v-list-item-icon>
              <v-avatar size="100">
                <img
                  :src="`https://willgroup.sharepoint.com/_layouts/15/userphoto.aspx?size=L&username=${properties.Email}`"
                />
              </v-avatar>
            </v-list-item-icon>
            <v-list-item-content>
              <v-list-item-title v-text="properties.DisplayName"></v-list-item-title>
              <v-list-item-subtitle
                v-text="`${properties.userProperties.Department}　${properties.Title}`"
              ></v-list-item-subtitle>
              <v-list-item-subtitle v-text="properties.userProperties.Office"></v-list-item-subtitle>
              <v-list-item-subtitle v-text="properties.Email"></v-list-item-subtitle>
            </v-list-item-content>
          </v-list-item>
          <v-subheader>基本情報</v-subheader>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>名前</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text v-html="properties.DisplayName" :class="$style.actiontext"></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>部署</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties['SPS-Department']"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>役職</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text v-html="properties.Title" :class="$style.actiontext"></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>説明</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties.AboutMe"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-subheader>連絡先の情報</v-subheader>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>勤務先の電子メール</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties.WorkEmail"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>携帯電話</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties.CellPhone"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>Fax</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties.Fax"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>オフィス</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties.Office"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>勤務先所在地</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties['SPS-Location']"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-subheader>詳細</v-subheader>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>以前のプロジェクト</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties['SPS-PastProjects']"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>スキル</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties['SPS-Skills']"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>学歴</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties['SPS-School']"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>誕生日</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties['SPS-Birthday']"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
          <v-divider></v-divider>
          <v-list-item>
            <v-list-item-content>
              <v-list-item-title>興味</v-list-item-title>
            </v-list-item-content>
            <v-list-item-action>
              <v-list-item-action-text
                v-html="properties.userProperties['SPS-Interests']"
                :class="$style.actiontext"
              ></v-list-item-action-text>
            </v-list-item-action>
          </v-list-item>
        </v-list-item-group>
      </v-list>
    </v-card>
  </v-flex>
</template>

<script>
import { sp, Search, Web } from "@pnp/sp";
export default {
  data() {
    return {
      properties: {
        Title: "",
        userProperties: {
          Department: ""
        }
      }
    };
  },
  methods: {},
  mounted() {
    sp.profiles.getPropertiesFor(this.$route.params.pid).then(profile => {
      // console.log(profile.DisplayName);
      // console.log(profile.Email);
      // console.log(profile.Title);
      // console.log(profile.UserProfileProperties.length);
      // Properties are stored in inconvenient Key/Value pairs,
      // so parse into an object called userProperties
      let properties = {};
      profile.UserProfileProperties.forEach(prop => {
        properties[prop.Key] = prop.Value;
      });
      profile.userProperties = properties;
      this.properties = profile;
      let scroll = document.querySelectorAll("[data-is-scrollable]")[0];
      scroll.scrollTop = -100;
    });
  },
  computed: {},
  watch: {}
};
</script>

<style lang="scss" module>
:global {
  .v-list-item__action {
    max-width: 60% !important;
  }
}
.mywrap {
  .actiontext {
    white-space: normal !important;
    word-wrap: break-word !important;
  }
}
</style>


