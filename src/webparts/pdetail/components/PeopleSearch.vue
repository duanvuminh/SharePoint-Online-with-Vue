<template>
  <v-flex style="min-height:50vh" :class="$style.mywrap">
    <v-card class="elevation-1">
      <v-toolbar color="primary" dark flat>
        <v-toolbar-title>社員検索</v-toolbar-title>
      </v-toolbar>
      <v-card-text>
        <v-form>
          <v-row no-gutters>
            <v-col cols="12" sm="5">
              <v-autocomplete
                label="会社名"
                :items="items"
                v-model="comapany_name"
                no-data-text="見つかりません"
                attach="#search2"
              ></v-autocomplete>
              <div id="search2" style="display:relative"></div>
            </v-col>
            <v-col cols="12" sm="5" offset="0" offset-sm="2">
              <v-text-field label="部署名" type="text" v-model="department_name"></v-text-field>
            </v-col>
          </v-row>
          <v-row no-gutters>
            <v-col cols="12" sm="5">
              <v-text-field label="社員名" type="text" v-model="people_name"></v-text-field>
            </v-col>
            <v-col cols="12" sm="5" offset="0" offset-sm="2">
              <v-text-field label="社員役職" type="text" v-model="people_job"></v-text-field>
            </v-col>
          </v-row>
        </v-form>
      </v-card-text>
      <v-card-actions>
        <v-btn
          color="primary"
          top
          :to="{ path: '/search', query: { a:comapany_name ,b:department_name,c:people_name,d:people_job }}"
        >検索</v-btn>
        <v-btn @click="clear" text>クリア</v-btn>
      </v-card-actions>
    </v-card>
    <div class="pt-5 pb-5">
      <v-data-table
        expand
        hide-default-header
        :items="search_result"
        :items-per-page="12"
        class="elevation-0"
        :footer-props="{'items-per-page-options': [12],
       'items-per-page-text': null,
       'disable-items-per-page': true,
       'showFirstLastPage': true }"
        v-if="searchdone"
      >
        <template slot="item" slot-scope="props">
          <tr>
            <td>
              <v-card class="mx-auto">
                <v-img
                  :src="`https://willgroup.sharepoint.com/_layouts/15/userphoto.aspx?size=L&username=${props.item.WorkEmail}`"
                  height="250px"
                ></v-img>

                <v-card-title>
                  <div class="subtitle-1">{{props.item.PreferredName?props.item.PreferredName:"-"}}</div>
                  <div style="width:100%"></div>
                  <span
                    class="grey--text subtitle-1"
                  >{{props.item.JobTitle?props.item.JobTitle:"-"}}</span>
                </v-card-title>
                <v-card-actions>
                  <router-link :to="{ name: 'detail', params: { pid: props.item.AccountName }}">
                    <v-btn text>詳細</v-btn>
                  </router-link>
                </v-card-actions>
              </v-card>
            </td>
          </tr>
        </template>
        <template slot="no-data">...</template>
        <template
          slot="footer.page-text"
          slot-scope="props"
        >{{ props.pageStart }} - {{ props.pageStop }} / {{ props.itemsLength }}</template>
      </v-data-table>
    </div>
  </v-flex>
</template>

<script>
import { sp, Search, Web } from "@pnp/sp";
export default {
  data() {
    return {
      comapany_name: "",
      department_name: "",
      people_name: "",
      people_job: "",
      // search
      searchService: {},
      search_result: [],
      kaishaCollection: [],
      items: [],
      //table
      searchdone: false
    };
  },
  methods: {
    searchAll(queryText, rowLimit, startRow, allResult) {
      let allResults = allResult || [];
      return this.searchService
        .execute({
          __metadata: {
            type: "Microsoft.Office.Server.Search.REST.SearchRequest"
          },
          Querytext: queryText,
          TrimDuplicates: false,
          SourceId: "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31",
          RowLimit: rowLimit,
          StartRow: startRow,
          TrimDuplicates: false
        })
        .then(data => {
          let relevantResults = data.PrimarySearchResults;
          allResults = allResults.concat(relevantResults);
          if (
            data.TotalRowsIncludingDuplicates >
            startRow + relevantResults.length
          ) {
            return this.searchAll(
              queryText,
              rowLimit,
              startRow + relevantResults.length,
              allResults
            );
          }
          return allResults;
        });
    },
    search() {
      let sqlText = "";
      this.comapany_name = this.$route.query.a ? this.$route.query.a : "";
      this.department_name = this.$route.query.b ? this.$route.query.b : "";
      this.people_name = this.$route.query.c ? this.$route.query.c : "";
      this.people_job = this.$route.query.d ? this.$route.query.d : "";
      if (
        this.comapany_name != "" ||
        this.department_name != "" ||
        this.people_name != "" ||
        this.people_job != ""
      ) {
        if (this.comapany_name != "") {
          if (sqlText == "") {
            sqlText += `OfficeNumber:${this.comapany_name}`;
          } else {
            sqlText += ` OfficeNumber:${this.comapany_name}`;
          }
        }
        if (this.department_name != "") {
          if (sqlText == "") {
            sqlText += `Department:${this.department_name}`;
          } else {
            sqlText += ` Department:${this.department_name}`;
          }
        }
        if (this.people_name != "") {
          if (sqlText == "") {
            sqlText += `PreferredName:${this.people_name}`;
          } else {
            sqlText += ` PreferredName:${this.people_name}`;
          }
        }
        if (this.people_job != "") {
          if (sqlText == "") {
            sqlText += `JobTitle:${this.people_job}`;
          } else {
            sqlText += ` JobTitle:${this.people_job}`;
          }
        }
      } else {
        return;
      }
      console.log(sqlText);
      this.searchAll(sqlText, 500, 0, 0).then(r => {
        console.log(r);
        console.log(sp);
        this.search_result = Object.values(r).sort((a, b) => {
          return a.PreferredName.localeCompare(b.PreferredName);
        });
        this.searchdone = true;
      });
      // this.searchService
      //   .execute({
      //     __metadata: {
      //       type: "Microsoft.Office.Server.Search.REST.SearchRequest"
      //     },
      //     Querytext: sqlText,
      //     TrimDuplicates: false,
      //     SourceId: "B09A7990-05EA-4AF9-81EF-EDFAB16C4E31",
      //     RowLimit: 500,
      //     TrimDuplicates: false
      //   })
      //   .then(r => {
      //     console.log(r);
      //     this.search_result = r.PrimarySearchResults;
      //     this.matchCount = this.search_result.length;
      //     console.log(this.search_result);
      //   });
    },
    clear() {
      this.comapany_name = "";
      this.department_name = "";
      this.people_name = "";
      this.people_job = "";
      this.search_result = [];
      this.matchCount = 0;
      this.searchdone = false;
      this.$router.push({
        path: "/search",
        query: { a: "", b: "", c: "", d: "" }
      });
    }
  },
  mounted() {
    this.searchService = new Search("https://willgroup.sharepoint.com");
    this.searchAll('JobTitle:"社長" JobTitle:"理事長"', 500, 0, 0).then(r => {
      this.kaishaCollection = Object.values(r).sort((a, b) => {
        return a.PreferredName.localeCompare(b.PreferredName);
      });
      Promise.all(
        this.kaishaCollection.map(x => {
          return sp.profiles.getPropertiesFor(x.AccountName).then(profile => {
            let properties = {};
            profile.UserProfileProperties.forEach(prop => {
              properties[prop.Key] = prop.Value;
            });
            return properties.Office.split(" ")[0];
          });
        })
      ).then(res => {
        this.items = res;
        console.log(res);
      });
    });
    this.search();
  },
  computed: {},
  watch: {
    $route(to, from) {
      this.search();
    }
  }
};
</script>

<style lang="scss" module>
:global {
  #search2,
  #search3,
  #search4 {
    position: relative;
    top: -64px;
  }
}
.mywrap {
  ul {
    list-style: none;
  }
  tr {
    float: left;
    margin-bottom: 16px;
    padding: 0 8px;
    /* Extra small devices (phones, 600px and down) */
    @media only screen and (max-width: 600px) {
      width: 100%;
      display: block;
    }

    /* Small devices (portrait tablets and large phones, 600px and up) */
    @media only screen and (min-width: 600px) {
      width: 50%;
    }

    /* Medium devices (landscape tablets, 768px and up) */
    @media only screen and (min-width: 768px) {
      width: 25%;
    }

    /* Large devices (laptops/desktops, 992px and up) */
    @media only screen and (min-width: 992px) {
      width: 25%;
    }

    /* Extra large devices (large laptops and desktops, 1200px and up) */
    @media only screen and (min-width: 1200px) {
      width: 25%;
    }
    td {
      display: block;
      height: auto;
      border-bottom: none;
      padding: 0;
    }
    &:hover {
      background: none !important;
    }
  }
}
</style>


