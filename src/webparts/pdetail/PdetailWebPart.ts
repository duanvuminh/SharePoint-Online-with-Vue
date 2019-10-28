import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'PdetailWebPartStrings';

// Importing Vue.js
import Vue from 'vue';
// Importing Vue.js SFC
import PdetailComponent from './components/Pdetail.vue';
import PeopleSearch from './components/PeopleSearch.vue';
import PeopleDetail from './components/PeopleDetail.vue';
// FontAwaresome
import Vuetify from 'vuetify';
import 'vuetify/dist/vuetify.min.css';
import VueRouter from 'vue-router';

Vue.use(VueRouter);
Vue.use(Vuetify);

export interface IPdetailWebPartProps {
  description: string;
}

export default class PdetailWebPart extends BaseClientSideWebPart<IPdetailWebPartProps> {

  public render(): void {
    const id: string = `wp-${this.instanceId}`;
    this.domElement.innerHTML = `<head><link href="https://fonts.googleapis.com/css?family=Roboto:100,300,400,500,700,900" rel="stylesheet"><link href="https://cdn.jsdelivr.net/npm/@mdi/font@4.x/css/materialdesignicons.min.css" rel="stylesheet"></head><div id="${id}"></div>`;

    let el = new Vue({
      el: `#${id}`,
      vuetify: new Vuetify(),
      router: new VueRouter({
        mode: 'history',
        base: "/sites/dev1/SitePages/wgpeople.aspx",
        routes: [
          {
            path: '/',
            name: 'home',
            component: PeopleSearch,
            children: [{
              path: '/search',
              component: PeopleSearch
            }]
          },
          {
            path: '/Detail/:pid',
            name: 'detail',
            component: PeopleDetail
          }
        ]
      }),
      render: h => h(PdetailComponent, {
        props: {
          description: this.properties.description
        }
      })
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
