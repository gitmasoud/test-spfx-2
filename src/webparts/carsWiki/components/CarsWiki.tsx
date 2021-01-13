import * as React from 'react';
import styles from './CarsWiki.module.scss';
import { ICarsWikiProps } from './ICarsWikiProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration } from '@microsoft/sp-http';  
import MockHttpClient from './MockHttpClient';
import {
  Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from "@pnp/sp/webs";

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Id: string;
}

export default class CarsWiki extends React.Component<ICarsWikiProps, {}> {
  private _renderListAsync(): void {
    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getMockListData().then((response) => {
        this._renderList(response.value);
      });
    }
    //else if (Environment.type == EnvironmentType.SharePoint ||
    //         Environment.type == EnvironmentType.ClassicSharePoint) {
    //  this._getListData()
    //    .then((response) => {
    //      this._renderList(response.value);
    //    });
    //}
  }
  private _renderList(items: ISPList[]): void {
    let html: string = '';
    items.forEach((item: ISPList) => {
      html += `
    <ul class="${styles}">
      <li class="${styles}">
        <span class="ms-font-l">${item.Title}</span>
      </li>
    </ul>`;
    });
  
    const listContainer: Element = document.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  //get data when not online
  private _getMockListData(): Promise<ISPLists> {
    return MockHttpClient.get()
      .then((data: ISPList[]) => {
        var listData: ISPLists = { value: data };
        return listData;
      }) as Promise<ISPLists>;
  }
  //get some data when online
  _getListData () {
   return  this.props.context.spHttpClient
  .get(`https://etelimitedcouk.sharepoint.com/sites/TestSite/_api/Web/Lists/getByTitle('Wiki%20articles')/Items?$select=Title`, SPHttpClient.configurations.v1)
  .then((res: SPHttpClientResponse): Promise<{ Title: string; }> => {
    return res.json();
  })
  .then((web: {Title: string}): void => {
    console.log(web.Title);
  });
  }

  getFetchStyle = async () => {
    const w = await sp.web.select("Title")();
    //get title
    console.log(w.Title);
    //get all lists on current site
    console.log(sp.web.lists.get());
    //'https://etelimitedcouk.sharepoint.com/sites/TestSite/Lists/WikiArticlesContent/AllItems.aspx'
     //https://etelimitedcouk.sharepoint.com/sites/TestSite/Lists/_api/web/lists/getByTitle('WikiArticlesContent')/items?$select=Title,FileRef,FieldValuesAsText%2FMetaInfo&$expand=FieldValuesAsText

    //const web: Web = new Web(this.props.context.web.absoluteUrl); 
    //web.select("Title").get().then(w => { });

     //console.log(sp.web.lists.getByTitle('WikiArticlesContent').items.get());

    // private web = Web("https://testinglala.sharepoint.com/sites/Test/");   
    //this.web.lists.getByTitle("Employee").items().then((items) => {  

    //get lists from other sp site
                      //https://etelimitedcouk.sharepoint.com/sites/TestSite/_api/Web/Lists/getByTitle('Wiki%20articles')/items/_api/web?$select=Title
    const wurl = Web("https://etelimitedcouk.sharepoint.com/sites/TestSite/");
    const r = wurl.lists.getByTitle('Wiki articles').items(); //<< if there spaces in lists leave them in
    console.log(r);

    const list = sp.web.lists.getById("7B89368552-3d8d-4680-af50-fab774baa367");
    console.log(list)
  }
  //init pnpjs
  pnpSetup = () => {
    sp.setup({
      spfxContext: this.props.context
    })
  }
  public render(): React.ReactElement<ICarsWikiProps> {
    this.pnpSetup();
    this.getFetchStyle();
    //debugger
    //this._renderListAsync();
    return (
      <div className={ styles.carsWiki }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
              
            </div>
          </div>
          <div id="spListContainer" />
          <div id="errorSection" />
        </div>
      </div>
    );
  }
}
