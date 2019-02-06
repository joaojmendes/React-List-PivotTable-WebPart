
/*
  Data Services
  Author: joao Mendes
  date 5/2/2019
*/
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from "@microsoft/sp-http";
import { sp, Web, StorageEntity, ClientSidePage, } from '@pnp/sp';
import * as $ from 'jquery';

export default class SPService {
  private saveData: StorageEntity = null;
  private currentPageName: string;
  private currentPage: any;


  constructor(private _context: WebPartContext | ApplicationCustomizerContext) {

  }

  // Get PivotData from page properties
  public async getSaveData() {
    const apiUrl = `${this._context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${this._context.pageContext.listItem.id})`;
    let returnProperties: any = null;

    const _data = await this._context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);
    if (_data.ok) {
      const results = await _data.json();
      console.log(results);
      // if (results && results.value && results.value.length > 0) {
      if (results) {
        const canvasContent = JSON.parse(results.CanvasContent1);
        for (const v of canvasContent) {
          if (v.id === this._context.instanceId) {
            returnProperties = v.webPartData.properties;
            console.log(v.webPartData.properties);
            break;
          }
        }
        // Save CurrentPage  Data
        this.currentPage = results;
      }
    }
    return returnProperties;
  }

  public async setSaveData(data: any) {

    let i: number = 0;
    // Is currentPage loaded ? if not return
    if (!this.currentPage) return;
    // Save Current CanvasContent
    let canvasContent1Updated: any = JSON.parse(this.currentPage.CanvasContent1);
    // Read CanvasContent object
    const canvasContent1 = JSON.parse(this.currentPage.CanvasContent1);
    for (const v of canvasContent1) {
      // Get data for current Instance of webpart in Page
      if (v.id === this._context.instanceId) {
        // Update Property pivotData
        v.webPartData.properties.pivotData = data;
        console.log("update " + JSON.stringify(v.webPartData.properties));
        // Update local copy of CanvasContent1
        canvasContent1Updated[i] = v;
        break;
      }
      i++;
    }
    // Update CanvasContent of current Page width new data
    this.currentPage.CanvasContent1 = JSON.stringify(canvasContent1Updated);
    //console.log(JSON.stringify("after" + this.currentPage.CanvasContent1));
    console.log("novo1:" + JSON.stringify(canvasContent1Updated));
    console.log("novo2:" + JSON.stringify(this.currentPage.CanvasContent1));
    const spOpts: ISPHttpClientOptions = {
      body: `{ "__metadata":{"type":"SP.Publishing.SitePage"},"CanvasContent1":${JSON.stringify(this.currentPage.CanvasContent1)}}`
    };
    // Checkout Page before save
    let apiUrl = `${this._context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${this._context.pageContext.listItem.id})/checkoutpage`;
    const pageCheckOut = await this._context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, {});
    // Save Page with new data

      apiUrl = `${this._context.pageContext.web.absoluteUrl}/_api/sitepages/pages(${this.currentPage.Id})/savepage`;
      const _data = await this._context.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts);
      if (_data.ok) {
        console.log("Page Updated with new properties.");
      } else {
        console.log(_data.statusText + _data.status);
      }

    return;
  }
  /**
   * Get List Items
   */

  public async getListItems(listId: string): Promise<any[]> {
    let spWeb: Web;
    spWeb = new Web(this._context.pageContext.web.absoluteUrl);

    if (!listId) return [];

    try {
      const items: any = await spWeb.lists.getById(listId).items.getAll();
      return items;
    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }

  // Get List Views
  public async getListViews(listId: string): Promise<any[]> {
    let spWeb: Web;
    spWeb = new Web(this._context.pageContext.web.absoluteUrl);
    try {
      const views: any = await spWeb.lists
        .getById(listId).views.get();
      return views;
    } catch (error) {
      console.dir(error);
      return Promise.reject(error);
    }
  }
}
