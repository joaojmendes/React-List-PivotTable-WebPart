
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { sp, Web } from '@pnp/sp';

export default class SPService {

  constructor(private _context: WebPartContext | ApplicationCustomizerContext) { }

  /**
   * Get List Items
   */

  public async getListItems(listId: string): Promise<any[]> {
    let spWeb: Web;
    spWeb = new Web(this._context.pageContext.web.absoluteUrl);

    if (!listId) return [];

    try {
      const items: any = await spWeb.lists
        .getById(listId)
        .items.getAll();
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
