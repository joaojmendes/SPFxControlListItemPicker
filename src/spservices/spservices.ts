// João Mendes
// Outubro 2019

import {
  HttpClient,
  HttpClientConfiguration,
  HttpClientResponse,
  IHttpClientOptions,
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";


// Class Services
export default class spservices {
  constructor(
    private _context:
       WebPartContext
      | ApplicationCustomizerContext
  ) {

  }

     /**
   * Gets list items for list item picker
   * @param filterText
   * @param listId
   * @param internalColumnName
   * @param [keyInternalColumnName]
   * @param [webUrl]
   * @param [filterList]
   * @returns list items for list item picker
   */
  public async getListItemsForListItemPicker(
    filterText: string,
    listId: string,
    internalColumnName: string,
    keyInternalColumnName?: string,
    webUrl?: string,
    filterList?: string
  ): Promise<any[]> {
    let _filter: string = `$filter=startswith(${internalColumnName},'${encodeURIComponent(
      filterText.replace("'", "''")
    )}') `;
    let costumfilter: string = filterList
      ? `and ${filterList}`
      : "";
    let _top = " &$top=2000";

    // test wild character "*"  if "*" load first 30 items
    if (
      (filterText.trim().indexOf("*") == 0 &&
        filterText.trim().length == 1) ||
      filterText.trim().length == 0
    ) {
      _filter = "";
      costumfilter = filterList ? `$filter=${filterList}&` : "";
      _top = `$top=500`;
    }

    try {
      const webAbsoluteUrl = !webUrl
        ? this._context.pageContext.web.absoluteUrl
        : webUrl;
      const apiUrl = `${webAbsoluteUrl}/_api/web/lists('${listId}')/items?$orderby=${internalColumnName}&$select=${keyInternalColumnName ||
        "Id"},${internalColumnName}&${_filter}${costumfilter}${_top}`;
      const data = await this._context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );
      if (data.ok) {
        const results = await data.json();
        if (
          results &&
          results.value &&
          results.value.length > 0
        ) {
          return results.value;
        }
      }
      return [];
    } catch (error) {
      return Promise.reject(error);
    }
  }
}
