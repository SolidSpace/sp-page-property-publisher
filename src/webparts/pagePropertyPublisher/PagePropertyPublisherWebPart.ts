import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import "./sass/style.scss";
import {
  IPropertyPaneConfiguration,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField,
  PropertyPaneHorizontalRule,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PagePropertyPublisherWebPartStrings';
import { Placeholder, IPlaceholderProps } from './components/Placeholder';
import { DisplayMode } from '@microsoft/sp-core-library';

import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  IDynamicDataPropertyDefinition,
  IDynamicDataCallables,
  IDynamicDataSource
} from '@microsoft/sp-dynamic-data';
import { PropertyPaneDescription } from 'PagePropertyPublisherWebPartStrings';

export interface IPagePropertyPublisherWebPartProps {
  skipSystemFields: string;
  multivalueLabel: string;
  multivalueOrOperator: string;
  multivalueAndOperator: string;
  multivalueSepOperator: string;
}

export default class PagePropertyPublisherWebPart extends BaseClientSideWebPart<IPagePropertyPublisherWebPartProps> implements IDynamicDataCallables {
  //private _prevMode: any;
  private _pageProperties: any = {};
  private _propsDefinition: any = [];
  //private _pageContextDataSrc: IDynamicDataSource = null;
  private _systemColumns: Array<string> = ["FileSystemObjectType", "Id", "ServerRedirectedEmbedUri", "ServerRedirectedEmbedUrl", "ContentTypeId", "ComplianceAssetId", "WikiField", "BannerImageUrl", "PromotedState", "FirstPublishedDate", "ID", "Created", "AuthorId", "Modified", "EditorId", "CheckoutUserId", "GUID", "CanvasContent1", "LayoutWebpartsContent"];
  private _unsafeColumns: Array<string> = ["CanvasContent1", "LayoutWebpartsContent"];

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      // other init code may be present
      sp.setup({
        spfxContext: this.context
      });
    }).then(_ => {
      this.context.dynamicDataSourceManager.initializeSource(this);
      //this.context.dynamicDataProvider.registerAvailableSourcesChanged(this._registerDynamicDataCallback.bind(that));
      //- this._pageContextDataSrc = this.context.dynamicDataProvider.tryGetSource("PageContext");
      //      (this._pageContextDataSrc)?this.context.dynamicDataProvider.registerSourceChanged("PageContext",this._pagePropertyChanged.bind(this)):console.log("PageContext n/a");
      /*
            this._pageContextDataSrc.getPropertyDefinitionsAsync().then((result)=>{
              result.forEach((property)=>{
                this._pageContextDataSrc.getPropertyValueAsync(property.id).then((value)=>{
                  console.log("property:");
                  console.log(value);
                });
              });
            });
      */
      this._onLoadPageProps().then((result) => {
        return this._updatePageProperties(result);
      });
    });
  }

  private _onRefreshProperties(forceRefresh?: boolean) {
    if (this.displayMode == DisplayMode.Read || forceRefresh) {
      this._onLoadPageProps().then((result) => {
        this.render();
        return this._updatePageProperties(result);
      }).catch((error) => {
        console.error(error);
      });
    } else {
      this.render();
    }
  }

  public onDisplayModeChanged() {
    this._onRefreshProperties();
  }

  public getPropertyDefinitions(): IDynamicDataPropertyDefinition[] {
    return this._propsDefinition;
  }

  public getPropertyValue(propertyId: string) {
    if (this._pageProperties.hasOwnProperty(propertyId)) {
      return this._pageProperties[propertyId];
    } else {
      throw new Error('Bad property id');
    }

  }

  private _updatePageProperties(result: any): Promise<boolean> {
    try {
      this._pageProperties = this._skipSystemColumns(result);
      const that = this;
      //small timeout needed. otherwise old props will be sent
      setTimeout(() => {
        this._propsDefinition.forEach((element, index: string) => {
          that.context.dynamicDataSourceManager.notifyPropertyChanged(element.id);
        });
        that.context.dynamicDataSourceManager.notifySourceChanged();
        console.log("mode change -> notified");
      }, 500);
      return Promise.resolve(true);
    } catch (e) {
      console.error(e);
      return Promise.reject(e);
    }
  }

  private _skipSystemColumns(pageProperties: any) {

    if (pageProperties.Title != undefined) {
      let skipColumns: Array<string> = (this.properties.skipSystemFields == '0') ? this._unsafeColumns.concat(...this._systemColumns) : this._unsafeColumns;
      skipColumns.forEach((keyName) => {
        delete pageProperties[keyName];
      });


    }
    return pageProperties;

  }

  private async _onLoadPageProps(): Promise<any> {
    const item: any = await sp.web.lists.getById(this.context.pageContext.list.id.toString()).items.getById(this.context.pageContext.listItem.id).get();
    let result = {};
    this._propsDefinition = [];


    for (var key in item) {
      let keyLower = key.toLowerCase();
      if (keyLower.substr(0, 1) != "_" && keyLower.search("odata") < 0) {

        result[key] = item[key];
        this._propsDefinition.push({ id: key, title: key });
        if (Array.isArray(item[key])) {
          if (this.properties.multivalueAndOperator != "") {
            result[key + "_OP_AND"] = key + ":"+ item[key].join(" AND ");
            this._propsDefinition.push({ id: key + "_OP_AND", title: key + " AND Operator" });
          }

          if (this.properties.multivalueOrOperator != "") {
            result[key + "_OP_OR"] = key + ":"+ item[key].join(" OR ");
            this._propsDefinition.push({ id: key + "_OP_OR", title: key + " OR Operator" });
          }

          if (this.properties.multivalueSepOperator != "") {
            result[key + "_SEPERATOR"] = item[key].join(";");
            this._propsDefinition.push({ id: key + "_SEPERATOR", title: key + " Seperator" });
          }

        }
      }
    }
    return result;
  }

  public onPropertyPaneFieldChanged(targetProperty: string, newValue: any) {
    if (targetProperty == "skipSystemFields" || targetProperty == "PropertyPaneSystemColumnsOff" || targetProperty == "PropertyPaneSystemColumnsOn") {
      this._onRefreshProperties();
    }
  }

  public render(): void {
    const element: React.ReactElement<IPlaceholderProps> = React.createElement(
      Placeholder, {
      pageProperties: this._pageProperties,
      displaymode: this.displayMode,
      cbRefreshProperties: this._onRefreshProperties.bind(this)
    }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.AboutText
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneLabel("multivalueLabel", {
                  text: strings.PropertyPaneMultivalueLabel
                }),
                PropertyPaneCheckbox('multivalueOrOperator', {
                  checked: true,
                  text: strings.PropertyPaneMultivalueOR
                }),
                PropertyPaneCheckbox('multivalueAndOperator', {
                  checked: true,
                  text: strings.PropertyPaneMultivalueAND
                }),
                PropertyPaneCheckbox('multivalueSepOperator', {
                  checked: true,
                  text: strings.PropertyPaneMultivalueSeperator
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneChoiceGroup('skipSystemFields', {
                  "label": strings['PropertyPaneSystemColumnsText'],
                  options: [
                    { key: '0', text: strings['PropertyPaneSystemColumnsOff'], checked: true },
                    { key: '1', text: strings['PropertyPaneSystemColumnsOn'] }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

