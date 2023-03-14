import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'CreateListWebPartStrings';
import CreateList from './components/CreateList';
import { ICreateListProps } from './components/ICreateListProps';
import { spfi, SPFI, SPFx } from "@pnp/sp";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import { IList } from '@pnp/sp/lists';

import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ICreateListWebPartProps {
  description: string;
  listExisits: boolean
  listName: string
  listObj: IList
}



export default class CreateListWebPart extends BaseClientSideWebPart<ICreateListWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {



    let element: React.ReactElement<ICreateListProps> = React.createElement(
      CreateList,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        uniqueWebpartId: this.properties.listName,
        doesListExist: this.properties.listExisits,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    const sp = spfi().using(SPFx(this.context))

    
    this.properties.listName = this.context.webPartTag.split('.')[3]

    this.properties.listExisits = false
    let listexists = false
    // Check if lists existis with name

    await sp.web.lists.ensure(this.properties.listName)
    .then(listEnsureResult => {
      listexists = true
      this.properties.listObj = listEnsureResult.list
      // check if the list was created, or if it already existed:
      if (listEnsureResult.created) {
        console.log("New List created for webpart");
      } else {
        console.log("Webpart list already existed");
      }
    })
    .catch(error => {
      console.log(`List creation failed with error: ${error}`);
    });

    this.properties.listExisits = listexists

    return await super.onInit();
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {

    // if (this.properties.listExisits)
    // {
    //   (async () => {
    //     // Perform async operations
    //     await this.properties.listObj.delete()
    //     .then(() => {
    //       console.log(`List '${this.properties.listName}' deleted successfully.`);
    //     })
    //     .catch(error => {
    //       console.log(`List deletion failed with error: ${error}`);
    //     })

    //     ReactDom.unmountComponentAtNode(this.domElement);

    //   })();
    // }
    // else
    // {
    //   ReactDom.unmountComponentAtNode(this.domElement);
    // }
    this.deleteListNonPnP()

    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async createList(): Promise<boolean> {
    console.log("Creating List")
    let listCreated = true
    await spfi().using(SPFx(this.context)).web.lists.add(this.properties.listName, 'Created from SPFx', 100)
    .then(result => {
      console.log(`List '${this.properties.listName}' created successfully.`);
    })
    .catch(error => {
      console.log(`List creation failed with error: ${error}`);
      listCreated = false
    });

    return Promise.resolve(listCreated)
  }

  private async deleteList(): Promise<boolean> {
    console.log("Deleting List")
    let listDeleted = false
    await spfi().using(SPFx(this.context)).web.lists.getByTitle(this.properties.listName).delete()
    .then(() => {
      console.log(`List '${this.properties.listName}' deleted successfully.`);
      listDeleted = true
    })
    .catch(error => {
      console.log(`List deletion failed with error: ${error}`);
    })
    return Promise.resolve(listDeleted);
  }

  private deleteListNonPnP(): void {
    const url: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listName}')`;

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
        headers: {
            "X-HTTP-Method": "DELETE",
            "IF-MATCH": "*"
        }
    }).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          console.log(`List ${this.properties.listName} deleted successfully`);
        }
        else {
          console.log(`Error: ${response.statusText}`);
        }
    });
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
                }),
                
                PropertyPaneButton("CreateBtn", {
                  text: strings.CreateBtn,
                  
                  onClick: async ()  => {
                    await this.createList()
                    .then(created => {
                      console.log(`List created: ${created.toString()}`)
                      this.properties.listExisits = created
                      this.render()
                    });
                  },
                }),
                PropertyPaneButton("DeleteBtn", {
                  text: strings.DeleteBtn,
                  onClick: async () => {
                    await this.deleteList()
                    .then(deleted => {
                      console.log(`List deleted: ${deleted.toString()}`)
                      this.properties.listExisits = !deleted
                      this.render()
                    });
                  },
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
