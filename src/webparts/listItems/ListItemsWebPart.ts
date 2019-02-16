import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ListItemsWebPartStrings';
import ListItems from './components/ListItems';
import { IListItemsProps } from './components/IListItemsProps';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import { SPHttpClientResponse, SPHttpClient } from '@microsoft/sp-http';


export interface IListItemsWebPartProps {
  listId: string;
  itemId: number;
}

export default class ListItemsWebPart extends BaseClientSideWebPart<IListItemsWebPartProps> {
  private itemsDropDown: PropertyPaneAsyncDropdown;

  public render(): void {
    const element: React.ReactElement<IListItemsProps> = React.createElement(
      ListItems,
      {
        listName: this.properties.listId, // temp
        item: this.properties.itemId ? this.properties.itemId.toString(): '' // temp
      }
    );

    ReactDom.render( element, this.domElement );
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode( this.domElement );
  }

  protected get dataVersion(): Version {
    return Version.parse( '1.0' );
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    // reference to item dropdown needed later after selecting a list
    this.itemsDropDown = new PropertyPaneAsyncDropdown( 'itemId', {
      label           : strings.ItemFieldLabel,
      loadOptions     : this.loadItems.bind( this ),
      onPropertyChange: this.onListItemChange.bind( this ),
      selectedKey     : this.properties.itemId,
      // should be disabled if no list has been selected
      disabled        : !this.properties.listId
    } );

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
                // PropertyPaneTextField( 'listName', {
                //   label: strings.ListFieldLabel
                // } )

                new PropertyPaneAsyncDropdown( 'listId', {
                  label           : strings.ListFieldLabel,
                  loadOptions     : this.loadLists.bind( this ),
                  onPropertyChange: this.onListChange.bind( this ),
                  selectedKey     : this.properties.listId
                } ),

                this.itemsDropDown

              ]
            }
          ]
        }
      ]
    };
  }

  private loadLists(): Promise<IDropdownOption[]> {

    if ( Environment.type === EnvironmentType.Local ) {
      return new Promise<IDropdownOption[]>( ( resolve: ( options: IDropdownOption[] ) => void, reject: ( error: any ) => void ) => {
        setTimeout( () => {
          resolve( [{
            key: 'sharedDocuments',
            text: 'Shared Documents'
          },
          {
            key: 'myDocuments',
            text: 'My Documents'
          }] );
        }, 2000 );
      } );
    }

    if ( Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint ) {
      return this.context.spHttpClient
        .get( `${ this.context.pageContext.web.absoluteUrl }/_api/web/lists?$filter=Hidden eq false&$select=Id,Title`, SPHttpClient.configurations.v1 )
        .then( ( response: SPHttpClientResponse ): Promise<any> => {
          return response.json();
        } )
        .then( ( response: any ) => {
          let lists: any[] = response.value;
          lists = lists.map( value => { return { key: value.Id, text: value.Title } } )
            .sort( ( a, b ) => {
              if ( a.text > b.text ) return 1;
              if ( a.text < b.text ) return -1;
              return 0;
             });

          return lists;
        });
    }
  }

  private onListChange( propertyPath: string, newValue: any ): void {
    const oldValue: any = get( this.properties, propertyPath );

    //debugger;
    // store new value in web part properties
    update( this.properties, propertyPath, (): any => { return newValue; } );

    // reset selected item
    this.properties.itemId = undefined;

    // store new value in web part properties
    update( this.properties, 'item', (): any => { return this.properties.itemId; } );

    // refresh web part
    this.render();

    // reset selected values in item dropdown
    this.itemsDropDown.properties.selectedKey = this.properties.itemId;

    // allow to load items
    this.itemsDropDown.properties.disabled = false;

    // load items and re-render items dropdown
    this.itemsDropDown.render();
  }

  private loadItems(): Promise<IDropdownOption[]> {
    if ( !this.properties.listId ) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }

    const wp: ListItemsWebPart = this;

    if ( Environment.type === EnvironmentType.Local ) {

      return new Promise<IDropdownOption[]>( ( resolve: ( options: IDropdownOption[] ) => void, reject: ( error: any ) => void ) => {
        setTimeout( () => {
          const items = {
            sharedDocuments: [
              {
                key: 'spfx_presentation.pptx',
                text: 'SPFx for the masses'
              },
              {
                key: 'hello-world.spapp',
                text: 'hello-world.spapp'
              }
            ],
            myDocuments: [
              {
                key: 'isaiah_cv.docx',
                text: 'Isaiah CV'
              },
              {
                key: 'isaiah_expenses.xlsx',
                text: 'Isaiah Expenses'
              }
            ]
          };
          resolve( items[wp.properties.listId] );
        }, 2000 );
      } );
    }

    if ( Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint ) {
      let query = wp.properties.listId !== '' ? `lists(guid'${ wp.properties.listId }')/items?$select=Id,Title` : null;

      if ( query ) {
        return this.context.spHttpClient
        .get( `${ this.context.pageContext.web.absoluteUrl }/_api/web/${ query }`, SPHttpClient.configurations.v1 )
          .then( ( response: SPHttpClientResponse ): Promise<any> => {
            return response.json();
          } )
          .then( ( response: any ) => {
            let items: any[] = response.value;
            debugger;
            items = items.map( value => { return { key: value.Id, text: value.Title } } )
              .sort( ( a, b ) => {
                if ( a.text > b.text ) return 1;
                if ( a.text < b.text ) return -1;
                return 0;
              } );

            return items;
          } )
          .catch( error => {
            console.log( 'error', error );
            return Promise.resolve();
          } );
      }
    }

  }

  private onListItemChange( propertyPath: string, newValue: any ): void {
    const oldValue: any = get( this.properties, propertyPath );
    // store new value in web part properties
    update( this.properties, propertyPath, (): any => { return newValue; } );
    // refresh web part
    this.render();
  }
}
