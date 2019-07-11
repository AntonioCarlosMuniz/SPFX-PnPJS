import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpPnPjscrudWebPart.module.scss';
import * as strings from 'SpPnPjscrudWebPartStrings';
import { IListItem } from './IListItem';
import pnp from "@pnp/pnpjs";  

export interface ISpPnPjscrudWebPartProps {
  listName: string;
}

export default class SpPnPjscrudWebPart extends BaseClientSideWebPart<ISpPnPjscrudWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.spPnPjscrud }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">CRUD operations</span>
              <p class="${ styles.subTitle }">SP PnP JS</p>
              <p class="${ styles.description }">Name: ${escape(this.properties.listName)}</p>
              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
                  <button class="${styles.button} create-Button">  
                    <span class="${styles.label}">Criar item</span>  
                  </button>  
                  <button class="${styles.button} read-Button">  
                    <span class="${styles.label}">Ler item</span>  
                  </button>  
                </div>  
              </div>  
  
              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
                  <button class="${styles.button} update-Button">  
                    <span class="${styles.label}">Atualizar item</span>  
                  </button>  
                  <button class="${styles.button} delete-Button">  
                    <span class="${styles.label}">Deletar item</span>  
                  </button>  
                </div>  
              </div>  
  
              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">  
                <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">  
                  <div class="status"></div>  
                  <ul class="items"><ul>  
                </div>  
              </div>                
  
            </div>  
          </div>  
        </div>  
      </div>`;  
      this.setButtonsEventHandlers();  
  }  
  
  private setButtonsEventHandlers(): void {  
    const webPart: SpPnPjscrudWebPart = this;  
    this.domElement.querySelector('button.create-Button').addEventListener('click', () => { webPart.createItem(); });  
    this.domElement.querySelector('button.read-Button').addEventListener('click', () => { webPart.lendoItem(); });  
    this.domElement.querySelector('button.update-Button').addEventListener('click', () => { webPart.updateItem(); });  
    this.domElement.querySelector('button.delete-Button').addEventListener('click', () => { webPart.deleteItem(); });  
  }  

  private createItem(): void{

    pnp.sp.web.lists.getByTitle('Pessoas').items.add({
      Title: this.properties.listName
  }).then(console.log);
}

  private updateItemsHtml(items: IListItem[]): void{
  }

  private lendoItem(): void {

    this.getLatestItemId()
    .then((itemId: number): Promise<IListItem> => {
      if(itemId === -1) {
        throw new Error ('Não existe item para ler');
      }

      return pnp.sp.web.lists.getByTitle(this.properties.listName)
      .items.getById(itemId).select('Title', 'Id').get();
    })
    .then((item: IListItem): void => {
    }, (error: any): void => {
    });
  }
  
  private updateItem(): void {  
    let latestItemId: number = undefined;  
    let etag: string = undefined;  
  
    this.getLatestItemId()  
      .then((itemId: number) => {  
        if (itemId === -1) {  
          throw new Error('No items found in the list');  
        }  
  
        latestItemId = itemId;  
        return pnp.sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(itemId).get(undefined, {  
            headers: {  
              'Accept': 'application/json;odata=minimalmetadata'  
            }  
          });  
      })  
      .then((item: any): Promise<IListItem> => {  
        etag = item["odata.etag"];  
        return Promise.resolve((item as any) as IListItem);  
      })  
      .then((item: IListItem) => {  
        return pnp.sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(item.Id).update({  
            'Title': `Updated Item ${new Date()}`  
          }, etag);  
      })  
      .then((result: any): void => {  
      }, (error: any): void => {  
      });  
  }   
  
  private deleteItem(): void {  
    if (!window.confirm('Are you sure you want to delete the latest item?')) {  
      return;  
    }  
    
    let latestItemId: number = undefined;  
    let etag: string = undefined;  
    this.getLatestItemId()  
      .then((itemId: any) => {  
        if (itemId === -1) {  
          throw new Error('No items found in the list');  
        }  
    
        latestItemId = itemId;  
        return pnp.sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(latestItemId).select('Id').get(undefined, {  
            headers: {  
              'Accept': 'application/json;odata=minimalmetadata'  
            }  
          });  
      })  
      .then((item: any): Promise<IListItem> => {  
        etag = item["odata.etag"];  
        return Promise.resolve((item as any) as IListItem);  
      })  
      .then((item: IListItem): Promise<void> => {  
        return pnp.sp.web.lists.getByTitle(this.properties.listName)  
          .items.getById(item.Id).delete(etag);  
      })  
      .then((): void => {  
      }, (error: any): void => {  
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
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getLatestItemId(): Promise<number> {  
    return new Promise<number>((resolve: (itemId: number) => void, reject: (error: any) => void): void => {  
      pnp.sp.web.lists.getByTitle(this.properties.listName)  
        .items.orderBy('Id', false).top(1).select('Id').get()  
        .then((items: { Id: number }[]): void => {  
          if (items.length === 0) {  
            resolve(-1);  
          }  
          else {  
            resolve(items[0].Id);  
          }  
        }, (error: any): void => {  
          reject(error);  
        });  
    });  
  }  
}