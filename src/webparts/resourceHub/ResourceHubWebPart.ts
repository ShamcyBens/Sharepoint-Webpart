import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import '@fortawesome/fontawesome-free/css/all.css';
import styles from './ResourceHubWebPart.module.scss';
import * as strings from 'ResourceHubWebPartStrings';

export interface IResourceHubWebPartProps {
  description: string;
}

export interface ISPList {
  Name: string;
  Audience: string;
  Category: string;
  Tags: string;
  Rating: number;
  Link: string;
}

export default class ResourceHubWebPart extends BaseClientSideWebPart<IResourceHubWebPartProps> {
  private selectedCategories: string[] = [];

  private async _getListData(): Promise<ISPList[]> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const listTitle = "ResourceHub";
    const apiUrl = `${siteUrl}/_api/web/lists/getByTitle('${listTitle}')/items?`;

    const response: SPHttpClientResponse = await this.context.spHttpClient.get(apiUrl, SPHttpClient.configurations.v1);

    if (response.ok) {
      const data = await response.json();
      return data.value.map((item: any) => ({
        Audience: item.Audience || 'Unknown',
        Category: item.Category || 'Unknown',
        Tags: item.Tags || 'Unknown',
        Rating: item.Rating || 0,
        Link: item.FileRef
      }));
    } else {
      throw new Error(`Failed to fetch lists: ${response.statusText}`);
    }
  }

  public async render(): Promise<void> {
    const lists: ISPList[] = await this._getListData();
    this.domElement.innerHTML = this._renderHTML(lists);
    this._attachEventHandlers(lists);
  }

  private _renderHTML(lists: ISPList[]): string {
    const categories: string[] = Array.from(new Set(lists.map(list => list.Category)));

    const checkboxesHTML = categories.map(category => `
      <label class="${styles.checkboxLabel}">
        <input type="checkbox" value="${category}" class="${styles.categoryCheckbox}"/>
        ${category}
      </label>
    `).join('');

    return `
      <div class="${styles.resourceHub} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
        <div class="${styles.intro}">
          <p>Welcome to our Resource Hub.
          Explore the comprehensive list of our resources organised by categories
          below to stay informed :</p>
        </div>
        <div class="${styles.filters}">
          <h3>Filter</h3>
          ${checkboxesHTML}
        </div>
        <div class="${styles.resources}" id="resourcesContainer">
          ${this._renderListsHTML(lists)}
        </div>
      </div>`;
  }

  private _renderListsHTML(lists: ISPList[]): string {
    let listsHTML: string = '';

    const categories: string[] = Array.from(new Set(lists.map(list => list.Category)));

    listsHTML += `<div class="${styles.categoriesContainer}">`;

    categories.forEach(category => {
      const categoryLists: ISPList[] = lists.filter(list => list.Category === category);
      listsHTML += `
        <div class="${styles.category}" data-category="${category}">
          <h2 class="${styles.categoryTitle}">${category}</h2>
          <ul class="${styles.list}">`;

      categoryLists.forEach(list => {
        listsHTML += `
          <li class="${styles.listItem}">
              ${escape(list.Name)}
            </a>
            <div>Audience: ${escape(list.Audience)}</div>
            <div>Tags: ${escape(list.Tags)}</div>
            <div>Rating: ${list.Rating}</div>
          </li>`;
      });

      listsHTML += `</ul></div>`;
    });

    listsHTML += `</div>`;

    return listsHTML;
  }

  private _attachEventHandlers(lists: ISPList[]): void {
    const checkboxes = this.domElement.querySelectorAll(`.${styles.categoryCheckbox}`);
    checkboxes.forEach(checkbox => {
      checkbox.addEventListener('change', () => {
        this._handleCheckboxChange(lists);
      });
    });
  }

  private _handleCheckboxChange(lists: ISPList[]): void {
    const selectedCheckboxes = this.domElement.querySelectorAll(`.${styles.categoryCheckbox}:checked`);
    this.selectedCategories = Array.from(selectedCheckboxes).map((checkbox: HTMLInputElement) => checkbox.value);
  
    const filteredLists = this.selectedCategories.length > 0
      ? lists.filter(list => this.selectedCategories.includes(list.Category))
      : lists;
  
    const resourcesContainer = this.domElement.querySelector('#resourcesContainer');
    if (resourcesContainer) {
      resourcesContainer.innerHTML = this._renderListsHTML(filteredLists);
    }
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
