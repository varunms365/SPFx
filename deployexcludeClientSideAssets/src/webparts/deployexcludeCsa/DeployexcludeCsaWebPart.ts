import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './DeployexcludeCsaWebPart.module.scss';
import * as strings from 'DeployexcludeCsaWebPartStrings';

export interface IDeployexcludeCsaWebPartProps {
  description: string;
}

const logo: any = require('./assets/vatluri1.png');

export default class DeployexcludeCsaWebPart extends BaseClientSideWebPart<IDeployexcludeCsaWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.deployexcludeCsa }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Deploy includeClientSideAssets as flase.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
          </div>
        </div>
        <img src="${logo}" alt="Varun Atluri logo" width="150" style="margin: 10px 0 0 30px" />
      </div>`;
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
