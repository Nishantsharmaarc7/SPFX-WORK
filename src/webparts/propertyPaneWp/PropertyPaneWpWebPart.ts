import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import styles from './PropertyPaneWpWebPart.module.scss';
import * as strings from 'PropertyPaneWpWebPartStrings';

export interface IPropertyPaneWpWebPartProps {
  description: string;
}
export default class PropertyPaneWpWebPart extends BaseClientSideWebPart<IPropertyPaneWpWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.propertyPaneWp}">
    </section>`;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
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
