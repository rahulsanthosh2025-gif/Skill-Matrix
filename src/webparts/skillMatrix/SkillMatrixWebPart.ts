import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SkillMatrixWebPartStrings';
import { SkillMatrix } from './components/SkillMatrix';
import { ISkillMatrixComponentProps } from './components/SkillMatrix';

export interface ISkillMatrixWebPartProps {
  description: string;
  listName: string;
}

export default class SkillMatrixWebPart extends BaseClientSideWebPart<ISkillMatrixWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISkillMatrixComponentProps> = React.createElement(
      SkillMatrix,
      {
        context: this.context,
        skillLogs: [],
        isLoading: true
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: 'SharePoint List Name',
                  value: 'SkillLogs'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}