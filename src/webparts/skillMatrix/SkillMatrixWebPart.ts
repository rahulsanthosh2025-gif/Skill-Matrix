import * as React from 'react';
import * as ReactDom from 'react-dom';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp/presets/all';

import SkillMatrix from './components/SkillMatrix';
import { ISkillMatrixProps } from './components/ISkillMatrixProps';

export interface ISkillMatrixWebPartProps {}

export default class SkillMatrixWebPart extends BaseClientSideWebPart<ISkillMatrixWebPartProps> {
  protected async onInit(): Promise<void> {
    await super.onInit();

    // Configure PnPjs to use SPFx context
    sp.setup({
      spfxContext: this.context,
    });
  }

  public render(): void {
    const element: React.ReactElement<ISkillMatrixProps> = React.createElement(
      SkillMatrix,
      {
        siteUrl: this.context.pageContext.web.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
}