import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EmployeeOtmWebPartStrings';
import EmployeeOtm from './components/EmployeeOtm';
import { IEmployeeOtmProps } from './components/IEmployeeOtmProps';

import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";
import "@pnp/sp/profiles";
import "@pnp/sp/site-users";

import { spfi } from '@pnp/sp/fi';
import { SPFx } from '@pnp/sp/behaviors/spfx';

export interface IEmployeeOtmWebPartProps {
  description: string;
}

export interface IEmployee {
  Name: string;
  LoginName: string;
  Reason: string;
  PicUrl: string;
}

export interface IRowEotm {
  EmployeeId: number;
  Month: string;
  Reason: string;
}

export default class EmployeeOtmWebPart extends BaseClientSideWebPart<IEmployeeOtmWebPartProps> {
  Eotm: IEmployee;

  constructor() {
    super();

    this.assignEotm = this.assignEotm.bind(this);
  }

  public async getEotm(): Promise<IEmployee> {
  const sp = spfi().using(SPFx(this.context));

    let endUser: any;
    const eotms: Array<IRowEotm> = await sp.web.lists.getByTitle('Employees').items();
    const currentMonth = ((new Date().getMonth())+1).toString();
    for (const eotm of eotms) {
      console.log(eotm);
      if (eotm.Month === currentMonth) {
        endUser = eotm;
        break;
      }
    }

    for (const user of await sp.web.siteUsers()) {
      if(user.Id !== endUser.EmployeeId) continue;
      else {
        const profile: IEmployee = {Name: null, LoginName: null, Reason: null, PicUrl: null};
        const properties = (await sp.profiles.getPropertiesFor(user.LoginName)).UserProfileProperties;

        profile.Name = user.Title;
        profile.Reason = endUser.Reason

      for (const property of properties) {
        switch (property.Key) {
          case 'AccountName':
            profile.LoginName = property.Value;
            break;

          case 'PictureURL':
            profile.PicUrl = property.Value;
            break;
        
          default:
            break;
        }
      }
      return profile;
    }
  }
  }

  public async getEmployees(): Promise<Array<IEmployee>> {

  const sp = spfi().using(SPFx(this.context));
    const rawEmployees = await sp.profiles();
    const employees = new Array<IEmployee>();
    console.log(rawEmployees); employees;
    for (const rawEmployee of rawEmployees) {
      rawEmployee;
    }
    return employees;
  }

  public async assignEotm(prevEotm: IEmployee, newEotm: IEmployee, reason: string): Promise<IEmployee | null> {
    return;
  }

  employees: Array<IEmployee>;

  public render(): void {
    const element: React.ReactElement<IEmployeeOtmProps> = React.createElement(
      EmployeeOtm, {Eotm: this.Eotm, assignEotm: this.assignEotm, rootLink: this.context.pageContext.web.absoluteUrl}
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void>{
    this.Eotm = await this.getEotm();
    this.getEmployees();
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
