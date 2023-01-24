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
  FirstName: string;
  LastName: string;
  LoginName: string;
  Eotm: string;
  PicUrl: string;
}

export default class EmployeeOtmWebPart extends BaseClientSideWebPart<IEmployeeOtmWebPartProps> {
  constructor() {
    super();

    this.assign_eotm = this.assign_eotm.bind(this);
  }

  public async getEmployees(): Promise<Array<IEmployee>> {
    const sp = spfi().using(SPFx(this.context));
    const profiles = new Array<IEmployee>();

    for (const user of await sp.web.siteUsers()) {
      if(user.PrincipalType !== 1) continue;

      let properties = await sp.profiles.getPropertiesFor(user.LoginName);
      if (properties['odata.null']) continue;
      else properties = properties.UserProfileProperties;

      const profile: IEmployee = {FirstName: null, LastName: null, LoginName: null, Eotm: null, PicUrl: null};

      for (const property of properties) {
        switch (property.Key) {
          case 'FirstName':
            profile.FirstName = property.Value;
            break;

          case 'LastName':
            profile.LastName = property.Value;
            break;

          case 'Eotm':
            profile.Eotm = property.Value;
            break;

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
      if (profile.FirstName) profiles.push(profile);
    }
    profiles.push({
      FirstName: 'None',
      LastName: '',
      LoginName: null,
      Eotm: null,
      PicUrl: null
    });
    return profiles;
  }

  public async assign_eotm(prevEotm: IEmployee, newEotm: IEmployee, reason: string): Promise<IEmployee | null> {

    const sp = spfi().using(SPFx(this.context));
    if (newEotm.FirstName === 'None') {
      return sp.profiles.setSingleValueProfileProperty(prevEotm.LoginName, 'Eotm', null)
      .then(
        ()=>null,
        (reason)=>{ console.log(reason); return prevEotm; }
      );
    }
    if (prevEotm) {
      return sp.profiles.setSingleValueProfileProperty(prevEotm.LoginName, 'Eotm', null)
      .then(
        () => {
          return sp.profiles.setSingleValueProfileProperty(newEotm.LoginName, 'Eotm', reason)
          .then(() => {
            return {
              FirstName: newEotm.FirstName,
              LastName: newEotm.LastName,
              LoginName: newEotm.LoginName,
              Eotm: reason,
              PicUrl: newEotm.PicUrl
              };
            },
            (reason) => {
              console.log(reason);
              sp.profiles.setSingleValueProfileProperty(prevEotm.LoginName, 'Eotm', prevEotm.Eotm);
              return prevEotm;
            }
          );
        },
        (reason) => {
          console.log(reason);
          return prevEotm;
        }
      );
    }
    else {
      return sp.profiles.setSingleValueProfileProperty(newEotm.LoginName, 'Eotm', reason)
      .then(() => {
        return {
          FirstName: newEotm.FirstName,
          LastName: newEotm.LastName,
          LoginName: newEotm.LoginName,
          Eotm: reason,
          PicUrl: newEotm.PicUrl
        };
      }
      );
    }

  }

  employees: Array<IEmployee>;

  public render(): void {
    const element: React.ReactElement<IEmployeeOtmProps> = React.createElement(
      EmployeeOtm, {employees: this.employees, assign_eotm: this.assign_eotm, rootLink: this.context.pageContext.web.absoluteUrl}
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void>{
    this.employees = await this.getEmployees();
    return Promise.resolve();
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
