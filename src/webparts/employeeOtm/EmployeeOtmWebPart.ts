import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneButton,
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
import "@pnp/sp/fields/list";
import "@pnp/sp/fields";
import "@pnp/sp/views/list";


import { spfi } from '@pnp/sp/fi';
import { SPFx } from '@pnp/sp/behaviors/spfx';
import { IListEnsureResult } from '@pnp/sp/lists';

export interface IEmployeeOtmWebPartProps {
  listName: string;
  listCreated: boolean;
}

export interface IEmployee {
  Name: string;
  LoginName: string;
  Reason: string;
  PicUrl: string;
  Month: string;
}

export interface IRowEotm {
  EmployeeId: number;
  Month: string;
  Reason: string;
}

export default class EmployeeOtmWebPart extends BaseClientSideWebPart<IEmployeeOtmWebPartProps> {
  Eotm: IEmployee = null;
  //listNameExistance: boolean = true;
  //prevListName: string;
  
  private async getEotm(): Promise<IEmployee | null> {
    if ( !this.properties.listName ) return null;

    const sp = spfi().using(SPFx(this.context));


    let endUser: any = null;
    const eotms: Array<IRowEotm> = await sp.web.lists.getByTitle(this.properties.listName).items();
    const currentMonth = ((new Date().getMonth())+1).toString();
    for (const eotm of eotms) {
      if (eotm.Month === currentMonth) {
        endUser = eotm;
        break;
      }
    }
    if ( !endUser ) {
      return null;
    }

    for (const user of await sp.web.siteUsers()) {
      if(user.Id !== endUser.EmployeeId) continue;
      else {
        const profile: IEmployee = {Name: user.Title, LoginName: null, Reason: endUser.Reason, PicUrl: null, Month: endUser.Month};
        const properties = (await sp.profiles.getPropertiesFor(user.LoginName)).UserProfileProperties;

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

  public render(): void {
    const element: React.ReactElement<IEmployeeOtmProps> = React.createElement(
      EmployeeOtm, {Eotm: this.Eotm, rootLink: this.context.pageContext.web.absoluteUrl, TeamsContext: this.context.sdks.microsoftTeams}
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void>{
    this.createList = this.createList.bind(this);
    this.Eotm = await this.getEotm();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

/*
  protected async checkListName (listName: string): Promise<string> {
    if (!listName) {
      this.prevListName = null;
      return '';
    }
    const sp = spfi().using(SPFx(this.context));

    return sp.web.lists.getByTitle(listName)().then(
        async()  =>  { 
          console.log("Your list has been found!");
          this.prevListName = this.properties.listName;
          return '';
        }, 
            ()  =>  { 
          console.log("There is no such a list!");
          this.properties.listName = this.prevListName;
          return "The list's name is invalid.";
        }
      );
  }
  */

  private async createList () {
    if (this.properties.listName) {
      this.properties.listCreated = true;

      const sp = spfi().using(SPFx(this.context));
      const listName = this.properties.listName;
      const listRes: IListEnsureResult = await sp.web.lists.ensure(listName);
      
      if (listRes.created) {
        await sp.web.lists.getByTitle(listName).fields.getByTitle('Title').update({Required: false});
  
        await sp.web.lists.getByTitle(listName).fields.addUser  ('Employee', {Required: true});
        await sp.web.lists.getByTitle(listName).fields.addChoice('Month', {Required: true, Choices: ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']});
        await sp.web.lists.getByTitle(listName).fields.addText  ('Reason', {Required: true});
  
        await sp.web.lists.getByTitle(listName).defaultView.fields.removeAll();

        await sp.web.lists.getByTitle(listName).defaultView.fields.add('Employee');
        await sp.web.lists.getByTitle(listName).defaultView.fields.add('Month');
        await sp.web.lists.getByTitle(listName).defaultView.fields.add('Reason');
      }
    }
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
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: strings.PropertyPaneListName,
                  disabled: this.properties.listCreated
                }),
                PropertyPaneButton('', {
                  text: strings.PropertyPaneButton,
                  disabled: this.properties.listCreated,
                  onClick: this.createList
                  }
                )
              ]
            } 
          ]
        }
      ]
    }
  }
}

// Property pane with a employee awarding possibility
/*
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      {
        header: {
          description: strings.PropertyPaneDescription
        },
        groups: [
          {
            groupFields: [
              PropertyPaneTextField('listName', {
                label: strings.PropertyPaneListName,
                onGetErrorMessage: this.checkListName
              }),
              PropertyPaneButton('collapsion', {
                text: strings.PropertyPanePanelButton,
                disabled: !this.properties.collapsion || !this.properties.listName,
                onClick: ()=>this.properties.collapsion = false
              })
            ]
          },
          {
            isCollapsed: this.properties.collapsion,
            groupFields: [
              PropertyPaneDropdown('employee', {
                label: strings.PropertyPaneEmployee,
                options: this.employeesOptions,
                disabled: this.blockade
              }),
              PropertyPaneDropdown('month', {
                label: strings.PropertyPaneMonth,
                options: [
                  {key: '1', text: 'Januar'},
                  {key: '2', text: 'Februar'},
                  {key: '3', text: 'March'},
                  {key: '4', text: 'April'},
                  {key: '5', text: 'May'},
                  {key: '6', text: 'Juni'},
                  {key: '7', text: 'July'},
                  {key: '8', text: 'August'},
                  {key: '9', text: 'September'},
                  {key: '10', text: 'October'},
                  {key: '11', text: 'November'},
                  {key: '12', text: 'December'}
                ],
                disabled: this.blockade
              }),
              PropertyPaneTextField('reason', {
                label: strings.PropertyPaneReason,
                disabled: this.blockade
              }),
              PropertyPaneButton('', {
                text: strings.PropertyPaneChoiceButton,
                onClick: async()=>{
                  if (this.properties.employee && this.properties.month && this.properties.reason) {
                    if(this.Eotm?.Month === this.properties.month) return;

                    const sp = spfi().using(SPFx(this.context));

                    sp.web.lists.getByTitle(this.properties.listName).items.add({
                      EmployeeId: this.properties.employee,
                      EmployeeStringId: this.properties.employee.toString(),
                      Month: this.properties.month,
                      Reason: this.properties.reason
                    });
                    
                    this.properties.employee = null;
                    this.properties.month = null;
                    this.properties.reason = null;
                    this.properties.collapsion = true;
                  }
                }
              })
            ]
          }
        ]
      }
    ]
  };
}
*/